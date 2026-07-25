// -----------------------------------------------------------------------------
// File: cpp\src\ui\supervisor.h
// Purpose: Native port of WorkerSupervisor — spawns XactCopyExecutive, drives
//          jobs over the named-pipe IPC contract, monitors heartbeats and job
//          activity, and restarts/resumes the worker after stalls, exits, or
//          channel loss (auto-recovery forces resume-from-journal).
// -----------------------------------------------------------------------------

#pragma once

#include <atomic>
#include <cstdint>
#include <functional>
#include <mutex>
#include <optional>
#include <string>
#include <thread>
#include <vector>

#include "../core/ipc.h"
#include "../core/pipe.h"
#include "../storage/stores.h"

namespace xact::ui {

using models::CopyJobOptions;
using models::CopyJobResult;
using models::CopyProgressSnapshot;

class WorkerSupervisor {
public:
    std::function<void(const std::string&)> on_log;
    std::function<void(const CopyProgressSnapshot&)> on_progress;
    std::function<void(const CopyJobResult&)> on_job_completed;
    std::function<void(const std::string&)> on_worker_state_changed;
    std::function<void(bool)> on_job_pause_state_changed;

    WorkerSupervisor() = default;
    ~WorkerSupervisor() { stop(); }

    WorkerSupervisor(const WorkerSupervisor&) = delete;
    WorkerSupervisor& operator=(const WorkerSupervisor&) = delete;

    bool is_job_running() const noexcept { return job_running_.load(); }
    bool is_job_paused() const noexcept { return job_paused_.load(); }
    std::string current_job_id() {
        std::lock_guard<std::mutex> guard(state_lock_);
        return current_job_id_;
    }

    // Starts a new job (throws when one is already running or spawn fails).
    void start_job(const CopyJobOptions& options) {
        std::lock_guard<std::mutex> guard(restart_lock_);
        if (stopping_.load()) throw std::runtime_error("Supervisor is stopping.");
        if (job_running_.load()) throw std::runtime_error("A copy job is already running.");

        {
            std::lock_guard<std::mutex> state_guard(state_lock_);
            current_options_ = options;
            current_job_id_ = storage::detail::new_guid_n();
        }
        job_running_.store(true);
        cancel_requested_.store(false);
        set_paused_state(false);
        auto_recover_.store(true);
        mark_job_activity(true);

        try {
            ensure_worker_connected();
            send_start_job_command();
        } catch (...) {
            // Roll back so the UI can retry instead of being wedged "running".
            job_running_.store(false);
            throw;
        }
        raise_state("Job running");
    }

    void cancel_job(const std::string& reason) {
        if (!job_running_.load()) return;
        auto_recover_.store(false);
        cancel_requested_.store(true);
        set_paused_state(false);

        ipc::CancelJobCommand command;
        command.JobId = current_job_id();
        command.Reason = reason.empty() ? "User requested cancellation." : reason;
        try_send(ipc::message_types::CancelJobCommand,
                 [&command](json::Writer& w) { command.to_json(w); });
    }

    void pause_job(const std::string& reason) {
        if (!job_running_.load() || job_paused_.load()) return;
        ipc::PauseJobCommand command;
        command.JobId = current_job_id();
        command.Reason = reason.empty() ? "User requested pause." : reason;
        try_send(ipc::message_types::PauseJobCommand,
                 [&command](json::Writer& w) { command.to_json(w); });
        set_paused_state(true);
        raise_state("Job paused");
    }

    void resume_job(const std::string& reason) {
        if (!job_running_.load() || !job_paused_.load()) return;
        ipc::ResumeJobCommand command;
        command.JobId = current_job_id();
        command.Reason = reason.empty() ? "User requested resume." : reason;
        try_send(ipc::message_types::ResumeJobCommand,
                 [&command](json::Writer& w) { command.to_json(w); });
        set_paused_state(false);
        raise_state("Job running");
    }

    // Graceful shutdown: ask the worker to exit, then force-kill leftovers.
    void stop() {
        bool expected = false;
        if (!stopped_.compare_exchange_strong(expected, true)) return;

        stopping_.store(true);
        auto_recover_.store(false);
        job_running_.store(false);
        cancel_requested_.store(false);
        set_paused_state(false);
        {
            std::lock_guard<std::mutex> guard(state_lock_);
            current_job_id_.clear();
            current_options_.reset();
        }

        ipc::ShutdownCommand shutdown;
        shutdown.Reason = "UI shutdown requested.";
        try_send(ipc::message_types::ShutdownCommand,
                 [&shutdown](json::Writer& w) { shutdown.to_json(w); });

        if (!shutdown_worker(true)) {
            raise_log("[Supervisor] One or more worker processes did not terminate during shutdown.");
        }
        stop_heartbeat_timer();
    }

    // Called by the UI once per second (or from the built-in timer thread).
    void check_heartbeat() {
        if (stopping_.load() || stopped_.load()) return;
        if (!pipe_connected_.load()) return;

        ULONGLONG now = GetTickCount64();
        ULONGLONG last_heartbeat = last_heartbeat_tick_.load();
        if (last_heartbeat != 0 && now - last_heartbeat <= HeartbeatTimeoutMs) {
            check_for_job_activity_stall(now);
            return;
        }
        if (last_heartbeat == 0) return;
        double stale_seconds = static_cast<double>(now - last_heartbeat) / 1000.0;
        char text[64];
        std::snprintf(text, sizeof(text), "%.1f", stale_seconds);
        queue_recovery("Worker heartbeat stale for " + std::string(text) + "s.");
    }

private:
    static constexpr ULONGLONG HeartbeatTimeoutMs = 10'000;
    static constexpr ULONGLONG MinimumJobActivityTimeoutMs = 45'000;
    static constexpr ULONGLONG MaximumJobActivityTimeoutMs = 600'000;

    std::mutex restart_lock_;
    std::mutex send_lock_;
    std::mutex state_lock_;

    pipe::MessagePipe pipe_;
    pipe::CancellationEvent pipe_cancel_;
    std::atomic<bool> pipe_connected_{false};
    std::thread receive_thread_;
    std::thread heartbeat_thread_;
    std::atomic<bool> heartbeat_stop_{false};

    HANDLE worker_process_ = nullptr;
    DWORD worker_pid_ = 0;

    std::optional<CopyJobOptions> current_options_;
    std::string current_job_id_;
    std::atomic<bool> job_running_{false};
    std::atomic<bool> job_paused_{false};
    std::atomic<bool> cancel_requested_{false};
    std::atomic<bool> auto_recover_{false};
    std::atomic<bool> stopping_{false};
    std::atomic<bool> stopped_{false};
    std::atomic<int> recovery_in_flight_{0};
    std::atomic<int> consecutive_malformed_{0};

    std::atomic<ULONGLONG> last_heartbeat_tick_{0};
    std::atomic<ULONGLONG> last_job_activity_tick_{0};
    std::atomic<ULONGLONG> last_worker_progress_tick_{0};

    // ---- Event raising -----------------------------------------------------

    void raise_log(const std::string& message) {
        if (on_log) on_log(message);
    }
    void raise_state(const std::string& state) {
        if (on_worker_state_changed) on_worker_state_changed(state);
    }
    void set_paused_state(bool value) {
        bool previous = job_paused_.exchange(value);
        if (previous != value && on_job_pause_state_changed) on_job_pause_state_changed(value);
    }

    // ---- Worker spawn ------------------------------------------------------

    static std::wstring own_directory() {
        wchar_t buffer[MAX_PATH];
        DWORD length = GetModuleFileNameW(nullptr, buffer, MAX_PATH);
        std::wstring path(buffer, length);
        return storage::fsutil::get_directory_name(path);
    }

    // Port of EnumerateWorkerCandidateDirectories: base dir + win-x64 +
    // publish, then up to four parent levels with the same suffixes.
    static std::wstring resolve_worker_executable() {
        std::vector<std::wstring> candidates;
        auto add = [&candidates](const std::wstring& dir) {
            if (dir.empty()) return;
            for (const auto& existing : candidates) {
                if (storage::fsutil::to_upper_invariant(existing) ==
                    storage::fsutil::to_upper_invariant(dir)) {
                    return;
                }
            }
            candidates.push_back(dir);
        };

        std::wstring base = own_directory();
        add(base);
        add(base + L"\\win-x64");
        add(base + L"\\publish");
        std::wstring cursor = base;
        for (int depth = 0; depth < 5; ++depth) {
            std::wstring parent = storage::fsutil::get_directory_name(cursor);
            if (parent.empty() || parent == cursor) break;
            add(parent);
            add(parent + L"\\win-x64");
            add(parent + L"\\publish");
            cursor = parent;
        }

        for (const auto& directory : candidates) {
            for (const wchar_t* name : {L"XactCopyExecutive.exe", L"XactCopy.Worker.exe"}) {
                std::wstring candidate = directory + L"\\" + name;
                if (storage::fsutil::file_exists(candidate)) return candidate;
            }
        }
        return std::wstring();
    }

    void ensure_worker_connected() {
        if (pipe_connected_.load() && pipe_.is_open()) return;

        std::wstring pipe_name = L"xactcopy-" +
                                 storage::fsutil::utf8_to_wide(storage::detail::new_guid_n());
        std::wstring worker_path = resolve_worker_executable();
        if (worker_path.empty()) {
            throw std::runtime_error(
                "Unable to locate worker binary (XactCopyExecutive.exe). Build the native worker "
                "so it sits next to the UI executable.");
        }

        std::wstring command_line = L"\"" + worker_path + L"\" --pipe \"" + pipe_name + L"\"";
        STARTUPINFOW startup{};
        startup.cb = sizeof(startup);
        startup.dwFlags = STARTF_USESHOWWINDOW;
        startup.wShowWindow = SW_HIDE;
        PROCESS_INFORMATION process{};
        std::vector<wchar_t> mutable_command(command_line.begin(), command_line.end());
        mutable_command.push_back(L'\0');
        if (!CreateProcessW(worker_path.c_str(), mutable_command.data(), nullptr, nullptr, FALSE,
                            CREATE_NO_WINDOW, nullptr, nullptr, &startup, &process)) {
            throw std::runtime_error("Failed to start worker process (Win32 " +
                                     std::to_string(GetLastError()) + ").");
        }
        CloseHandle(process.hThread);
        worker_process_ = process.hProcess;
        worker_pid_ = process.dwProcessId;

        pipe_cancel_.reset();
        pipe_ = pipe::MessagePipe::connect_client(pipe_name, 15'000);
        pipe_connected_.store(true);
        last_heartbeat_tick_.store(GetTickCount64());
        consecutive_malformed_.store(0);

        if (receive_thread_.joinable()) receive_thread_.join();
        receive_thread_ = std::thread([this] { receive_loop(); });
        start_heartbeat_timer();

        raise_state("Worker connected (PID " + std::to_string(worker_pid_) + ")");
    }

    void start_heartbeat_timer() {
        if (heartbeat_thread_.joinable()) return;
        heartbeat_stop_.store(false);
        heartbeat_thread_ = std::thread([this] {
            while (!heartbeat_stop_.load()) {
                Sleep(1000);
                if (heartbeat_stop_.load()) return;
                check_heartbeat();
            }
        });
    }

    void stop_heartbeat_timer() {
        heartbeat_stop_.store(true);
        if (heartbeat_thread_.joinable()) heartbeat_thread_.join();
        if (receive_thread_.joinable()) receive_thread_.join();
    }

    // ---- Sending -----------------------------------------------------------

    void try_send(std::string_view message_type,
                  const std::function<void(json::Writer&)>& payload) {
        if (!pipe_connected_.load() || !pipe_.is_open()) return;
        std::string message = ipc::serialize_envelope(message_type, payload);
        std::lock_guard<std::mutex> guard(send_lock_);
        try {
            pipe::CancellationEvent no_cancel;
            pipe_.write_message(message, no_cancel);
        } catch (const std::exception& ex) {
            raise_log("[Supervisor] Send failed: " + std::string(ex.what()));
        }
    }

    void send_start_job_command() {
        ipc::StartJobCommand command;
        {
            std::lock_guard<std::mutex> guard(state_lock_);
            command.JobId = current_job_id_;
            command.Options = current_options_.value_or(CopyJobOptions{});
        }
        try_send(ipc::message_types::StartJobCommand,
                 [&command](json::Writer& w) { command.to_json(w); });
    }

    // ---- Receiving ---------------------------------------------------------

    void receive_loop() {
        try {
            while (true) {
                std::optional<std::string> raw = pipe_.read_message(pipe_cancel_);
                if (!raw.has_value()) break;
                handle_incoming(*raw);
            }
        } catch (const std::exception& ex) {
            if (!stopping_.load()) {
                raise_log("[Supervisor] Receive loop error: " + std::string(ex.what()));
            }
        }
        pipe_connected_.store(false);
        if (!stopping_.load()) {
            queue_recovery("Worker disconnected.");
        }
    }

    void handle_incoming(const std::string& raw) {
        std::string message_type;
        if (!ipc::try_read_message_type(raw, message_type)) {
            raise_log("[Supervisor] Received malformed message.");
            if (consecutive_malformed_.fetch_add(1) + 1 >= 3 && job_running_.load()) {
                queue_recovery("Worker IPC stream desynchronized (malformed messages).");
            }
            return;
        }
        consecutive_malformed_.store(0);

        if (message_type == ipc::message_types::WorkerHeartbeatEvent) {
            auto [header, payload] = ipc::parse_envelope(raw);
            auto heartbeat = ipc::WorkerHeartbeatEvent::from_json(payload);
            last_heartbeat_tick_.store(GetTickCount64());
            if (!(heartbeat.LastProgressUtc == time::DateTimeOffset::min_value())) {
                last_worker_progress_tick_.store(GetTickCount64());
            }
            set_paused_state(heartbeat.IsJobPaused);
            return;
        }
        if (message_type == ipc::message_types::WorkerProgressEvent) {
            auto [header, payload] = ipc::parse_envelope(raw);
            auto progress = ipc::WorkerProgressEvent::from_json(payload);
            last_heartbeat_tick_.store(GetTickCount64());
            mark_job_activity(true);
            if (!job_running_.load()) return;
            std::string expected_job = current_job_id();
            if (!expected_job.empty() &&
                !models::detail::equals_ignore_case(progress.JobId, expected_job)) {
                return;
            }
            if (on_progress) on_progress(progress.Snapshot);
            return;
        }
        if (message_type == ipc::message_types::WorkerLogEvent) {
            auto [header, payload] = ipc::parse_envelope(raw);
            auto log_event = ipc::WorkerLogEvent::from_json(payload);
            mark_job_activity(false);
            raise_log(log_event.Message);
            return;
        }
        if (message_type == ipc::message_types::WorkerJobResultEvent) {
            auto [header, payload] = ipc::parse_envelope(raw);
            auto result_event = ipc::WorkerJobResultEvent::from_json(payload);
            complete_job_state();
            if (on_job_completed) on_job_completed(result_event.Result);
            raise_state("Job finished");
            return;
        }
        if (message_type == ipc::message_types::WorkerFatalEvent) {
            auto [header, payload] = ipc::parse_envelope(raw);
            auto fatal = ipc::WorkerFatalEvent::from_json(payload);
            CopyJobResult failure;
            failure.Succeeded = false;
            failure.Cancelled = false;
            failure.ErrorMessage = fatal.ErrorMessage;
            complete_job_state();
            raise_log("[Worker Fatal] " + fatal.ErrorMessage);
            if (on_job_completed) on_job_completed(failure);
            raise_state("Worker fatal error");
            return;
        }
    }

    void complete_job_state() {
        job_running_.store(false);
        set_paused_state(false);
        auto_recover_.store(false);
        cancel_requested_.store(false);
        {
            std::lock_guard<std::mutex> guard(state_lock_);
            current_job_id_.clear();
            current_options_.reset();
        }
        last_job_activity_tick_.store(0);
        last_worker_progress_tick_.store(0);
    }

    void mark_job_activity(bool mark_progress) {
        ULONGLONG now = GetTickCount64();
        last_job_activity_tick_.store(now);
        if (mark_progress) last_worker_progress_tick_.store(now);
    }

    // ---- Stall detection + recovery ---------------------------------------

    ULONGLONG resolve_job_activity_timeout_ms() {
        double operation_timeout_seconds = 10.0;
        double max_retry_delay_seconds = 2.0;
        {
            std::lock_guard<std::mutex> guard(state_lock_);
            if (current_options_.has_value()) {
                if (current_options_->OperationTimeout.ticks > 0) {
                    operation_timeout_seconds = current_options_->OperationTimeout.total_seconds();
                }
                if (current_options_->MaxRetryDelay.ticks > 0) {
                    max_retry_delay_seconds = current_options_->MaxRetryDelay.total_seconds();
                }
            }
        }
        double computed_seconds = operation_timeout_seconds * 4.0 + max_retry_delay_seconds * 2.0 + 10.0;
        double bounded = std::max(static_cast<double>(MinimumJobActivityTimeoutMs) / 1000.0,
                                  std::min(static_cast<double>(MaximumJobActivityTimeoutMs) / 1000.0,
                                           computed_seconds));
        return static_cast<ULONGLONG>(bounded * 1000.0);
    }

    void check_for_job_activity_stall(ULONGLONG now) {
        if (!job_running_.load() || job_paused_.load() || !auto_recover_.load()) return;

        ULONGLONG last_activity = last_job_activity_tick_.load();
        ULONGLONG last_progress = last_worker_progress_tick_.load();
        if (last_activity == 0) last_activity = last_progress;
        if (last_activity == 0) return;

        ULONGLONG timeout = resolve_job_activity_timeout_ms();
        if (now - last_activity <= timeout) return;

        double activity_seconds = static_cast<double>(now - last_activity) / 1000.0;
        double progress_seconds = last_progress == 0
                                      ? activity_seconds
                                      : static_cast<double>(now - last_progress) / 1000.0;
        char text[128];
        std::snprintf(text, sizeof(text),
                      "Worker activity stalled for %.1fs (last progress %.1fs ago).",
                      activity_seconds, progress_seconds);
        queue_recovery(text);
    }

    void queue_recovery(const std::string& reason) {
        int expected = 0;
        if (!recovery_in_flight_.compare_exchange_strong(expected, 1)) return;
        std::thread([this, reason]() {
            try {
                recover_worker(reason);
            } catch (const std::exception& ex) {
                raise_log("[Supervisor] Recovery failed: " + std::string(ex.what()));
            }
            recovery_in_flight_.store(0);
        }).detach();
    }

    void recover_worker(const std::string& reason) {
        if (!job_running_.load()) return;
        std::lock_guard<std::mutex> guard(restart_lock_);
        if (stopping_.load() || stopped_.load()) return;

        if (!auto_recover_.load()) {
            bool clean = shutdown_worker(true);
            std::string final_reason = clean
                                           ? reason
                                           : reason + " Worker process did not terminate cleanly.";
            complete_active_job_after_channel_loss(final_reason);
            return;
        }

        raise_log("[Supervisor] " + reason + " Restarting worker and resuming from journal.");
        if (!shutdown_worker(true)) {
            auto_recover_.store(false);
            complete_active_job_after_channel_loss(
                reason + " Worker process did not terminate cleanly, so auto-recovery was aborted "
                         "to prevent duplicate stuck workers.");
            return;
        }

        ensure_worker_connected();
        mark_job_activity(true);

        if (auto_recover_.load() && job_running_.load()) {
            {
                std::lock_guard<std::mutex> state_guard(state_lock_);
                if (current_options_.has_value() && !current_options_->ResumeFromJournal) {
                    current_options_->ResumeFromJournal = true;
                    raise_log("[Supervisor] Forced resume-from-journal after worker recovery.");
                }
            }
            send_start_job_command();
            if (job_paused_.load()) {
                ipc::PauseJobCommand pause;
                pause.JobId = current_job_id();
                pause.Reason = "Reapplying paused state after worker recovery.";
                try_send(ipc::message_types::PauseJobCommand,
                         [&pause](json::Writer& w) { pause.to_json(w); });
            }
        }
    }

    void complete_active_job_after_channel_loss(const std::string& reason) {
        if (!job_running_.load()) return;

        bool cancelled = cancel_requested_.load();
        std::string message =
            cancelled ? "Cancellation was requested, but the worker channel closed before "
                        "acknowledgement. " + reason
                      : "Worker channel lost: " + reason;

        CopyJobResult failure;
        failure.Succeeded = false;
        failure.Cancelled = cancelled;
        failure.ErrorMessage = message;

        complete_job_state();
        consecutive_malformed_.store(0);

        raise_log("[Supervisor] " + message);
        if (on_job_completed) on_job_completed(failure);
        raise_state(cancelled ? "Job cancelled" : "Worker disconnected");
    }

    // Returns true when the worker terminated (or none was running).
    bool shutdown_worker(bool kill_if_still_running) {
        pipe_cancel_.cancel();
        pipe_.close();
        pipe_connected_.store(false);
        if (receive_thread_.joinable()) receive_thread_.join();

        if (worker_process_ == nullptr) return true;

        bool clean = true;
        DWORD pid = worker_pid_;
        if (kill_if_still_running) {
            if (WaitForSingleObject(worker_process_, 2000) != WAIT_OBJECT_0) {
                TerminateProcess(worker_process_, 1);
                if (WaitForSingleObject(worker_process_, 3000) != WAIT_OBJECT_0 && pid != 0) {
                    // taskkill /T catches any children TerminateProcess missed.
                    std::wstring kill_command =
                        L"taskkill.exe /PID " + std::to_wstring(pid) + L" /T /F";
                    STARTUPINFOW startup{};
                    startup.cb = sizeof(startup);
                    startup.dwFlags = STARTF_USESHOWWINDOW;
                    startup.wShowWindow = SW_HIDE;
                    PROCESS_INFORMATION process{};
                    std::vector<wchar_t> mutable_command(kill_command.begin(), kill_command.end());
                    mutable_command.push_back(L'\0');
                    if (CreateProcessW(nullptr, mutable_command.data(), nullptr, nullptr, FALSE,
                                       CREATE_NO_WINDOW, nullptr, nullptr, &startup, &process)) {
                        WaitForSingleObject(process.hProcess, 3000);
                        CloseHandle(process.hThread);
                        CloseHandle(process.hProcess);
                    }
                }
                if (WaitForSingleObject(worker_process_, 0) != WAIT_OBJECT_0) {
                    clean = false;
                    raise_log("[Supervisor] Worker PID " + std::to_string(pid) +
                              " did not terminate promptly; continuing with recovery.");
                }
            }
        }
        CloseHandle(worker_process_);
        worker_process_ = nullptr;
        worker_pid_ = 0;
        return clean;
    }
};

} // namespace xact::ui
