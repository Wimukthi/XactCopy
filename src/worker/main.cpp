// -----------------------------------------------------------------------------
// File: src\worker\main.cpp
// Purpose: XactCopyExecutive (native) — out-of-process copy worker. Speaks the
//          .NET worker's IPC contract (connection log, 1 Hz heartbeats, ping/
//          pause/resume/cancel/shutdown) and executes copy jobs through the
//          native ResilientCopyEngine with throttled progress/log telemetry.
// -----------------------------------------------------------------------------

#include <atomic>
#include <cstdio>
#include <mutex>
#include <optional>
#include <string>
#include <string_view>
#include <thread>
#include <vector>

#include "../core/ipc.h"
#include "../core/pipe.h"
#include "engine.h"

namespace {

using namespace xact;

class WorkerHost {
public:
    int run(const std::wstring& pipe_name, DWORD expected_parent_pid) {
        try {
            pipe_ = pipe::MessagePipe::create_server(pipe_name);
        } catch (const std::exception& ex) {
            std::fprintf(stderr, "XactCopyExecutive fatal error: %s\n", ex.what());
            return 1;
        }

        try {
            pipe_.wait_for_connection(lifetime_);
            ULONG client_pid = 0;
            if (expected_parent_pid == 0 ||
                !GetNamedPipeClientProcessId(pipe_.native_handle(), &client_pid) ||
                client_pid != expected_parent_pid) {
                throw std::runtime_error(
                    "Named-pipe client is not the supervisor process that launched this worker.");
            }
            send_log("", "Worker connected to supervisor.");

            std::thread heartbeat([this] { heartbeat_loop(); });

            try {
                receive_loop();
            } catch (const std::exception&) {
                // Receive-loop failures end the session; shutdown continues below.
            }

            lifetime_.cancel();
            job_cancel_.store(true);
            control_.resume();
            if (job_thread_.joinable()) job_thread_.join();
            heartbeat.join();
            pipe_.close();
            return 0;
        } catch (const std::exception& ex) {
            std::fprintf(stderr, "XactCopyExecutive fatal error: %s\n", ex.what());
            lifetime_.cancel();
            if (job_thread_.joinable()) job_thread_.join();
            return 1;
        }
    }

private:
    pipe::MessagePipe pipe_;
    pipe::CancellationEvent lifetime_;
    std::mutex send_lock_;

    // Job state.
    engine::ExecutionControl control_;
    std::thread job_thread_;
    std::atomic<bool> job_running_{false};
    std::atomic<bool> job_paused_{false};
    std::atomic<bool> job_cancel_{false};
    std::mutex job_state_lock_;
    std::string active_job_id_;
    time::DateTimeOffset last_progress_utc_ = time::DateTimeOffset::now_utc();
    std::optional<models::CopyProgressSnapshot> last_progress_snapshot_;

    // Telemetry policy (mirrors ConfigureRuntimeTelemetryPolicy).
    std::int32_t progress_interval_ms_ = 75;
    std::int32_t max_logs_per_second_ = 100;
    ULONGLONG last_progress_sent_tick_ = 0;
    std::optional<models::CopyProgressSnapshot> pending_progress_;
    ULONGLONG log_window_start_tick_ = 0;
    std::int32_t log_window_sent_ = 0;
    std::int64_t suppressed_logs_ = 0;
    std::mutex telemetry_lock_;

    // ---- Sending ----------------------------------------------------------

    void send_envelope(std::string_view message_type, const std::function<void(json::Writer&)>& payload) {
        if (!pipe_.is_open()) return;
        std::string message = ipc::serialize_envelope(message_type, payload);
        std::lock_guard<std::mutex> guard(send_lock_);
        // Frame writes are never cancelled mid-write; a partial frame would
        // desynchronize the supervisor's parser (matches the .NET worker).
        pipe::CancellationEvent no_cancel;
        pipe_.write_message(message, no_cancel);
    }

    void send_log(std::string_view job_id, std::string_view text) {
        ipc::WorkerLogEvent log_event;
        log_event.JobId = std::string(job_id);
        log_event.Message = std::string(text);
        log_event.TimestampUtc = time::DateTimeOffset::now_utc();
        send_envelope(ipc::message_types::WorkerLogEvent,
                      [&log_event](json::Writer& w) { log_event.to_json(w); });
    }

    void send_heartbeat() {
        ipc::WorkerHeartbeatEvent heartbeat;
        heartbeat.WorkerProcessId = static_cast<std::int32_t>(GetCurrentProcessId());
        heartbeat.IsJobRunning = job_running_.load();
        heartbeat.IsJobPaused = job_paused_.load();
        {
            std::lock_guard<std::mutex> guard(job_state_lock_);
            heartbeat.ActiveJobId = active_job_id_;
            heartbeat.LastProgressUtc = last_progress_utc_;
        }
        heartbeat.TimestampUtc = time::DateTimeOffset::now_utc();
        send_envelope(ipc::message_types::WorkerHeartbeatEvent,
                      [&heartbeat](json::Writer& w) { heartbeat.to_json(w); });
    }

    void send_progress_event(const std::string& job_id, const models::CopyProgressSnapshot& snapshot) {
        ipc::WorkerProgressEvent progress;
        progress.JobId = job_id;
        progress.Snapshot = snapshot;
        progress.TimestampUtc = time::DateTimeOffset::now_utc();
        send_envelope(ipc::message_types::WorkerProgressEvent,
                      [&progress](json::Writer& w) { progress.to_json(w); });
    }

    // ---- Telemetry policy -------------------------------------------------

    void configure_telemetry(const models::CopyJobOptions& options) {
        std::int32_t default_interval = 75;
        std::int32_t default_logs = 100;
        switch (options.WorkerTelemetryProfileValue) {
            case models::WorkerTelemetryProfile::Verbose:
                default_interval = 40;
                default_logs = 300;
                break;
            case models::WorkerTelemetryProfile::Debug:
                default_interval = 20;
                default_logs = 0;
                break;
            default:
                break;
        }

        std::int32_t base_interval = options.WorkerProgressEmitIntervalMs > 0
                                         ? options.WorkerProgressEmitIntervalMs
                                         : default_interval;
        base_interval = std::max(20, std::min(1000, base_interval));
        std::int32_t base_logs = options.WorkerMaxLogsPerSecond >= 0
                                     ? std::min(5000, options.WorkerMaxLogsPerSecond)
                                     : default_logs;

        switch (options.WorkerTelemetryProfileValue) {
            case models::WorkerTelemetryProfile::Verbose:
                progress_interval_ms_ = std::max(20, base_interval / 2);
                max_logs_per_second_ = base_logs <= 0 ? 0 : std::min(5000, base_logs * 3);
                break;
            case models::WorkerTelemetryProfile::Debug:
                progress_interval_ms_ = std::max(20, base_interval / 3);
                max_logs_per_second_ = 0;
                break;
            default:
                progress_interval_ms_ = base_interval;
                max_logs_per_second_ = base_logs;
                break;
        }

        std::lock_guard<std::mutex> guard(telemetry_lock_);
        last_progress_sent_tick_ = 0;
        pending_progress_.reset();
        log_window_start_tick_ = 0;
        log_window_sent_ = 0;
        suppressed_logs_ = 0;
    }

    // Rate-limited progress with the file-transition guarantee: when the
    // current file changes, the previous pending snapshot is sent first so the
    // UI always observes each file's completion event.
    void on_engine_progress(const std::string& job_id, const models::CopyProgressSnapshot& snapshot) {
        {
            std::lock_guard<std::mutex> guard(job_state_lock_);
            last_progress_utc_ = time::DateTimeOffset::now_utc();
            last_progress_snapshot_ = snapshot;
        }

        // Coalesce to the send interval. The engine emits a snapshot per file,
        // so a whole-drive scan of many small files would otherwise flood the
        // pipe (thousands of frames/sec). The newest un-sent snapshot is kept in
        // pending_progress_ and delivered by flush_pending_progress at job end.
        bool send_now = false;
        {
            std::lock_guard<std::mutex> guard(telemetry_lock_);
            ULONGLONG now = GetTickCount64();
            if (last_progress_sent_tick_ == 0 ||
                now - last_progress_sent_tick_ >= static_cast<ULONGLONG>(progress_interval_ms_)) {
                send_now = true;
                last_progress_sent_tick_ = now;
                pending_progress_.reset();
            } else {
                pending_progress_ = snapshot;
            }
        }

        if (send_now) send_progress_event(job_id, snapshot);
    }

    void flush_pending_progress(const std::string& job_id) {
        std::optional<models::CopyProgressSnapshot> pending;
        {
            std::lock_guard<std::mutex> guard(telemetry_lock_);
            pending = std::move(pending_progress_);
            pending_progress_.reset();
        }
        if (pending.has_value()) send_progress_event(job_id, *pending);
    }

    void on_engine_log(const std::string& job_id, const std::string& message) {
        std::int64_t suppressed_summary = 0;
        bool should_send = false;
        {
            std::lock_guard<std::mutex> guard(telemetry_lock_);
            ULONGLONG now = GetTickCount64();
            if (log_window_start_tick_ == 0 || now - log_window_start_tick_ >= 1000) {
                suppressed_summary = suppressed_logs_;
                suppressed_logs_ = 0;
                log_window_start_tick_ = now;
                log_window_sent_ = 0;
            }
            if (max_logs_per_second_ <= 0 || log_window_sent_ < max_logs_per_second_) {
                log_window_sent_ += 1;
                should_send = true;
            } else {
                suppressed_logs_ += 1;
            }
        }

        if (suppressed_summary > 0) {
            send_log(job_id, "[WorkerTelemetry] Suppressed " + std::to_string(suppressed_summary) +
                                 " log message(s) due to telemetry rate policy.");
        }
        if (should_send) send_log(job_id, message);
    }

    // ---- Job lifecycle ----------------------------------------------------

    void handle_start_job(const ipc::StartJobCommand& command) {
        if (job_running_.load()) {
            send_log(active_job_id_snapshot(),
                     "Job start ignored because another job is already running.");
            return;
        }
        if (job_thread_.joinable()) job_thread_.join();

        {
            std::lock_guard<std::mutex> guard(job_state_lock_);
            active_job_id_ = command.JobId;
            last_progress_utc_ = time::DateTimeOffset::now_utc();
            last_progress_snapshot_.reset();
        }
        job_cancel_.store(false);
        job_paused_.store(false);
        control_.resume();
        job_running_.store(true);
        configure_telemetry(command.Options);
        apply_process_priority(command.Options.WorkerProcessPriorityClass);

        send_log(command.JobId, "Job accepted by worker.");

        std::string job_id = command.JobId;
        models::CopyJobOptions options = command.Options;
        job_thread_ = std::thread([this, job_id, options]() { run_job(job_id, options); });
    }

    void run_job(const std::string& job_id, const models::CopyJobOptions& options) {
        std::optional<ipc::WorkerJobResultEvent> result_event;
        std::optional<ipc::WorkerFatalEvent> fatal_event;

        try {
            engine::ResilientCopyEngine copy_engine(
                options, &control_,
                [this, job_id](const models::CopyProgressSnapshot& snapshot) {
                    on_engine_progress(job_id, snapshot);
                },
                [this, job_id](const std::string& message) { on_engine_log(job_id, message); });

            models::CopyJobResult result = copy_engine.run(job_cancel_);
            if (job_cancel_.load()) {
                result.Cancelled = true;
                apply_last_progress(result);
            }
            ipc::WorkerJobResultEvent event;
            event.JobId = job_id;
            event.Result = result;
            event.TimestampUtc = time::DateTimeOffset::now_utc();
            result_event = std::move(event);
        } catch (const engine::OperationCanceled&) {
            models::CopyJobResult result;
            result.Succeeded = false;
            result.Cancelled = true;
            result.ErrorMessage = "Job cancelled by supervisor.";
            result.TransferEnginePolicyValue = options.TransferEnginePolicyValue;
            result.JournalPath = journal_path_for_options(options);
            apply_last_progress(result);
            ipc::WorkerJobResultEvent event;
            event.JobId = job_id;
            event.Result = result;
            event.TimestampUtc = time::DateTimeOffset::now_utc();
            result_event = std::move(event);
        } catch (const std::exception& ex) {
            ipc::WorkerFatalEvent event;
            event.JobId = job_id;
            event.ErrorMessage = ex.what();
            event.StackTraceText = ex.what();
            event.TimestampUtc = time::DateTimeOffset::now_utc();
            fatal_event = std::move(event);
        }

        {
            std::lock_guard<std::mutex> guard(job_state_lock_);
            active_job_id_.clear();
        }
        job_running_.store(false);
        job_paused_.store(false);

        try {
            flush_pending_progress(job_id);
        } catch (const std::exception&) {
        }

        try {
            if (result_event.has_value()) {
                send_envelope(ipc::message_types::WorkerJobResultEvent,
                              [&](json::Writer& w) { result_event->to_json(w); });
            }
            if (fatal_event.has_value()) {
                send_envelope(ipc::message_types::WorkerFatalEvent,
                              [&](json::Writer& w) { fatal_event->to_json(w); });
            }
        } catch (const std::exception&) {
        }
    }

    static std::string journal_path_for_options(const models::CopyJobOptions& options) {
        if (!options.ResumeJournalPathHint.empty() &&
            storage::fsutil::file_exists(
                storage::fsutil::utf8_to_wide(options.ResumeJournalPathHint))) {
            return options.ResumeJournalPathHint;
        }
        std::wstring source = storage::fsutil::get_full_path(
            storage::fsutil::utf8_to_wide(options.SourceRoot));
        std::wstring destination = storage::fsutil::get_full_path(
            storage::fsutil::utf8_to_wide(options.DestinationRoot));
        if (options.OperationMode == models::JobOperationMode::ScanOnly &&
            options.DestinationRoot.empty()) {
            destination = source;
        }
        if (source.empty() || destination.empty()) return std::string();
        const std::string journal_id =
            storage::JobJournalStore::build_job_id(source, destination);
        return storage::fsutil::wide_to_utf8(
            storage::JobJournalStore::get_default_journal_path(journal_id));
    }

    void apply_last_progress(models::CopyJobResult& result) {
        std::lock_guard<std::mutex> guard(job_state_lock_);
        if (!last_progress_snapshot_.has_value()) return;
        const auto& snapshot = *last_progress_snapshot_;
        result.TotalFiles = std::max(result.TotalFiles, snapshot.TotalFiles);
        result.CompletedFiles = std::max(result.CompletedFiles, snapshot.CompletedFiles);
        result.FailedFiles = std::max(result.FailedFiles, snapshot.FailedFiles);
        result.RecoveredFiles = std::max(result.RecoveredFiles, snapshot.RecoveredFiles);
        result.SkippedFiles = std::max(result.SkippedFiles, snapshot.SkippedFiles);
        result.TotalBytes = std::max(result.TotalBytes, snapshot.TotalBytes);
        result.WorkBytesCompleted =
            std::max(result.WorkBytesCompleted, snapshot.WorkBytesCompleted);
        result.BytesRead = std::max(result.BytesRead, snapshot.BytesRead);
        result.BytesWritten = std::max(result.BytesWritten, snapshot.BytesWritten);
        result.BytesVerified = std::max(result.BytesVerified, snapshot.BytesVerified);
        result.BytesSkipped = std::max(result.BytesSkipped, snapshot.BytesSkipped);
        result.BytesReused = std::max(result.BytesReused, snapshot.BytesReused);
        // TotalBytesCopied is logical progress and includes skipped/reused work.
        // CopiedBytes is retained for compatibility but must reflect actual
        // destination writes in results emitted from a cancelled operation.
        result.CopiedBytes = std::max(result.CopiedBytes, snapshot.BytesWritten);
    }

    void apply_process_priority(const std::string& requested) {
        std::string normalized;
        for (char c : requested) {
            if (c == ' ' || c == '_') continue;
            normalized.push_back(c >= 'a' && c <= 'z' ? static_cast<char>(c - 'a' + 'A') : c);
        }
        DWORD priority_class = 0;
        if (normalized == "IDLE") priority_class = IDLE_PRIORITY_CLASS;
        else if (normalized == "BELOWNORMAL") priority_class = BELOW_NORMAL_PRIORITY_CLASS;
        else if (normalized == "NORMAL") priority_class = NORMAL_PRIORITY_CLASS;
        else if (normalized == "ABOVENORMAL") priority_class = ABOVE_NORMAL_PRIORITY_CLASS;
        else if (normalized == "HIGH") priority_class = HIGH_PRIORITY_CLASS;
        if (priority_class == 0) return;
        if (GetPriorityClass(GetCurrentProcess()) == priority_class) return;
        if (SetPriorityClass(GetCurrentProcess(), priority_class)) {
            std::string label = normalized == "BELOWNORMAL" ? "BelowNormal"
                                : normalized == "ABOVENORMAL" ? "AboveNormal"
                                : normalized == "IDLE" ? "Idle"
                                : normalized == "HIGH" ? "High" : "Normal";
            send_log(active_job_id_snapshot(), "Worker process priority set to " + label + ".");
        }
    }

    std::string active_job_id_snapshot() {
        std::lock_guard<std::mutex> guard(job_state_lock_);
        return active_job_id_;
    }

    // ---- Loops ------------------------------------------------------------

    void heartbeat_loop() {
        while (!lifetime_.is_cancelled()) {
            try {
                send_heartbeat();
            } catch (const std::exception&) {
                return;
            }
            if (WaitForSingleObject(lifetime_.handle(), 1000) == WAIT_OBJECT_0) {
                return;
            }
        }
    }

    void receive_loop() {
        while (pipe_.is_open() && !lifetime_.is_cancelled()) {
            std::optional<std::string> incoming = pipe_.read_message(lifetime_);
            if (!incoming.has_value()) {
                return; // Peer closed the pipe.
            }
            if (!handle_incoming(*incoming)) {
                return; // Shutdown requested.
            }
        }
    }

    // Returns false when the worker should shut down.
    bool handle_incoming(const std::string& incoming_json) {
        std::string message_type;
        if (!ipc::try_read_message_type(incoming_json, message_type)) {
            send_log(active_job_id_snapshot(), "Received malformed IPC message.");
            return true;
        }

        if (message_type == ipc::message_types::StartJobCommand) {
            auto [header, payload] = ipc::parse_envelope(incoming_json);
            handle_start_job(ipc::StartJobCommand::from_json(payload));
            return true;
        }
        if (message_type == ipc::message_types::CancelJobCommand) {
            auto [header, payload] = ipc::parse_envelope(incoming_json);
            auto command = ipc::CancelJobCommand::from_json(payload);
            if (!job_running_.load()) {
                send_log(command.JobId, "Cancel command ignored because no active job is running.");
                return true;
            }
            control_.resume();
            job_paused_.store(false);
            job_cancel_.store(true);
            send_log(command.JobId, "Cancel requested: " + command.Reason);
            return true;
        }
        if (message_type == ipc::message_types::PauseJobCommand) {
            auto [header, payload] = ipc::parse_envelope(incoming_json);
            auto command = ipc::PauseJobCommand::from_json(payload);
            if (!job_running_.load()) {
                send_log(command.JobId, "Pause command ignored because no active job is running.");
                return true;
            }
            control_.pause();
            job_paused_.store(true);
            send_log(command.JobId, "Pause requested: " + command.Reason);
            return true;
        }
        if (message_type == ipc::message_types::ResumeJobCommand) {
            auto [header, payload] = ipc::parse_envelope(incoming_json);
            auto command = ipc::ResumeJobCommand::from_json(payload);
            if (!job_running_.load()) {
                send_log(command.JobId, "Resume command ignored because no active job is running.");
                return true;
            }
            control_.resume();
            job_paused_.store(false);
            send_log(command.JobId, "Resume requested: " + command.Reason);
            return true;
        }
        if (message_type == ipc::message_types::PingCommand) {
            send_heartbeat();
            return true;
        }
        if (message_type == ipc::message_types::ShutdownCommand) {
            send_log(active_job_id_snapshot(), "Shutdown command received.");
            return false;
        }

        send_log(active_job_id_snapshot(), "Unsupported message type: " + message_type);
        return true;
    }
};

std::wstring read_pipe_name(int argc, wchar_t** argv) {
    for (int i = 1; i + 1 < argc; ++i) {
        if (_wcsicmp(argv[i], L"--pipe") == 0) {
            return argv[i + 1];
        }
    }
    return std::wstring();
}

DWORD read_parent_pid(int argc, wchar_t** argv) {
    for (int i = 1; i + 1 < argc; ++i) {
        if (_wcsicmp(argv[i], L"--parent-pid") == 0) {
            wchar_t* end = nullptr;
            unsigned long value = std::wcstoul(argv[i + 1], &end, 10);
            if (end != argv[i + 1] && end != nullptr && *end == L'\0' && value <= MAXDWORD) {
                return static_cast<DWORD>(value);
            }
        }
    }
    return 0;
}

void notify_manual_launch() {
    static const wchar_t* message =
        L"XactCopyExecutive is a background worker and must be started by XactCopy.\r\n"
        L"Use XactCopy.exe to start copy jobs.";
    std::fwprintf(stdout, L"%s\n", message);
    MessageBoxW(nullptr, message, L"XactCopyExecutive", MB_OK | MB_ICONINFORMATION);
}

} // namespace

int wmain(int argc, wchar_t** argv) {
    std::wstring pipe_name = read_pipe_name(argc, argv);
    DWORD parent_pid = read_parent_pid(argc, argv);
    if (pipe_name.empty() || parent_pid == 0) {
        notify_manual_launch();
        return 2;
    }

    WorkerHost host;
    return host.run(pipe_name, parent_pid);
}
