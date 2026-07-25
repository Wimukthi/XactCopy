// -----------------------------------------------------------------------------
// File: cpp\src\ui\recovery.h
// Purpose: Native port of the crash-recovery layer: RecoveryStateStore
//          (runtime\recovery-state.json, .NET-compatible), RecoveryService
//          (interrupted-run detection, prompt/auto-resume policy, run
//          touch/end/interrupt marks), and the RunOnce autostart hook.
// -----------------------------------------------------------------------------

#pragma once

#include <mutex>
#include <optional>
#include <string>
#include <vector>

#include "../core/models.h"
#include "../storage/stores.h"
#include "settings.h"

namespace xact::ui {

struct LaunchOptions {
    bool is_recovery_autostart = false;
    bool force_resume_prompt = false;
    bool explorer_scan_mode = false; // launched via the "Scan for Bad Blocks" verb
    std::string explorer_folder_path;
    std::vector<std::string> explorer_source_paths;
};

struct RecoveryActiveRun {
    std::string RunId;
    std::string JobId;
    std::string JobName;
    std::string Trigger = "manual";
    models::CopyJobOptions Options;
    time::DateTimeOffset StartedUtc = time::DateTimeOffset::now_utc();
    time::DateTimeOffset LastHeartbeatUtc = time::DateTimeOffset::now_utc();
    time::DateTimeOffset LastProgressUtc = time::DateTimeOffset::now_utc();
    std::string JournalPath;

    void to_json(json::Writer& w) const {
        w.begin_object();
        w.key("RunId"); w.value(RunId);
        w.key("JobId"); w.value(JobId);
        w.key("JobName"); w.value(JobName);
        w.key("Trigger"); w.value(Trigger);
        w.key("Options");
        Options.to_json(w);
        w.key("StartedUtc"); w.value_literal_string(StartedUtc.to_string());
        w.key("LastHeartbeatUtc"); w.value_literal_string(LastHeartbeatUtc.to_string());
        w.key("LastProgressUtc"); w.value_literal_string(LastProgressUtc.to_string());
        w.key("JournalPath"); w.value(JournalPath);
        w.end_object();
    }

    static RecoveryActiveRun from_json(const json::Value& value) {
        RecoveryActiveRun run;
        const auto* obj = value.as_object();
        if (obj == nullptr) return run;
        models::detail::read(obj, "RunId", run.RunId);
        models::detail::read(obj, "JobId", run.JobId);
        models::detail::read(obj, "JobName", run.JobName);
        models::detail::read(obj, "Trigger", run.Trigger);
        if (const auto* v = obj->find("Options"); v != nullptr) {
            run.Options = models::CopyJobOptions::from_json(*v);
        }
        models::detail::read(obj, "StartedUtc", run.StartedUtc);
        models::detail::read(obj, "LastHeartbeatUtc", run.LastHeartbeatUtc);
        models::detail::read(obj, "LastProgressUtc", run.LastProgressUtc);
        models::detail::read(obj, "JournalPath", run.JournalPath);
        return run;
    }
};

struct RecoveryState {
    std::string ProcessSessionId;
    time::DateTimeOffset LastStartUtc = time::DateTimeOffset::min_value();
    time::DateTimeOffset LastExitUtc = time::DateTimeOffset::min_value();
    bool CleanShutdown = true;
    bool PendingResumePrompt = false;
    std::string LastInterruptionReason;
    std::optional<RecoveryActiveRun> ActiveRun;

    void to_json(json::Writer& w) const {
        w.begin_object();
        w.key("ProcessSessionId"); w.value(ProcessSessionId);
        w.key("LastStartUtc"); w.value_literal_string(LastStartUtc.to_string());
        w.key("LastExitUtc"); w.value_literal_string(LastExitUtc.to_string());
        w.key("CleanShutdown"); w.value(CleanShutdown);
        w.key("PendingResumePrompt"); w.value(PendingResumePrompt);
        w.key("LastInterruptionReason"); w.value(LastInterruptionReason);
        w.key("ActiveRun");
        if (ActiveRun.has_value()) ActiveRun->to_json(w);
        else w.value(nullptr);
        w.end_object();
    }

    static RecoveryState from_json(const json::Value& value) {
        RecoveryState state;
        const auto* obj = value.as_object();
        if (obj == nullptr) return state;
        models::detail::read(obj, "ProcessSessionId", state.ProcessSessionId);
        models::detail::read(obj, "LastStartUtc", state.LastStartUtc);
        models::detail::read(obj, "LastExitUtc", state.LastExitUtc);
        models::detail::read(obj, "CleanShutdown", state.CleanShutdown);
        models::detail::read(obj, "PendingResumePrompt", state.PendingResumePrompt);
        models::detail::read(obj, "LastInterruptionReason", state.LastInterruptionReason);
        if (const auto* v = obj->find("ActiveRun"); v != nullptr && v->is_object()) {
            state.ActiveRun = RecoveryActiveRun::from_json(*v);
        }
        return state;
    }
};

class RecoveryStateStore {
public:
    static std::wstring default_path() {
        return storage::fsutil::local_app_data() + L"\\XactCopy\\runtime\\recovery-state.json";
    }

    explicit RecoveryStateStore(std::wstring path = default_path()) : path_(std::move(path)) {}

    const std::wstring& state_path() const noexcept { return path_; }

    RecoveryState load() const {
        auto bytes = storage::fsutil::read_all_bytes(path_);
        if (!bytes.has_value() || bytes->empty()) return RecoveryState{};
        try {
            std::string_view text(reinterpret_cast<const char*>(bytes->data()), bytes->size());
            return RecoveryState::from_json(json::parse(text));
        } catch (const std::exception&) {
            return RecoveryState{};
        }
    }

    void save(const RecoveryState& state) const {
        std::wstring directory = storage::fsutil::get_directory_name(path_);
        if (!directory.empty()) storage::fsutil::create_directories(directory);

        json::Writer w(/*indented*/ true);
        state.to_json(w);
        std::string text = w.take();

        // Matches the .NET store's fixed ".tmp" temp name + replace move.
        std::wstring temp_path = path_ + L".tmp";
        DeleteFileW(temp_path.c_str());
        if (storage::fsutil::write_file_raw(temp_path,
                                            reinterpret_cast<const unsigned char*>(text.data()),
                                            text.size(), false, true)) {
            MoveFileExW(temp_path.c_str(), path_.c_str(), MOVEFILE_REPLACE_EXISTING);
        }
    }

private:
    std::wstring path_;
};

// RunOnce registration so an interrupted session relaunches the UI once after
// a crash/reboot (port of WindowsStartupService).
class WindowsStartupService {
public:
    void register_recovery_run_once() {
        wchar_t module_path[MAX_PATH];
        DWORD length = GetModuleFileNameW(nullptr, module_path, MAX_PATH);
        std::wstring command = L"\"" + std::wstring(module_path, length) + L"\" --recovery-autostart";
        HKEY key = nullptr;
        if (RegOpenKeyExW(HKEY_CURRENT_USER,
                          L"Software\\Microsoft\\Windows\\CurrentVersion\\RunOnce", 0,
                          KEY_SET_VALUE, &key) == ERROR_SUCCESS) {
            RegSetValueExW(key, L"XactCopyRecovery", 0, REG_SZ,
                           reinterpret_cast<const BYTE*>(command.c_str()),
                           static_cast<DWORD>((command.size() + 1) * sizeof(wchar_t)));
            RegCloseKey(key);
        }
    }

    void clear_recovery_run_once() {
        HKEY key = nullptr;
        if (RegOpenKeyExW(HKEY_CURRENT_USER,
                          L"Software\\Microsoft\\Windows\\CurrentVersion\\RunOnce", 0,
                          KEY_SET_VALUE, &key) == ERROR_SUCCESS) {
            RegDeleteValueW(key, L"XactCopyRecovery");
            RegCloseKey(key);
        }
    }
};

struct RecoveryStartupInfo {
    std::optional<RecoveryActiveRun> interrupted_run;
    std::string interruption_reason;
    bool should_prompt = false;
    bool should_auto_resume = false;
    bool was_recovery_autostart = false;

    bool has_interrupted_run() const { return interrupted_run.has_value(); }
};

class RecoveryService {
public:
    explicit RecoveryService(RecoveryStateStore store = RecoveryStateStore{},
                             bool manage_run_once = true)
        : store_(std::move(store)), manage_run_once_(manage_run_once) {
        state_ = store_.load();
    }

    RecoveryStartupInfo initialize_session(const AppSettings& settings,
                                           const LaunchOptions& launch) {
        std::lock_guard<std::mutex> guard(lock_);
        state_ = store_.load();

        RecoveryStartupInfo info;
        bool has_pending = state_.ActiveRun.has_value() && state_.PendingResumePrompt;
        bool interrupted = (!state_.CleanShutdown && state_.ActiveRun.has_value()) || has_pending;

        if (interrupted) {
            info.interrupted_run = state_.ActiveRun;
            info.interruption_reason = state_.LastInterruptionReason.empty()
                                           ? "The previous copy session ended unexpectedly."
                                           : state_.LastInterruptionReason;
            state_.PendingResumePrompt = true;
            state_.LastInterruptionReason = info.interruption_reason;
        }

        bool force_prompt = launch.is_recovery_autostart || launch.force_resume_prompt;
        bool prompt_pending = info.has_interrupted_run() && state_.PendingResumePrompt;
        info.should_auto_resume = prompt_pending && settings.auto_resume_after_crash();
        info.should_prompt = !info.should_auto_resume && prompt_pending &&
                             (force_prompt || settings.prompt_resume_after_crash());
        info.was_recovery_autostart = launch.is_recovery_autostart;

        state_.ProcessSessionId = storage::detail::new_guid_n();
        state_.LastStartUtc = time::DateTimeOffset::now_utc();
        state_.CleanShutdown = false;
        save_no_throw();
        return info;
    }

    void mark_resume_prompt_deferred(const AppSettings& settings, bool suppress_for_this_run) {
        std::lock_guard<std::mutex> guard(lock_);
        if (!state_.ActiveRun.has_value()) return;
        state_.PendingResumePrompt =
            suppress_for_this_run ? false : settings.keep_resume_prompt_until_resolved();
        save_no_throw();
    }

    void mark_job_started(const std::string& run_id, const std::string& job_name,
                          const models::CopyJobOptions& options, const std::string& journal_path,
                          const AppSettings& settings) {
        std::lock_guard<std::mutex> guard(lock_);
        time::DateTimeOffset now = time::DateTimeOffset::now_utc();

        RecoveryActiveRun run;
        run.RunId = run_id;
        run.JobId = run_id;
        run.JobName = job_name;
        run.Trigger = "manual";
        run.Options = options;
        run.StartedUtc = now;
        run.LastHeartbeatUtc = now;
        run.LastProgressUtc = now;
        run.JournalPath = journal_path;
        state_.ActiveRun = std::move(run);

        state_.PendingResumePrompt = false;
        state_.LastInterruptionReason.clear();
        state_.CleanShutdown = false;
        state_.LastStartUtc = now;
        last_touch_tick_ = GetTickCount64();
        touch_interval_ms_ =
            static_cast<ULONGLONG>(std::max(1, settings.recovery_touch_interval_seconds())) * 1000;

        save_no_throw();
        if (manage_run_once_) {
            if (settings.enable_recovery_autostart()) startup_.register_recovery_run_once();
            else startup_.clear_recovery_run_once();
        }
    }

    void touch_active_run(const std::string& run_id) {
        std::lock_guard<std::mutex> guard(lock_);
        if (!state_.ActiveRun.has_value()) return;
        if (!models::detail::equals_ignore_case(state_.ActiveRun->RunId, run_id)) return;

        ULONGLONG now = GetTickCount64();
        if (last_touch_tick_ != 0 && now - last_touch_tick_ < touch_interval_ms_) return;

        time::DateTimeOffset now_utc = time::DateTimeOffset::now_utc();
        state_.ActiveRun->LastHeartbeatUtc = now_utc;
        state_.ActiveRun->LastProgressUtc = now_utc;
        last_touch_tick_ = now;
        save_no_throw();
    }

    void mark_job_ended(const std::string& run_id) {
        std::lock_guard<std::mutex> guard(lock_);
        if (state_.ActiveRun.has_value() &&
            models::detail::equals_ignore_case(state_.ActiveRun->RunId, run_id)) {
            state_.ActiveRun.reset();
            state_.PendingResumePrompt = false;
            state_.LastInterruptionReason.clear();
            save_no_throw();
        }
        if (manage_run_once_) startup_.clear_recovery_run_once();
    }

    void mark_run_interrupted(const std::string& reason) {
        std::lock_guard<std::mutex> guard(lock_);
        if (!state_.ActiveRun.has_value()) return;
        state_.PendingResumePrompt = true;
        state_.LastInterruptionReason =
            reason.empty() ? "Copy session was interrupted." : reason;
        save_no_throw();
    }

    void mark_clean_shutdown() {
        std::lock_guard<std::mutex> guard(lock_);
        state_.CleanShutdown = true;
        state_.LastExitUtc = time::DateTimeOffset::now_utc();
        save_no_throw();
        if (manage_run_once_) startup_.clear_recovery_run_once();
    }

    std::optional<RecoveryActiveRun> get_pending_interrupted_run() {
        std::lock_guard<std::mutex> guard(lock_);
        return state_.ActiveRun;
    }

private:
    RecoveryStateStore store_;
    WindowsStartupService startup_;
    bool manage_run_once_;
    std::mutex lock_;
    RecoveryState state_;
    ULONGLONG last_touch_tick_ = 0;
    ULONGLONG touch_interval_ms_ = 2000;

    void save_no_throw() {
        try {
            store_.save(state_);
        } catch (const std::exception&) {
        }
    }
};

} // namespace xact::ui
