// -----------------------------------------------------------------------------
// File: tests\test_supervisor.cpp
// Purpose: Headless tests for the native UI core: WorkerSupervisor driving the
//          native worker end-to-end (clean job, kill-mid-job auto-recovery
//          with journal resume), RecoveryService state round-trip, and
//          settings DOM preservation. Runs from the build output directory so
//          the supervisor resolves XactCopyExecutive.exe beside this test exe.
// -----------------------------------------------------------------------------

#include <atomic>
#include <chrono>
#include <condition_variable>
#include <cstdio>
#include <mutex>
#include <sstream>
#include <string>
#include <vector>

#include "../src/ui/elevation.h"
#include "../src/ui/task_file.h"
#include "../src/ui/task_model.h"
#include "../src/ui/job_manager.h"
#include "../src/ui/journal_report.h"
#include "../src/ui/log_model.h"
#include "../src/ui/range_map.h"
#include "../src/ui/recovery.h"
#include "../src/ui/settings.h"
#include "../src/ui/supervisor.h"

#include <tlhelp32.h>

namespace {

using namespace xact;

int g_failures = 0;
int g_checks = 0;

void check(bool condition, const std::string& label) {
    ++g_checks;
    std::printf("  %s: %s\n", condition ? "ok" : "FAIL", label.c_str());
    if (!condition) ++g_failures;
}

std::wstring make_work_dir() {
    wchar_t temp_root[MAX_PATH];
    GetTempPathW(MAX_PATH, temp_root);
    std::wstring dir = std::wstring(temp_root) + L"xactcopy-sup-" +
                       storage::fsutil::random_temp_suffix().substr(0, 12);
    storage::fsutil::create_directories(dir);
    return dir;
}

void write_pattern_file(const std::wstring& path, std::size_t size, unsigned seed) {
    std::vector<unsigned char> data(size);
    unsigned state = seed;
    for (auto& byte : data) {
        state = state * 1664525u + 1013904223u;
        byte = static_cast<unsigned char>(state >> 24);
    }
    storage::fsutil::write_file_raw(path, data.data(), data.size(), false, false);
}

bool files_equal(const std::wstring& a, const std::wstring& b) {
    auto left = storage::fsutil::read_all_bytes(a);
    auto right = storage::fsutil::read_all_bytes(b);
    return left.has_value() && right.has_value() && *left == *right;
}

// Collects supervisor events with simple waiting helpers.
struct EventSink {
    std::mutex lock;
    std::condition_variable signal;
    std::vector<std::string> logs;
    std::vector<std::string> states;
    std::vector<models::CopyJobResult> results;
    std::atomic<int> progress_events{0};

    void attach(ui::WorkerSupervisor& supervisor) {
        supervisor.on_log = [this](const std::string& message) {
            std::lock_guard<std::mutex> guard(lock);
            logs.push_back(message);
            signal.notify_all();
        };
        supervisor.on_worker_state_changed = [this](const std::string& state) {
            std::lock_guard<std::mutex> guard(lock);
            states.push_back(state);
            signal.notify_all();
        };
        supervisor.on_progress = [this](const models::CopyProgressSnapshot&) {
            progress_events.fetch_add(1);
            signal.notify_all();
        };
        supervisor.on_job_completed = [this](const models::CopyJobResult& result) {
            std::lock_guard<std::mutex> guard(lock);
            results.push_back(result);
            signal.notify_all();
        };
    }

    bool wait_for_result(int timeout_seconds) {
        std::unique_lock<std::mutex> guard(lock);
        return signal.wait_for(guard, std::chrono::seconds(timeout_seconds),
                               [this] { return !results.empty(); });
    }

    bool wait_for_progress(int timeout_seconds) {
        std::unique_lock<std::mutex> guard(lock);
        return signal.wait_for(guard, std::chrono::seconds(timeout_seconds),
                               [this] { return progress_events.load() > 0; });
    }

    bool wait_for_log(const std::string& fragment, int timeout_seconds) {
        std::unique_lock<std::mutex> guard(lock);
        return signal.wait_for(guard, std::chrono::seconds(timeout_seconds), [this, &fragment] {
            for (const auto& line : logs) {
                if (line.find(fragment) != std::string::npos) return true;
            }
            return false;
        });
    }

    bool any_log_contains(const std::string& fragment) {
        std::lock_guard<std::mutex> guard(lock);
        for (const auto& line : logs) {
            if (line.find(fragment) != std::string::npos) return true;
        }
        return false;
    }

    bool any_state_contains(const std::string& fragment, std::string* matched = nullptr) {
        std::lock_guard<std::mutex> guard(lock);
        for (const auto& state : states) {
            if (state.find(fragment) != std::string::npos) {
                if (matched != nullptr) *matched = state;
                return true;
            }
        }
        return false;
    }
};

void test_automatic_recovery_budget() {
    std::printf("--- supervisor: bounded automatic recovery policy ---\n");
    std::atomic<int> attempts{0};
    for (int expected = 1; expected <= ui::supervisor_detail::MaximumAutomaticRecoveries;
         ++expected) {
        int claimed = 0;
        check(ui::supervisor_detail::try_claim_automatic_recovery(attempts, claimed) &&
                  claimed == expected,
              "automatic recovery attempt " + std::to_string(expected) + " is permitted");
    }
    int rejected = 0;
    check(!ui::supervisor_detail::try_claim_automatic_recovery(attempts, rejected) &&
              rejected == ui::supervisor_detail::MaximumAutomaticRecoveries + 1,
          "automatic recovery stops after the configured budget");
}

void test_clean_job() {
    std::printf("--- supervisor: clean job ---\n");
    std::wstring work = make_work_dir();
    std::wstring source = work + L"\\src";
    std::wstring destination = work + L"\\dst";
    storage::fsutil::create_directories(source);
    write_pattern_file(source + L"\\one.bin", 512 * 1024, 11);
    write_pattern_file(source + L"\\two.bin", 64 * 1024, 22);

    ui::WorkerSupervisor supervisor;
    EventSink sink;
    sink.attach(supervisor);

    models::CopyJobOptions options;
    options.SourceRoot = storage::fsutil::wide_to_utf8(source);
    options.DestinationRoot = storage::fsutil::wide_to_utf8(destination);
    supervisor.start_job(options);

    check(sink.wait_for_result(60), "clean job completes");
    {
        std::lock_guard<std::mutex> guard(sink.lock);
        if (!sink.results.empty()) {
            const auto& result = sink.results.front();
            check(result.Succeeded, "clean job succeeded (" + result.ErrorMessage + ")");
            check(result.CompletedFiles == 2,
                  "clean job completed files (" + std::to_string(result.CompletedFiles) + ")");
        }
    }
    check(sink.any_state_contains("Worker connected (PID"), "worker-connected state raised");
    check(sink.any_state_contains("Job finished"), "job-finished state raised");
    check(files_equal(source + L"\\one.bin", destination + L"\\one.bin"), "destination one.bin equal");
    check(files_equal(source + L"\\two.bin", destination + L"\\two.bin"), "destination two.bin equal");

    supervisor.stop();
    check(true, "supervisor stopped cleanly");
}

void test_kill_recovery() {
    std::printf("--- supervisor: kill-mid-job auto-recovery ---\n");
    std::wstring work = make_work_dir();
    std::wstring source = work + L"\\src";
    std::wstring destination = work + L"\\dst";
    storage::fsutil::create_directories(source);
    // 24 MB throttled to 2 MB/s -> ~12 s run; plenty of time to kill mid-copy.
    write_pattern_file(source + L"\\big.bin", 24 * 1024 * 1024, 33);

    ui::WorkerSupervisor supervisor;
    EventSink sink;
    sink.attach(supervisor);

    models::CopyJobOptions options;
    options.SourceRoot = storage::fsutil::wide_to_utf8(source);
    options.DestinationRoot = storage::fsutil::wide_to_utf8(destination);
    options.TransferEnginePolicyValue = models::TransferEnginePolicy::ManagedRescue;
    options.MaxThroughputBytesPerSecond = 2 * 1024 * 1024;
    options.ResumeFromJournal = true;
    supervisor.start_job(options);

    check(sink.wait_for_progress(30), "progress observed before kill");

    // Extract the worker PID from the state message and kill it.
    std::string connected_state;
    check(sink.any_state_contains("Worker connected (PID", &connected_state),
          "worker PID state available");
    DWORD pid = 0;
    if (std::size_t open = connected_state.find("PID "); open != std::string::npos) {
        pid = static_cast<DWORD>(std::strtoul(connected_state.c_str() + open + 4, nullptr, 10));
    }
    check(pid != 0, "worker PID parsed (" + std::to_string(pid) + ")");
    if (pid != 0) {
        Sleep(1500); // let some bytes land so resume has coverage to keep
        HANDLE process = OpenProcess(PROCESS_TERMINATE, FALSE, pid);
        check(process != nullptr, "opened worker process for kill");
        if (process != nullptr) {
            TerminateProcess(process, 9);
            CloseHandle(process);
        }
    }

    check(sink.wait_for_result(180), "job completes after recovery");
    {
        std::lock_guard<std::mutex> guard(sink.lock);
        if (!sink.results.empty()) {
            const auto& result = sink.results.front();
            check(result.Succeeded,
                  "recovered job succeeded (" + result.ErrorMessage + ")");
        }
    }
    check(sink.any_log_contains("Restarting worker and resuming from journal"),
          "recovery restart log present");
    check(sink.any_log_contains("Loaded journal:"), "worker resumed from journal");
    check(files_equal(source + L"\\big.bin", destination + L"\\big.bin"),
          "destination equals source after recovery");

    supervisor.stop();
    check(true, "supervisor stopped cleanly after recovery");
}

void test_task_model() {
    std::printf("--- task model: source groups, tasks, persistence ---\n");
    auto item = [](const wchar_t* path, bool directory = false, bool contents = false) {
        ui::SourceItem source;
        source.path = path;
        source.is_directory = directory;
        source.contents_only = contents;
        return source;
    };

    // Items sharing a parent become one job rooted at that parent.
    auto groups = ui::group_sources({item(L"C:\\Users\\me\\Photos", true),
                                     item(L"C:\\Users\\me\\notes.txt"),
                                     item(L"D:\\Media\\Photos", true),
                                     item(L"C:\\Users\\me\\Archive", true, true)});
    check(groups.size() == 3, "three source groups (shared parent, other drive, contents-only)");
    check(groups[0].root == L"C:\\Users\\me" && groups[0].relative_paths.size() == 2 &&
              groups[0].relative_paths[0] == "Photos" && groups[0].relative_paths[1] == "notes.txt",
          "items in one folder keep their own names under one root");
    check(groups[1].root == L"D:\\Media" && groups[1].relative_paths.size() == 1,
          "an item from another drive is its own job");
    check(groups[2].contents_only && groups[2].root == L"C:\\Users\\me\\Archive" &&
              groups[2].relative_paths.empty(),
          "a contents-only folder copies everything beneath it");
    auto root_item = ui::group_sources({item(L"E:\\file.bin")});
    check(root_item.size() == 1 && root_item[0].root == L"E:\\",
          "a file at a drive root is rooted at the drive, not \"E:\"");
    auto duplicates = ui::group_sources({item(L"C:\\a\\x.txt"), item(L"c:\\A\\X.TXT")});
    check(duplicates.size() == 1 && duplicates[0].relative_paths.size() == 1,
          "the same item twice is listed once");

    const auto collisions = ui::find_destination_collisions(groups);
    check(collisions.size() == 1 && _wcsicmp(collisions[0].c_str(), L"Photos") == 0,
          "two different Photos folders are reported as a destination collision");
    check(ui::find_destination_collisions(
              ui::group_sources({item(L"C:\\a\\one"), item(L"C:\\b\\two")})).empty(),
          "distinct names from different folders do not collide");

    models::CopyJobOptions options;
    options.SourceRoot = "C:\\Users\\me";
    options.SelectedRelativePaths = {"Photos", "notes.txt"};
    auto restored = ui::items_from_options(options);
    check(restored.size() == 2 && restored[0].path == L"C:\\Users\\me\\Photos",
          "job options map back to source items");
    options.SelectedRelativePaths.clear();
    restored = ui::items_from_options(options);
    check(restored.size() == 1 && restored[0].contents_only,
          "a whole-folder job maps back to a contents-only folder");

    // Result mapping.
    models::CopyJobResult result;
    result.Succeeded = true;
    check(ui::state_from_result(result) == ui::TaskState::Completed, "clean success is Completed");
    result.RecoveredFiles = 1;
    check(ui::state_from_result(result) == ui::TaskState::CompletedWithRecovery,
          "recovered data is not reported as a plain completion");
    result = {};
    result.Cancelled = true;
    check(ui::state_from_result(result) == ui::TaskState::Cancelled, "cancel maps to Cancelled");
    result = {};
    result.ErrorMessage = "boom";
    check(ui::state_from_result(result) == ui::TaskState::Failed, "an error maps to Failed");

    // Run records map onto the states the panes draw.
    storage::ManagedJobRun run;
    run.Status = storage::ManagedJobRunStatus::Completed;
    check(ui::state_from_run(run) == ui::TaskState::Completed, "a completed run without a result");
    models::CopyJobResult recovered;
    recovered.Succeeded = true;
    recovered.RecoveredFiles = 2;
    run.Result = recovered;
    check(ui::state_from_run(run) == ui::TaskState::CompletedWithRecovery,
          "a completed run keeps its recovered-data outcome");
    run.Status = storage::ManagedJobRunStatus::Interrupted;
    check(ui::state_from_run(run) == ui::TaskState::Incomplete, "an interrupted run is incomplete");

    // Latest run per job: runs are stored newest first.
    std::vector<storage::ManagedJobRun> runs(4);
    runs[0].JobId = "job-b";
    runs[1].JobId = "";        // ad-hoc run, no job
    runs[2].JobId = "JOB-A";
    runs[3].JobId = "job-b";   // older run of job-b
    const auto latest = ui::latest_run_per_job(runs);
    check(latest.size() == 2 && latest[0].first == "job-b" && latest[0].second == 0 &&
              latest[1].second == 2,
          "the newest run wins per job and ad-hoc runs are skipped");

    // On-exit policy setting.
    check(ui::parse_task_exit_policy("Keep") == ui::TaskExitPolicy::Keep &&
              ui::parse_task_exit_policy("discard") == ui::TaskExitPolicy::Discard &&
              ui::parse_task_exit_policy("") == ui::TaskExitPolicy::Ask &&
              ui::parse_task_exit_policy("bogus") == ui::TaskExitPolicy::Ask,
          "on-exit policy parses with Ask as the fallback");
    check(std::string(ui::to_setting(ui::TaskExitPolicy::Discard)) == "discard",
          "on-exit policy round-trips to its setting value");

    models::CopyJobOptions copy;
    copy.SourceRoot = "C:\\a";
    copy.DestinationRoot = "E:\\Backup";
    check(ui::task_title(copy) == "Copy contents of C:\\a \xE2\x86\x92 E:\\Backup",
          "task title for a whole folder");
    models::CopyJobOptions scan;
    scan.OperationMode = models::JobOperationMode::ScanOnly;
    scan.UseExperimentalRawDiskScan = true;
    scan.SourceRoot = "D:\\";
    scan.SelectedRelativePaths = {"Media"};
    check(ui::task_title(scan) == "Assess Media (direct NTFS)", "task title for a direct-NTFS assessment");
    check(ui::summarize_options(models::CopyJobOptions{}).find("verify full") != std::string::npos,
          "option summary names the verification mode");

    // Source batches: one row per add action, in first-appearance order.
    std::vector<ui::SourceItem> batched(4);
    batched[0].path = L"C:\\a"; batched[0].batch = 7;
    batched[1].path = L"D:\\b"; batched[1].batch = 3;
    batched[2].path = L"C:\\c"; batched[2].batch = 7;
    batched[3].path = L"E:\\d"; batched[3].batch = 3;
    const auto batches = ui::source_batches(batched);
    check(batches.size() == 2 && batches[0].id == 7 && batches[0].items == std::vector<std::size_t>{0, 2} &&
              batches[1].items == std::vector<std::size_t>{1, 3},
          "items group into one row per add action");
    models::CopyJobOptions multi;
    multi.DestinationRoot = "F:\\Out";
    check(ui::task_title_for_parts(multi, 3, 2) == "Copy 3 items from 2 locations \xE2\x86\x92 F:\\Out",
          "a multi-part task is titled by items and locations");
    multi.SourceRoot = "C:\\x";
    check(ui::task_title_for_parts(multi, 1, 1) == ui::task_title(multi),
          "a one-part task keeps the plain title");

    // The legacy session-task file is still readable for the one-time migration.
    const std::string legacy =
        R"({"SchemaVersion":1,"ListRunning":true,"Tasks":[{"DisplayName":"Old","Options":{"SourceRoot":"C:\\b"}},{"Options":{}}]})";
    const auto migrated = ui::pending_tasks_from_json(legacy);
    check(migrated.size() == 1 && migrated[0].options.SourceRoot == "C:\\b" &&
              migrated[0].display_name == "Old",
          "legacy session tasks migrate, skipping entries without a source");
    check(ui::pending_tasks_from_json("{not json").empty(),
          "a corrupt legacy file migrates nothing");
    const std::wstring file = make_work_dir() + L"\\session-tasks.json";
    check(storage::fsutil::write_atomic_bytes(file, legacy) && ui::load_pending_tasks(file).size() == 1,
          "the legacy file loads from disk");

    // Journal folder: serialized only when set; the path follows it.
    models::CopyJobOptions journal_options;
    json::Writer plain(false);
    journal_options.to_json(plain);
    check(plain.take().find("JournalDirectory") == std::string::npos,
          "default runs keep the journal folder out of the wire format");
    journal_options.JournalDirectory = "D:\\Journals";
    json::Writer custom(false);
    journal_options.to_json(custom);
    check(models::CopyJobOptions::from_json(json::parse(custom.take())).JournalDirectory == "D:\\Journals",
          "a custom journal folder round-trips");
    const std::wstring journals = make_work_dir() + L"\\journals here";
    check(storage::JobJournalStore::get_default_journal_path("abc", journals + L"\\") ==
              journals + L"\\job-abc.json",
          "the journal path follows the configured folder");
    check(storage::JobJournalStore::get_default_journal_path("abc").find(L"\\XactCopy\\journals\\job-abc.json") !=
              std::wstring::npos,
          "an empty folder means the default location");
}

void test_elevation_handoff() {
    std::printf("--- elevation: direct-NTFS handoff validation ---\n");
    namespace el = ui::elevation;
    const std::string sid = el::current_user_sid();
    check(!sid.empty(), "current user SID resolves");

    el::HandoffRequest request;
    request.Options.OperationMode = models::JobOperationMode::ScanOnly;
    request.Options.UseExperimentalRawDiskScan = true;
    request.Options.SourceRoot = "C:\\Data";
    request.Options.SelectedRelativePaths = {"a.bin", "nested\\b.bin"};
    request.FromRecovery = true;
    request.ManagedRunId = "managed-1";
    request.DisplayName = "Nightly assessment";
    request.RecoveryRunId = "recovery-1";
    request.UserSid = sid;
    request.CreatedUnixMs = el::unix_now_ms();

    std::string token = el::write_handoff(request);
    check(el::is_handoff_token(token), "handoff token is a SHA-256 hex digest");
    check(el::handoff_pending(token), "written handoff is visible to the waiting window");
    el::HandoffRequest loaded;
    check(el::consume_handoff(token, loaded) == el::HandoffStatus::Ok, "valid handoff is accepted");
    check(!el::handoff_pending(token), "accepting a handoff removes it (the acceptance signal)");
    check(loaded.Options.SourceRoot == "C:\\Data" && loaded.Options.UseExperimentalRawDiskScan &&
              loaded.Options.SelectedRelativePaths.size() == 2 && loaded.FromRecovery &&
              loaded.ManagedRunId == "managed-1" && loaded.DisplayName == "Nightly assessment" &&
              loaded.RecoveryRunId == "recovery-1",
          "handoff round-trips the exact run request");
    check(el::consume_handoff(token, loaded) == el::HandoffStatus::Missing,
          "a consumed handoff cannot be replayed");

    // Content swapped after approval: the approved token no longer matches.
    token = el::write_handoff(request);
    el::HandoffRequest swapped = request;
    swapped.Options.SourceRoot = "D:\\Other";
    const std::string swapped_payload = swapped.to_json();
    storage::fsutil::write_file_raw(el::handoff_path(token),
                                    reinterpret_cast<const unsigned char*>(swapped_payload.data()),
                                    swapped_payload.size(), false, false);
    check(el::consume_handoff(token, loaded) == el::HandoffStatus::Tampered,
          "a request rewritten after approval is refused");

    token = el::write_handoff(request);
    check(el::consume_handoff(token, loaded, el::unix_now_ms(), "S-1-5-21-0-0-0-1001") ==
              el::HandoffStatus::WrongUser,
          "a handoff for another account is refused");

    el::HandoffRequest stale = request;
    stale.CreatedUnixMs = el::unix_now_ms() - el::HandoffMaximumAgeMs - 1000;
    token = el::write_handoff(stale);
    check(el::consume_handoff(token, loaded) == el::HandoffStatus::Expired,
          "an old handoff is refused");

    el::HandoffRequest copy_request = request;
    copy_request.Options.OperationMode = models::JobOperationMode::Copy;
    token = el::write_handoff(copy_request);
    check(el::consume_handoff(token, loaded) == el::HandoffStatus::Ok,
          "a copy with direct NTFS reads may be handed to the elevated window");
    copy_request.Options.UseExperimentalRawDiskScan = false;
    token = el::write_handoff(copy_request);
    check(el::consume_handoff(token, loaded) == el::HandoffStatus::NotRawScan,
          "elevation never starts a run that does not read the raw volume");

    check(el::consume_handoff("../../evil", loaded) == el::HandoffStatus::Invalid,
          "a malformed token cannot name another file");

    // Needs elevation only for a local NTFS source, and only when not elevated.
    models::CopyJobOptions raw;
    raw.OperationMode = models::JobOperationMode::ScanOnly;
    raw.UseExperimentalRawDiskScan = true;
    wchar_t windows_dir[MAX_PATH]{};
    GetWindowsDirectoryW(windows_dir, MAX_PATH);
    raw.SourceRoot = storage::fsutil::wide_to_utf8(windows_dir);
    std::string unsupported;
    check(el::raw_scan_requires_elevation(raw, false, unsupported),
          "a local NTFS raw assessment needs elevation when unelevated");
    check(!el::raw_scan_requires_elevation(raw, true, unsupported),
          "an elevated process needs no relaunch");
    raw.SourceRoot = "\\\\server\\share\\data";
    check(!el::raw_scan_requires_elevation(raw, false, unsupported) && !unsupported.empty(),
          "a UNC source never prompts for elevation and explains why");
    raw.SourceRoot = storage::fsutil::wide_to_utf8(windows_dir);
    raw.OperationMode = models::JobOperationMode::Copy;
    check(el::raw_scan_requires_elevation(raw, false, unsupported),
          "a copy with direct NTFS reads needs elevation too");
    raw.UseExperimentalRawDiskScan = false;
    check(!el::raw_scan_requires_elevation(raw, false, unsupported),
          "a copy with standard reads never requests elevation");
}

// Freezes every thread of a process, as a system sleep does, or thaws it.
void set_process_frozen(DWORD pid, bool frozen) {
    HANDLE snapshot = CreateToolhelp32Snapshot(TH32CS_SNAPTHREAD, 0);
    if (snapshot == INVALID_HANDLE_VALUE) return;
    THREADENTRY32 entry{};
    entry.dwSize = sizeof(entry);
    for (BOOL more = Thread32First(snapshot, &entry); more; more = Thread32Next(snapshot, &entry)) {
        if (entry.th32OwnerProcessID != pid) continue;
        HANDLE thread = OpenThread(THREAD_SUSPEND_RESUME, FALSE, entry.th32ThreadID);
        if (thread == nullptr) continue;
        if (frozen) SuspendThread(thread);
        else ResumeThread(thread);
        CloseHandle(thread);
    }
    CloseHandle(snapshot);
}

void test_sleep_resume_does_not_restart_worker() {
    std::printf("--- supervisor: system sleep/resume keeps the worker ---\n");
    std::wstring work = make_work_dir();
    std::wstring source = work + L"\\src";
    std::wstring destination = work + L"\\dst";
    storage::fsutil::create_directories(source);
    write_pattern_file(source + L"\\sleepy.bin", 12 * 1024 * 1024, 41);

    ui::WorkerSupervisor supervisor;
    EventSink sink;
    sink.attach(supervisor);

    models::CopyJobOptions options;
    options.SourceRoot = storage::fsutil::wide_to_utf8(source);
    options.DestinationRoot = storage::fsutil::wide_to_utf8(destination);
    options.TransferEnginePolicyValue = models::TransferEnginePolicy::ManagedRescue;
    options.MaxThroughputBytesPerSecond = 2 * 1024 * 1024;
    options.ResumeFromJournal = true;
    supervisor.start_job(options);

    check(sink.wait_for_progress(30), "progress observed before simulated sleep");
    std::string connected_state;
    DWORD pid = 0;
    if (sink.any_state_contains("Worker connected (PID", &connected_state)) {
        if (std::size_t open = connected_state.find("PID "); open != std::string::npos) {
            pid = static_cast<DWORD>(std::strtoul(connected_state.c_str() + open + 4, nullptr, 10));
        }
    }
    check(pid != 0, "worker PID parsed for simulated sleep");
    if (pid != 0) {
        // The UI receives PBT_APMSUSPEND, then the whole system stops. Freeze
        // the worker past the 10 s heartbeat timeout while the supervisor's
        // own watchdog thread keeps running, which is harsher than a real
        // sleep, where the supervisor stops too.
        supervisor.notify_system_suspending();
        set_process_frozen(pid, true);
        Sleep(12'000);
        set_process_frozen(pid, false);
        supervisor.notify_system_resumed();
    }

    check(sink.wait_for_result(120), "job completes after simulated sleep");
    {
        std::lock_guard<std::mutex> guard(sink.lock);
        if (!sink.results.empty()) {
            const auto& result = sink.results.front();
            check(result.Succeeded, "job succeeded across sleep (" + result.ErrorMessage + ")");
        }
    }
    check(!sink.any_log_contains("Restarting worker"),
          "no worker restart was triggered by the sleep gap");
    check(files_equal(source + L"\\sleepy.bin", destination + L"\\sleepy.bin"),
          "destination equals source after sleep/resume");
    supervisor.stop();
}

void test_stop_joins_in_flight_recovery() {
    std::printf("--- supervisor: shutdown joins in-flight recovery ---\n");
    std::wstring work = make_work_dir();
    std::wstring source = work + L"\\src";
    std::wstring destination = work + L"\\dst";
    storage::fsutil::create_directories(source);
    write_pattern_file(source + L"\\big.bin", 8 * 1024 * 1024, 44);

    ui::WorkerSupervisor supervisor;
    EventSink sink;
    sink.attach(supervisor);

    models::CopyJobOptions options;
    options.SourceRoot = storage::fsutil::wide_to_utf8(source);
    options.DestinationRoot = storage::fsutil::wide_to_utf8(destination);
    options.TransferEnginePolicyValue = models::TransferEnginePolicy::ManagedRescue;
    options.MaxThroughputBytesPerSecond = 1024 * 1024;
    options.ResumeFromJournal = true;
    supervisor.start_job(options);
    check(sink.wait_for_progress(30), "progress observed before shutdown-race kill");

    std::string connected_state;
    check(sink.any_state_contains("Worker connected (PID", &connected_state),
          "shutdown-race worker PID state available");
    DWORD pid = 0;
    if (std::size_t open = connected_state.find("PID "); open != std::string::npos) {
        pid = static_cast<DWORD>(std::strtoul(connected_state.c_str() + open + 4, nullptr, 10));
    }
    HANDLE process = pid == 0 ? nullptr : OpenProcess(PROCESS_TERMINATE, FALSE, pid);
    check(process != nullptr, "shutdown-race worker opened for kill");
    if (process != nullptr) {
        TerminateProcess(process, 9);
        CloseHandle(process);
    }

    check(sink.wait_for_log("Restarting worker and resuming from journal", 30),
          "automatic recovery entered before supervisor shutdown");
    auto started = std::chrono::steady_clock::now();
    supervisor.stop();
    double seconds = std::chrono::duration<double>(
                         std::chrono::steady_clock::now() - started)
                         .count();
    check(seconds < 15.0, "shutdown joins recovery without a deadlock or detached callback");
}

void test_recovery_service() {
    std::printf("--- recovery service: state round-trip ---\n");
    std::wstring work = make_work_dir();
    std::wstring state_path = work + L"\\recovery-state.json";
    std::wstring settings_path = work + L"\\settings.json";

    ui::AppSettings settings(settings_path);
    ui::LaunchOptions launch;

    {
        ui::RecoveryService service{ui::RecoveryStateStore(state_path), /*manage_run_once*/ false};
        ui::RecoveryStartupInfo info = service.initialize_session(settings, launch);
        check(!info.has_interrupted_run(), "fresh state has no interrupted run");

        models::CopyJobOptions options;
        options.SourceRoot = "D:\\Recover\\Src";
        options.DestinationRoot = "E:\\Recover\\Dst";
        service.mark_job_started("run-123", "Recover job", options, "C:\\journal.json", settings,
                                 "managed-run-456");
        // Simulate a crash: no mark_job_ended / mark_clean_shutdown.
    }

    {
        ui::RecoveryService service{ui::RecoveryStateStore(state_path), false};
        ui::RecoveryStartupInfo info = service.initialize_session(settings, launch);
        check(info.has_interrupted_run(), "interrupted run detected after unclean exit");
        check(info.should_prompt, "prompt requested (default settings)");
        if (info.has_interrupted_run()) {
            check(info.interrupted_run->RunId == "run-123", "run id round-trips");
            check(info.interrupted_run->Options.SourceRoot == "D:\\Recover\\Src",
                  "options round-trip through recovery state");
            check(info.interrupted_run->ManagedRunId == "managed-run-456",
                  "managed history run id round-trips for restart resume");
        }
        service.mark_job_ended("run-123");
        service.mark_clean_shutdown();
    }

    {
        ui::RecoveryService service{ui::RecoveryStateStore(state_path), false};
        ui::RecoveryStartupInfo info = service.initialize_session(settings, launch);
        check(!info.has_interrupted_run(), "clean shutdown leaves no interrupted run");
    }

    settings.set_bool("AutoResumeAfterCrash", true);
    settings.save();
    {
        ui::RecoveryService service{ui::RecoveryStateStore(state_path), false};
        models::CopyJobOptions options;
        options.SourceRoot = "D:\\Recover\\Src";
        options.DestinationRoot = "E:\\Recover\\Dst";
        service.mark_job_started("run-unbound", "Unbound recovery", options,
                                 "C:\\journal.json", settings);
    }
    {
        ui::RecoveryService service{ui::RecoveryStateStore(state_path), false};
        ui::RecoveryStartupInfo info = service.initialize_session(settings, launch);
        check(info.has_interrupted_run() && !info.should_auto_resume && info.should_prompt,
              "auto-resume is blocked when persisted media identity is missing");
        service.mark_job_ended("run-unbound");
        service.mark_clean_shutdown();
    }
    {
        ui::RecoveryService service{ui::RecoveryStateStore(state_path), false};
        models::CopyJobOptions options;
        options.SourceRoot = "D:\\Recover\\Src";
        options.DestinationRoot = "E:\\Recover\\Dst";
        service.mark_job_started("run-bound", "Bound recovery", options,
                                 "C:\\journal.json", settings);
        service.update_media_identities("run-bound", "vol:SOURCE:00000001",
                                        "vol:DESTINATION:00000002");
    }
    {
        ui::RecoveryService service{ui::RecoveryStateStore(state_path), false};
        ui::RecoveryStartupInfo info = service.initialize_session(settings, launch);
        check(info.has_interrupted_run() && info.should_auto_resume &&
                  info.interrupted_run->Options.ExpectedSourceIdentity ==
                      "vol:SOURCE:00000001" &&
                  info.interrupted_run->Options.ExpectedDestinationIdentity ==
                      "vol:DESTINATION:00000002",
              "worker-confirmed media identities make auto-resume eligible");
        service.mark_job_ended("run-bound");
        service.mark_clean_shutdown();
    }
    {
        ui::RecoveryService service{ui::RecoveryStateStore(state_path), false};
        models::CopyJobOptions options;
        options.SourceRoot = "D:\\Recover\\Src";
        options.DestinationRoot = "E:\\Recover\\Dst";
        options.ContinueOnFileError = true;
        service.mark_job_started("run-risky", "Risky recovery", options,
                                 "C:\\journal.json", settings);
        service.update_media_identities("run-risky", "vol:SOURCE:00000001",
                                        "vol:DESTINATION:00000002");
    }
    {
        ui::RecoveryService service{ui::RecoveryStateStore(state_path), false};
        ui::RecoveryStartupInfo info = service.initialize_session(settings, launch);
        check(info.has_interrupted_run() && !info.should_auto_resume && info.should_prompt,
              "attended-only integrity settings block automatic crash recovery");
        service.mark_job_ended("run-risky");
        service.mark_clean_shutdown();
    }

    const std::string invalid_state = "{ invalid recovery state";
    check(storage::fsutil::write_atomic_bytes(state_path, invalid_state),
          "recovery corruption fixture is written");
    {
        ui::RecoveryService service{ui::RecoveryStateStore(state_path), false};
        std::vector<std::string> warnings;
        service.on_persistence_warning =
            [&warnings](const std::string& warning) { warnings.push_back(warning); };
        ui::RecoveryStartupInfo info = service.initialize_session(settings, launch);
        WIN32_FIND_DATAW preserved{};
        HANDLE search = FindFirstFileW((state_path + L".corrupt-*").c_str(), &preserved);
        const bool corrupt_preserved = search != INVALID_HANDLE_VALUE;
        if (search != INVALID_HANDLE_VALUE) FindClose(search);
        check(!info.has_interrupted_run() && !warnings.empty() && corrupt_preserved,
              "invalid recovery state is preserved and reported instead of silently discarded");
        service.mark_clean_shutdown();
    }
}

void test_recovery_writer_coalesces() {
    std::printf("--- recovery writer: coalescing and shutdown ---\n");
    std::mutex mutex;
    std::condition_variable signal;
    bool entered = false;
    bool release = false;
    std::vector<ui::RecoveryState> saved;
    ui::AppSettings settings(make_work_dir() + L"\\settings.json");
    {
        ui::RecoveryService service(ui::RecoveryStateStore(make_work_dir() + L"\\state.json"), false,
            [&](const ui::RecoveryState& state) {
                std::unique_lock<std::mutex> guard(mutex);
                saved.push_back(state);
                if (saved.size() == 1) {
                    entered = true;
                    signal.notify_all();
                    signal.wait(guard, [&] { return release; });
                }
            });
        models::CopyJobOptions options;
        options.SelectedRelativePaths = {"first.bin", "folder\\second.bin"};
        service.mark_job_started("coalesce", "Test", options, "journal", settings);
        {
            std::unique_lock<std::mutex> guard(mutex);
            check(signal.wait_for(guard, std::chrono::seconds(5), [&] { return entered; }),
                  "background writer begins without waiting on the UI");
        }
        for (int i = 0; i < 100; ++i)
            service.update_media_identities("coalesce", "source-" + std::to_string(i), "dest");
        {
            std::lock_guard<std::mutex> guard(mutex);
            release = true;
        }
        signal.notify_all();
    }
    check(saved.size() == 2 && saved.back().ActiveRun.has_value() &&
              saved.back().ActiveRun->Options.ExpectedSourceIdentity == "source-99" &&
              saved.back().ActiveRun->Options.SelectedRelativePaths ==
                  std::vector<std::string>({"first.bin", "folder\\second.bin"}),
          "blocked writes coalesce to the latest state and destructor joins the writer");

    const auto path = make_work_dir() + L"\\recovery.json";
    {
        ui::RecoveryService service(ui::RecoveryStateStore(path), false);
        models::CopyJobOptions options;
        service.mark_job_started("final", "Test", options, "journal", settings);
        service.update_media_identities("final", "latest", "dest");
        service.mark_clean_shutdown();
    }
    const auto final_state = ui::RecoveryStateStore(path).load();
    check(final_state.CleanShutdown && final_state.ActiveRun.has_value() &&
              final_state.ActiveRun->Options.ExpectedSourceIdentity == "latest",
          "clean shutdown persists the final state after pending writes");
}

void test_log_model() {
    std::printf("--- log model: ring, filtering and search ---\n");
    ui::LogModel log(3);
    const auto timestamp = std::chrono::system_clock::time_point(std::chrono::seconds(42));
    log.append(L"first", ui::LogSeverity::Normal, timestamp);
    log.append(L"WARNING: second", ui::classify_log_line("WARNING: second"), timestamp);
    log.append(L"ERROR: third", ui::classify_log_line("ERROR: third"), timestamp);
    check(log.size() == 3 && log.size(ui::LogFilter::WarningsAndErrors) == 2 &&
              log.size(ui::LogFilter::ErrorsOnly) == 1 && log.at(ui::LogFilter::All, 0)->timestamp == timestamp,
          "model retains timestamp and maintains each severity view");
    log.append(L"fourth", ui::LogSeverity::Normal, timestamp);
    check(log.size() == 3 && log.at(ui::LogFilter::All, 0)->text == L"WARNING: second" &&
              log.size(ui::LogFilter::ErrorsOnly) == 1,
          "oldest entry is evicted without losing retained filtered entries");
    check(log.find_next(ui::LogFilter::WarningsAndErrors, L"error", 1) == 1 &&
              log.find_next(ui::LogFilter::All, L"warning", 2) == 0,
          "case-insensitive search wraps in the selected view");
    log.set_capacity(1);
    check(log.size() == 1 && log.size(ui::LogFilter::WarningsAndErrors) == 0 &&
              log.at(ui::LogFilter::All, 0)->text == L"fourth",
          "capacity changes evict exactly the excess lines");
    log.clear();
    check(log.size() == 0, "clear empties every view");
    ui::LogModel concurrent(2);
    concurrent.append(L"old one", ui::LogSeverity::Normal, timestamp);
    concurrent.append(L"old two", ui::LogSeverity::Warning, timestamp);
    auto snapshot = concurrent.begin_snapshot();
    concurrent.append(L"new three", ui::LogSeverity::Error, timestamp);
    concurrent.clear();
    concurrent.append(L"new job", ui::LogSeverity::Normal, timestamp);
    const auto first = concurrent.next_snapshot(snapshot);
    const auto second = concurrent.next_snapshot(snapshot);
    check(first && second && first->text == L"old one" && second->text == L"old two" &&
              !concurrent.next_snapshot(snapshot),
          "background snapshot keeps its original lines across eviction and reset");
    concurrent.end_snapshot(snapshot);
    check(concurrent.size() == 1 && concurrent.at(ui::LogFilter::All, 0)->text == L"new job",
          "snapshot does not disturb the live ring after a reset");
    ui::LogModel large(20001);
    for (int i = 0; i < 20002; ++i)
        large.append(std::to_wstring(i), ui::LogSeverity::Normal, timestamp);
    check(large.size() == 20001 && large.at(ui::LogFilter::All, 0)->text == L"1",
          "configured capacity above 20000 is honored exactly");
}

void test_journal_report_and_range_map() {
    std::printf("--- journal report and range-map geometry ---\n");
    storage::JobJournal journal;
    journal.RunOptions = models::CopyJobOptions{};
    storage::JournalFileEntry entry;
    entry.RelativePath = "folder/a,\"b\".txt";
    entry.SourceLength = 1000;
    entry.BytesCopied = 900;
    entry.LastError = "line one\r\nline two";
    entry.RescueRanges = {{0, 400, storage::RescueRangeState::Good},
                          {400, 100, storage::RescueRangeState::Recovered},
                          {500, 50, storage::RescueRangeState::Bad},
                          {550, 50, storage::RescueRangeState::KnownBad}};
    entry.RecoveredRanges = {{420, 20}};
    journal.Files.set(entry.RelativePath, entry);
    std::ostringstream output;
    ui::write_journal_report_csv(output, journal);
    const auto csv = output.str();
    check(csv.rfind("\xEF\xBB\xBF", 0) == 0 &&
              csv.find("\"folder/a,\"\"b\"\".txt\"") != std::string::npos &&
              csv.find("\"line one\r\nline two\"") != std::string::npos,
          "CSV has UTF-8 BOM and RFC 4180 quoting for paths and multiline errors");
    check(csv.find(",120,100,500:100,400:100,420:20,") != std::string::npos,
          "report merges bad ranges and separately identifies synthetic fill");
    journal.RunOptions->OperationMode = models::JobOperationMode::ScanOnly;
    std::ostringstream scan_output;
    ui::write_journal_report_csv(scan_output, journal);
    check(scan_output.str().find(",120,100,420:20; 500:100,400:100,,") != std::string::npos,
          "scan journals treat legacy recovered-range entries as unreadable ranges");
    journal.RunOptions.reset();
    std::ostringstream legacy_output;
    ui::write_journal_report_csv(legacy_output, journal);
    check(legacy_output.str().find(",100,100,500:100,400:100,,420:20\r\n") != std::string::npos,
          "ambiguous legacy ranges are reported separately instead of guessed");
    journal.RunOptions = models::CopyJobOptions{};
    entry.RescueRanges.clear();
    journal.Files.set(entry.RelativePath, entry);
    std::ostringstream legacy_copy_output;
    ui::write_journal_report_csv(legacy_copy_output, journal);
    check(legacy_copy_output.str().find(",20,20,,420:20,420:20,") != std::string::npos,
          "legacy copy fill contributes to recovered and unreadable bytes");
    const auto spans = ui::range_map_spans(1000000,
        {{0, 1000000, storage::RescueRangeState::Good},
         {500000, 1, storage::RescueRangeState::Bad},
         {700000, 1, storage::RescueRangeState::KnownBad}}, 100);
    bool bad_visible = false, known_visible = false;
    for (const auto& span : spans) {
        if (span.state == storage::RescueRangeState::Bad && span.right > span.left &&
            span.offset == 500000 && span.length == 1) bad_visible = true;
        if (span.state == storage::RescueRangeState::KnownBad && span.right > span.left)
            known_visible = true;
    }
    check(bad_visible && known_visible, "sub-pixel bad and known-bad ranges stay visible");
    check(ui::range_map_spans(0, entry.RescueRanges, 100).empty() &&
              ui::range_map_spans(1000, entry.RescueRanges, 0).empty(),
          "empty files and empty viewports have no spans");
    const auto tiny = ui::range_map_spans(1, {}, 10);
    check(tiny.size() == 1 && tiny[0].left == 0 && tiny[0].right == 10 &&
              tiny[0].offset == 0 && tiny[0].length == 1,
          "pending pixels in tiny files retain a nonempty byte interval");
    const auto clipped = ui::range_map_spans(100, {{95, INT64_MAX,
        storage::RescueRangeState::Bad}}, 10);
    check(!clipped.empty() && clipped.back().state == storage::RescueRangeState::Bad &&
              clipped.back().offset == 95 && clipped.back().length == 5,
          "range geometry clips overflowing lengths at the file boundary");
}

void test_settings_preservation() {
    std::printf("--- settings: DOM preservation ---\n");
    std::wstring work = make_work_dir();
    std::wstring path = work + L"\\settings.json";

    // Simulate a .NET-authored settings file with keys the native UI ignores.
    std::string seeded =
        "{\r\n  \"Theme\": \"dark\",\r\n  \"AccentColorHex\": \"#5A78C8\",\r\n"
        "  \"DefaultBufferSizeMb\": 8,\r\n  \"SomeFutureSetting\": [1, 2, 3],\r\n"
        "  \"DefaultVerificationMode\": \"sampled\"\r\n}";
    storage::fsutil::write_file_raw(path, reinterpret_cast<const unsigned char*>(seeded.data()),
                                    seeded.size(), false, false);

    ui::AppSettings settings(path);
    check(settings.get_bool("ProgressBarShowPercentage", true),
          "new settings default to progress percentages enabled");
    settings.set_bool("ProgressBarShowPercentage", false);
    check(!settings.get_bool("ProgressBarShowPercentage", true),
          "explicitly disabled progress percentages override the new default");
    check(settings.theme() == "dark", "theme read");
    check(settings.get_int("DefaultBufferSizeMb", 1, 256, 4) == 8, "int read");
    auto options = settings.build_default_options();
    check(options.BufferSizeBytes == 8 * 1024 * 1024, "options buffer from settings");
    check(options.VerificationModeValue == models::VerificationMode::Sampled,
          "options verification kebab parse");

    settings.set_string("Theme", "light");
    settings.set_int("DefaultBufferSizeMb", 16);
    settings.save();

    ui::AppSettings reloaded(path);
    check(!reloaded.get_bool("ProgressBarShowPercentage", true),
          "saved disabled progress percentages remain disabled after reload");
    check(reloaded.theme() == "light", "edited theme persisted");
    check(reloaded.get_int("DefaultBufferSizeMb", 1, 256, 4) == 16, "edited int persisted");
    check(reloaded.get_string("AccentColorHex", "") == "#5A78C8", "untouched key preserved");
    const json::Value* future = nullptr;
    {
        auto bytes = storage::fsutil::read_all_bytes(path);
        check(bytes.has_value(), "settings file readable");
        if (bytes.has_value()) {
            std::string text(reinterpret_cast<const char*>(bytes->data()), bytes->size());
            check(text.find("SomeFutureSetting") != std::string::npos,
                  "unknown keys survive native save");
        }
    }
    (void)future;
}

void test_settings_failure_reporting() {
    std::printf("--- settings: corruption and save failures ---\n");
    std::wstring work = make_work_dir();
    std::wstring corrupt_path = work + L"\\corrupt-settings.json";
    const std::string invalid = "[]";
    check(storage::fsutil::write_atomic_bytes(corrupt_path, invalid),
          "non-object settings fixture is written");

    ui::AppSettings corrupt(corrupt_path);
    WIN32_FIND_DATAW preserved{};
    HANDLE search = FindFirstFileW((corrupt_path + L".corrupt-*").c_str(), &preserved);
    const bool found_preserved_copy = search != INVALID_HANDLE_VALUE;
    if (search != INVALID_HANDLE_VALUE) FindClose(search);
    check(!corrupt.load_warning().empty() && found_preserved_copy,
          "non-object settings are reported and preserved before defaults are used");

    const std::wstring parent_file = work + L"\\not-a-directory";
    const unsigned char marker = 0x5A;
    check(storage::fsutil::write_file_raw(parent_file, &marker, 1, false, false),
          "save failure fixture is written");
    ui::AppSettings unwritable(parent_file + L"\\settings.json");
    const bool transaction_saved = unwritable.update_and_save([](ui::AppSettings& target) {
        target.set_string("Theme", "light");
    });
    check(!transaction_saved && !unwritable.last_save_error().empty() &&
              unwritable.theme() == "dark",
          "atomic settings save failure is diagnosed and rolls back in-memory edits");

    models::CopyJobOptions failed_defaults;
    failed_defaults.BufferSizeBytes = 16 * 1024 * 1024;
    check(!unwritable.save_run_defaults(failed_defaults) &&
              unwritable.get_int("DefaultBufferSizeMb", 1, 256, 4) == 4,
          "failed Save Defaults cannot leak into a later unrelated settings save");
}

void test_settings_option_mapping() {
    std::printf("--- settings: complete option mapping ---\n");
    std::wstring work = make_work_dir();
    ui::AppSettings settings(work + L"\\settings.json");

    settings.set_string("DefaultOverwritePolicy", "ask");
    settings.set_string("DefaultTransferEnginePolicy", "native");
    settings.set_string("DefaultSymlinkHandling", "follow");
    settings.set_string("DefaultVerificationMode", "sampled");
    settings.set_string("DefaultVerificationHashAlgorithm", "sha512");
    settings.set_string("DefaultSalvageFillPattern", "random");
    settings.set_string("DefaultSourceMutationPolicy", "wait-for-reappearance");
    settings.set_string("DefaultScanPerformanceProfile", "precise");
    settings.set_string("WorkerTelemetryProfile", "debug");

    settings.set_bool("DefaultResumeFromJournal", false);
    settings.set_bool("DefaultSalvageUnreadableBlocks", false);
    settings.set_bool("DefaultContinueOnFileError", false);
    settings.set_bool("DefaultVerifyAfterCopy", true);
    settings.set_bool("DefaultUseBadRangeMap", false);
    settings.set_bool("DefaultSkipKnownBadRanges", false);
    settings.set_bool("DefaultUpdateBadRangeMapFromRun", false);
    settings.set_bool("DefaultUseExperimentalRawDiskScan", true);
    settings.set_bool("DefaultUseAdaptiveBuffer", true);
    settings.set_bool("DefaultWaitForMediaAvailability", true);
    settings.set_bool("DefaultWaitForFileLockRelease", true);
    settings.set_bool("DefaultTreatAccessDeniedAsContention", true);
    settings.set_bool("DefaultFragileMediaMode", true);
    settings.set_bool("DefaultSkipFileOnFirstReadError", false);
    settings.set_bool("DefaultPersistFragileSkipsAcrossResume", false);
    settings.set_bool("DefaultPreserveTimestamps", false);
    settings.set_bool("DefaultCopyEmptyDirectories", false);
    settings.set_bool("DefaultAllowRecoveredOverwriteExisting", true);
    settings.set_bool("DefaultAllowDecryptedDestination", true);

    settings.set_int("DefaultBadRangeMapMaxAgeDays", 7);
    settings.set_int("DefaultLockContentionProbeIntervalMs", 750);
    settings.set_int("DefaultFragileFailureWindowSeconds", 31);
    settings.set_int("DefaultFragileFailureThreshold", 5);
    settings.set_int("DefaultFragileCooldownSeconds", 9);
    settings.set_int("DefaultBufferSizeMb", 16);
    settings.set_int("DefaultMaxRetries", 23);
    settings.set_int("DefaultOperationTimeoutSeconds", 17);
    settings.set_int("DefaultPerFileTimeoutSeconds", 29);
    settings.set_int("DefaultMaxThroughputMbPerSecond", 123);
    settings.set_int("DefaultParallelSmallFileWorkers", 7);
    settings.set_int("DefaultParallelScanWorkers", 6);
    settings.set_int("DefaultSmallFileThresholdKb", 512);
    settings.set_string("DefaultWorkerProcessPriorityClass", "BelowNormal");
    settings.set_int("WorkerProgressIntervalMs", 125);
    settings.set_int("WorkerMaxLogsPerSecond", 321);
    settings.set_int("DefaultRescueFastScanChunkKb", 64);
    settings.set_int("DefaultRescueTrimChunkKb", 65);
    settings.set_int("DefaultRescueScrapeChunkKb", 66);
    settings.set_int("DefaultRescueRetryChunkKb", 67);
    settings.set_int("DefaultRescueSplitMinimumKb", 68);
    settings.set_int("DefaultRescueFastScanRetries", 8);
    settings.set_int("DefaultRescueTrimRetries", 9);
    settings.set_int("DefaultRescueScrapeRetries", 10);
    settings.set_int("DefaultSampleVerificationChunkKb", 256);
    settings.set_int("DefaultSampleVerificationChunkCount", 11);

    const models::CopyJobOptions options = settings.build_default_options();
    check(options.OverwritePolicyValue == models::OverwritePolicy::Ask,
          "overwrite policy maps");
    check(options.TransferEnginePolicyValue == models::TransferEnginePolicy::NativeFast,
          "transfer engine maps");
    check(options.SymlinkHandling == models::SymlinkHandlingMode::Follow,
          "symlink policy maps");
    check(options.VerificationModeValue == models::VerificationMode::Sampled &&
              options.VerifyAfterCopy,
          "verification mode and enable flag map");
    check(ui::verification_combo_index(options) == 1,
          "sampled verification selects the sampled UI option");
    models::CopyJobOptions full_verification = options;
    full_verification.VerificationModeValue = models::VerificationMode::Full;
    check(ui::verification_combo_index(full_verification) == 2,
          "full verification selects the full UI option");
    full_verification.VerifyAfterCopy = false;
    check(ui::verification_combo_index(full_verification) == 0,
          "disabled verification selects the none UI option");
    check(options.VerificationHashAlgorithmValue == models::VerificationHashAlgorithm::Sha512,
          "verification hash maps");
    check(options.SalvageFillPatternValue == models::SalvageFillPattern::Random,
          "salvage fill maps");
    check(options.SourceMutationPolicyValue == models::SourceMutationPolicy::WaitForReappearance,
          "source mutation policy maps");
    check(options.ScanPerformanceProfileValue == models::ScanPerformanceProfile::Precise,
          "scan profile maps");
    check(options.WorkerTelemetryProfileValue == models::WorkerTelemetryProfile::Debug,
          "telemetry profile maps");

    check(!options.ResumeFromJournal && !options.SalvageUnreadableBlocks &&
              !options.ContinueOnFileError && !options.UseBadRangeMap &&
              !options.SkipKnownBadRanges && !options.UpdateBadRangeMapFromRun,
          "copy/map boolean defaults map");
    check(options.UseExperimentalRawDiskScan && options.UseAdaptiveBufferSizing &&
              options.WaitForMediaAvailability && options.WaitForFileLockRelease &&
              options.TreatAccessDeniedAsContention,
          "contention and raw-scan booleans map");
    check(options.FragileMediaMode && !options.SkipFileOnFirstReadError &&
              !options.PersistFragileSkipAcrossResume && !options.PreserveTimestamps &&
              !options.CopyEmptyDirectories && options.AllowRecoveredOverwriteExisting &&
              options.AllowDecryptedDestination,
          "fragile and copy behavior booleans map");

    check(options.BadRangeMapMaxAgeDays == 7 &&
              options.LockContentionProbeInterval == time::TimeSpan::from_milliseconds(750) &&
              options.FragileFailureWindowSeconds == 31 &&
              options.FragileFailureThreshold == 5 && options.FragileCooldownSeconds == 9,
          "map, contention, and fragile numeric settings map");
    check(options.BufferSizeBytes == 16 * 1024 * 1024 && options.MaxRetries == 23 &&
              options.OperationTimeout == time::TimeSpan::from_seconds(17) &&
              options.PerFileTimeout == time::TimeSpan::from_seconds(29),
          "buffer, retry, and timeout settings map");
    check(options.MaxThroughputBytesPerSecond == 123LL * 1024 * 1024 &&
              options.ParallelSmallFileWorkers == 7 && options.ParallelScanWorkers == 6 &&
              options.SmallFileThresholdBytes == 512 * 1024 &&
              options.WorkerProcessPriorityClass == "BelowNormal",
          "performance settings map");
    check(options.WorkerProgressEmitIntervalMs == 125 && options.WorkerMaxLogsPerSecond == 321,
          "worker telemetry numeric settings map");
    check(options.RescueFastScanChunkBytes == 64 * 1024 &&
              options.RescueTrimChunkBytes == 65 * 1024 &&
              options.RescueScrapeChunkBytes == 66 * 1024 &&
              options.RescueRetryChunkBytes == 67 * 1024 &&
              options.RescueSplitMinimumBytes == 68 * 1024 &&
              options.RescueFastScanRetries == 8 && options.RescueTrimRetries == 9 &&
              options.RescueScrapeRetries == 10,
          "rescue tuning settings map");
    check(options.SampleVerificationChunkBytes == 256 * 1024 &&
              options.SampleVerificationChunkCount == 11,
          "sample verification settings map");

    ui::AppSettings auto_settings(work + L"\\auto-settings.json");
    const models::CopyJobOptions safe_defaults = auto_settings.build_default_options();
    check(!safe_defaults.SalvageUnreadableBlocks && !safe_defaults.ContinueOnFileError &&
              safe_defaults.VerifyAfterCopy && !safe_defaults.UseBadRangeMap &&
              !safe_defaults.SkipKnownBadRanges && !safe_defaults.UpdateBadRangeMapFromRun &&
              !safe_defaults.AllowRecoveredOverwriteExisting &&
              !safe_defaults.AllowDecryptedDestination && safe_defaults.MaxRetries == 2,
          "missing settings use the integrity-first safety profile");
    check(safe_defaults.ParallelSmallFileWorkers == 0,
          "zero small-file workers remains the documented auto value");
}

void test_job_catalog_store() {
    std::printf("--- job catalog: store round-trip + normalization ---\n");
    std::wstring work = make_work_dir();
    std::wstring path = work + L"\\catalog.json";

    storage::JobCatalogStore store(path);
    check(store.load().Jobs.empty(), "missing file loads default catalog");

    storage::JobCatalog catalog;
    storage::ManagedJob job;
    job.JobId = "1234567890abcdef1234567890abcdef";
    job.Name = "Nightly Backup";
    job.Options.SourceRoot = "C:\\Data";
    job.Options.DestinationRoot = "D:\\Backup";
    job.Options.MaxRetries = 7;
    catalog.Jobs.push_back(job);

    storage::ManagedJobQueueEntry entry;
    entry.QueueEntryId = "feedfacefeedfacefeedfacefeedface";
    entry.JobId = job.JobId;
    entry.Trigger = "manual";
    entry.EnqueuedBy = "tests";
    catalog.QueueEntries.push_back(entry);

    storage::ManagedJobRun run;
    run.RunId = "0011223344556677889900aabbccddee";
    run.JobId = job.JobId;
    run.DisplayName = "Nightly Backup";
    run.SourceRoot = job.Options.SourceRoot;
    run.DestinationRoot = job.Options.DestinationRoot;
    run.Status = storage::ManagedJobRunStatus::Completed;
    run.FinishedUtc = time::DateTimeOffset::now_utc();
    models::CopyJobResult result;
    result.Succeeded = true;
    result.TotalFiles = 12;
    result.CompletedFiles = 12;
    result.JournalPath = "C:\\journals\\x.journal.json";
    run.Result = result;
    catalog.Runs.push_back(run);

    check(store.save(catalog), "save catalog");

    storage::JobCatalog loaded = store.load();
    check(loaded.SchemaVersion == 4, "schema version 4");
    check(loaded.Jobs.size() == 1 && loaded.Jobs[0].Name == "Nightly Backup", "job round-trip");
    check(loaded.Jobs[0].Pinned, "a saved job is pinned by default");
    check(loaded.Jobs[0].Options.MaxRetries == 7, "job options round-trip");
    check(loaded.QueueEntries.size() == 1 && !loaded.QueueEntries[0].LastAttemptUtc.has_value(),
          "queue entry nullable LastAttemptUtc stays null");
    check(loaded.Runs.size() == 1 && loaded.Runs[0].Status == storage::ManagedJobRunStatus::Completed,
          "run status enum string round-trip");
    check(loaded.Runs[0].Result.has_value() && loaded.Runs[0].Result->CompletedFiles == 12,
          "run result round-trip");
    check(loaded.Runs[0].FinishedUtc.has_value(), "run FinishedUtc round-trip");

    {
        auto bytes = storage::fsutil::read_all_bytes(path);
        check(bytes.has_value(), "catalog file readable");
        if (bytes.has_value()) {
            std::string text(reinterpret_cast<const char*>(bytes->data()), bytes->size());
            check(text.find("\"Status\": \"Completed\"") != std::string::npos,
                  "status serialized as member-name string");
            check(text.find("\"LastAttemptUtc\": null") != std::string::npos,
                  "null nullable serialized as null literal");
        }
    }

    // Legacy schema-v1 queue migration + duplicate queue-entry removal.
    storage::JobCatalog legacy;
    legacy.SchemaVersion = 1;
    legacy.Jobs.push_back(job);
    legacy.QueuedJobIds.push_back(" " + job.JobId + " ");
    storage::JobCatalogStore::normalize(legacy);
    check(legacy.SchemaVersion == 4, "legacy schema upgraded");
    {
        // Schema <= 3 jobs carry no Pinned field: they were saved on purpose.
        const json::Value parsed = json::parse(R"({"JobId":"a","Name":"Old","Options":{}})");
        check(storage::ManagedJob::from_json(parsed).Pinned, "a pre-v4 job reads as pinned");
        storage::ManagedJob task;
        task.JobId = "b";
        task.Name = "Task";
        task.Pinned = false;
        json::Writer writer(false);
        task.to_json(writer);
        check(!storage::ManagedJob::from_json(json::parse(writer.take())).Pinned,
              "an unpinned task round-trips");
    }
    check(legacy.QueueEntries.size() == 1 && legacy.QueueEntries[0].JobId == job.JobId &&
              legacy.QueueEntries[0].Trigger == "legacy-migration" &&
              legacy.QueueEntries[0].EnqueuedBy == "migration",
          "legacy queue ids migrated to entries");

    storage::JobCatalog dupes;
    dupes.Jobs.push_back(job);
    dupes.QueueEntries.push_back(entry);
    dupes.QueueEntries.push_back(entry);
    storage::JobCatalogStore::normalize(dupes);
    check(dupes.QueueEntries.size() == 1, "duplicate queue entries removed");

    // A corrupted primary recovers from the independently committed mirror.
    std::string garbage = "{ not json";
    storage::fsutil::write_file_raw(path, reinterpret_cast<const unsigned char*>(garbage.data()),
                                    garbage.size(), false, false);
    check(store.load().Jobs.size() == 1 && store.load().Jobs[0].Name == "Nightly Backup",
          "corrupted catalog primary loads the trusted mirror");
    storage::fsutil::delete_file(store.mirror_path());
    check(store.load().Jobs.empty(),
          "catalog defaults safely when no authenticated snapshot remains");
}

void test_job_manager_service() {
    std::printf("--- job manager: service behaviors ---\n");
    std::wstring work = make_work_dir();
    std::wstring path = work + L"\\catalog.json";

    ui::JobManagerService manager(path);

    models::CopyJobOptions options;
    options.SourceRoot = "C:\\Src";
    options.DestinationRoot = "D:\\Dst";
    options.ExpectedSourceIdentity = "identity-src";
    options.ExpectedDestinationIdentity = "identity-dst";
    options.ResumeJournalPathHint = "C:\\journals\\hint.json";
    options.AllowJournalRootRemap = true;

    auto saved = manager.save_job("  Weekly Sync  ", options);
    check(saved.has_value() && saved->Name == "Weekly Sync", "save_job trims name");
    check(saved.has_value() &&
              saved->Options.ExpectedSourceIdentity == "identity-src" &&
              saved->Options.ExpectedDestinationIdentity == "identity-dst" &&
              saved->Options.ResumeJournalPathHint.empty() &&
              !saved->Options.AllowJournalRootRemap,
          "save_job retains unattended media binding and scrubs transient journal fields");

    check(!manager.save_job("   ", options).has_value(), "blank name rejected");

    auto duplicate = manager.duplicate_job(saved->JobId, "Weekly Sync Copy");
    check(duplicate.has_value() && duplicate->JobId != saved->JobId, "duplicate creates new id");

    auto jobs = manager.get_jobs();
    check(jobs.size() == 2 && jobs[0].Name == "Weekly Sync" && jobs[1].Name == "Weekly Sync Copy",
          "get_jobs sorted by name");

    check(manager.queue_job(saved->JobId), "queue job");
    check(!manager.queue_job(saved->JobId), "duplicate queue rejected without allow_duplicate");
    check(manager.queue_job(saved->JobId, /*allow_duplicate*/ true, "job-manager"), "duplicate allowed");
    check(manager.queue_job(duplicate->JobId), "queue second job");

    auto queue = manager.get_queue_entries();
    check(queue.size() == 3 && queue[0].Position == 1 && queue[0].JobName == "Weekly Sync",
          "queue views ordered with positions");
    check(manager.mark_queue_entry_blocked(queue[0].QueueEntryId,
                                           "full verification is required"),
          "blocked unattended queue entry records its reason");
    auto blocked_queue = manager.get_queue_entries();
    check(blocked_queue.size() == 3 &&
              blocked_queue[0].LastErrorMessage.find("full verification") != std::string::npos,
          "blocked unattended queue entry remains queued and visible");
    check(!manager.mark_queue_entry_blocked(queue[0].QueueEntryId,
                                            "full verification is required"),
          "repeated unattended deferral does not rewrite an unchanged queue record");

    check(manager.move_queue_entry(queue[2].QueueEntryId, ui::QueueMoveDirection::Top), "move to top");
    auto moved = manager.get_queue_entries();
    check(moved.size() == 3 && moved[0].QueueEntryId == queue[2].QueueEntryId, "queue order updated");

    ui::QueuedJobWorkItem work_item;
    check(manager.try_dequeue_next_job(work_item), "dequeue next");
    check(work_item.Attempt == 1 && work_item.Job.JobId == duplicate->JobId,
          "dequeue returns front entry with attempt count");
    check(manager.get_queue_entries().size() == 2, "dequeue removed entry");

    auto run = manager.create_run_for_job(saved->JobId, "queued-manual", work_item.QueueEntryId, work_item.Attempt);
    check(run.has_value() && run->Status == storage::ManagedJobRunStatus::Queued && run->Summary == "Queued",
          "create_run_for_job initial state");
    check(run.has_value() && run->Options.has_value() &&
              run->Options->ExpectedSourceIdentity == "identity-src" &&
              run->Options->ExpectedDestinationIdentity == "identity-dst",
          "run history keeps the immutable launch options needed for resume");

    manager.mark_run_running(run->RunId, "C:\\journals\\run.json");
    manager.mark_run_paused(run->RunId);
    manager.mark_run_resumed(run->RunId);

    models::CopyJobResult result;
    result.Succeeded = true;
    result.TotalFiles = 10;
    result.CompletedFiles = 10;
    result.AverageBytesPerSecond = 2.5 * 1024 * 1024;
    result.JournalPath = "C:\\journals\\run.json";
    manager.mark_run_completed(run->RunId, result);

    auto runs = manager.get_recent_runs();
    check(!runs.empty() && runs[0].RunId == run->RunId, "newest run first");
    check(runs[0].Status == storage::ManagedJobRunStatus::Completed &&
              runs[0].Summary == "Completed: 10/10 files at 2.5 MB/s.",
          "completion summary text matches .NET format");
    check(runs[0].FinishedUtc.has_value(), "completed run has FinishedUtc");

    models::CopyJobResult assessment_result;
    assessment_result.Succeeded = true;
    assessment_result.TotalFiles = 10;
    assessment_result.CompletedFiles = 10;
    assessment_result.RecoveredFiles = 2;
    assessment_result.BytesRead = 1024;
    models::CopyJobOptions assessment_options = options;
    assessment_options.OperationMode = models::JobOperationMode::ScanOnly;
    assessment_options.DestinationRoot.clear();
    assessment_options.UseExperimentalRawDiskScan = true;
    const std::wstring assessment_journal_path = work + L"\\assessment-journal.json";
    storage::JobJournal assessment_journal;
    assessment_journal.JobId = "assessment-journal";
    assessment_journal.SourceRoot = assessment_options.SourceRoot;
    assessment_journal.DestinationRoot = assessment_options.SourceRoot;
    assessment_journal.RunOptions = assessment_options;
    storage::JournalFileEntry bad_entry;
    bad_entry.RelativePath = "bad-show.mkv";
    bad_entry.SourceLength = 1024;
    bad_entry.BytesCopied = 1024;
    bad_entry.State = storage::FileCopyState::CompletedWithRecovery;
    bad_entry.RescueRanges.push_back(
        {512, 128, storage::RescueRangeState::Bad});
    assessment_journal.Files.set(bad_entry.RelativePath, bad_entry);
    storage::JournalFileEntry pending_entry;
    pending_entry.RelativePath = "pending-show.mkv";
    pending_entry.SourceLength = 2048;
    assessment_journal.Files.set(pending_entry.RelativePath, pending_entry);
    storage::JournalFileEntry unlocalized_entry;
    unlocalized_entry.RelativePath = "fragile-skip.mkv";
    unlocalized_entry.SourceLength = 4096;
    unlocalized_entry.State = storage::FileCopyState::Failed;
    unlocalized_entry.LastError =
        "Fragile mode stopped this older run before range localization.";
    assessment_journal.Files.set(unlocalized_entry.RelativePath, unlocalized_entry);
    storage::JobJournalStore assessment_store;
    check(assessment_store.save(assessment_journal_path, assessment_journal),
          "assessment journal fixture saved");

    auto assessment_run = manager.create_ad_hoc_run(
        assessment_options, "Readability Assessment", "scan");
    manager.mark_run_running(
        assessment_run.RunId,
        storage::fsutil::wide_to_utf8(assessment_journal_path),
        &assessment_options);
    manager.mark_run_completed(assessment_run.RunId, assessment_result);
    const auto report_journal = manager.load_run_journal(assessment_run.RunId);
    check(report_journal.has_value() && report_journal->Files.size() == 3,
          "report loader authenticates and loads the selected run's journal");
    auto wrong_options = assessment_options;
    wrong_options.SourceRoot = "C:\\different-source";
    const auto wrong_run = manager.create_ad_hoc_run(wrong_options, "Wrong roots", "scan");
    manager.mark_run_running(wrong_run.RunId,
                             storage::fsutil::wide_to_utf8(assessment_journal_path), &wrong_options);
    check(!manager.load_run_journal(wrong_run.RunId).has_value(),
          "report loader rejects an authenticated journal belonging to different roots");
    manager.delete_run(wrong_run.RunId);
    auto runs_after_assessment = manager.get_recent_runs();
    check(!runs_after_assessment.empty() &&
              runs_after_assessment[0].Status == storage::ManagedJobRunStatus::Completed &&
              runs_after_assessment[0].Summary.find("completed with findings") !=
                  std::string::npos &&
              runs_after_assessment[0].Summary.find("unreadable ranges in 2") !=
                  std::string::npos,
          "successful assessment history distinguishes unreadable findings from recovery output");
    auto inspection = manager.inspect_run_journal(assessment_run.RunId);
    check(inspection.Loaded && inspection.ScanMode && inspection.TotalFiles == 3 &&
              inspection.CompletedFiles == 1 && inspection.PendingFiles == 1 &&
              inspection.FailedFiles == 1 && inspection.Findings.size() == 2 &&
              inspection.Findings[0].RelativePath == "bad-show.mkv" &&
              inspection.Findings[0].UnreadableBytes == 128 &&
              inspection.Findings[1].RelativePath == "fragile-skip.mkv" &&
              inspection.Findings[1].UnreadableRangeCount == 0 &&
              !inspection.Findings[1].ErrorMessage.empty(),
          "assessment inspection exposes localized and unlocalized failed filenames");
    bool resume_exact = false;
    auto resume_options = manager.get_resume_options(
        assessment_run.RunId, &resume_exact, nullptr);
    check(resume_exact && resume_options.has_value() &&
              resume_options->OperationMode == models::JobOperationMode::ScanOnly &&
              resume_options->UseExperimentalRawDiskScan,
          "run resume restores exact scan mode and direct NTFS backend");

    models::CopyJobOptions legacy_scan_options = assessment_options;
    legacy_scan_options.SourceRoot = storage::fsutil::wide_to_utf8(work + L"\\legacy-source");
    legacy_scan_options.DestinationRoot.clear();
    storage::fsutil::create_directories(
        storage::fsutil::utf8_to_wide(legacy_scan_options.SourceRoot));
    const std::wstring legacy_source = storage::fsutil::get_full_path(
        storage::fsutil::utf8_to_wide(legacy_scan_options.SourceRoot));
    const std::wstring legacy_journal_path =
        storage::JobJournalStore::get_default_journal_path(
            storage::JobJournalStore::build_job_id(legacy_source, legacy_source));
    storage::JobJournal legacy_journal;
    legacy_journal.JobId = "legacy-journal";
    legacy_journal.SourceRoot = legacy_scan_options.SourceRoot;
    legacy_journal.DestinationRoot = legacy_scan_options.SourceRoot;
    check(assessment_store.save(legacy_journal_path, legacy_journal),
          "legacy journal fixture saved at its deterministic path");
    auto legacy_run = manager.create_ad_hoc_run(
        legacy_scan_options, "Readability Assessment", "scan");
    storage::JobCatalog legacy_catalog = storage::JobCatalogStore(path).load();
    for (auto& catalog_run : legacy_catalog.Runs) {
        if (catalog_run.RunId != legacy_run.RunId) continue;
        catalog_run.Options.reset();
        catalog_run.JournalPath.clear();
    }
    check(storage::JobCatalogStore(path).save(legacy_catalog),
          "legacy run fixture removes its per-run path and options");
    ui::JobManagerService legacy_manager(path);
    check(legacy_manager.get_resume_journal_path(
              legacy_run.RunId, /*validate_roots*/ true) ==
              storage::fsutil::wide_to_utf8(legacy_journal_path),
          "legacy run safely reconstructs a missing deterministic journal path");
    auto legacy_inspection = legacy_manager.inspect_run_journal(legacy_run.RunId);
    check(legacy_inspection.Loaded && legacy_inspection.ScanMode,
          "legacy run with no destination is restored as assessment mode");

    models::CopyJobResult incomplete_result;
    incomplete_result.TotalFiles = 10;
    incomplete_result.CompletedFiles = 9;
    incomplete_result.FailedFiles = 1;
    incomplete_result.ErrorMessage = "one file failed";
    check(incomplete_result.is_incomplete(), "partial result is classified as incomplete");
    auto incomplete_run = manager.create_ad_hoc_run(options, "Partial Copy", "manual");
    manager.mark_run_running(incomplete_run.RunId, "");
    manager.mark_run_completed(incomplete_run.RunId, incomplete_result);
    auto runs_after_incomplete = manager.get_recent_runs();
    check(!runs_after_incomplete.empty() &&
              runs_after_incomplete[0].Status == storage::ManagedJobRunStatus::Incomplete &&
              runs_after_incomplete[0].Summary.rfind("Incomplete:", 0) == 0,
          "partial run is persisted with a distinct incomplete status");

    auto interrupted = manager.create_ad_hoc_run(options, "Manual Copy", "manual");
    manager.mark_run_running(interrupted.RunId, "C:\\journals\\interrupted.json");
    check(manager.mark_any_running_runs_interrupted("Application closed.") == 1,
          "running run marked interrupted");

    models::CopyJobResult cancelled_result;
    cancelled_result.Cancelled = true;
    cancelled_result.ErrorMessage = "cancelled";
    manager.mark_run_completed(interrupted.RunId, cancelled_result);
    auto cancelled_run = manager.get_run_by_id(interrupted.RunId);
    check(cancelled_run.has_value() &&
              cancelled_run->JournalPath == "C:\\journals\\interrupted.json",
          "empty cancellation result cannot erase the resumable journal path");
    manager.mark_run_running(interrupted.RunId, "C:\\journals\\interrupted.json", &options);
    auto restarted_run = manager.get_run_by_id(interrupted.RunId);
    check(restarted_run.has_value() && !restarted_run->Result.has_value() &&
              !restarted_run->FinishedUtc.has_value() &&
              restarted_run->Status == storage::ManagedJobRunStatus::Running,
          "resumed run clears its stale terminal result while running");

    storage::JobJournalStore::remove_journal_set(assessment_journal_path);
    storage::JobJournalStore::remove_journal_set(legacy_journal_path);

    check(manager.delete_job(duplicate->JobId), "delete job");
    check(manager.get_queue_entries().empty() ||
              manager.get_queue_entries()[0].JobId != duplicate->JobId,
          "delete job drops its queue entries");

    check(manager.clear_run_history() == 5, "clear history removes runs");
    check(manager.get_recent_runs().empty(), "history empty after clear");

    // A fresh service instance sees the persisted state.
    ui::JobManagerService reloaded(path);
    check(reloaded.get_jobs().size() == 1 && reloaded.get_jobs()[0].Name == "Weekly Sync",
          "persisted catalog reloads");
}

} // namespace

void test_task_file() {
    std::printf("--- tasks file: export, import, and hand-written files ---\n");
    models::CopyJobOptions defaults;
    defaults.MaxRetries = 3;
    defaults.BufferSizeBytes = 8 * 1024 * 1024;

    // Export then import gives the same tasks back.
    storage::ManagedJob copy;
    copy.Name = "Photos";
    copy.Pinned = true;
    copy.Options.OperationMode = models::JobOperationMode::Copy;
    copy.Options.SourceRoot = "D:\\Data";
    copy.Options.SelectedRelativePaths = {"Photos", "notes.txt"};
    copy.Options.DestinationRoot = "F:\\Backup";
    copy.Options.MaxRetries = 1;
    copy.Options.ExpectedSourceIdentity = "machine-specific";
    copy.Parts = {{"E:\\Camera", {}}};
    storage::ManagedJob scan;
    scan.Name = "Check D";
    scan.Pinned = false;
    scan.Options.OperationMode = models::JobOperationMode::ScanOnly;
    scan.Options.SourceRoot = "D:\\";
    const std::string text = ui::write_task_file({copy, scan});
    check(text.find("machine-specific") == std::string::npos && text.find("ExpectedSourceIdentity") == std::string::npos,
          "export leaves out media identities and other machine-specific fields");
    const auto back = ui::read_task_file(text, defaults);
    check(back.error.empty() && back.problems.empty() && back.tasks.size() == 2, "an exported file imports cleanly");
    if (back.tasks.size() == 2) {
        const auto& photos = back.tasks[0];
        check(photos.name == "Photos" && photos.pinned &&
                  photos.options.OperationMode == models::JobOperationMode::Copy &&
                  photos.options.SourceRoot == "D:\\Data" &&
                  photos.options.SelectedRelativePaths == std::vector<std::string>{"Photos", "notes.txt"} &&
                  photos.options.DestinationRoot == "F:\\Backup" && photos.options.MaxRetries == 1 &&
                  photos.options.ExpectedSourceIdentity.empty(),
              "a copy task round-trips its items, destination, pin, and options");
        check(photos.parts.size() == 1 && photos.parts[0].SourceRoot == "E:\\Camera" &&
                  photos.parts[0].SelectedRelativePaths.empty(),
              "a contents-only folder part keeps its meaning");
        check(back.tasks[1].options.OperationMode == models::JobOperationMode::ScanOnly &&
                  back.tasks[1].options.SourceRoot == "D:\\" && back.tasks[1].options.DestinationRoot.empty() &&
                  !back.tasks[1].pinned,
              "an assessment of a drive round-trips");
    }

    // A hand-written file: missing options come from the defaults, and a
    // file cannot set machine-specific fields through Options.
    const std::string written = R"({
      "Format": "XactCopy.Tasks", "Version": 1,
      "Tasks": [
        { "Sources": ["C:\\Work\\Projects\\"], "Destination": "G:\\Copy",
          "Options": { "VerificationMode": "Full", "SourceRoot": "X:\\Evil", "ExpectedSourceIdentity": "x" } },
        { "Action": "Assess", "Sources": [{ "Path": "C:\\Work", "ContentsOnly": true }] },
        { "Sources": ["C:\\a.txt"] },
        { "Action": "Delete", "Sources": ["C:\\a.txt"], "Destination": "D:\\" },
        { "Destination": "D:\\" }
      ]
    })";
    const auto read = ui::read_task_file(written, defaults);
    check(read.error.empty() && read.tasks.size() == 2 && read.problems.size() == 3,
          "valid tasks import and the three broken ones are reported");
    if (read.tasks.size() == 2) {
        const auto& first = read.tasks[0];
        check(first.options.SourceRoot == "C:\\Work" &&
                  first.options.SelectedRelativePaths == std::vector<std::string>{"Projects"} &&
                  first.options.MaxRetries == 3 && first.options.BufferSizeBytes == 8 * 1024 * 1024 &&
                  first.options.VerificationModeValue == models::VerificationMode::Full &&
                  first.options.ExpectedSourceIdentity.empty() && !first.name.empty(),
              "options merge over the defaults, paths come only from Sources, and a name is made");
        check(read.tasks[1].options.SourceRoot == "C:\\Work" && read.tasks[1].options.SelectedRelativePaths.empty() &&
                  read.tasks[1].options.OperationMode == models::JobOperationMode::ScanOnly,
              "a ContentsOnly source becomes a whole-folder root");
    }
    check(!ui::read_task_file(R"({"Format":"Other","Version":1,"Tasks":[]})", defaults).error.empty(),
          "a file of another format is refused");
    check(!ui::read_task_file(R"({"Format":"XactCopy.Tasks","Version":2,"Tasks":[]})", defaults).error.empty(),
          "a newer version is refused");
    check(!ui::read_task_file("not json", defaults).error.empty(), "invalid JSON is refused");
    check(ui::is_task_file_path(L"C:\\x\\My.XCTASKS") && !ui::is_task_file_path(L"C:\\x\\a.json"),
          "tasks files are recognised by their extension");
}

void test_job_manager_tasks() {
    std::printf("--- job manager: tasks, pins, and the shared queue ---\n");
    std::wstring work = make_work_dir();
    std::wstring path = work + L"\\catalog.json";
    {
        ui::JobManagerService manager(path);
        models::CopyJobOptions options;
        options.SourceRoot = "C:\\Src";
        options.DestinationRoot = "D:\\Dst";
        auto saved = manager.save_job("Saved", options);
        auto task_a = manager.save_job("Task A", options, std::string(), /*pinned*/ false);
        auto task_b = manager.save_job("Task B", options, std::string(), /*pinned*/ false);
        check(saved.has_value() && saved->Pinned, "save_job pins by default");
        check(task_a.has_value() && !task_a->Pinned, "a builder task starts unpinned");
        check(manager.count_unpinned_jobs() == 2, "two unpinned tasks");

        auto ordered = manager.get_jobs_in_order();
        check(ordered.size() == 3 && ordered[0].Name == "Saved" && ordered[2].Name == "Task B",
              "tasks list in the order they were added");

        // Re-saving an existing job keeps its pin unless one is given.
        check(manager.save_job("Task A renamed", options, task_a->JobId).has_value() &&
                  !manager.get_job_by_id(task_a->JobId)->Pinned,
              "editing a task keeps it unpinned");

        // Shared queue: back, front, and a job queued twice keeps one entry.
        const std::string entry_a = manager.enqueue_job(task_a->JobId, /*front*/ false);
        const std::string entry_b = manager.enqueue_job(task_b->JobId, /*front*/ false);
        check(!entry_a.empty() && !entry_b.empty(), "tasks enqueue");
        check(manager.enqueue_job(task_a->JobId, false) == entry_a &&
                  manager.get_queue_entries().size() == 2,
              "queuing an already-queued task keeps its single entry");
        check(manager.enqueue_job(task_b->JobId, /*front*/ true) == entry_b &&
                  manager.get_queue_entries()[0].QueueEntryId == entry_b,
              "Run moves an already-queued task to the front");

        check(manager.set_job_pinned(task_b->JobId, true), "pin a task");
        check(!manager.set_job_pinned(task_b->JobId, true), "pinning twice changes nothing");

        // Discard removes unpinned tasks and their queue entries, not history.
        auto run = manager.create_run_for_job(task_a->JobId, "manual");
        check(run.has_value(), "a run for the unpinned task");
        check(manager.count_unpinned_jobs() == 1, "one unpinned task before discard");
        check(manager.delete_unpinned_jobs() == 1, "discard removes the one unpinned task");
        check(manager.get_jobs_in_order().size() == 2, "pinned jobs stay");
        auto queue = manager.get_queue_entries();
        check(queue.size() == 1 && queue[0].JobId == task_b->JobId,
              "the discarded task's queue entry is gone");
        check(manager.get_run_by_id(run->RunId).has_value(), "run history survives a discard");
        // A task with several parts (one per root) runs them as one task.
        models::CopyJobOptions first;
        first.SourceRoot = "C:\\";
        first.SelectedRelativePaths = {"Games"};
        first.DestinationRoot = "D:\\Dst";
        first.ExpectedSourceIdentity = "c-drive";
        auto multi = manager.add_task("Two drives", first, {{"E:\\", {"Media"}}});
        check(multi.has_value() && !multi->Pinned && multi->Parts.size() == 1, "multi-part task saved");
        check(ui::JobManagerService::part_count(*multi) == 2, "two parts");
        const auto second = ui::JobManagerService::part_options(*multi, 1);
        check(second.SourceRoot == "E:\\" && second.SelectedRelativePaths == std::vector<std::string>{"Media"} &&
                  second.DestinationRoot == "D:\\Dst" && second.ExpectedSourceIdentity.empty(),
              "part options swap the root and items, keep the rest, and drop the media binding");
        check(ui::JobManagerService::find_part(*multi, second) == std::optional<std::size_t>(1),
              "a running part is matched back to its index");
        auto part_run = manager.create_run_for_job_part(multi->JobId, 1, "queued-manual", "entry-x", 1);
        check(part_run.has_value() && part_run->DisplayName == "Two drives (part 2/2)" &&
                  part_run->SourceRoot == "E:\\" && part_run->QueueEntryId == "entry-x",
              "a part's run is named and rooted per part");
        auto copy = manager.duplicate_job(multi->JobId, "Two drives (copy)");
        check(copy.has_value() && copy->Parts.size() == 1 && !copy->Pinned, "duplicates keep parts and pin");

        // A task edited in the builder and added again updates itself.
        manager.set_job_pinned(copy->JobId, true);
        models::CopyJobOptions edited = first;
        edited.SelectedRelativePaths = {"Games", "Saves"};
        const std::size_t before = manager.get_jobs_in_order().size();
        auto updated = manager.add_task("Generated title", edited, {}, /*pinned*/ false, copy->JobId);
        check(updated.has_value() && updated->JobId == copy->JobId &&
                  manager.get_jobs_in_order().size() == before,
              "adding an edited task updates it instead of making another");
        check(updated->Name == "Two drives (copy)" && updated->Pinned && updated->Parts.empty() &&
                  updated->Options.SelectedRelativePaths.size() == 2,
              "the update keeps the name and pin and takes the new items and parts");
        auto orphan = manager.add_task("Orphan", edited, {}, /*pinned*/ false, "no-such-job");
        check(orphan.has_value() && orphan->JobId != copy->JobId && !orphan->Pinned,
              "an edit of a task removed meanwhile becomes a new unpinned task");
        manager.delete_job(orphan->JobId);
        manager.delete_job(copy->JobId);
        manager.delete_job(multi->JobId);

        // Runs of one source and destination share a journal: the latest owns it.
        const std::wstring journal = work + L"\\job-shared.json";
        write_pattern_file(journal, 16, 7);
        const std::string journal_utf8 = storage::fsutil::wide_to_utf8(journal);
        auto older = manager.create_run_for_job(task_b->JobId, "manual");
        manager.mark_run_running(older->RunId, journal_utf8);
        Sleep(20);
        auto newer = manager.create_run_for_job(task_b->JobId, "manual");
        manager.mark_run_running(newer->RunId, journal_utf8);
        using JournalState = ui::JobManagerService::RunJournalState;
        check(manager.run_journal_state(*manager.get_run_by_id(newer->RunId)) == JournalState::Available,
              "the latest run owns the shared journal");
        check(manager.run_journal_state(*manager.get_run_by_id(older->RunId)) == JournalState::Replaced,
              "an older run of the same pair reports its journal as replaced");
        check(manager.get_resume_journal_path(older->RunId).empty(),
              "an older run does not resume from the later run's journal");
        DeleteFileW(journal.c_str());
        check(manager.run_journal_state(*manager.get_run_by_id(newer->RunId)) == JournalState::Removed,
              "a deleted journal reports as removed");
        manager.delete_run(older->RunId);
        manager.delete_run(newer->RunId);
    }
    // Everything above was persisted through the signed catalog.
    ui::JobManagerService reopened(path);
    auto jobs = reopened.get_jobs_in_order();
    check(jobs.size() == 2 && jobs[1].Name == "Task B" && jobs[1].Pinned,
          "pins and order survive a reload");
    check(reopened.get_queue_entries().size() == 1, "the queue survives a reload");
}

int main() {
    test_automatic_recovery_budget();
    test_settings_preservation();
    test_settings_failure_reporting();
    test_settings_option_mapping();
    test_recovery_service();
    test_recovery_writer_coalesces();
    test_log_model();
    test_journal_report_and_range_map();
    test_job_catalog_store();
    test_job_manager_service();
    test_job_manager_tasks();
    test_task_file();
    test_clean_job();
    test_kill_recovery();
    test_sleep_resume_does_not_restart_worker();
    test_elevation_handoff();
    test_task_model();
    test_stop_joins_in_flight_recovery();

    if (g_failures == 0) {
        std::printf("SUPERVISOR PASS: %d checks\n", g_checks);
        return 0;
    }
    std::printf("SUPERVISOR FAILED: %d of %d checks\n", g_failures, g_checks);
    return 1;
}
