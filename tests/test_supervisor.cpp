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
#include <string>
#include <vector>

#include "../src/ui/job_manager.h"
#include "../src/ui/recovery.h"
#include "../src/ui/settings.h"
#include "../src/ui/supervisor.h"

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
        service.mark_job_started("run-123", "Recover job", options, "C:\\journal.json", settings);
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
        }
        service.mark_job_ended("run-123");
        service.mark_clean_shutdown();
    }

    {
        ui::RecoveryService service{ui::RecoveryStateStore(state_path), false};
        ui::RecoveryStartupInfo info = service.initialize_session(settings, launch);
        check(!info.has_interrupted_run(), "clean shutdown leaves no interrupted run");
    }
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
    check(loaded.SchemaVersion == 2, "schema version 2");
    check(loaded.Jobs.size() == 1 && loaded.Jobs[0].Name == "Nightly Backup", "job round-trip");
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
    check(legacy.SchemaVersion == 2, "legacy schema upgraded");
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

    // Corrupted file falls back to a default catalog like the .NET store.
    std::string garbage = "{ not json";
    storage::fsutil::write_file_raw(path, reinterpret_cast<const unsigned char*>(garbage.data()),
                                    garbage.size(), false, false);
    check(store.load().Jobs.empty(), "corrupted catalog loads default");
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
    check(saved.has_value() && saved->Options.ExpectedSourceIdentity.empty() &&
              saved->Options.ResumeJournalPathHint.empty() && !saved->Options.AllowJournalRootRemap,
          "save_job scrubs per-run template fields");

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

    auto interrupted = manager.create_ad_hoc_run(options, "Manual Copy", "manual");
    manager.mark_run_running(interrupted.RunId, "");
    check(manager.mark_any_running_runs_interrupted("Application closed.") == 1,
          "running run marked interrupted");

    check(manager.delete_job(duplicate->JobId), "delete job");
    check(manager.get_queue_entries().empty() ||
              manager.get_queue_entries()[0].JobId != duplicate->JobId,
          "delete job drops its queue entries");

    check(manager.clear_run_history() == 2, "clear history removes runs");
    check(manager.get_recent_runs().empty(), "history empty after clear");

    // A fresh service instance sees the persisted state.
    ui::JobManagerService reloaded(path);
    check(reloaded.get_jobs().size() == 1 && reloaded.get_jobs()[0].Name == "Weekly Sync",
          "persisted catalog reloads");
}

} // namespace

int main() {
    test_settings_preservation();
    test_recovery_service();
    test_job_catalog_store();
    test_job_manager_service();
    test_clean_job();
    test_kill_recovery();

    if (g_failures == 0) {
        std::printf("SUPERVISOR PASS: %d checks\n", g_checks);
        return 0;
    }
    std::printf("SUPERVISOR FAILED: %d of %d checks\n", g_failures, g_checks);
    return 1;
}
