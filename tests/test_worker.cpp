// -----------------------------------------------------------------------------
// File: tests\test_worker.cpp
// Purpose: Headless tests for worker-engine building blocks that don't need a
//          full IPC round-trip. Currently guards scan_source: O(1) ordinal-
//          ignore-case dedup (regression for the O(n^2) whole-drive hang) and
//          responsive cancellation during enumeration.
// -----------------------------------------------------------------------------

#include <atomic>
#include <chrono>
#include <cstdio>
#include <cstring>
#include <string>
#include <thread>
#include <vector>

#include "../src/ui/selection.h"
#include "../src/worker/engine.h"
#include "../src/worker/engine_support.h"

namespace {

using namespace xact;
using namespace xact::engine;

int g_failures = 0;
int g_checks = 0;

void check(bool condition, const std::string& label) {
    ++g_checks;
    std::printf("  %s: %s\n", condition ? "ok" : "FAIL", label.c_str());
    if (!condition) ++g_failures;
}

std::wstring make_tree(int dirs, int per_dir) {
    wchar_t temp_root[MAX_PATH];
    GetTempPathW(MAX_PATH, temp_root);
    std::wstring root = std::wstring(temp_root) + L"xactscan-" +
                        storage::fsutil::random_temp_suffix().substr(0, 10);
    storage::fsutil::create_directories(root);
    for (int d = 0; d < dirs; ++d) {
        std::wstring sub = root + L"\\dir" + std::to_wstring(d);
        storage::fsutil::create_directories(sub);
        for (int f = 0; f < per_dir; ++f) {
            std::wstring path = sub + L"\\file" + std::to_wstring(f) + L".bin";
            unsigned char byte = static_cast<unsigned char>(f & 0xFF);
            storage::fsutil::write_file_raw(path, &byte, 1, false, false);
        }
    }
    return root;
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

bool files_equal(const std::wstring& left_path, const std::wstring& right_path) {
    auto left = storage::fsutil::read_all_bytes(left_path);
    auto right = storage::fsutil::read_all_bytes(right_path);
    return left.has_value() && right.has_value() && *left == *right;
}

void remove_tree(const std::wstring& root) {
    std::wstring cmd = L"cmd /c rmdir /s /q \"" + root + L"\"";
    _wsystem(cmd.c_str());
}

void test_scan_source_fast_and_complete() {
    std::printf("--- scan_source: fast O(1) enumeration ---\n");
    const int dirs = 50, per_dir = 400; // 20,000 files
    std::wstring root = make_tree(dirs, per_dir);

    int heartbeats = 0;
    auto log = [&heartbeats](const std::string& text) {
        if (text.rfind("Enumerating source:", 0) == 0) ++heartbeats;
    };

    auto t0 = std::chrono::steady_clock::now();
    SourceScanResult scan = scan_source(root, {}, models::SymlinkHandlingMode::Skip, true, log);
    double seconds = std::chrono::duration<double>(std::chrono::steady_clock::now() - t0).count();
    std::printf("  enumerated %zu files in %.3fs\n", scan.files.size(), seconds);

    check(scan.files.size() == static_cast<std::size_t>(dirs * per_dir),
          "all files enumerated");
    // The old O(n^2) dedup made 20k files take many seconds; O(1) is well under.
    check(seconds < 10.0, "20k-file enumeration completes quickly (no O(n^2) dedup)");

    remove_tree(root);
}

void test_scan_source_cancellation() {
    std::printf("--- scan_source: responsive cancellation ---\n");
    std::wstring root = make_tree(40, 400); // 16,000 files

    std::atomic<bool> flag{false};
    std::thread flip([&flag] {
        std::this_thread::sleep_for(std::chrono::milliseconds(1));
        flag.store(true);
    });

    bool threw_user_cancel = false;
    try {
        scan_source(root, {}, models::SymlinkHandlingMode::Skip, true, nullptr,
                    [&flag] { return flag.load(); });
    } catch (const OperationCanceled& oc) {
        threw_user_cancel = oc.user_requested;
    }
    flip.join();

    check(threw_user_cancel, "scan_source throws user-requested OperationCanceled when cancelled");

    remove_tree(root);
}

void test_raw_read_plan_mapping() {
    std::printf("--- raw volume: file-to-volume extent mapping ---\n");
    std::vector<RawDiskExtent> extents{
        RawDiskExtent{0, 8192, 4096, false},
        RawDiskExtent{4096, 0, 4096, true},
        RawDiskExtent{8192, 16384, 8192, false},
    };
    std::vector<RawDiskReadSegment> plan;
    std::string reason;
    bool planned = build_raw_read_plan(extents, 16384, 2048, 10240, plan, reason);
    check(planned && reason.empty(), "raw extent plan accepts a bounded request");
    check(plan.size() == 3, "raw extent plan preserves mapped and sparse segments");
    if (plan.size() == 3) {
        check(!plan[0].Sparse && plan[0].VolumeOffsetBytes == 10240 && plan[0].Length == 2048,
              "raw extent plan translates the first file offset");
        check(plan[1].Sparse && plan[1].Length == 4096,
              "raw extent plan represents the sparse hole");
        check(!plan[2].Sparse && plan[2].VolumeOffsetBytes == 16384 && plan[2].Length == 4096,
              "raw extent plan translates a later extent");
    }

    plan.clear();
    reason.clear();
    check(!build_raw_read_plan(extents, 16384, 16000, 1024, plan, reason) &&
              reason.find("exceeds file bounds") != std::string::npos,
          "raw extent plan rejects reads beyond the file");

    plan.clear();
    reason.clear();
    std::vector<RawDiskExtent> overlapping{
        RawDiskExtent{0, 0, 8192, false},
        RawDiskExtent{4096, 16384, 4096, false},
    };
    check(!build_raw_read_plan(overlapping, 12288, 0, 4096, plan, reason),
          "raw extent plan rejects overlapping extents");
}

void test_exact_item_selection() {
    std::printf("--- selection: exact Explorer item ---\n");
    std::wstring root = make_tree(2, 2);
    std::wstring selected_folder = root + L"\\dir0";
    std::wstring selected_file = root + L"\\dir1\\file0.bin";

    ui::SelectionModel folder_selection;
    check(folder_selection.add(selected_folder), "selected folder accepted");
    check(_wcsicmp(folder_selection.display_path().c_str(), selected_folder.c_str()) == 0,
          "selected folder remains the displayed source");
    std::wstring folder_root = folder_selection.common_root();
    check(_wcsicmp(folder_root.c_str(), root.c_str()) == 0,
          "selected folder uses its parent only as the worker root");
    auto folder_relative = folder_selection.relative_paths(folder_root);
    check(folder_relative.size() == 1 &&
              models::detail::equals_ignore_case(folder_relative[0], "dir0"),
          "selected folder becomes the exact worker filter");

    SourceScanResult folder_scan =
        scan_source(folder_root, folder_relative, models::SymlinkHandlingMode::Skip,
                    true, nullptr);
    check(folder_scan.files.size() == 2 &&
              engine::detail::starts_with_ignore_case(
                  folder_scan.files[0].relative_path, "dir0\\") &&
              engine::detail::starts_with_ignore_case(
                  folder_scan.files[1].relative_path, "dir0\\"),
          "folder filter excludes sibling folders");

    ui::SelectionModel file_selection;
    check(file_selection.add(selected_file), "selected file accepted");
    check(_wcsicmp(file_selection.display_path().c_str(), selected_file.c_str()) == 0,
          "selected file remains the displayed source");
    std::wstring file_root = file_selection.common_root();
    auto file_relative = file_selection.relative_paths(file_root);
    SourceScanResult file_scan =
        scan_source(file_root, file_relative, models::SymlinkHandlingMode::Skip,
                    true, nullptr);
    check(file_scan.files.size() == 1 &&
              models::detail::equals_ignore_case(
                  file_scan.files[0].relative_path, "file0.bin"),
          "file filter copies only the exact selected file");

    remove_tree(root);
}

void remove_result_journal(const models::CopyJobResult& result) {
    if (!result.JournalPath.empty()) {
        storage::JobJournalStore::remove_journal_set(
            storage::fsutil::utf8_to_wide(result.JournalPath));
    }
}

void test_engine_copy_and_sampled_verification() {
    std::printf("--- engine: copy + sampled verification ---\n");
    std::wstring work = make_tree(0, 0);
    std::wstring source = work + L"\\source";
    std::wstring destination = work + L"\\destination";
    storage::fsutil::create_directories(source);
    std::wstring source_file = source + L"\\payload.bin";
    write_pattern_file(source_file, 192 * 1024, 101);

    models::CopyJobOptions options;
    options.SourceRoot = storage::fsutil::wide_to_utf8(source);
    options.DestinationRoot = storage::fsutil::wide_to_utf8(destination);
    options.TransferEnginePolicyValue = models::TransferEnginePolicy::NativeFast;
    options.ResumeFromJournal = false;
    options.UseBadRangeMap = false;
    options.UpdateBadRangeMapFromRun = false;
    options.CopyEmptyDirectories = false;
    options.ParallelSmallFileWorkers = 0; // exercise the engine's auto resolution path
    options.VerifyAfterCopy = true;
    options.VerificationModeValue = models::VerificationMode::Sampled;
    options.VerificationHashAlgorithmValue = models::VerificationHashAlgorithm::Sha512;
    options.SampleVerificationChunkBytes = 4096;
    options.SampleVerificationChunkCount = 3;

    std::atomic<bool> cancel{false};
    ExecutionControl control;
    std::vector<std::string> logs;
    ResilientCopyEngine engine(options, &control, nullptr,
                               [&logs](const std::string& message) { logs.push_back(message); });
    const models::CopyJobResult result = engine.run(cancel);

    bool verified_log = false;
    for (const auto& line : logs) {
        if (line.find("Verifying (sampled): payload.bin") != std::string::npos) {
            verified_log = true;
            break;
        }
    }
    check(result.Succeeded, "native engine copy succeeds");
    check(result.CompletedFiles == 1, "native engine reports one completed file");
    check(files_equal(source_file, destination + L"\\payload.bin"),
          "native engine destination matches source");
    check(verified_log, "sampled verification path is executed");
    remove_result_journal(result);

    models::CopyJobOptions full_options = options;
    full_options.DestinationRoot = storage::fsutil::wide_to_utf8(work + L"\\full-destination");
    full_options.VerificationModeValue = models::VerificationMode::Full;
    std::vector<std::string> full_logs;
    ResilientCopyEngine full_engine(
        full_options, &control, nullptr,
        [&full_logs](const std::string& message) { full_logs.push_back(message); });
    const models::CopyJobResult full_result = full_engine.run(cancel);
    bool full_verified_log = false;
    for (const auto& line : full_logs) {
        if (line.find("Verifying (full hash): payload.bin") != std::string::npos) {
            full_verified_log = true;
            break;
        }
    }
    check(full_result.Succeeded, "native engine full verification copy succeeds");
    check(full_verified_log, "full verification path is executed");
    remove_result_journal(full_result);
    remove_tree(work);
}

void test_engine_fast_scan_profile() {
    std::printf("--- engine: fast scan profile ---\n");
    std::wstring work = make_tree(2, 3);
    std::wstring destination = work + L"\\unused-destination";

    models::CopyJobOptions options;
    options.SourceRoot = storage::fsutil::wide_to_utf8(work);
    options.DestinationRoot = storage::fsutil::wide_to_utf8(destination);
    options.OperationMode = models::JobOperationMode::ScanOnly;
    options.ResumeFromJournal = false;
    options.UseBadRangeMap = false;
    options.UpdateBadRangeMapFromRun = false;
    options.ScanPerformanceProfileValue = models::ScanPerformanceProfile::Fast;
    options.ParallelScanWorkers = 2;

    std::atomic<bool> cancel{false};
    ExecutionControl control;
    std::vector<std::string> logs;
    ResilientCopyEngine engine(options, &control, nullptr,
                               [&logs](const std::string& message) { logs.push_back(message); });
    const models::CopyJobResult result = engine.run(cancel);

    bool profile_log = false;
    for (const auto& line : logs) {
        if (line.find("Scan performance profile: Fast; workers=2.") != std::string::npos) {
            profile_log = true;
            break;
        }
    }
    check(result.Succeeded, "fast scan succeeds");
    check(result.TotalFiles == 6, "fast scan reports all files");
    check(profile_log, "fast scan profile and worker count are applied");
    remove_result_journal(result);
    remove_tree(work);
}

void test_engine_raw_scan_backend_or_fallback() {
    std::printf("--- engine: raw volume scan backend or safe fallback ---\n");
    std::wstring work = make_tree(0, 0);
    std::wstring source = work + L"\\source";
    std::wstring destination = work + L"\\unused-destination";
    storage::fsutil::create_directories(source);
    write_pattern_file(source + L"\\payload.bin", 96 * 1024, 707);

    models::CopyJobOptions options;
    options.SourceRoot = storage::fsutil::wide_to_utf8(source);
    options.DestinationRoot = storage::fsutil::wide_to_utf8(destination);
    options.OperationMode = models::JobOperationMode::ScanOnly;
    options.UseExperimentalRawDiskScan = true;
    options.ResumeFromJournal = false;
    options.UseBadRangeMap = false;
    options.UpdateBadRangeMapFromRun = false;
    options.ScanPerformanceProfileValue = models::ScanPerformanceProfile::Precise;

    std::atomic<bool> cancel{false};
    ExecutionControl control;
    std::vector<std::string> logs;
    ResilientCopyEngine engine(options, &control, nullptr,
                               [&logs](const std::string& message) { logs.push_back(message); });
    const models::CopyJobResult result = engine.run(cancel);

    bool backend_decision_logged = false;
    bool old_native_fallback_logged = false;
    for (const auto& line : logs) {
        if (line.find("Scan backend: Raw disk") != std::string::npos) {
            backend_decision_logged = true;
        }
        if (line.find("not ported in native build") != std::string::npos) {
            old_native_fallback_logged = true;
        }
    }
    check(result.Succeeded, "raw scan succeeds whether raw access is enabled or unavailable");
    check(backend_decision_logged, "raw scan logs its enabled or fallback backend decision");
    check(!old_native_fallback_logged, "raw scan no longer reports the native backend as unported");
    remove_result_journal(result);
    remove_tree(work);
}

void test_raw_disk_elevated_integration() {
    wchar_t enabled[8]{};
    DWORD length = GetEnvironmentVariableW(L"XACTCOPY_RAW_DISK_TEST", enabled,
                                           static_cast<DWORD>(sizeof(enabled) / sizeof(enabled[0])));
    if (length == 0 || _wcsicmp(enabled, L"1") != 0) {
        std::printf("--- raw volume: elevated integration skipped (set XACTCOPY_RAW_DISK_TEST=1) ---\n");
        return;
    }

    std::printf("--- raw volume: elevated integration ---\n");
    std::wstring work = make_tree(0, 0);
    std::wstring source = work + L"\\source";
    storage::fsutil::create_directories(source);
    std::wstring file = source + L"\\payload.bin";
    write_pattern_file(file, 128 * 1024, 1701);

    std::string reason;
    auto context = RawDiskScanContext::try_create(source, reason);
    check(context != nullptr, "elevated raw context opens the local NTFS source volume");
    if (context) {
        auto expected = storage::fsutil::read_all_bytes(file);
        std::vector<unsigned char> actual(64 * 1024);
        std::atomic<bool> cancel{false};
        CancelContext read_cancel{&cancel, 0};
        RawDiskReadResult result = context->read_chunk(
            file, "payload.bin", 0, static_cast<std::int32_t>(actual.size()), actual.data(),
            10000, read_cancel);
        check(result.state == RawDiskReadResult::State::Success && result.BytesRead == actual.size(),
              "elevated raw context reads a file extent successfully");
        check(expected.has_value() && expected->size() >= actual.size() &&
                  std::memcmp(expected->data(), actual.data(), actual.size()) == 0,
              "elevated raw read matches standard file bytes");
    } else {
        std::printf("  raw context reason: %s\n", reason.c_str());
    }

    remove_tree(work);
}

} // namespace

int main() {
    test_scan_source_fast_and_complete();
    test_scan_source_cancellation();
    test_raw_read_plan_mapping();
    test_exact_item_selection();
    test_engine_copy_and_sampled_verification();
    test_engine_fast_scan_profile();
    test_engine_raw_scan_backend_or_fallback();
    test_raw_disk_elevated_integration();

    if (g_failures == 0) {
        std::printf("WORKER PASS: %d checks\n", g_checks);
        return 0;
    }
    std::printf("WORKER FAILED: %d of %d checks\n", g_failures, g_checks);
    return 1;
}
