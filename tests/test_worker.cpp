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
#include <filesystem>
#include <string>
#include <system_error>
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

bool write_named_stream(const std::wstring& file_path, std::wstring_view stream_name,
                        const std::vector<unsigned char>& bytes) {
    std::wstring stream_path = file_path + std::wstring(stream_name);
    return storage::fsutil::write_file_raw(stream_path, bytes.data(), bytes.size(), false, false);
}

void remove_tree(const std::wstring& root) {
    std::error_code cleanup_error;
    std::filesystem::remove_all(std::filesystem::path(root), cleanup_error);
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

void test_scan_source_partial_enumeration_is_visible() {
    std::printf("--- scan_source: partial enumeration is never reported complete ---\n");
    std::wstring root = make_tree(2, 2);

    SourceScanResult start_failure = scan_source(
        root, {}, models::SymlinkHandlingMode::Skip, true, nullptr, nullptr,
        [&root](const std::wstring& current, bool at_start) -> DWORD {
            if (at_start && _wcsicmp(current.c_str(), root.c_str()) == 0) return ERROR_ACCESS_DENIED;
            return ERROR_SUCCESS;
        });
    check(!start_failure.complete && !start_failure.errors.empty() && start_failure.files.empty(),
          "enumeration start failure produces an incomplete result");

    const std::wstring failed_directory = root + L"\\dir0";
    SourceScanResult end_failure = scan_source(
        root, {}, models::SymlinkHandlingMode::Skip, true, nullptr, nullptr,
        [&failed_directory](const std::wstring& current, bool at_start) -> DWORD {
            if (!at_start && _wcsicmp(current.c_str(), failed_directory.c_str()) == 0) {
                return ERROR_GEN_FAILURE;
            }
            return ERROR_SUCCESS;
        });
    check(!end_failure.complete && !end_failure.errors.empty() && !end_failure.files.empty(),
          "enumeration end failure keeps partial files but marks the result incomplete");

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

    std::vector<std::string> filter_logs;
    SelectionFilter component_filter = SelectionFilter::create(
        root, {"report..txt", "folder\\..\\escape.bin"},
        [&filter_logs](const std::string& line) { filter_logs.push_back(line); });
    check(component_filter.should_include_file("report..txt"),
          "selection filter permits harmless filenames containing two dots");
    check(!component_filter.should_include_file("escape.bin") && !filter_logs.empty(),
          "selection filter rejects an actual parent-directory component");

    remove_tree(root);
}

void test_scan_symlink_scope_policy() {
    std::printf("--- scan_source: symbolic-link scope policy ---\n");
    std::wstring work = make_tree(0, 0);
    std::wstring root = work + L"\\root";
    std::wstring internal = root + L"\\internal";
    std::wstring external = work + L"\\external";
    storage::fsutil::create_directories(internal);
    storage::fsutil::create_directories(external);
    write_pattern_file(internal + L"\\inside.bin", 4096, 1101);
    write_pattern_file(external + L"\\outside.bin", 4096, 1102);

    constexpr DWORD allow_unprivileged_create = 0x2;
    bool internal_link = CreateSymbolicLinkW(
                             (root + L"\\internal-link").c_str(), internal.c_str(),
                             SYMBOLIC_LINK_FLAG_DIRECTORY | allow_unprivileged_create) != FALSE;
    bool external_link = CreateSymbolicLinkW(
                             (root + L"\\external-link").c_str(), external.c_str(),
                             SYMBOLIC_LINK_FLAG_DIRECTORY | allow_unprivileged_create) != FALSE;
    if (internal_link && external_link) {
        const bool cycle_link = CreateSymbolicLinkW(
                                    (internal + L"\\cycle-to-root").c_str(), root.c_str(),
                                    SYMBOLIC_LINK_FLAG_DIRECTORY | allow_unprivileged_create) != FALSE;
        std::vector<std::string> logs;
        SourceScanResult scoped = scan_source(
            root, {}, models::SymlinkHandlingMode::FollowInternal, true,
            [&logs](const std::string& line) { logs.push_back(line); });
        bool copied_external = false;
        bool copied_internal_path = false;
        bool copied_internal_alias = false;
        for (const auto& file : scoped.files) {
            if (engine::detail::starts_with_ignore_case(file.relative_path,
                                                        "external-link\\")) {
                copied_external = true;
            }
            if (models::detail::equals_ignore_case(file.relative_path,
                                                   "internal\\inside.bin")) {
                copied_internal_path = true;
            }
            if (models::detail::equals_ignore_case(file.relative_path,
                                                   "internal-link\\inside.bin")) {
                copied_internal_alias = true;
            }
        }
        bool logged_scope_rejection = false;
        bool logged_cycle_rejection = false;
        for (const auto& line : logs) {
            if (line.find("Skipping external symbolic-link target") != std::string::npos) {
                logged_scope_rejection = true;
            }
            if (line.find("Skipping symbolic-link directory cycle") != std::string::npos) {
                logged_cycle_rejection = true;
            }
        }
        check(scoped.complete && !copied_external && logged_scope_rejection,
              "internal-only link policy refuses an external resolved target");
        check(copied_internal_path && copied_internal_alias,
              "internal link following preserves both logical source paths");
        if (cycle_link) {
            check(logged_cycle_rejection && scoped.files.size() == 2,
                  "symbolic-link cycles stop without suppressing legitimate aliases");
        }

        SourceScanResult followed = scan_source(
            root, {}, models::SymlinkHandlingMode::Follow, true, nullptr);
        bool followed_external = false;
        for (const auto& file : followed.files) {
            if (engine::detail::starts_with_ignore_case(file.relative_path,
                                                        "external-link\\")) {
                followed_external = true;
            }
        }
        check(followed.complete && followed_external,
              "expert external-link policy follows an explicitly allowed target");
    } else {
        std::printf("  symbolic-link creation unavailable (Win32 %lu); scope integration skipped\n",
                    static_cast<unsigned long>(GetLastError()));
    }
    remove_tree(work);
}

void remove_result_journal(const models::CopyJobResult& result) {
    if (!result.JournalPath.empty()) {
        storage::JobJournalStore::remove_journal_set(
            storage::fsutil::utf8_to_wide(result.JournalPath));
    }
}

void remove_bad_range_map_set(const std::wstring& map_path) {
    const std::wstring mirror =
        storage::detail::build_mirror_path(map_path, L"badmaps-mirror", L"badmap");
    storage::fsutil::delete_file(map_path);
    storage::fsutil::delete_file(mirror);
    for (int generation = 1; generation <= storage::BadRangeMapStore::BackupGenerationCount;
         ++generation) {
        storage::fsutil::delete_file(storage::detail::backup_path(map_path, generation));
        storage::fsutil::delete_file(storage::detail::backup_path(mirror, generation));
    }
}

void test_file_stability_detects_same_size_timestamp_mutation() {
    std::printf("--- source stability: same-size/timestamp mutation ---\n");
    std::wstring work = make_tree(0, 0);
    std::wstring file = work + L"\\payload.bin";
    write_pattern_file(file, 64 * 1024, 1801);

    FileStabilitySnapshot before;
    check(try_get_file_stability(file, before),
          "source stability snapshot captures identity and change time");
    Sleep(20);
    HANDLE handle = CreateFileW(file.c_str(), GENERIC_WRITE, FILE_SHARE_READ,
                                nullptr, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, nullptr);
    bool mutated = false;
    if (handle != INVALID_HANDLE_VALUE) {
        const unsigned char replacement = 0xA5;
        DWORD written = 0;
        mutated = WriteFile(handle, &replacement, 1, &written, nullptr) != FALSE && written == 1;
        FlushFileBuffers(handle);
        CloseHandle(handle);
    }
    check(mutated && set_last_write_time_utc(file, before.last_write_utc_ticks),
          "test mutation restores the enumerated size and last-write timestamp");

    FileStabilitySnapshot after;
    check(try_get_file_stability(file, after),
          "source stability snapshot can be recaptured after mutation");
    check(after.identity.file_index == before.identity.file_index &&
              after.identity.volume_serial == before.identity.volume_serial &&
              after.length == before.length,
          "mutation keeps the same file identity and length");
    check(!file_stability_matches(before, after),
          "change-time stability rejects content mutation hidden behind a restored timestamp");
    remove_tree(work);
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
    check(result.IntegrityNotice.find("Only sampled destination verification") !=
              std::string::npos,
          "sampled verification limitation is carried into the result summary");
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

void test_engine_preserves_metadata_and_cleans_abandoned_stage() {
    std::printf("--- engine: metadata preservation + abandoned-stage cleanup ---\n");
    std::wstring work = make_tree(0, 0);
    std::wstring source = work + L"\\source";
    std::wstring destination = work + L"\\destination";
    storage::fsutil::create_directories(source);
    storage::fsutil::create_directories(destination);
    std::wstring source_empty_directory = source + L"\\empty-folder";
    storage::fsutil::create_directories(source_empty_directory);
    WIN32_FILE_ATTRIBUTE_DATA source_directory_before{};
    check(GetFileAttributesExW(source_empty_directory.c_str(), GetFileExInfoStandard,
                               &source_directory_before) != FALSE,
          "source empty-directory metadata is readable");
    FILETIME directory_write_time = ticks_to_filetime(
        time::DateTimeOffset::now_utc().utc_ticks() - 24 * time::TicksPerHour);
    check(set_file_times_utc(source_empty_directory,
                             source_directory_before.ftCreationTime,
                             source_directory_before.ftLastAccessTime,
                             directory_write_time),
          "source empty-directory timestamp can be pinned");
    SetFileAttributesW(source_empty_directory.c_str(),
                       FILE_ATTRIBUTE_HIDDEN | FILE_ATTRIBUTE_READONLY);

    std::wstring source_file = source + L"\\payload.bin";
    std::wstring destination_file = destination + L"\\payload.bin";
    write_pattern_file(source_file, 96 * 1024, 911);
    const std::vector<unsigned char> stream_bytes{
        0x78, 0x61, 0x63, 0x74, 0x2D, 0x73, 0x74, 0x72, 0x65, 0x61, 0x6D};
    std::wstring test_stream = L":xactcopy-test:$DATA";
    if (!write_named_stream(source_file, test_stream, stream_bytes)) {
        DWORD error = GetLastError();
        if (error == ERROR_INVALID_FUNCTION || error == ERROR_NOT_SUPPORTED ||
            error == ERROR_INVALID_PARAMETER) {
            std::printf("  alternate streams unsupported on this test volume; metadata test skipped\n");
            remove_tree(work);
            return;
        }
        check(false, "source alternate stream can be created");
        remove_tree(work);
        return;
    }
    SetFileAttributesW(source_file.c_str(),
                       FILE_ATTRIBUTE_HIDDEN | FILE_ATTRIBUTE_READONLY | FILE_ATTRIBUTE_ARCHIVE);

    const unsigned char old_bytes[] = {0x4F, 0x4C, 0x44};
    storage::fsutil::write_file_raw(destination_file, old_bytes, sizeof(old_bytes), false, false);
    const std::vector<unsigned char> stale_bytes{0x73, 0x74, 0x61, 0x6C, 0x65};
    write_named_stream(destination_file, L":stale:$DATA", stale_bytes);
    check(SetFileAttributesW(destination_file.c_str(), FILE_ATTRIBUTE_READONLY) != FALSE,
          "existing destination can be marked read-only");

    std::wstring abandoned_stage =
        destination_file + L".xactcopy-stage.0123456789abcdef0123456789abcdef";
    std::wstring marker_decoy = destination + L"\\user.xactcopy-stage.stale";
    storage::fsutil::write_file_raw(abandoned_stage, old_bytes, sizeof(old_bytes), false, false);
    storage::fsutil::write_file_raw(marker_decoy, old_bytes, sizeof(old_bytes), false, false);
    check(set_last_write_time_utc(
              abandoned_stage,
              time::DateTimeOffset::now_utc().utc_ticks() - 2 * 60 * 60 * time::TicksPerSecond),
          "abandoned stage timestamp can be aged");
    check(set_last_write_time_utc(
              marker_decoy,
              time::DateTimeOffset::now_utc().utc_ticks() - 2 * 60 * 60 * time::TicksPerSecond),
          "ordinary marker-containing file timestamp can be aged");

    ExistingFileMetadata source_metadata;
    check(try_get_existing_file_metadata(source_file, source_metadata),
          "source metadata is readable before managed copy");

    models::CopyJobOptions options;
    options.SourceRoot = storage::fsutil::wide_to_utf8(source);
    options.DestinationRoot = storage::fsutil::wide_to_utf8(destination);
    options.TransferEnginePolicyValue = models::TransferEnginePolicy::ManagedRescue;
    options.ResumeFromJournal = false;
    options.UseBadRangeMap = false;
    options.UpdateBadRangeMapFromRun = false;
    options.CopyEmptyDirectories = true;
    options.VerifyAfterCopy = true;
    options.VerificationModeValue = models::VerificationMode::Full;
    options.SalvageUnreadableBlocks = false;
    options.ContinueOnFileError = false;

    std::atomic<bool> cancel{false};
    ExecutionControl control;
    std::vector<std::string> metadata_logs;
    ResilientCopyEngine engine(
        options, &control, nullptr,
        [&metadata_logs](const std::string& message) { metadata_logs.push_back(message); });
    const models::CopyJobResult result = engine.run(cancel);

    auto copied_stream = storage::fsutil::read_all_bytes(destination_file + test_stream);
    ExistingFileMetadata destination_metadata;
    DWORD destination_attributes = GetFileAttributesW(destination_file.c_str());
    check(result.Succeeded, "managed copy with metadata succeeds");
    check(files_equal(source_file, destination_file), "managed destination data matches source");
    check(copied_stream.has_value() && *copied_stream == stream_bytes,
          "alternate data stream bytes are preserved");
    bool alternate_stream_verified = false;
    for (const auto& line : metadata_logs) {
        if (line.find("Verifying alternate stream payload.bin") != std::string::npos) {
            alternate_stream_verified = true;
            break;
        }
    }
    check(alternate_stream_verified,
          "full verification hashes alternate data stream contents");
    check(GetFileAttributesW((destination_file + L":stale:$DATA").c_str()) == INVALID_FILE_ATTRIBUTES,
          "stale destination alternate stream is removed");
    check((destination_attributes & FILE_ATTRIBUTE_HIDDEN) != 0,
          "source hidden attribute is preserved");
    check((destination_attributes & FILE_ATTRIBUTE_READONLY) != 0,
          "source read-only attribute is preserved");
    check(try_get_existing_file_metadata(destination_file, destination_metadata) &&
              destination_metadata.last_write_utc_ticks == source_metadata.last_write_utc_ticks,
          "source last-write timestamp is preserved");
    check(GetFileAttributesW(abandoned_stage.c_str()) == INVALID_FILE_ATTRIBUTES,
          "abandoned staging file is cleaned before the copy");
    check(GetFileAttributesW(marker_decoy.c_str()) != INVALID_FILE_ATTRIBUTES,
          "cleanup never deletes an ordinary file that merely contains the stage marker");
    std::wstring destination_empty_directory = destination + L"\\empty-folder";
    WIN32_FILE_ATTRIBUTE_DATA destination_directory_metadata{};
    DWORD destination_directory_attributes =
        GetFileAttributesW(destination_empty_directory.c_str());
    check(destination_directory_attributes != INVALID_FILE_ATTRIBUTES &&
              (destination_directory_attributes & FILE_ATTRIBUTE_HIDDEN) != 0 &&
              (destination_directory_attributes & FILE_ATTRIBUTE_READONLY) != 0,
          "empty-directory basic attributes are preserved");
    check(GetFileAttributesExW(destination_empty_directory.c_str(), GetFileExInfoStandard,
                               &destination_directory_metadata) != FALSE &&
              filetime_to_ticks(destination_directory_metadata.ftLastWriteTime.dwLowDateTime,
                                destination_directory_metadata.ftLastWriteTime.dwHighDateTime) ==
                  filetime_to_ticks(directory_write_time.dwLowDateTime,
                                    directory_write_time.dwHighDateTime),
          "empty-directory last-write timestamp is preserved after child publication");

    remove_result_journal(result);
    SetFileAttributesW(source_empty_directory.c_str(), FILE_ATTRIBUTE_NORMAL);
    SetFileAttributesW(destination_empty_directory.c_str(), FILE_ATTRIBUTE_NORMAL);
    remove_tree(work);
}

void test_engine_reports_post_publish_metadata_failure_without_rolling_back_bytes() {
    std::printf("--- engine: post-publish metadata failure is not a false rollback ---\n");
    std::wstring work = make_tree(0, 0);
    std::wstring source = work + L"\\source";
    std::wstring destination = work + L"\\destination";
    storage::fsutil::create_directories(source);
    storage::fsutil::create_directories(destination);
    std::wstring source_file = source + L"\\payload.bin";
    std::wstring destination_file = destination + L"\\payload.bin";
    write_pattern_file(source_file, 48 * 1024, 1811);
    write_pattern_file(destination_file, 48 * 1024, 1812);
    SetFileAttributesW(source_file.c_str(), FILE_ATTRIBUTE_READONLY | FILE_ATTRIBUTE_ARCHIVE);

    models::CopyJobOptions options;
    options.SourceRoot = storage::fsutil::wide_to_utf8(source);
    options.DestinationRoot = storage::fsutil::wide_to_utf8(destination);
    options.TransferEnginePolicyValue = models::TransferEnginePolicy::ManagedRescue;
    options.ResumeFromJournal = false;
    options.UseBadRangeMap = false;
    options.UpdateBadRangeMapFromRun = false;
    options.CopyEmptyDirectories = false;
    options.VerifyAfterCopy = true;
    options.VerificationModeValue = models::VerificationMode::Full;

    SetEnvironmentVariableW(L"XACTCOPY_DEV_FAIL_POST_PUBLISH_ATTRIBUTES", L"1");
    std::atomic<bool> cancel{false};
    ExecutionControl control;
    ResilientCopyEngine engine(options, &control, nullptr, nullptr);
    const models::CopyJobResult result = engine.run(cancel);
    SetEnvironmentVariableW(L"XACTCOPY_DEV_FAIL_POST_PUBLISH_ATTRIBUTES", nullptr);

    check(result.Succeeded && files_equal(source_file, destination_file),
          "durably published bytes remain a successful exact copy");
    check(result.MetadataNotice.find("Published file attributes could not be fully restored") !=
              std::string::npos,
          "post-publish attribute failure is surfaced as a metadata warning");
    remove_result_journal(result);
    SetFileAttributesW(source_file.c_str(), FILE_ATTRIBUTE_NORMAL);
    SetFileAttributesW(destination_file.c_str(), FILE_ATTRIBUTE_NORMAL);
    remove_tree(work);
}

void test_engine_efs_policy_when_supported() {
    std::printf("--- engine: EFS preservation and explicit plaintext policy ---\n");
    std::wstring work = make_tree(0, 0);
    std::wstring source = work + L"\\source";
    storage::fsutil::create_directories(source);
    std::wstring source_file = source + L"\\secret.bin";
    write_pattern_file(source_file, 48 * 1024, 1821);
    if (!EncryptFileW(source_file.c_str())) {
        std::printf("  EFS unavailable for this account/volume (Win32 %lu); test skipped\n",
                    static_cast<unsigned long>(GetLastError()));
        remove_tree(work);
        return;
    }

    models::CopyJobOptions options;
    options.SourceRoot = storage::fsutil::wide_to_utf8(source);
    options.DestinationRoot = storage::fsutil::wide_to_utf8(work + L"\\encrypted-destination");
    options.TransferEnginePolicyValue = models::TransferEnginePolicy::ManagedRescue;
    options.ResumeFromJournal = false;
    options.UseBadRangeMap = false;
    options.UpdateBadRangeMapFromRun = false;
    options.CopyEmptyDirectories = false;
    options.VerifyAfterCopy = true;
    options.VerificationModeValue = models::VerificationMode::Full;
    options.AllowDecryptedDestination = false;

    std::atomic<bool> cancel{false};
    ExecutionControl control;
    ResilientCopyEngine encrypted_engine(options, &control, nullptr, nullptr);
    const models::CopyJobResult encrypted_result = encrypted_engine.run(cancel);
    std::wstring encrypted_destination =
        work + L"\\encrypted-destination\\secret.bin";
    DWORD encrypted_attributes = GetFileAttributesW(encrypted_destination.c_str());
    check(encrypted_result.Succeeded && files_equal(source_file, encrypted_destination),
          "managed copy preserves encrypted source bytes exactly");
    check(encrypted_attributes != INVALID_FILE_ATTRIBUTES &&
              (encrypted_attributes & FILE_ATTRIBUTE_ENCRYPTED) != 0,
          "EFS encryption is preserved unless plaintext is explicitly allowed");

    options.DestinationRoot = storage::fsutil::wide_to_utf8(work + L"\\plaintext-destination");
    options.AllowDecryptedDestination = true;
    ResilientCopyEngine plaintext_engine(options, &control, nullptr, nullptr);
    const models::CopyJobResult plaintext_result = plaintext_engine.run(cancel);
    std::wstring plaintext_destination =
        work + L"\\plaintext-destination\\secret.bin";
    DWORD plaintext_attributes = GetFileAttributesW(plaintext_destination.c_str());
    check(plaintext_result.Succeeded && files_equal(source_file, plaintext_destination),
          "explicit plaintext override still copies source bytes exactly");
    check(plaintext_attributes != INVALID_FILE_ATTRIBUTES &&
              (plaintext_attributes & FILE_ATTRIBUTE_ENCRYPTED) == 0 &&
              plaintext_result.IntegrityNotice.find("EFS encryption was explicitly removed") !=
                  std::string::npos,
          "plaintext EFS downgrade is explicit and reported in the result");

    remove_result_journal(plaintext_result);
    remove_result_journal(encrypted_result);
    remove_tree(work);
}

void test_engine_reports_hard_link_topology_limit() {
    std::printf("--- engine: hard-link topology limitation is explicit ---\n");
    std::wstring work = make_tree(0, 0);
    std::wstring source = work + L"\\source";
    std::wstring destination = work + L"\\destination";
    storage::fsutil::create_directories(source);
    std::wstring source_file = source + L"\\payload.bin";
    std::wstring source_link = source + L"\\payload-link.bin";
    write_pattern_file(source_file, 24 * 1024, 1841);
    if (!CreateHardLinkW(source_link.c_str(), source_file.c_str(), nullptr)) {
        std::printf("  hard links unsupported on this test volume (Win32 %lu); test skipped\n",
                    static_cast<unsigned long>(GetLastError()));
        remove_tree(work);
        return;
    }

    models::CopyJobOptions options;
    options.SourceRoot = storage::fsutil::wide_to_utf8(source);
    options.DestinationRoot = storage::fsutil::wide_to_utf8(destination);
    options.TransferEnginePolicyValue = models::TransferEnginePolicy::ManagedRescue;
    options.ResumeFromJournal = false;
    options.UseBadRangeMap = false;
    options.UpdateBadRangeMapFromRun = false;
    options.CopyEmptyDirectories = false;
    options.VerifyAfterCopy = true;
    options.VerificationModeValue = models::VerificationMode::Full;

    std::atomic<bool> cancel{false};
    ExecutionControl control;
    ResilientCopyEngine engine(options, &control, nullptr, nullptr);
    const models::CopyJobResult result = engine.run(cancel);
    FileIdentity destination_identity;
    FileIdentity destination_link_identity;
    const bool independent_destinations =
        try_get_file_identity(destination + L"\\payload.bin", destination_identity) &&
        try_get_file_identity(destination + L"\\payload-link.bin", destination_link_identity) &&
        destination_identity.file_index != destination_link_identity.file_index;
    check(result.Succeeded && independent_destinations,
          "hard-linked source names remain exact independent destination files");
    check(result.MetadataNotice.find("hard-link topology was not preserved") !=
              std::string::npos,
          "hard-link fidelity limitation is reported in the result");

    remove_result_journal(result);
    remove_tree(work);
}

void test_engine_parallel_native_requires_full_verification() {
    std::printf("--- engine: parallel native path is staged and fully verified ---\n");
    std::wstring work = make_tree(0, 0);
    std::wstring source = work + L"\\source";
    std::wstring destination = work + L"\\destination";
    storage::fsutil::create_directories(source);
    for (int index = 0; index < 4; ++index) {
        write_pattern_file(source + L"\\small" + std::to_wstring(index) + L".bin",
                           12 * 1024, static_cast<unsigned>(900 + index));
    }

    models::CopyJobOptions options;
    options.SourceRoot = storage::fsutil::wide_to_utf8(source);
    options.DestinationRoot = storage::fsutil::wide_to_utf8(destination);
    options.TransferEnginePolicyValue = models::TransferEnginePolicy::NativeFast;
    options.ResumeFromJournal = false;
    options.UseBadRangeMap = false;
    options.UpdateBadRangeMapFromRun = false;
    options.CopyEmptyDirectories = false;
    options.ParallelSmallFileWorkers = 2;
    options.SmallFileThresholdBytes = 64 * 1024;
    options.VerifyAfterCopy = true;
    options.VerificationModeValue = models::VerificationMode::Full;

    std::atomic<bool> cancel{false};
    ExecutionControl control;
    std::vector<std::string> logs;
    ResilientCopyEngine engine(options, &control, nullptr,
                               [&logs](const std::string& message) { logs.push_back(message); });
    const models::CopyJobResult result = engine.run(cancel);

    bool parallel_log = false;
    for (const auto& line : logs) {
        if (line.find("Parallel small file phase complete: 4 succeeded") != std::string::npos) {
            parallel_log = true;
            break;
        }
    }
    check(result.Succeeded, "parallel native copy succeeds after full verification");
    check(result.ParallelNativeFastPathFiles == 4,
          "parallel native path reports every small file");
    check(result.BytesReused == 0,
          "fresh parallel native copies are not misreported as journal reuse");
    check(parallel_log, "parallel native phase reports completed workers");
    for (int index = 0; index < 4; ++index) {
        check(files_equal(source + L"\\small" + std::to_wstring(index) + L".bin",
                          destination + L"\\small" + std::to_wstring(index) + L".bin"),
              "parallel native destination matches source " + std::to_string(index));
    }
    remove_result_journal(result);
    remove_tree(work);
}

void test_engine_refuses_ask_and_preserves_destination() {
    std::printf("--- engine: Ask conflict is explicit and non-destructive ---\n");
    std::wstring work = make_tree(0, 0);
    std::wstring source = work + L"\\source";
    std::wstring destination = work + L"\\destination";
    storage::fsutil::create_directories(source);
    storage::fsutil::create_directories(destination);
    write_pattern_file(source + L"\\payload.bin", 64 * 1024, 901);
    const unsigned char sentinel[] = {0x53, 0x41, 0x46, 0x45, 0x21};
    std::wstring destination_file = destination + L"\\payload.bin";
    storage::fsutil::write_file_raw(destination_file, sentinel, sizeof(sentinel), false, false);

    models::CopyJobOptions options;
    options.SourceRoot = storage::fsutil::wide_to_utf8(source);
    options.DestinationRoot = storage::fsutil::wide_to_utf8(destination);
    options.TransferEnginePolicyValue = models::TransferEnginePolicy::ManagedRescue;
    options.OverwritePolicyValue = models::OverwritePolicy::Ask;
    options.ResumeFromJournal = false;
    options.UseBadRangeMap = false;
    options.UpdateBadRangeMapFromRun = false;
    options.CopyEmptyDirectories = false;
    options.SalvageUnreadableBlocks = false;
    options.ContinueOnFileError = false;

    std::atomic<bool> cancel{false};
    ExecutionControl control;
    std::vector<std::string> logs;
    ResilientCopyEngine engine(options, &control, nullptr,
                               [&logs](const std::string& message) { logs.push_back(message); });
    const models::CopyJobResult result = engine.run(cancel);

    auto preserved = storage::fsutil::read_all_bytes(destination_file);
    check(!result.Succeeded, "Ask conflict is not reported as a clean success");
    check(result.FailedFiles == 1 && result.SkippedFiles == 0,
          "Ask conflict is counted as an explicit file failure");
    check(preserved.has_value() && preserved->size() == sizeof(sentinel) &&
              std::memcmp(preserved->data(), sentinel, sizeof(sentinel)) == 0,
          "Ask conflict preserves the existing destination bytes");

    bool refusal_logged = false;
    for (const auto& line : logs) {
        if (line.find("requires explicit confirmation") != std::string::npos) {
            refusal_logged = true;
            break;
        }
    }
    check(refusal_logged, "Ask conflict explains why overwrite was refused");
    remove_result_journal(result);
    remove_tree(work);
}

void test_engine_stages_destination_on_failure() {
    std::printf("--- engine: failed copy preserves previous destination ---\n");
    std::wstring work = make_tree(0, 0);
    std::wstring source = work + L"\\source";
    std::wstring destination = work + L"\\destination";
    storage::fsutil::create_directories(source);
    storage::fsutil::create_directories(destination);
    write_pattern_file(source + L"\\payload.bin", 64 * 1024, 902);
    const unsigned char sentinel[] = {0x4F, 0x4C, 0x44, 0x21};
    std::wstring destination_file = destination + L"\\payload.bin";
    storage::fsutil::write_file_raw(destination_file, sentinel, sizeof(sentinel), false, false);

    SetEnvironmentVariableW(L"XACTCOPY_DEV_FAULT_RULES", L"read,0,0,always,io");
    models::CopyJobOptions options;
    options.SourceRoot = storage::fsutil::wide_to_utf8(source);
    options.DestinationRoot = storage::fsutil::wide_to_utf8(destination);
    options.TransferEnginePolicyValue = models::TransferEnginePolicy::ManagedRescue;
    options.OverwritePolicyValue = models::OverwritePolicy::Overwrite;
    options.ResumeFromJournal = false;
    options.UseBadRangeMap = false;
    options.UpdateBadRangeMapFromRun = false;
    options.CopyEmptyDirectories = false;
    options.MaxRetries = 0;
    options.InitialRetryDelay = time::TimeSpan::from_milliseconds(1);
    options.MaxRetryDelay = time::TimeSpan::from_milliseconds(1);
    options.SalvageUnreadableBlocks = false;
    options.ContinueOnFileError = false;
    options.SmallFileThresholdBytes = 4096;

    std::atomic<bool> cancel{false};
    ExecutionControl control;
    ResilientCopyEngine engine(options, &control, nullptr, nullptr);
    const models::CopyJobResult result = engine.run(cancel);
    SetEnvironmentVariableW(L"XACTCOPY_DEV_FAULT_RULES", nullptr);

    auto preserved = storage::fsutil::read_all_bytes(destination_file);
    check(!result.Succeeded && result.FailedFiles == 1,
          "injected source failure is reported as a failed copy");
    check(preserved.has_value() && preserved->size() == sizeof(sentinel) &&
              std::memcmp(preserved->data(), sentinel, sizeof(sentinel)) == 0,
          "failed copy leaves the previous destination intact");
    remove_result_journal(result);
    remove_tree(work);
}

void test_engine_fault_injection_timeout_and_offline_are_safe() {
    std::printf("--- engine: timeout retry and offline failure safety ---\n");
    std::wstring work = make_tree(0, 0);
    std::wstring source = work + L"\\source";
    std::wstring destination = work + L"\\destination";
    storage::fsutil::create_directories(source);
    storage::fsutil::create_directories(destination);
    std::wstring source_file = source + L"\\payload.bin";
    std::wstring destination_file = destination + L"\\payload.bin";
    write_pattern_file(source_file, 64 * 1024, 914);

    models::CopyJobOptions options;
    options.SourceRoot = storage::fsutil::wide_to_utf8(source);
    options.DestinationRoot = storage::fsutil::wide_to_utf8(destination);
    options.TransferEnginePolicyValue = models::TransferEnginePolicy::ManagedRescue;
    options.ResumeFromJournal = false;
    options.UseBadRangeMap = false;
    options.UpdateBadRangeMapFromRun = false;
    options.CopyEmptyDirectories = false;
    options.MaxRetries = 1;
    options.InitialRetryDelay = time::TimeSpan::from_milliseconds(1);
    options.MaxRetryDelay = time::TimeSpan::from_milliseconds(1);
    options.SalvageUnreadableBlocks = false;
    options.ContinueOnFileError = false;

    SetEnvironmentVariableW(L"XACTCOPY_DEV_FAULT_RULES", L"read,0,0,once,timeout");
    std::vector<std::string> retry_logs;
    std::atomic<bool> cancel{false};
    ExecutionControl control;
    ResilientCopyEngine retry_engine(
        options, &control, nullptr,
        [&retry_logs](const std::string& message) { retry_logs.push_back(message); });
    const models::CopyJobResult retry_result = retry_engine.run(cancel);
    SetEnvironmentVariableW(L"XACTCOPY_DEV_FAULT_RULES", nullptr);

    bool saw_retry = false;
    for (const auto& line : retry_logs) {
        if (line.find("Read retry") != std::string::npos) {
            saw_retry = true;
            break;
        }
    }
    check(retry_result.Succeeded && saw_retry,
          "a transient timeout is retried and completes exactly");
    check(files_equal(source_file, destination_file),
          "timeout retry destination matches source");

    const unsigned char sentinel[] = {0x53, 0x41, 0x46, 0x45};
    storage::fsutil::write_file_raw(destination_file, sentinel, sizeof(sentinel), false, false);
    options.MaxRetries = 0;
    SetEnvironmentVariableW(L"XACTCOPY_DEV_FAULT_RULES", L"write,0,0,once,offline");
    ResilientCopyEngine offline_engine(options, &control, nullptr, nullptr);
    const models::CopyJobResult offline_result = offline_engine.run(cancel);
    SetEnvironmentVariableW(L"XACTCOPY_DEV_FAULT_RULES", nullptr);

    auto preserved = storage::fsutil::read_all_bytes(destination_file);
    check(!offline_result.Succeeded && offline_result.FailedFiles == 1,
          "offline destination failure is reported as incomplete");
    check(preserved.has_value() && preserved->size() == sizeof(sentinel) &&
              std::memcmp(preserved->data(), sentinel, sizeof(sentinel)) == 0,
          "offline destination failure leaves the previous file intact");

    remove_result_journal(offline_result);
    remove_result_journal(retry_result);
    remove_tree(work);
}

void test_engine_fragile_first_read_findings_are_durable() {
    std::printf("--- engine: fragile first-read findings survive into journal and map ---\n");
    std::wstring work = make_tree(0, 0);
    std::wstring scan_source = work + L"\\scan-source";
    std::wstring scan_destination = work + L"\\scan-unused";
    storage::fsutil::create_directories(scan_source);
    write_pattern_file(scan_source + L"\\payload.bin", 64 * 1024, 1954);

    const std::wstring scan_map_path = storage::BadRangeMapStore::get_default_map_path(
        storage::fsutil::get_full_path(scan_source));
    remove_bad_range_map_set(scan_map_path);

    models::CopyJobOptions scan_options;
    scan_options.SourceRoot = storage::fsutil::wide_to_utf8(scan_source);
    scan_options.DestinationRoot = storage::fsutil::wide_to_utf8(scan_destination);
    scan_options.OperationMode = models::JobOperationMode::ScanOnly;
    scan_options.ScanPerformanceProfileValue = models::ScanPerformanceProfile::Precise;
    scan_options.ResumeFromJournal = false;
    scan_options.UseBadRangeMap = true;
    scan_options.SkipKnownBadRanges = true;
    scan_options.UpdateBadRangeMapFromRun = true;
    scan_options.FragileMediaMode = true;
    scan_options.SkipFileOnFirstReadError = true;
    scan_options.PersistFragileSkipAcrossResume = true;
    scan_options.WaitForMediaAvailability = false;
    scan_options.ContinueOnFileError = true;
    scan_options.MaxRetries = 4;
    scan_options.UseAdaptiveBufferSizing = false;

    std::atomic<bool> cancel{false};
    ExecutionControl control;
    SetEnvironmentVariableW(L"XACTCOPY_DEV_FAULT_RULES", L"read,0,0,always,io");
    ResilientCopyEngine first_scan_engine(scan_options, &control, nullptr, nullptr);
    const models::CopyJobResult first_scan = first_scan_engine.run(cancel);
    SetEnvironmentVariableW(L"XACTCOPY_DEV_FAULT_RULES", nullptr);

    storage::JobJournalStore journal_store;
    auto first_journal = journal_store.load(
        storage::fsutil::utf8_to_wide(first_scan.JournalPath));
    const storage::JournalFileEntry* first_entry =
        first_journal.has_value() ? first_journal->Files.find("payload.bin") : nullptr;
    std::int64_t journal_bad_bytes = 0;
    std::int32_t journal_bad_ranges = 0;
    if (first_entry != nullptr) {
        for (const auto& range : first_entry->RescueRanges) {
            if (range.State == storage::RescueRangeState::Bad && range.Length > 0) {
                ++journal_bad_ranges;
                journal_bad_bytes += range.Length;
            }
        }
    }

    storage::BadRangeMapStore map_store;
    auto first_map = map_store.load(scan_map_path);
    const storage::BadRangeMapFileEntry* first_map_entry =
        first_map.has_value() ? first_map->Files.find("payload.bin") : nullptr;
    check(first_scan.SkippedFiles == 1 && first_scan.FailedFiles == 0,
          "fragile assessment reports the intentionally skipped file");
    check(first_entry != nullptr && first_entry->State == storage::FileCopyState::Failed &&
              first_entry->DoNotRetry && journal_bad_ranges == 1 && journal_bad_bytes > 0,
          "fragile assessment journal retains the failed filename and attempted byte range");
    check(first_map_entry != nullptr && first_map_entry->BadRanges.size() == 1 &&
              first_map_entry->BadRanges[0].Offset == 0 &&
              first_map_entry->BadRanges[0].Length == journal_bad_bytes &&
              first_map_entry->ConfirmationCount == 1,
          "first fragile observation is visible in the bad-range map but remains unconfirmed");

    SetEnvironmentVariableW(L"XACTCOPY_DEV_FAULT_RULES", L"read,0,0,always,io");
    ResilientCopyEngine confirming_scan_engine(scan_options, &control, nullptr, nullptr);
    const models::CopyJobResult confirming_scan = confirming_scan_engine.run(cancel);
    SetEnvironmentVariableW(L"XACTCOPY_DEV_FAULT_RULES", nullptr);
    auto confirmed_map = map_store.load(scan_map_path);
    const storage::BadRangeMapFileEntry* confirmed_entry =
        confirmed_map.has_value() ? confirmed_map->Files.find("payload.bin") : nullptr;
    check(confirming_scan.SkippedFiles == 1 && confirmed_entry != nullptr &&
              confirmed_entry->ConfirmationCount >= 2,
          "a second matching fragile observation promotes the range to a trusted hint");

    std::wstring fast_source = work + L"\\fast-source";
    storage::fsutil::create_directories(fast_source);
    write_pattern_file(fast_source + L"\\fast.bin", 64 * 1024, 1957);
    const std::wstring fast_map_path = storage::BadRangeMapStore::get_default_map_path(
        storage::fsutil::get_full_path(fast_source));
    remove_bad_range_map_set(fast_map_path);
    models::CopyJobOptions fast_options = scan_options;
    fast_options.SourceRoot = storage::fsutil::wide_to_utf8(fast_source);
    fast_options.DestinationRoot = fast_options.SourceRoot;
    fast_options.ScanPerformanceProfileValue = models::ScanPerformanceProfile::Fast;

    SetEnvironmentVariableW(L"XACTCOPY_DEV_FAULT_RULES", L"read,0,0,once,io");
    std::vector<std::string> fast_logs;
    ResilientCopyEngine fast_engine(
        fast_options, &control, nullptr,
        [&fast_logs](const std::string& message) { fast_logs.push_back(message); });
    const models::CopyJobResult fast_result = fast_engine.run(cancel);
    SetEnvironmentVariableW(L"XACTCOPY_DEV_FAULT_RULES", nullptr);
    auto fast_map = map_store.load(fast_map_path);
    const storage::BadRangeMapFileEntry* fast_entry =
        fast_map.has_value() ? fast_map->Files.find("fast.bin") : nullptr;
    bool precise_reread_queued = false;
    for (const auto& line : fast_logs) {
        if (line.find("Fast scan fallback queued: fast.bin") != std::string::npos) {
            precise_reread_queued = true;
            break;
        }
    }
    check(fast_result.SkippedFiles == 1 && fast_entry != nullptr &&
              fast_entry->ConfirmationCount == 1,
          "fast fragile assessment preserves a one-shot first-read finding");
    check(!precise_reread_queued,
          "fast fragile assessment does not reread the failed range through precise fallback");

    std::wstring copy_source = work + L"\\copy-source";
    std::wstring copy_destination = work + L"\\copy-destination";
    storage::fsutil::create_directories(copy_source);
    storage::fsutil::create_directories(copy_destination);
    write_pattern_file(copy_source + L"\\copy.bin", 64 * 1024, 1955);
    const std::wstring copy_map_path = storage::BadRangeMapStore::get_default_map_path(
        storage::fsutil::get_full_path(copy_source));
    remove_bad_range_map_set(copy_map_path);

    models::CopyJobOptions copy_options = scan_options;
    copy_options.SourceRoot = storage::fsutil::wide_to_utf8(copy_source);
    copy_options.DestinationRoot = storage::fsutil::wide_to_utf8(copy_destination);
    copy_options.OperationMode = models::JobOperationMode::Copy;
    copy_options.TransferEnginePolicyValue = models::TransferEnginePolicy::ManagedRescue;
    copy_options.SmallFileThresholdBytes = 4096;
    copy_options.UseBadRangeMap = false;
    copy_options.SkipKnownBadRanges = false;

    SetEnvironmentVariableW(L"XACTCOPY_DEV_FAULT_RULES", L"read,0,0,always,io");
    ResilientCopyEngine copy_engine(copy_options, &control, nullptr, nullptr);
    const models::CopyJobResult copy_result = copy_engine.run(cancel);
    SetEnvironmentVariableW(L"XACTCOPY_DEV_FAULT_RULES", nullptr);
    auto copy_map = map_store.load(copy_map_path);
    const storage::BadRangeMapFileEntry* copy_map_entry =
        copy_map.has_value() ? copy_map->Files.find("copy.bin") : nullptr;
    auto copy_journal = journal_store.load(
        storage::fsutil::utf8_to_wide(copy_result.JournalPath));
    const storage::JournalFileEntry* copy_entry =
        copy_journal.has_value() ? copy_journal->Files.find("copy.bin") : nullptr;
    check(copy_result.SkippedFiles == 1 && copy_result.FailedFiles == 0 &&
              copy_entry != nullptr && copy_entry->State == storage::FileCopyState::Failed,
          "fragile copy retains the skipped source filename in its journal");
    check(copy_map_entry != nullptr && copy_map_entry->BadRanges.size() == 1 &&
              copy_map_entry->ConfirmationCount == 1,
          "fragile copy persists its first unreadable primary-data range");
    check(GetFileAttributesW((copy_destination + L"\\copy.bin").c_str()) ==
              INVALID_FILE_ATTRIBUTES,
          "fragile read failure does not publish a partial destination file");

    std::wstring offline_source = work + L"\\offline-source";
    storage::fsutil::create_directories(offline_source);
    write_pattern_file(offline_source + L"\\offline.bin", 32 * 1024, 1956);
    const std::wstring offline_map_path = storage::BadRangeMapStore::get_default_map_path(
        storage::fsutil::get_full_path(offline_source));
    remove_bad_range_map_set(offline_map_path);
    models::CopyJobOptions offline_options = scan_options;
    offline_options.SourceRoot = storage::fsutil::wide_to_utf8(offline_source);
    offline_options.DestinationRoot = offline_options.SourceRoot;
    offline_options.MaxRetries = 0;

    SetEnvironmentVariableW(L"XACTCOPY_DEV_FAULT_RULES", L"read,0,0,always,offline");
    ResilientCopyEngine offline_engine(offline_options, &control, nullptr, nullptr);
    const models::CopyJobResult offline_result = offline_engine.run(cancel);
    SetEnvironmentVariableW(L"XACTCOPY_DEV_FAULT_RULES", nullptr);
    auto offline_map = map_store.load(offline_map_path);
    check(offline_result.FailedFiles == 1 && offline_result.SkippedFiles == 0,
          "an unavailable device is reported as an I/O failure, not a fragile bad-range skip");
    check(offline_map.has_value() && offline_map->Files.empty(),
          "device unavailability does not create a physical bad-range finding");

    remove_result_journal(first_scan);
    remove_result_journal(confirming_scan);
    remove_result_journal(fast_result);
    remove_result_journal(copy_result);
    remove_result_journal(offline_result);
    remove_bad_range_map_set(scan_map_path);
    remove_bad_range_map_set(fast_map_path);
    remove_bad_range_map_set(copy_map_path);
    remove_bad_range_map_set(offline_map_path);
    remove_tree(work);
}

void test_engine_salvage_is_not_success() {
    std::printf("--- engine: salvage is visibly non-exact ---\n");
    std::wstring work = make_tree(0, 0);
    std::wstring source = work + L"\\source";
    std::wstring destination = work + L"\\destination";
    storage::fsutil::create_directories(source);
    write_pattern_file(source + L"\\payload.bin", 64 * 1024, 903);
    storage::fsutil::create_directories(destination);
    write_pattern_file(destination + L"\\payload.bin", 64 * 1024, 1903);
    auto original_destination =
        storage::fsutil::read_all_bytes(destination + L"\\payload.bin");
    const std::wstring map_path = storage::BadRangeMapStore::get_default_map_path(
        storage::fsutil::normalize_path(source));
    remove_bad_range_map_set(map_path);

    SetEnvironmentVariableW(L"XACTCOPY_DEV_FAULT_RULES", L"read,0,0,always,io");
    models::CopyJobOptions options;
    options.SourceRoot = storage::fsutil::wide_to_utf8(source);
    options.DestinationRoot = storage::fsutil::wide_to_utf8(destination);
    options.TransferEnginePolicyValue = models::TransferEnginePolicy::ManagedRescue;
    options.ResumeFromJournal = false;
    options.UseBadRangeMap = false;
    options.UpdateBadRangeMapFromRun = true;
    options.CopyEmptyDirectories = false;
    options.MaxRetries = 0;
    options.InitialRetryDelay = time::TimeSpan::from_milliseconds(1);
    options.MaxRetryDelay = time::TimeSpan::from_milliseconds(1);
    options.SalvageUnreadableBlocks = true;
    options.SalvageFillPatternValue = models::SalvageFillPattern::Ones;
    options.ContinueOnFileError = false;

    std::atomic<bool> cancel{false};
    ExecutionControl control;
    std::vector<std::string> logs;
    ResilientCopyEngine engine(
        options, &control, nullptr,
        [&logs](const std::string& message) { logs.push_back(message); });
    const models::CopyJobResult result = engine.run(cancel);
    SetEnvironmentVariableW(L"XACTCOPY_DEV_FAULT_RULES", nullptr);

    check(!result.Succeeded && result.RecoveredFiles == 1,
          "salvaged copy is not reported as an exact success");
    auto preserved_destination =
        storage::fsutil::read_all_bytes(destination + L"\\payload.bin");
    check(original_destination.has_value() && preserved_destination == original_destination,
          "salvage preserves a better pre-existing destination");
    WIN32_FIND_DATAW sidecar_find{};
    HANDLE sidecar_search = FindFirstFileW(
        (destination + L"\\payload.xactcopy-recovered.*.bin").c_str(), &sidecar_find);
    bool sidecar_exists = sidecar_search != INVALID_HANDLE_VALUE;
    std::wstring sidecar_path;
    if (sidecar_exists) {
        sidecar_path = destination + L"\\" + sidecar_find.cFileName;
        FindClose(sidecar_search);
    }
    check(sidecar_exists && !files_equal(source + L"\\payload.bin", sidecar_path),
          "non-exact salvage is visibly named while preserving its extension");
    const std::wstring manifest_path = sidecar_path + L".recovery.json";
    auto manifest_bytes = storage::fsutil::read_all_bytes(manifest_path);
    bool manifest_valid = false;
    if (manifest_bytes.has_value()) {
        try {
            std::string manifest_text(reinterpret_cast<const char*>(manifest_bytes->data()),
                                      manifest_bytes->size());
            json::Value manifest_value = json::parse(manifest_text);
            const json::Object* manifest = manifest_value.as_object();
            manifest_valid = manifest != nullptr &&
                             manifest->find("SourceRelativePath") != nullptr &&
                             manifest->find("SourceRelativePath")->as_string() == "payload.bin" &&
                             manifest->find("FillPattern") != nullptr &&
                             manifest->find("FillPattern")->as_string() == "0xFF" &&
                             manifest->find("SyntheticRanges") != nullptr &&
                             manifest->find("SyntheticRanges")->as_array() != nullptr &&
                             !manifest->find("SyntheticRanges")->as_array()->empty();
        } catch (const std::exception&) {
            manifest_valid = false;
        }
    }
    check(manifest_valid,
          "recovered output has a readable manifest with source, fill, and synthetic ranges");
    bool sidecar_logged = false;
    for (const auto& line : logs) {
        if (line.find("RECOVERED OUTPUT") != std::string::npos) sidecar_logged = true;
    }
    check(sidecar_logged, "salvage sidecar location is reported in the operation log");
    storage::BadRangeMapStore map_store;
    auto persisted_map = map_store.load(map_path);
    check(persisted_map.has_value() &&
              persisted_map->Files.find("payload.bin") == nullptr,
          "synthetic recovered ranges are not promoted to future skip hints");

    SetEnvironmentVariableW(L"XACTCOPY_DEV_FAULT_RULES", L"read,0,0,always,io");
    options.AllowRecoveredOverwriteExisting = true;
    ResilientCopyEngine override_engine(options, &control, nullptr, nullptr);
    const models::CopyJobResult override_result = override_engine.run(cancel);
    SetEnvironmentVariableW(L"XACTCOPY_DEV_FAULT_RULES", nullptr);
    const std::wstring override_manifest =
        destination + L"\\payload.bin.recovery.json";
    check(!override_result.Succeeded && override_result.RecoveredFiles == 1 &&
              storage::fsutil::file_exists(override_manifest),
          "expert recovered-overwrite output still receives an adjacent manifest");

    options.SalvageUnreadableBlocks = false;
    ResilientCopyEngine exact_engine(options, &control, nullptr, nullptr);
    const models::CopyJobResult exact_result = exact_engine.run(cancel);
    check(exact_result.Succeeded && files_equal(source + L"\\payload.bin",
                                                destination + L"\\payload.bin"),
          "a later exact copy replaces expert recovery output with source bytes");
    check(!storage::fsutil::file_exists(override_manifest),
          "a later exact copy removes its owned obsolete recovery manifest");
    remove_result_journal(result);
    remove_result_journal(override_result);
    remove_result_journal(exact_result);
    remove_bad_range_map_set(map_path);
    remove_tree(work);
}

void test_engine_rejects_unverified_journal_completion() {
    std::printf("--- engine: journal completion requires content validation ---\n");
    std::wstring work = make_tree(0, 0);
    std::wstring source = work + L"\\source";
    std::wstring destination = work + L"\\destination";
    storage::fsutil::create_directories(source);
    write_pattern_file(source + L"\\payload.bin", 96 * 1024, 904);

    models::CopyJobOptions first;
    first.SourceRoot = storage::fsutil::wide_to_utf8(source);
    first.DestinationRoot = storage::fsutil::wide_to_utf8(destination);
    first.TransferEnginePolicyValue = models::TransferEnginePolicy::ManagedRescue;
    first.ResumeFromJournal = false;
    first.UseBadRangeMap = false;
    first.UpdateBadRangeMapFromRun = false;
    first.CopyEmptyDirectories = false;
    first.SalvageUnreadableBlocks = false;
    first.ContinueOnFileError = false;
    first.ParallelSmallFileWorkers = 1;

    std::atomic<bool> cancel{false};
    ExecutionControl control;
    ResilientCopyEngine first_engine(first, &control, nullptr, nullptr);
    const models::CopyJobResult first_result = first_engine.run(cancel);
    check(first_result.Succeeded, "initial journal-backed copy succeeds");

    write_pattern_file(destination + L"\\payload.bin", 96 * 1024, 905);

    models::CopyJobOptions resumed = first;
    resumed.ResumeFromJournal = true;
    resumed.VerifyAfterCopy = true;
    resumed.VerificationModeValue = models::VerificationMode::Full;
    std::vector<std::string> logs;
    ResilientCopyEngine resumed_engine(
        resumed, &control, nullptr,
        [&logs](const std::string& message) { logs.push_back(message); });
    const models::CopyJobResult resumed_result = resumed_engine.run(cancel);

    bool rejected = false;
    for (const auto& line : logs) {
        if (line.find("Journal completion rejected") != std::string::npos) {
            rejected = true;
            break;
        }
    }
    check(resumed_result.Succeeded, "resume repairs a same-length corrupted destination");
    check(files_equal(source + L"\\payload.bin", destination + L"\\payload.bin"),
          "resume restores the source bytes instead of trusting length");
    check(rejected, "resume logs rejection of an invalid journal completion");

    std::vector<std::string> reuse_logs;
    ResilientCopyEngine reuse_engine(
        resumed, &control, nullptr,
        [&reuse_logs](const std::string& message) { reuse_logs.push_back(message); });
    const models::CopyJobResult reuse_result = reuse_engine.run(cancel);
    bool reused_log = false;
    for (const auto& line : reuse_logs) {
        if (line.find("Reused: payload.bin (verified journal completion)") != std::string::npos) {
            reused_log = true;
            break;
        }
    }
    check(reuse_result.Succeeded && reuse_result.CompletedFiles == 1 &&
              reuse_result.SkippedFiles == 0 && reuse_result.BytesReused == 96 * 1024 &&
              reuse_result.BytesWritten == 0,
          "verified journal reuse remains a clean completion");
    check(reused_log, "verified journal reuse is reported distinctly from a policy skip");
    remove_result_journal(reuse_result);
    remove_result_journal(resumed_result);
    remove_result_journal(first_result);
    remove_tree(work);
}

void test_engine_tracks_source_identity_changes() {
    std::printf("--- engine: journal tracks source identity changes ---\n");
    std::wstring work = make_tree(0, 0);
    std::wstring source = work + L"\\source";
    std::wstring destination = work + L"\\destination";
    storage::fsutil::create_directories(source);
    std::wstring source_file = source + L"\\payload.bin";
    write_pattern_file(source_file, 80 * 1024, 906);

    ExistingFileMetadata original_metadata;
    FileIdentity original_identity;
    check(try_get_existing_file_metadata(source_file, original_metadata) &&
              try_get_file_identity(source_file, original_identity),
          "source identity is readable before journaling");

    models::CopyJobOptions options;
    options.SourceRoot = storage::fsutil::wide_to_utf8(source);
    options.DestinationRoot = storage::fsutil::wide_to_utf8(destination);
    options.TransferEnginePolicyValue = models::TransferEnginePolicy::ManagedRescue;
    options.ResumeFromJournal = false;
    options.UseBadRangeMap = false;
    options.UpdateBadRangeMapFromRun = false;
    options.CopyEmptyDirectories = false;
    options.SalvageUnreadableBlocks = false;
    options.ContinueOnFileError = false;
    options.ParallelSmallFileWorkers = 1;

    std::atomic<bool> cancel{false};
    ExecutionControl control;
    ResilientCopyEngine first_engine(options, &control, nullptr, nullptr);
    const models::CopyJobResult first_result = first_engine.run(cancel);
    check(first_result.Succeeded, "initial identity-tracked copy succeeds");

    storage::JobJournalStore store;
    auto first_journal = store.load(storage::fsutil::utf8_to_wide(first_result.JournalPath));
    const auto* first_entry = first_journal.has_value()
                                  ? first_journal->Files.find("payload.bin")
                                  : nullptr;
    check(first_entry != nullptr && first_entry->SourceChangeUtcTicks != 0 &&
              first_entry->SourceFileIndex != 0 &&
              first_entry->SourceVolumeSerial != 0,
          "journal persists source file identity and NTFS change time");

    std::wstring old_source = source_file + L".old";
    MoveFileExW(source_file.c_str(), old_source.c_str(), MOVEFILE_REPLACE_EXISTING);
    write_pattern_file(source_file, 80 * 1024, 907);
    set_last_write_time_utc(source_file, original_metadata.last_write_utc_ticks);
    FileIdentity replacement_identity;
    bool replacement_identity_valid = try_get_file_identity(source_file, replacement_identity);
    check(replacement_identity_valid &&
              (replacement_identity.file_index != original_identity.file_index ||
               replacement_identity.volume_serial != original_identity.volume_serial),
          "replacement source has a different file identity despite matching metadata");

    models::CopyJobOptions resumed = options;
    resumed.ResumeFromJournal = true;
    std::vector<std::string> logs;
    ResilientCopyEngine resumed_engine(
        resumed, &control, nullptr,
        [&logs](const std::string& message) { logs.push_back(message); });
    const models::CopyJobResult resumed_result = resumed_engine.run(cancel);

    bool reset_logged = false;
    for (const auto& line : logs) {
        if (line.find("Journal source binding changed") != std::string::npos) {
            reset_logged = true;
            break;
        }
    }
    check(resumed_result.Succeeded, "copy succeeds after a same-metadata source replacement");
    check(files_equal(source_file, destination + L"\\payload.bin"),
          "replacement source bytes are copied instead of stale journal bytes");
    check(reset_logged, "journal coverage reset is logged for source identity change");
    DeleteFileW(old_source.c_str());
    remove_result_journal(resumed_result);
    remove_result_journal(first_result);
    remove_tree(work);
}

void test_engine_scan_resume_binds_change_time() {
    std::printf("--- engine: scan resume binds NTFS change time ---\n");
    std::wstring work = make_tree(0, 0);
    std::wstring source = work + L"\\source";
    storage::fsutil::create_directories(source);
    const std::wstring source_file = source + L"\\payload.bin";
    constexpr std::int64_t FileLength = 96 * 1024;
    write_pattern_file(source_file, static_cast<std::size_t>(FileLength), 1908);

    FileStabilitySnapshot original;
    check(try_get_file_stability(source_file, original),
          "scan resume: original source stability is readable");

    models::CopyJobOptions options;
    options.OperationMode = models::JobOperationMode::ScanOnly;
    options.SourceRoot = storage::fsutil::wide_to_utf8(source);
    options.TransferEnginePolicyValue = models::TransferEnginePolicy::ManagedRescue;
    options.ScanPerformanceProfileValue = models::ScanPerformanceProfile::Fast;
    options.ParallelScanWorkers = 1;
    options.ResumeFromJournal = false;
    options.UseBadRangeMap = false;
    options.UpdateBadRangeMapFromRun = false;
    options.ContinueOnFileError = false;
    options.MaxRetries = 0;

    std::atomic<bool> cancel{false};
    ExecutionControl control;
    ResilientCopyEngine first_engine(options, &control, nullptr, nullptr);
    const models::CopyJobResult first_result = first_engine.run(cancel);
    check(first_result.Succeeded && first_result.BytesRead >= FileLength,
          "scan resume: initial assessment reads the source");

    storage::JobJournalStore journal_store;
    auto first_journal =
        journal_store.load(storage::fsutil::utf8_to_wide(first_result.JournalPath));
    const auto* first_entry = first_journal.has_value()
                                  ? first_journal->Files.find("payload.bin")
                                  : nullptr;
    check(first_entry != nullptr && first_entry->SourceChangeUtcTicks != 0 &&
              first_entry->SourceFileIndex != 0 &&
              first_entry->SourceVolumeSerial != 0,
          "scan resume: completed entry persists identity and change time");

    HANDLE mutation = CreateFileW(source_file.c_str(), GENERIC_WRITE, FILE_SHARE_READ, nullptr,
                                  OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, nullptr);
    bool mutated = false;
    if (mutation != INVALID_HANDLE_VALUE) {
        LARGE_INTEGER position{};
        position.QuadPart = FileLength / 2;
        unsigned char replacement = 0xA7;
        DWORD written = 0;
        mutated = SetFilePointerEx(mutation, position, nullptr, FILE_BEGIN) &&
                  WriteFile(mutation, &replacement, 1, &written, nullptr) && written == 1 &&
                  FlushFileBuffers(mutation);
        CloseHandle(mutation);
    }
    check(mutated && set_last_write_time_utc(source_file, original.last_write_utc_ticks),
          "scan resume: mutation restores size and last-write metadata");
    FileStabilitySnapshot changed;
    check(try_get_file_stability(source_file, changed) &&
              changed.change_utc_ticks != original.change_utc_ticks,
          "scan resume: NTFS change time exposes the in-place mutation");

    models::CopyJobOptions resumed = options;
    resumed.ResumeFromJournal = true;
    resumed.ResumeJournalPathHint = first_result.JournalPath;
    std::vector<std::string> resumed_logs;
    ResilientCopyEngine resumed_engine(
        resumed, &control, nullptr,
        [&resumed_logs](const std::string& message) { resumed_logs.push_back(message); });
    const models::CopyJobResult resumed_result = resumed_engine.run(cancel);
    bool binding_reset = false;
    for (const auto& line : resumed_logs) {
        if (line.find("Journal source binding changed") != std::string::npos) {
            binding_reset = true;
            break;
        }
    }
    check(resumed_result.Succeeded && resumed_result.BytesReused == 0 &&
              resumed_result.BytesRead >= FileLength,
          "scan resume: stale completion is reread instead of reused");
    check(binding_reset, "scan resume: stale source binding reset is visible in the log");

    resumed.ResumeJournalPathHint = resumed_result.JournalPath;
    ResilientCopyEngine stable_resume_engine(resumed, &control, nullptr, nullptr);
    const models::CopyJobResult stable_resume_result = stable_resume_engine.run(cancel);
    check(stable_resume_result.Succeeded &&
              stable_resume_result.BytesReused == FileLength &&
              stable_resume_result.BytesRead == 0,
          "scan resume: unchanged completed assessment is reused without rereading");

    remove_result_journal(stable_resume_result);
    remove_result_journal(resumed_result);
    remove_result_journal(first_result);
    remove_tree(work);
}

void test_engine_precise_scan_resumes_partial_coverage() {
    std::printf("--- engine: precise scan resumes bound partial coverage ---\n");
    std::wstring work = make_tree(0, 0);
    std::wstring source = work + L"\\source";
    storage::fsutil::create_directories(source);
    const std::wstring source_file = source + L"\\large.bin";
    constexpr std::int64_t FileLength = 4LL * 1024 * 1024;
    write_pattern_file(source_file, static_cast<std::size_t>(FileLength), 1920);

    models::CopyJobOptions options;
    options.OperationMode = models::JobOperationMode::ScanOnly;
    options.SourceRoot = storage::fsutil::wide_to_utf8(source);
    options.ScanPerformanceProfileValue = models::ScanPerformanceProfile::Precise;
    options.BufferSizeBytes = models::CopyJobOptions::MinimumBufferSizeBytes;
    options.UseAdaptiveBufferSizing = false;
    options.RescueFastScanChunkBytes = 64 * 1024;
    options.ParallelScanWorkers = 1;
    options.ResumeFromJournal = false;
    options.UseBadRangeMap = false;
    options.UpdateBadRangeMapFromRun = false;
    options.MaxRetries = 0;

    std::atomic<bool> cancel{false};
    bool cancellation_armed = false;
    ExecutionControl control;
    ResilientCopyEngine interrupted_engine(
        options, &control,
        [&cancel, &cancellation_armed](const models::CopyProgressSnapshot& snapshot) {
            if (!cancellation_armed && snapshot.CurrentFileBytesCopied >= 128 * 1024) {
                cancellation_armed = true;
                // Let the adaptive journal interval expire before the callback
                // returns; execute_rescue_pass checkpoints this chunk next.
                Sleep(1100);
                cancel.store(true);
            }
        },
        nullptr);
    bool cancelled = false;
    try {
        (void)interrupted_engine.run(cancel);
    } catch (const OperationCanceled& oc) {
        cancelled = oc.user_requested;
    }

    const std::wstring full_source = storage::fsutil::get_full_path(source);
    const std::string job_id = storage::JobJournalStore::build_job_id(
        full_source, full_source);
    const std::wstring journal_path =
        storage::JobJournalStore::get_default_journal_path(job_id);
    storage::JobJournalStore store;
    auto interrupted_journal = store.load(journal_path);
    const auto* interrupted_entry = interrupted_journal.has_value()
                                        ? interrupted_journal->Files.find("large.bin")
                                        : nullptr;
    const std::int64_t checkpointed = interrupted_entry != nullptr
                                          ? interrupted_entry->BytesCopied
                                          : 0;
    check(cancelled && checkpointed > 0 && checkpointed < FileLength,
          "precise scan interruption durably checkpoints partial file coverage");
    check(interrupted_entry != nullptr && interrupted_entry->SourceChangeUtcTicks != 0 &&
              interrupted_entry->SourceFileIndex != 0 &&
              interrupted_entry->SourceVolumeSerial != 0,
          "partial precise coverage is bound to the exact source generation");

    cancel.store(false);
    models::CopyJobOptions resumed = options;
    resumed.ResumeFromJournal = true;
    resumed.ResumeJournalPathHint = storage::fsutil::wide_to_utf8(journal_path);
    ResilientCopyEngine resumed_engine(resumed, &control, nullptr, nullptr);
    const models::CopyJobResult result = resumed_engine.run(cancel);
    check(result.Succeeded && result.BytesReused >= checkpointed &&
              result.BytesRead < FileLength,
          "precise scan resumes from checkpointed ranges instead of rereading the file");

    remove_result_journal(result);
    remove_tree(work);
}

void test_engine_fast_scan_cancellation_preserves_completed_files() {
    std::printf("--- engine: fast scan cancellation preserves completed files ---\n");
    std::wstring work = make_tree(0, 0);
    std::wstring source = work + L"\\source";
    storage::fsutil::create_directories(source);
    constexpr int FileCount = 8;
    constexpr std::int64_t FileLength = 512 * 1024;
    for (int index = 0; index < FileCount; ++index) {
        write_pattern_file(source + L"\\file" + std::to_wstring(index) + L".bin",
                           static_cast<std::size_t>(FileLength), 1930 + index);
    }

    models::CopyJobOptions options;
    options.OperationMode = models::JobOperationMode::ScanOnly;
    options.SourceRoot = storage::fsutil::wide_to_utf8(source);
    options.ScanPerformanceProfileValue = models::ScanPerformanceProfile::Fast;
    options.ParallelScanWorkers = 1;
    options.ResumeFromJournal = false;
    options.UseBadRangeMap = false;
    options.UpdateBadRangeMapFromRun = false;
    options.MaxRetries = 0;

    std::atomic<bool> cancel{false};
    ExecutionControl control;
    ResilientCopyEngine interrupted_engine(
        options, &control,
        [&cancel](const models::CopyProgressSnapshot& snapshot) {
            if (snapshot.CompletedFiles >= 2) cancel.store(true);
        },
        nullptr);
    bool cancelled = false;
    try {
        (void)interrupted_engine.run(cancel);
    } catch (const OperationCanceled& oc) {
        cancelled = oc.user_requested;
    }

    const std::wstring full_source = storage::fsutil::get_full_path(source);
    const std::string job_id = storage::JobJournalStore::build_job_id(
        full_source, full_source);
    const std::wstring journal_path =
        storage::JobJournalStore::get_default_journal_path(job_id);
    storage::JobJournalStore store;
    auto interrupted_journal = store.load(journal_path);
    std::int32_t completed = 0;
    if (interrupted_journal.has_value()) {
        for (const auto& item : interrupted_journal->Files.entries) {
            if (item.second.State == storage::FileCopyState::Completed ||
                item.second.State == storage::FileCopyState::CompletedWithRecovery) {
                ++completed;
            }
        }
    }
    check(cancelled && completed >= 2 && completed < FileCount,
          "fast scan cancellation writes completed entries before unwinding");
    check(interrupted_journal.has_value() && interrupted_journal->RunOptions.has_value() &&
              interrupted_journal->RunOptions->OperationMode ==
                  models::JobOperationMode::ScanOnly &&
              interrupted_journal->RunOptions->ScanPerformanceProfileValue ==
                  models::ScanPerformanceProfile::Fast,
          "fast scan journal retains the exact task definition");

    cancel.store(false);
    models::CopyJobOptions resumed = options;
    resumed.ResumeFromJournal = true;
    resumed.ResumeJournalPathHint = storage::fsutil::wide_to_utf8(journal_path);
    ResilientCopyEngine resumed_engine(resumed, &control, nullptr, nullptr);
    const models::CopyJobResult result = resumed_engine.run(cancel);
    check(result.Succeeded && result.BytesReused >= completed * FileLength &&
              result.BytesRead <= (FileCount - completed) * FileLength,
          "fast scan resumes with completed files reused rather than restarted");

    remove_result_journal(result);
    remove_tree(work);
}

void test_engine_rejects_unsafe_options() {
    std::printf("--- engine: unsafe option combinations are rejected ---\n");
    std::wstring work = make_tree(0, 0);
    std::wstring source = work + L"\\source";
    std::wstring destination = work + L"\\destination";
    storage::fsutil::create_directories(source);
    write_pattern_file(source + L"\\payload.bin", 4096, 908);

    models::CopyJobOptions options;
    options.SourceRoot = storage::fsutil::wide_to_utf8(source);
    options.DestinationRoot = storage::fsutil::wide_to_utf8(destination);
    options.ResumeFromJournal = false;

    std::atomic<bool> cancel{false};
    ExecutionControl control;
    bool root_remap_rejected = false;
    options.AllowJournalRootRemap = true;
    try {
        ResilientCopyEngine engine(options, &control, nullptr, nullptr);
        (void)engine.run(cancel);
    } catch (const IoError& ex) {
        root_remap_rejected = std::string(ex.what()).find("AllowJournalRootRemap") != std::string::npos;
    }
    check(root_remap_rejected, "journal root remap is rejected before copying");

    bool access_denied_salvage_rejected = false;
    options.AllowJournalRootRemap = false;
    options.TreatAccessDeniedAsContention = true;
    options.SalvageUnreadableBlocks = true;
    try {
        ResilientCopyEngine engine(options, &control, nullptr, nullptr);
        (void)engine.run(cancel);
    } catch (const IoError& ex) {
        access_denied_salvage_rejected =
            std::string(ex.what()).find("TreatAccessDeniedAsContention") != std::string::npos;
    }
    check(access_denied_salvage_rejected,
          "access-denied contention cannot be combined with salvage");

    bool nested_destination_rejected = false;
    options.TreatAccessDeniedAsContention = false;
    options.SalvageUnreadableBlocks = false;
    options.DestinationRoot = storage::fsutil::wide_to_utf8(source + L"\\nested-destination");
    try {
        ResilientCopyEngine engine(options, &control, nullptr, nullptr);
        (void)engine.run(cancel);
    } catch (const IoError& ex) {
        nested_destination_rejected =
            std::string(ex.what()).find("inside the source tree") != std::string::npos;
    }
    check(nested_destination_rejected,
          "destination inside source tree is rejected before enumeration");

    bool random_fill_rejected = false;
    options.DestinationRoot = storage::fsutil::wide_to_utf8(destination);
    options.SalvageUnreadableBlocks = true;
    options.SalvageFillPatternValue = models::SalvageFillPattern::Random;
    try {
        ResilientCopyEngine engine(options, &control, nullptr, nullptr);
        (void)engine.run(cancel);
    } catch (const IoError& ex) {
        random_fill_rejected =
            std::string(ex.what()).find("Random salvage fill") != std::string::npos;
    }
    check(random_fill_rejected,
          "non-deterministic random salvage fill is rejected before copying");

    bool excessive_retries_rejected = false;
    options.SalvageUnreadableBlocks = false;
    options.SalvageFillPatternValue = models::SalvageFillPattern::Zero;
    options.MaxRetries = 33;
    try {
        ResilientCopyEngine engine(options, &control, nullptr, nullptr);
        (void)engine.run(cancel);
    } catch (const IoError& ex) {
        excessive_retries_rejected =
            std::string(ex.what()).find("safety limit of 32") != std::string::npos;
    }
    check(excessive_retries_rejected,
          "retry amplification above the safety limit is rejected");

    options.MaxRetries = 2;
    auto current_options_rejected_with = [&](std::string_view expected) {
        try {
            ResilientCopyEngine engine(options, &control, nullptr, nullptr);
            (void)engine.run(cancel);
        } catch (const IoError& ex) {
            return std::string(ex.what()).find(expected) != std::string::npos;
        }
        return false;
    };

    options.SampleVerificationChunkCount = 65;
    check(current_options_rejected_with("cannot exceed 64"),
          "sample verification count is bounded before allocation");
    options.SampleVerificationChunkCount = 3;
    options.ParallelScanWorkers = 65;
    check(current_options_rejected_with("worker counts cannot exceed 64"),
          "explicit worker counts cannot exceed the supported concurrency bound");
    options.ParallelScanWorkers = 0;
    options.LockContentionProbeInterval = time::TimeSpan::from_milliseconds(1);
    check(current_options_rejected_with("100 to 10000 milliseconds"),
          "contention polling interval cannot become a busy loop");
    options.LockContentionProbeInterval = time::TimeSpan::from_milliseconds(500);
    options.MaxRetryDelay = time::TimeSpan::from_milliseconds(100);
    options.InitialRetryDelay = time::TimeSpan::from_milliseconds(250);
    check(current_options_rejected_with("shorter than InitialRetryDelay"),
          "retry backoff bounds reject an inverted delay policy");
    options.InitialRetryDelay = time::TimeSpan::from_milliseconds(250);
    options.MaxRetryDelay = time::TimeSpan::from_seconds(8);

    std::wstring alias = work + L"\\source-alias";
    constexpr DWORD AllowUnprivilegedSymlinkCreate = 0x2;
    if (CreateSymbolicLinkW(alias.c_str(), source.c_str(),
                            SYMBOLIC_LINK_FLAG_DIRECTORY | AllowUnprivilegedSymlinkCreate)) {
        bool physical_alias_rejected = false;
        options.DestinationRoot = storage::fsutil::wide_to_utf8(alias);
        try {
            ResilientCopyEngine engine(options, &control, nullptr, nullptr);
            (void)engine.run(cancel);
        } catch (const IoError& ex) {
            physical_alias_rejected =
                std::string(ex.what()).find("same physical path") != std::string::npos;
        }
        check(physical_alias_rejected,
              "destination alias resolving to the source is rejected before enumeration");
        RemoveDirectoryW(alias.c_str());
    } else {
        std::printf("  directory symlink unavailable (Win32 %lu); alias integration skipped\n",
                    static_cast<unsigned long>(GetLastError()));
    }
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

void test_engine_parallel_scan_wait_for_media_is_thread_safe() {
    std::printf("--- engine: parallel scan media-identity synchronization ---\n");
    constexpr int directories = 16;
    constexpr int files_per_directory = 128;
    std::wstring work = make_tree(directories, files_per_directory);

    models::CopyJobOptions options;
    options.SourceRoot = storage::fsutil::wide_to_utf8(work);
    // Scan jobs use the source as their non-writing destination identity. This
    // reproduces the whole-drive path that previously let every fast-scan
    // thread assign the same shared identity string concurrently.
    options.DestinationRoot = options.SourceRoot;
    options.OperationMode = models::JobOperationMode::ScanOnly;
    options.ResumeFromJournal = false;
    options.UseBadRangeMap = false;
    options.UpdateBadRangeMapFromRun = false;
    options.WaitForMediaAvailability = true;
    options.ScanPerformanceProfileValue = models::ScanPerformanceProfile::Fast;
    options.ParallelScanWorkers = 16;

    std::atomic<bool> cancel{false};
    ExecutionControl control;
    std::atomic<int> sixteen_worker_progress_events{0};
    auto started = std::chrono::steady_clock::now();
    ResilientCopyEngine engine(
        options, &control,
        [&sixteen_worker_progress_events](const models::CopyProgressSnapshot& snapshot) {
            if (snapshot.ScanWorkerCount == 16) sixteen_worker_progress_events.fetch_add(1);
        },
        nullptr);
    const models::CopyJobResult result = engine.run(cancel);
    double seconds = std::chrono::duration<double>(
                         std::chrono::steady_clock::now() - started)
                         .count();

    check(result.Succeeded, "16-worker scan with media waiting succeeds without heap corruption");
    check(result.TotalFiles == directories * files_per_directory,
          "16-worker scan reports every file");
    check(sixteen_worker_progress_events.load() > 0,
          "16-worker scan reports the configured parallelism");
    check(seconds < 15.0,
          "cached media identity checks avoid a per-file volume-query slowdown");
    remove_result_journal(result);
    remove_tree(work);
}

void test_engine_large_fast_scan_checkpoint_is_cancelable() {
    std::printf("--- engine: large fast scan starts workers before checkpoint stalls ---\n");
    std::wstring work = make_tree(8, 256); // Exactly the large-scan checkpoint threshold.
    std::wstring destination = work + L"\\unused-destination";

    models::CopyJobOptions options;
    options.SourceRoot = storage::fsutil::wide_to_utf8(work);
    options.DestinationRoot = storage::fsutil::wide_to_utf8(destination);
    options.OperationMode = models::JobOperationMode::ScanOnly;
    options.ScanPerformanceProfileValue = models::ScanPerformanceProfile::Fast;
    options.ParallelScanWorkers = 1;
    options.ResumeFromJournal = false;
    options.UseBadRangeMap = false;
    options.UpdateBadRangeMapFromRun = false;
    options.ContinueOnFileError = true;

    std::atomic<bool> cancel{false};
    bool saw_background_checkpoint = false;
    bool saw_fast_engine = false;
    std::vector<std::string> logs;
    ExecutionControl control;
    ResilientCopyEngine engine(
        options, &control, nullptr,
        [&logs, &cancel, &saw_background_checkpoint, &saw_fast_engine](const std::string& message) {
            logs.push_back(message);
            if (message.find("Large fast scan: initial checkpoint deferred to the live progress writer") !=
                std::string::npos) {
                saw_background_checkpoint = true;
            }
            if (message.find("Fast scan engine:") != std::string::npos) {
                saw_fast_engine = true;
                // Cancel at the hand-off point. This proves the worker pool is
                // reached and that cancellation is no longer trapped behind a
                // synchronous pre-scan journal write.
                cancel.store(true);
            }
        });

    bool cancelled = false;
    try {
        (void)engine.run(cancel);
    } catch (const OperationCanceled& oc) {
        cancelled = oc.user_requested;
    }

    check(saw_background_checkpoint,
          "large fast scan defers its initial checkpoint to the live progress writer");
    check(saw_fast_engine, "large fast scan reaches the worker pool");
    check(cancelled, "large fast scan cancellation reaches the engine");

    const std::string job_id = storage::JobJournalStore::build_job_id(
        storage::fsutil::get_full_path(work), storage::fsutil::get_full_path(destination));
    storage::JobJournalStore::remove_journal_set(
        storage::JobJournalStore::get_default_journal_path(job_id));
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
        if (line.find("Assessment backend: Raw volume") != std::string::npos) {
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

void test_engine_does_not_persist_bad_map_when_disabled() {
    std::printf("--- engine: bad-range map update flag is authoritative ---\n");
    std::wstring work = make_tree(0, 0);
    std::wstring source = work + L"\\source";
    std::wstring destination = work + L"\\unused-destination";
    storage::fsutil::create_directories(source);
    write_pattern_file(source + L"\\payload.bin", 32 * 1024, 1708);

    const std::wstring map_path = storage::BadRangeMapStore::get_default_map_path(
        storage::fsutil::normalize_path(source));
    storage::fsutil::delete_file(map_path);

    models::CopyJobOptions options;
    options.SourceRoot = storage::fsutil::wide_to_utf8(source);
    options.DestinationRoot = storage::fsutil::wide_to_utf8(destination);
    options.OperationMode = models::JobOperationMode::ScanOnly;
    options.UseBadRangeMap = true;
    options.SkipKnownBadRanges = true;
    options.UpdateBadRangeMapFromRun = false;
    options.ResumeFromJournal = false;
    options.ScanPerformanceProfileValue = models::ScanPerformanceProfile::Precise;

    std::atomic<bool> cancel{false};
    ExecutionControl control;
    ResilientCopyEngine engine(options, &control, nullptr, nullptr);
    const models::CopyJobResult result = engine.run(cancel);
    check(result.Succeeded, "scan remains successful with map updates disabled");
    check(GetFileAttributesW(map_path.c_str()) == INVALID_FILE_ATTRIBUTES,
          "disabled map updates do not create a map snapshot");
    remove_result_journal(result);
    storage::fsutil::delete_file(map_path);
    remove_tree(work);
}

void test_engine_requires_repeated_bad_map_confirmation() {
    std::printf("--- engine: bad-range hints require repeated confirmation ---\n");
    std::wstring work = make_tree(0, 0);
    std::wstring source = work + L"\\source";
    std::wstring destination = work + L"\\unused-destination";
    storage::fsutil::create_directories(source);
    write_pattern_file(source + L"\\payload.bin", 32 * 1024, 1831);

    SourceScanResult scan =
        scan_source(source, {}, models::SymlinkHandlingMode::Skip, true, nullptr);
    check(scan.files.size() == 1, "bad-map fixture source is enumerated");
    if (scan.files.size() != 1) {
        remove_tree(work);
        return;
    }
    const SourceFileDescriptor& descriptor = scan.files.front();
    FileStabilitySnapshot identity;
    check(try_get_file_stability(descriptor.full_path, identity),
          "bad-map fixture captures a stable file identity and change time");

    std::wstring normalized_source = storage::fsutil::get_full_path(source);
    std::wstring map_path =
        storage::BadRangeMapStore::get_default_map_path(normalized_source);
    remove_bad_range_map_set(map_path);

    char fingerprint[64];
    std::snprintf(fingerprint, sizeof(fingerprint), "%016llX:%016llX:%016llX",
                  static_cast<unsigned long long>(descriptor.length),
                  static_cast<unsigned long long>(descriptor.last_write_utc_ticks),
                  static_cast<unsigned long long>(identity.change_utc_ticks));
    storage::BadRangeMapFileEntry map_entry;
    map_entry.RelativePath = descriptor.relative_path;
    map_entry.SourceLength = descriptor.length;
    map_entry.LastWriteUtcTicks = descriptor.last_write_utc_ticks;
    map_entry.SourceFileIndex = static_cast<std::int64_t>(identity.identity.file_index);
    map_entry.SourceVolumeSerial = static_cast<std::int64_t>(identity.identity.volume_serial);
    map_entry.FileFingerprint = fingerprint;
    map_entry.BadRanges.push_back(storage::ByteRange{0, 4096});
    map_entry.ConfirmationCount = 1;
    map_entry.LastScanUtc = time::DateTimeOffset::now_utc();

    storage::BadRangeMap map;
    map.SchemaVersion = 1;
    map.SourceRoot = storage::fsutil::wide_to_utf8(normalized_source);
    map.SourceIdentity = storage::fsutil::resolve_media_identity(normalized_source);
    map.UpdatedUtc = time::DateTimeOffset::now_utc();
    map.Files.set(descriptor.relative_path, map_entry);
    storage::BadRangeMapStore map_store;
    map_store.save(map_path, map);

    models::CopyJobOptions options;
    options.SourceRoot = storage::fsutil::wide_to_utf8(source);
    options.DestinationRoot = storage::fsutil::wide_to_utf8(destination);
    options.OperationMode = models::JobOperationMode::ScanOnly;
    options.ResumeFromJournal = false;
    options.UseBadRangeMap = true;
    options.SkipKnownBadRanges = true;
    options.UpdateBadRangeMapFromRun = false;
    options.ScanPerformanceProfileValue = models::ScanPerformanceProfile::Precise;

    auto run_scan = [&](std::vector<std::string>& logs) {
        std::atomic<bool> cancel{false};
        ExecutionControl control;
        ResilientCopyEngine engine(
            options, &control, nullptr,
            [&logs](const std::string& message) { logs.push_back(message); });
        return engine.run(cancel);
    };
    auto has_applied_hint_log = [](const std::vector<std::string>& logs) {
        for (const auto& line : logs) {
            if (line.find("Applied bad-range map hints") != std::string::npos) return true;
        }
        return false;
    };

    std::vector<std::string> first_logs;
    models::CopyJobResult first_result = run_scan(first_logs);
    check(first_result.Succeeded && !has_applied_hint_log(first_logs),
          "one observation is retained for diagnostics but not trusted as a skip hint");
    remove_result_journal(first_result);

    map_entry.ConfirmationCount = 2;
    map.Files.set(descriptor.relative_path, map_entry);
    map_store.save(map_path, map);
    std::vector<std::string> confirmed_logs;
    models::CopyJobResult confirmed_result = run_scan(confirmed_logs);
    check(confirmed_result.Succeeded && has_applied_hint_log(confirmed_logs),
          "matching repeated observations enable the bad-range read hint");

    remove_result_journal(confirmed_result);

    Sleep(20);
    HANDLE mutate = CreateFileW(descriptor.full_path.c_str(), GENERIC_WRITE, FILE_SHARE_READ,
                                nullptr, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, nullptr);
    bool mutated = false;
    if (mutate != INVALID_HANDLE_VALUE) {
        const unsigned char replacement = 0xD7;
        DWORD written = 0;
        mutated = WriteFile(mutate, &replacement, 1, &written, nullptr) != FALSE && written == 1;
        FlushFileBuffers(mutate);
        CloseHandle(mutate);
    }
    check(mutated && set_last_write_time_utc(
                         descriptor.full_path, descriptor.last_write_utc_ticks),
          "bad-map mutation restores size and last-write metadata");
    std::vector<std::string> changed_logs;
    models::CopyJobResult changed_result = run_scan(changed_logs);
    check(changed_result.Succeeded && !has_applied_hint_log(changed_logs),
          "change-time mismatch rejects a stale physical-range hint");
    remove_result_journal(changed_result);

    remove_bad_range_map_set(map_path);
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

        // Replace the file with same-length content and restore its timestamp.
        // A cached extent map must not be reused just because size/time match.
        ExistingFileMetadata before_replacement;
        check(try_get_existing_file_metadata(file, before_replacement),
              "elevated raw test captures replacement metadata");
        std::wstring old_file = file + L".old";
        MoveFileExW(file.c_str(), old_file.c_str(), MOVEFILE_REPLACE_EXISTING);
        write_pattern_file(file, 128 * 1024, 1702);
        set_last_write_time_utc(file, before_replacement.last_write_utc_ticks);
        auto expected_replacement = storage::fsutil::read_all_bytes(file);
        std::vector<unsigned char> replacement(actual.size());
        RawDiskReadResult replacement_result = context->read_chunk(
            file, "payload.bin", 0, static_cast<std::int32_t>(replacement.size()), replacement.data(),
            10000, read_cancel);
        check(replacement_result.state == RawDiskReadResult::State::Success &&
                  replacement_result.BytesRead == replacement.size(),
              "elevated raw context rebuilds a same-metadata replacement layout");
        check(expected_replacement.has_value() && expected_replacement->size() >= replacement.size() &&
                  std::memcmp(expected_replacement->data(), replacement.data(), replacement.size()) == 0,
              "elevated raw replacement read matches new file bytes");
        DeleteFileW(old_file.c_str());

        // Exercise the sparse-hole path used by the raw extent planner. The
        // standard read is the oracle; the raw result must synthesize the same
        // zero-filled hole rather than returning stale volume bytes.
        std::wstring sparse_file = source + L"\\sparse.bin";
        write_pattern_file(sparse_file, 256 * 1024, 1703);
        HANDLE sparse_handle = CreateFileW(sparse_file.c_str(), GENERIC_READ | GENERIC_WRITE,
                                           FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE,
                                           nullptr, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, nullptr);
        bool sparse_ready = false;
        if (sparse_handle != INVALID_HANDLE_VALUE) {
            DWORD returned = 0;
            sparse_ready = DeviceIoControl(sparse_handle, FSCTL_SET_SPARSE, nullptr, 0,
                                            nullptr, 0, &returned, nullptr) != FALSE;
            FILE_ZERO_DATA_INFORMATION zero_range{};
            zero_range.FileOffset.QuadPart = 64 * 1024;
            zero_range.BeyondFinalZero.QuadPart = 192 * 1024;
            if (sparse_ready) {
                sparse_ready = DeviceIoControl(sparse_handle, FSCTL_SET_ZERO_DATA,
                                                &zero_range, sizeof(zero_range), nullptr, 0,
                                                &returned, nullptr) != FALSE;
            }
            CloseHandle(sparse_handle);
        }
        if (sparse_ready) {
            auto expected_sparse = storage::fsutil::read_all_bytes(sparse_file);
            std::vector<unsigned char> sparse_actual(256 * 1024);
            RawDiskReadResult sparse_result = context->read_chunk(
                sparse_file, "sparse.bin", 0, static_cast<std::int32_t>(sparse_actual.size()),
                sparse_actual.data(), 10000, read_cancel);
            check(sparse_result.state == RawDiskReadResult::State::Success &&
                      sparse_result.BytesRead == sparse_actual.size(),
                  "elevated raw context reads sparse file extents");
            check(expected_sparse.has_value() && *expected_sparse == sparse_actual,
                  "elevated raw sparse-hole bytes match standard reads");
        } else {
            std::printf("  sparse FSCTL integration skipped on this volume\n");
        }
    } else {
        std::printf("  raw context reason: %s\n", reason.c_str());
    }

    remove_tree(work);
}

} // namespace

int main() {
    test_scan_source_fast_and_complete();
    test_scan_source_cancellation();
    test_scan_source_partial_enumeration_is_visible();
    test_raw_read_plan_mapping();
    test_exact_item_selection();
    test_scan_symlink_scope_policy();
    test_file_stability_detects_same_size_timestamp_mutation();
    test_engine_copy_and_sampled_verification();
    test_engine_preserves_metadata_and_cleans_abandoned_stage();
    test_engine_reports_post_publish_metadata_failure_without_rolling_back_bytes();
    test_engine_efs_policy_when_supported();
    test_engine_reports_hard_link_topology_limit();
    test_engine_parallel_native_requires_full_verification();
    test_engine_refuses_ask_and_preserves_destination();
    test_engine_stages_destination_on_failure();
    test_engine_fault_injection_timeout_and_offline_are_safe();
    test_engine_fragile_first_read_findings_are_durable();
    test_engine_salvage_is_not_success();
    test_engine_rejects_unverified_journal_completion();
    test_engine_tracks_source_identity_changes();
    test_engine_scan_resume_binds_change_time();
    test_engine_precise_scan_resumes_partial_coverage();
    test_engine_fast_scan_cancellation_preserves_completed_files();
    test_engine_rejects_unsafe_options();
    test_engine_fast_scan_profile();
    test_engine_parallel_scan_wait_for_media_is_thread_safe();
    test_engine_large_fast_scan_checkpoint_is_cancelable();
    test_engine_raw_scan_backend_or_fallback();
    test_engine_does_not_persist_bad_map_when_disabled();
    test_engine_requires_repeated_bad_map_confirmation();
    test_raw_disk_elevated_integration();

    if (g_failures == 0) {
        std::printf("WORKER PASS: %d checks\n", g_checks);
        return 0;
    }
    std::printf("WORKER FAILED: %d of %d checks\n", g_failures, g_checks);
    return 1;
}
