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
#include <string>
#include <thread>

#include "../src/ui/selection.h"
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

} // namespace

int main() {
    test_scan_source_fast_and_complete();
    test_scan_source_cancellation();
    test_exact_item_selection();

    if (g_failures == 0) {
        std::printf("WORKER PASS: %d checks\n", g_checks);
        return 0;
    }
    std::printf("WORKER FAILED: %d of %d checks\n", g_failures, g_checks);
    return 1;
}
