// -----------------------------------------------------------------------------
// File: src\ui\run_history.h
// Purpose: The History pane's run actions, lifted from the retired Job Manager
//          dialog: open a run's authenticated journal as readable JSON, export
//          a CSV report, show its bad-range map, and build the run's detail
//          lines (including journal findings). Journal I/O runs on worker
//          threads that post completion messages to the owner window.
// -----------------------------------------------------------------------------

#pragma once

#ifndef NOMINMAX
#define NOMINMAX
#endif
#include <windows.h>
#include <shobjidl.h>

#include <atomic>
#include <filesystem>
#include <fstream>
#include <memory>
#include <optional>
#include <string>
#include <thread>
#include <vector>

#include "dialogs.h"
#include "elevation.h"
#include "job_manager.h"
#include "journal_report.h"
#include "range_map_dialog.h"

namespace xact::ui {

// Local "YYYY-MM-DD HH:MM:SS" for a catalog timestamp, or "-".
inline std::wstring format_local_time(const DateTimeOffset& value) {
    // .NET ticks (0001-01-01) -> FILETIME ticks (1601-01-01).
    constexpr std::int64_t FileTimeEpochTicks = 504911232000000000LL;
    std::int64_t utc_ticks = value.utc_ticks();
    if (utc_ticks < FileTimeEpochTicks) return L"-";
    std::int64_t file_ticks = utc_ticks - FileTimeEpochTicks;
    FILETIME utc_file_time;
    utc_file_time.dwLowDateTime = static_cast<DWORD>(file_ticks & 0xFFFFFFFF);
    utc_file_time.dwHighDateTime = static_cast<DWORD>(file_ticks >> 32);
    SYSTEMTIME utc_system{};
    if (!FileTimeToSystemTime(&utc_file_time, &utc_system)) return L"-";
    SYSTEMTIME local{};
    if (!SystemTimeToTzSpecificLocalTime(nullptr, &utc_system, &local)) return L"-";
    wchar_t buffer[32];
    swprintf(buffer, 32, L"%04u-%02u-%02u %02u:%02u:%02u", local.wYear, local.wMonth,
             local.wDay, local.wHour, local.wMinute, local.wSecond);
    return buffer;
}

class RunHistory {
public:
    // Posted to the owner; lParam carries the result, handed back to finish_*.
    static constexpr UINT ExportDoneMessage = WM_APP + 0x531;
    static constexpr UINT FindingsDoneMessage = WM_APP + 0x532;

    explicit RunHistory(JobManagerService& manager) : manager_(manager) {}
    ~RunHistory() { shutdown(); }
    RunHistory(const RunHistory&) = delete;
    RunHistory& operator=(const RunHistory&) = delete;

    void set_owner(HWND owner) { owner_ = owner; }

    // A journal export is in flight (one at a time).
    bool busy() const { return journal_worker_.joinable(); }

    // A run can be resumed when it stopped short and has journal state.
    bool can_resume(const storage::ManagedJobRun& run) const {
        using storage::ManagedJobRunStatus;
        if (run.Status == ManagedJobRunStatus::Completed ||
            run.Status == ManagedJobRunStatus::Running ||
            run.Status == ManagedJobRunStatus::Paused ||
            run.Status == ManagedJobRunStatus::Queued) {
            return false;
        }
        if (!storage::catalog_detail::is_blank(run.JournalPath)) return true;
        // Legacy paths are discovered by the findings worker, never on the
        // UI thread.
        return findings_cache_.has_value() && findings_cache_->Loaded &&
               models::detail::equals_ignore_case(findings_cache_row_id_, run.RunId) &&
               findings_cache_run_updated_ticks_ == run.LastUpdatedUtc.utc_ticks();
    }

    void open_journal(const std::string& run_id) {
        if (busy() || owner_ == nullptr) return;
        closing_ = false;
        HWND owner = owner_;
        journal_worker_ = std::thread([this, owner, run_id] {
            auto result = std::make_unique<ExportResult>();
            try {
                auto journal = manager_.load_run_journal(run_id);
                const auto path = manager_.get_resume_journal_path(run_id);
                if (!journal.has_value() || path.empty()) {
                    result->error = "No authenticated journal matching this run could be loaded.";
                } else {
                    const std::wstring file_name = storage::fsutil::get_file_name(
                        storage::fsutil::utf8_to_wide(path));
                    std::wstring base_name, extension;
                    storage::fsutil::split_extension(file_name, base_name, extension);
                    const std::wstring view_root =
                        storage::fsutil::local_app_data() + L"\\XactCopy\\journal-views";
                    result->readable_path = view_root + L"\\" + base_name + L"-readable.json";
                    storage::fsutil::create_directories(view_root);
                    json::Writer writer(/*indented*/ true);
                    journal->to_json(writer);
                    result->exported =
                        storage::fsutil::write_atomic_bytes(result->readable_path, writer.take());
                    if (!result->exported) result->error = "The readable journal could not be written.";
                }
            } catch (const std::exception& error) {
                result->error = error.what();
            }
            post(owner, ExportDoneMessage, std::move(result));
        });
    }

    // Asks for a destination, then writes the CSV report in the background.
    void export_report(const std::string& run_id) {
        if (busy() || owner_ == nullptr) return;
        const std::wstring destination = pick_report_path(owner_);
        if (destination.empty()) return;
        start_report_or_map(run_id, destination, /*map*/ false);
    }

    void show_range_map(const std::string& run_id) {
        if (busy() || owner_ == nullptr) return;
        start_report_or_map(run_id, std::wstring(), /*map*/ true);
    }

    // Handles ExportDoneMessage: reports the outcome or shows the result.
    void finish_export(LPARAM lparam, const ThemePalette& theme) {
        std::unique_ptr<ExportResult> result(reinterpret_cast<ExportResult*>(lparam));
        if (journal_worker_.joinable()) journal_worker_.join();
        if (result == nullptr || owner_ == nullptr) return;
        if (!result->exported) {
            MessageDialog::show(owner_, theme, L"XactCopy",
                                L"The journal could not be validated and exported:\n\n" +
                                    storage::fsutil::utf8_to_wide(result->error),
                                MessageIcon::Warning);
            return;
        }
        if (result->kind == ExportResult::Kind::RangeMap) {
            if (result->map_files.empty()) {
                MessageDialog::show(owner_, theme, L"XactCopy",
                                    L"This run's journal recorded no byte ranges, so there is no map to show.\n\n"
                                    L"Fast and direct-NTFS assessments record ranges only for files with "
                                    L"unreadable areas, and a file that failed as a whole (for example, one "
                                    L"that could not be opened) has none. The run's details list every "
                                    L"failed file.");
            } else {
                RangeMapDialog::show(owner_, theme, std::move(result->map_files));
            }
            return;
        }
        if (result->kind == ExportResult::Kind::Report) {
            MessageDialog::show(owner_, theme, L"XactCopy",
                                L"Report saved to:\n" + result->readable_path);
            return;
        }
        // As the signed-in user, so an elevated XactCopy does not open the
        // viewer or Explorer as Administrator.
        if (!elevation::open_as_user(owner_, result->readable_path)) {
            std::wstring arguments = L"/select,\"" + result->readable_path + L"\"";
            elevation::open_as_user(owner_, L"explorer.exe", arguments);
        }
    }

    // Detail lines for a run. Journal findings load in the background; until
    // they arrive the last line says so, and FindingsDoneMessage asks the
    // owner to call details() again.
    std::vector<std::wstring> details(const storage::ManagedJobRun& run) {
        std::vector<std::wstring> lines;
        auto line = [&lines](const wchar_t* label, const std::wstring& value) {
            lines.push_back(std::wstring(label) + (value.empty() ? L"-" : value));
        };
        auto wide = [](const std::string& text) { return storage::fsutil::utf8_to_wide(text); };
        line(L"Name: ", wide(run.DisplayName));
        line(L"Status: ", wide(std::string(storage::to_string(run.Status))));
        line(L"Trigger: ", wide(run.Trigger));
        line(L"Started: ", format_local_time(run.StartedUtc));
        line(L"Finished: ", run.FinishedUtc.has_value() ? format_local_time(*run.FinishedUtc)
                                                        : std::wstring());
        line(L"Source: ", wide(run.SourceRoot));
        line(L"Destination: ", wide(run.DestinationRoot));
        line(L"Summary: ", wide(run.Summary));
        line(L"Journal: ", wide(run.JournalPath));
        if (!run.ErrorMessage.empty()) line(L"Error: ", wide(run.ErrorMessage));
        if (run.Result.has_value() && !run.Result->IntegrityNotice.empty()) {
            line(L"Integrity notice: ", wide(run.Result->IntegrityNotice));
        }
        if (run.Result.has_value() && !run.Result->MetadataNotice.empty()) {
            line(L"Metadata notice: ", wide(run.Result->MetadataNotice));
        }
        append_findings(run, lines);
        return lines;
    }

    // Handles FindingsDoneMessage. Returns true when the shown run's details
    // should be rebuilt.
    bool finish_findings(LPARAM lparam, const storage::ManagedJobRun* shown_run) {
        std::unique_ptr<FindingsResult> result(reinterpret_cast<FindingsResult*>(lparam));
        if (findings_worker_.joinable()) findings_worker_.join();
        if (result == nullptr) return false;
        findings_loading_row_id_.clear();
        findings_cache_row_id_ = result->row_id;
        findings_cache_path_ = result->journal_path;
        findings_cache_run_updated_ticks_ =
            shown_run != nullptr && models::detail::equals_ignore_case(shown_run->RunId, result->row_id)
                ? shown_run->LastUpdatedUtc.utc_ticks()
                : 0;
        findings_cache_loaded_tick_ = GetTickCount64();
        findings_cache_ = std::move(result->inspection);
        return shown_run != nullptr &&
               models::detail::equals_ignore_case(shown_run->RunId, findings_cache_row_id_);
    }

    // Stops accepting results and waits for workers (window teardown).
    void shutdown() {
        closing_ = true;
        if (journal_worker_.joinable()) journal_worker_.join();
        if (findings_worker_.joinable()) findings_worker_.join();
        if (owner_ != nullptr) {
            MSG pending{};
            while (PeekMessageW(&pending, owner_, ExportDoneMessage, ExportDoneMessage, PM_REMOVE)) {
                delete reinterpret_cast<ExportResult*>(pending.lParam);
            }
            while (PeekMessageW(&pending, owner_, FindingsDoneMessage, FindingsDoneMessage, PM_REMOVE)) {
                delete reinterpret_cast<FindingsResult*>(pending.lParam);
            }
        }
    }

private:
    struct ExportResult {
        enum class Kind { OpenJournal, Report, RangeMap } kind = Kind::OpenJournal;
        bool exported = false;
        std::wstring readable_path;
        std::string error;
        std::vector<RangeMapFile> map_files;
    };

    struct FindingsResult {
        std::string row_id;
        std::string journal_path;
        RunJournalInspection inspection;
    };

    JobManagerService& manager_;
    HWND owner_ = nullptr;
    std::atomic<bool> closing_{false};
    std::thread journal_worker_;
    std::thread findings_worker_;
    std::optional<RunJournalInspection> findings_cache_;
    std::string findings_cache_row_id_;
    std::string findings_cache_path_;
    std::int64_t findings_cache_run_updated_ticks_ = 0;
    ULONGLONG findings_cache_loaded_tick_ = 0;
    std::string findings_loading_row_id_;

    template <typename Result>
    void post(HWND owner, UINT message, std::unique_ptr<Result> result) {
        if (!closing_.load() &&
            PostMessageW(owner, message, 0, reinterpret_cast<LPARAM>(result.get()))) {
            result.release();
        }
    }

    static std::wstring pick_report_path(HWND owner) {
        std::wstring destination;
        IFileSaveDialog* dialog = nullptr;
        if (FAILED(CoCreateInstance(CLSID_FileSaveDialog, nullptr, CLSCTX_INPROC_SERVER,
                                    IID_PPV_ARGS(&dialog)))) {
            return destination;
        }
        const COMDLG_FILTERSPEC filters[] = {{L"CSV files", L"*.csv"}, {L"All files", L"*.*"}};
        dialog->SetFileTypes(2, filters);
        dialog->SetDefaultExtension(L"csv");
        dialog->SetFileName(L"XactCopy-report.csv");
        if (SUCCEEDED(dialog->Show(owner))) {
            IShellItem* item = nullptr;
            if (SUCCEEDED(dialog->GetResult(&item))) {
                PWSTR path = nullptr;
                if (SUCCEEDED(item->GetDisplayName(SIGDN_FILESYSPATH, &path))) {
                    destination = path;
                    CoTaskMemFree(path);
                }
                item->Release();
            }
        }
        dialog->Release();
        return destination;
    }

    void start_report_or_map(const std::string& run_id, const std::wstring& destination, bool map) {
        closing_ = false;
        HWND owner = owner_;
        journal_worker_ = std::thread([this, owner, run_id, destination, map] {
            auto result = std::make_unique<ExportResult>();
            result->kind = map ? ExportResult::Kind::RangeMap : ExportResult::Kind::Report;
            try {
                auto journal = manager_.load_run_journal(run_id);
                if (!journal.has_value()) {
                    result->error = "No authenticated journal matching this run could be loaded.";
                } else {
                    // Older journals may omit options even when the signed run
                    // record retains them; those options disambiguate legacy ranges.
                    if (!journal->RunOptions.has_value()) {
                        const auto run = manager_.get_run_by_id(run_id);
                        if (run.has_value() && run->Options.has_value()) journal->RunOptions = run->Options;
                    }
                    if (map) {
                        for (auto& [key, entry] : journal->Files.entries) {
                            if (entry.RescueRanges.empty() && !entry.RecoveredRanges.empty() &&
                                journal->RunOptions.has_value()) {
                                const auto state =
                                    journal->RunOptions->OperationMode == models::JobOperationMode::ScanOnly
                                        ? storage::RescueRangeState::Bad
                                        : storage::RescueRangeState::Recovered;
                                for (const auto& range : entry.RecoveredRanges) {
                                    entry.RescueRanges.push_back({range.Offset, range.Length, state});
                                }
                            }
                            if (entry.RescueRanges.empty()) continue;
                            RangeMapFile file;
                            file.path = storage::fsutil::utf8_to_wide(
                                entry.RelativePath.empty() ? key : entry.RelativePath);
                            file.size = entry.SourceLength;
                            file.ranges = std::move(entry.RescueRanges);
                            result->map_files.push_back(std::move(file));
                        }
                        std::stable_partition(result->map_files.begin(), result->map_files.end(),
                                              [](const RangeMapFile& file) { return file.has_findings(); });
                        result->exported = true;
                    } else {
                        const auto journal_path = manager_.get_resume_journal_path(run_id);
                        const auto normalized_output = storage::fsutil::normalize_path(destination);
                        const auto normalized_journal = storage::fsutil::normalize_path(
                            storage::fsutil::utf8_to_wide(journal_path));
                        if (!journal_path.empty() &&
                            (normalized_output == normalized_journal ||
                             normalized_output.starts_with(normalized_journal + L"."))) {
                            result->error = "The report destination is part of the journal set.";
                        } else {
                            std::ofstream stream(std::filesystem::path(destination),
                                                 std::ios::binary | std::ios::trunc);
                            if (!stream) throw std::runtime_error("The report file could not be opened.");
                            write_journal_report_csv(stream, *journal);
                            stream.close();
                            if (!stream) throw std::runtime_error("The report could not be fully written.");
                            result->exported = true;
                            result->readable_path = destination;
                        }
                    }
                }
            } catch (const std::exception& error) {
                result->error = error.what();
            }
            post(owner, ExportDoneMessage, std::move(result));
        });
    }

    void append_findings(const storage::ManagedJobRun& run, std::vector<std::wstring>& lines) {
        if (storage::catalog_detail::is_blank(run.JournalPath) &&
            storage::catalog_detail::is_blank(run.SourceRoot)) {
            lines.push_back(L"Findings: No resumable journal could be located for this run.");
            return;
        }
        const std::string journal_path = storage::catalog_detail::is_blank(run.JournalPath)
                                             ? "legacy:" + run.SourceRoot + "|" + run.DestinationRoot
                                             : run.JournalPath;
        const bool cached =
            findings_cache_.has_value() &&
            models::detail::equals_ignore_case(findings_cache_row_id_, run.RunId) &&
            models::detail::equals_ignore_case(findings_cache_path_, journal_path) &&
            findings_cache_run_updated_ticks_ == run.LastUpdatedUtc.utc_ticks() &&
            ((run.Status != storage::ManagedJobRunStatus::Running &&
              run.Status != storage::ManagedJobRunStatus::Paused) ||
             GetTickCount64() - findings_cache_loaded_tick_ < 10000);
        if (!cached) {
            if (!findings_worker_.joinable() && owner_ != nullptr) {
                closing_ = false;
                findings_loading_row_id_ = run.RunId;
                HWND owner = owner_;
                findings_worker_ = std::thread([this, owner, run_id = run.RunId, path = journal_path] {
                    auto result = std::make_unique<FindingsResult>();
                    result->row_id = run_id;
                    result->journal_path = path;
                    result->inspection = manager_.inspect_run_journal(run_id);
                    post(owner, FindingsDoneMessage, std::move(result));
                });
            }
            lines.push_back(models::detail::equals_ignore_case(findings_loading_row_id_, run.RunId)
                                ? L"Findings: Loading authenticated journal details..."
                                : L"Findings: Waiting for the selected journal inspection...");
            return;
        }

        const RunJournalInspection& inspection = *findings_cache_;
        if (!inspection.Loaded) {
            lines.push_back(L"Findings unavailable: " +
                            storage::fsutil::utf8_to_wide(inspection.ErrorMessage));
            return;
        }
        lines.push_back(
            L"Journal progress: " + std::to_wstring(inspection.CompletedFiles) + L"/" +
            std::to_wstring(inspection.TotalFiles) + L" completed; " +
            std::to_wstring(inspection.PendingFiles) + L" pending/in progress; " +
            std::to_wstring(inspection.FailedFiles) + L" failed; " +
            storage::fsutil::utf8_to_wide(JobManagerService::format_bytes(inspection.ProcessedBytes)) +
            L" of " +
            storage::fsutil::utf8_to_wide(JobManagerService::format_bytes(inspection.TotalBytes)) +
            L" checkpointed.");
        if (inspection.Findings.empty()) {
            lines.push_back(L"Findings: No unreadable or failed files are recorded in this checkpoint.");
            return;
        }
        lines.push_back((inspection.ScanMode ? L"Assessment findings (" : L"Copy/recovery findings (") +
                        std::to_wstring(inspection.Findings.size()) + L"):");
        for (const auto& finding : inspection.Findings) {
            std::wstring prefix = finding.UnreadableRangeCount > 0
                                      ? (inspection.ScanMode ? L"UNREADABLE: " : L"RECOVERED/UNREADABLE: ")
                                      : L"FAILED: ";
            std::wstring detail = prefix + storage::fsutil::utf8_to_wide(finding.RelativePath);
            if (finding.UnreadableRangeCount > 0) {
                detail += L" \x2014 " + std::to_wstring(finding.UnreadableRangeCount) + L" range(s), " +
                          storage::fsutil::utf8_to_wide(
                              JobManagerService::format_bytes(finding.UnreadableBytes));
            }
            if (!finding.ErrorMessage.empty()) {
                detail += L" \x2014 " + storage::fsutil::utf8_to_wide(finding.ErrorMessage);
            }
            lines.push_back(std::move(detail));
        }
    }
};

} // namespace xact::ui
