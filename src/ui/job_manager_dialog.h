// Job Manager dialog (port of XactCopy.UI JobManagerForm.vb) plus the small
// TextPromptDialog used by save-as-job / rename / duplicate. The console shows
// saved jobs, queue entries, and run history in one grid with view / status /
// search filters, a details pane, and the full action set. "Run Now" closes
// the dialog with a requested action the main window then dispatches.
#pragma once

#include <commctrl.h>
#include <shellapi.h>

#include "dialogs.h"
#include "job_manager.h"

namespace xact::ui {

// ---------------------------------------------------------------------------
// TextPromptDialog — one-line text input (mirrors .NET TextPromptDialog).
// ---------------------------------------------------------------------------

class TextPromptDialog {
public:
    static std::optional<std::string> show(HWND owner, const ThemePalette& theme,
                                           const std::wstring& title, const std::wstring& prompt,
                                           const std::string& initial_value) {
        TextPromptDialog dialog(theme, prompt, initial_value);
        dialog.host_.set_density_percent(theme.density_percent);
        dialog.host_.set_scale_percent(theme.scale_percent);
        dialog.host_.create(owner, L"XactCopyTextPrompt", title.c_str(), 380, 150,
                            [&dialog](HWND h, UINT m, WPARAM w, LPARAM l, bool& handled) {
                                return dialog.proc(h, m, w, l, handled);
                            });
        dialog.host_.run_modal();
        if (dialog.font_ != nullptr) DeleteObject(dialog.font_);
        if (dialog.window_brush_ != nullptr) DeleteObject(dialog.window_brush_);
        if (dialog.edit_brush_ != nullptr) DeleteObject(dialog.edit_brush_);
        return dialog.result_;
    }

private:
    static constexpr int IdEdit = 2001;
    static constexpr int IdOk = 2002;
    static constexpr int IdCancel = 2003;

    TextPromptDialog(const ThemePalette& theme, std::wstring prompt, std::string initial)
        : theme_(theme), prompt_(std::move(prompt)), initial_(std::move(initial)) {}

    LRESULT proc(HWND hwnd, UINT message, WPARAM wparam, LPARAM lparam, bool& handled) {
        switch (message) {
            case WM_CREATE: {
                set_dark_title_bar(hwnd, theme_.dark && theme_.themed_chrome);
                apply_window_icons(hwnd);
                window_brush_ = CreateSolidBrush(theme_.window);
                edit_brush_ = CreateSolidBrush(theme_.edit);
                UINT dpi = ui_layout_dpi(GetDpiForWindow(hwnd), theme_.density_percent,
                                         theme_.scale_percent);
                auto scale = [dpi](int value) { return MulDiv(value, static_cast<int>(dpi), 96); };
                font_ = CreateFontW(-MulDiv(9, static_cast<int>(dpi), 72), 0, 0, 0, FW_NORMAL,
                                    FALSE, FALSE, FALSE, DEFAULT_CHARSET, OUT_DEFAULT_PRECIS,
                                    CLIP_DEFAULT_PRECIS, CLEARTYPE_QUALITY, DEFAULT_PITCH,
                                    L"Segoe UI");
                RECT client;
                GetClientRect(hwnd, &client);
                HWND label = CreateWindowExW(0, L"STATIC", prompt_.c_str(),
                                             WS_CHILD | WS_VISIBLE | SS_NOPREFIX, scale(12),
                                             scale(10), client.right - scale(24), scale(18), hwnd,
                                             nullptr, GetModuleHandleW(nullptr), nullptr);
                SendMessageW(label, WM_SETFONT, reinterpret_cast<WPARAM>(font_), TRUE);
                edit_ = CreateWindowExW(0, L"EDIT",
                                        storage::fsutil::utf8_to_wide(initial_).c_str(),
                                        WS_CHILD | WS_VISIBLE | WS_TABSTOP | ES_AUTOHSCROLL |
                                            WS_BORDER,
                                        scale(12), scale(32), client.right - scale(24), scale(24),
                                        hwnd, reinterpret_cast<HMENU>(static_cast<INT_PTR>(IdEdit)),
                                        GetModuleHandleW(nullptr), nullptr);
                SendMessageW(edit_, WM_SETFONT, reinterpret_cast<WPARAM>(font_), TRUE);
                apply_control_theme(edit_, theme_.dark);
                SendMessageW(edit_, EM_SETSEL, 0, -1);
                auto make_button = [&](int id, const wchar_t* text, int x) {
                    HWND button = CreateWindowExW(
                        0, L"BUTTON", text, WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_OWNERDRAW, x,
                        client.bottom - scale(36), scale(84), scale(26), hwnd,
                        reinterpret_cast<HMENU>(static_cast<INT_PTR>(id)),
                        GetModuleHandleW(nullptr), nullptr);
                    SendMessageW(button, WM_SETFONT, reinterpret_cast<WPARAM>(font_), TRUE);
                };
                make_button(IdOk, L"OK", client.right - scale(188));
                make_button(IdCancel, L"Cancel", client.right - scale(96));
                SetFocus(edit_);
                handled = true;
                return 0;
            }
            case WM_ERASEBKGND: {
                RECT client;
                GetClientRect(hwnd, &client);
                themedraw::fill_rect(reinterpret_cast<HDC>(wparam), client, theme_.window);
                handled = true;
                return 1;
            }
            case WM_CTLCOLORSTATIC: {
                HDC dc = reinterpret_cast<HDC>(wparam);
                SetTextColor(dc, theme_.text);
                SetBkColor(dc, theme_.window);
                handled = true;
                return reinterpret_cast<LRESULT>(window_brush_);
            }
            case WM_CTLCOLOREDIT: {
                HDC dc = reinterpret_cast<HDC>(wparam);
                SetTextColor(dc, theme_.text);
                SetBkColor(dc, theme_.edit);
                handled = true;
                return reinterpret_cast<LRESULT>(edit_brush_);
            }
            case WM_DRAWITEM:
                themedraw::draw_button(*reinterpret_cast<DRAWITEMSTRUCT*>(lparam), theme_);
                handled = true;
                return TRUE;
            case WM_COMMAND: {
                int id = LOWORD(wparam);
                if (id == IdOk || (id == IDOK && HIWORD(wparam) == 0)) {
                    wchar_t buffer[512]{};
                    GetWindowTextW(edit_, buffer, 511);
                    result_ = storage::fsutil::wide_to_utf8(buffer);
                    DestroyWindow(hwnd);
                    handled = true;
                    return 0;
                }
                if (id == IdCancel || id == IDCANCEL) {
                    DestroyWindow(hwnd);
                    handled = true;
                    return 0;
                }
                return 0;
            }
            case WM_CLOSE:
                DestroyWindow(hwnd);
                handled = true;
                return 0;
            case WM_DESTROY:
                handled = true;
                return 0;
            default:
                return 0;
        }
        (void)lparam;
    }

    dialog_detail::ModalHost host_;
    ThemePalette theme_;
    std::wstring prompt_;
    std::string initial_;
    std::optional<std::string> result_;
    HWND edit_ = nullptr;
    HFONT font_ = nullptr;
    HBRUSH window_brush_ = nullptr;
    HBRUSH edit_brush_ = nullptr;
};

// ---------------------------------------------------------------------------
// Job Manager dialog
// ---------------------------------------------------------------------------

enum class JobManagerRequestAction { None, RunSavedJob, RunQueuedEntry };

struct JobManagerRequest {
    JobManagerRequestAction action = JobManagerRequestAction::None;
    std::string job_id;
    std::string queue_entry_id;
};

class JobManagerDialog {
public:
    static JobManagerRequest show(HWND owner, JobManagerService& manager, const ThemePalette& theme,
                                  AppSettings& settings) {
        JobManagerDialog dialog(manager, theme, settings);
        int width = settings.get_int("JobManagerWidth", 700, 20000, 1080);
        int height = settings.get_int("JobManagerHeight", 460, 20000, 640);
        dialog.host_.set_density_percent(theme.density_percent);
        dialog.host_.set_scale_percent(theme.scale_percent);
        dialog.host_.create(owner, L"XactCopyJobManager", L"Job Manager", width, height,
                            [&dialog](HWND h, UINT m, WPARAM w, LPARAM l, bool& handled) {
                                return dialog.proc(h, m, w, l, handled);
                            },
                            /*resizable*/ true);
        dialog.host_.run_modal();
        dialog.destroy_resources();
        return dialog.request_;
    }

private:
    enum class RowKind { SavedJob, QueueEntry, RunHistory };

    struct Row {
        RowKind kind = RowKind::SavedJob;
        std::string primary_id; // job id / queue entry id / run id
        std::string job_id;
        storage::ManagedJobRunStatus run_status = storage::ManagedJobRunStatus::Queued;
        std::wstring cells[10]; // Type, Name, State, Queue, Trigger, Started, Updated, Source, Destination, Summary
        std::string state_utf8;
    };

    static constexpr int IdList = 2101;
    static constexpr int IdViewCombo = 2102;
    static constexpr int IdStatusCombo = 2103;
    static constexpr int IdSearchEdit = 2104;
    static constexpr int IdRefresh = 2105;
    static constexpr int IdRunNow = 2110;
    static constexpr int IdQueue = 2111;
    static constexpr int IdRemoveQueue = 2112;
    static constexpr int IdMoveTop = 2113;
    static constexpr int IdMoveUp = 2114;
    static constexpr int IdMoveDown = 2115;
    static constexpr int IdMoveBottom = 2116;
    static constexpr int IdRename = 2117;
    static constexpr int IdDuplicate = 2118;
    static constexpr int IdDeleteJob = 2119;
    static constexpr int IdDeleteRun = 2120;
    static constexpr int IdOpenJournal = 2121;
    static constexpr int IdClearQueue = 2122;
    static constexpr int IdClearHistory = 2123;
    static constexpr int IdClose = 2124;
    static constexpr int IdDetailsList = 2130;
    static constexpr UINT_PTR AutoRefreshTimerId = 1;
    static constexpr UINT_PTR SearchDebounceTimerId = 2;

    JobManagerDialog(JobManagerService& manager, const ThemePalette& theme, AppSettings& settings)
        : manager_(manager), theme_(theme), settings_(settings) {
        grid_alternating_rows_ = settings_.get_bool("GridAlternatingRows", true);
        grid_row_height_ = settings_.get_int("GridRowHeight", 18, 48, 24);
        grid_header_style_ = settings_.get_string("GridHeaderStyle", "default");
    }

    void save_placement(HWND hwnd) {
        RECT r;
        if (!GetWindowRect(hwnd, &r)) return;
        settings_.set_int("JobManagerWidth", r.right - r.left);
        settings_.set_int("JobManagerHeight", r.bottom - r.top);
        settings_.save();
    }

    void destroy_resources() {
        if (font_ != nullptr) DeleteObject(font_);
        if (title_font_ != nullptr) DeleteObject(title_font_);
        if (window_brush_ != nullptr) DeleteObject(window_brush_);
        if (edit_brush_ != nullptr) DeleteObject(edit_brush_);
    }

    // .NET grid time format: local "yyyy-MM-dd HH:mm:ss".
    static std::wstring format_local_time(const DateTimeOffset& value) {
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

    static bool contains_ignore_case(const std::string& haystack, const std::string& needle) {
        if (needle.empty()) return true;
        if (haystack.size() < needle.size()) return false;
        auto fold = [](unsigned char c) {
            return (c >= 'A' && c <= 'Z') ? static_cast<unsigned char>(c - 'A' + 'a') : c;
        };
        for (std::size_t i = 0; i + needle.size() <= haystack.size(); ++i) {
            bool match = true;
            for (std::size_t j = 0; j < needle.size(); ++j) {
                if (fold(static_cast<unsigned char>(haystack[i + j])) !=
                    fold(static_cast<unsigned char>(needle[j]))) {
                    match = false;
                    break;
                }
            }
            if (match) return true;
        }
        return false;
    }

    template <typename... Fields>
    static bool matches_search(const std::string& search, const Fields&... fields) {
        if (search.empty()) return true;
        return (contains_ignore_case(fields, search) || ...);
    }

    static std::string build_saved_job_summary(const storage::ManagedJob& job) {
        std::string mode = job.Options.UseAdaptiveBufferSizing ? "Adaptive auto buffer" : "Manual buffer";
        std::string verify = job.Options.VerifyAfterCopy ? "Verify" : "No verify";
        return mode + "; retries " + std::to_string(std::max(0, job.Options.MaxRetries)) + "; " + verify + ".";
    }

    static std::string build_queue_summary(const JobQueueEntryView& entry) {
        std::string base = entry.EnqueuedBy.empty() ? std::string("Queued") : "Queued by " + entry.EnqueuedBy;
        if (entry.LastErrorMessage.empty()) return base;
        return base + "; last error: " + entry.LastErrorMessage;
    }

    void refresh_data(bool keep_selection) {
        saved_jobs_ = manager_.get_jobs();
        queue_entries_ = manager_.get_queue_entries();
        runs_ = manager_.get_recent_runs(500);
        populate_grid(keep_selection);

        std::wstring summary = L"Saved: " + std::to_wstring(saved_jobs_.size()) + L" | Queued: " +
                               std::to_wstring(queue_entries_.size()) + L" | Runs: " +
                               std::to_wstring(runs_.size());
        SetWindowTextW(summary_label_, summary.c_str());
    }

    // View filter: 0 All Items, 1 Saved Jobs, 2 Queue, 3 Run History.
    // Status filter index 0 = all, otherwise ManagedJobRunStatus by ordinal.
    void populate_grid(bool keep_selection) {
        std::string previous_primary;
        RowKind previous_kind = RowKind::SavedJob;
        if (keep_selection) {
            if (const Row* selected = selected_row()) {
                previous_primary = selected->primary_id;
                previous_kind = selected->kind;
            }
        }

        int view = static_cast<int>(SendMessageW(view_combo_, CB_GETCURSEL, 0, 0));
        int status_index = static_cast<int>(SendMessageW(status_combo_, CB_GETCURSEL, 0, 0));
        wchar_t search_buffer[256]{};
        GetWindowTextW(search_edit_, search_buffer, 255);
        std::string search =
            storage::catalog_detail::trim_copy(storage::fsutil::wide_to_utf8(search_buffer));

        rows_.clear();
        auto wide = [](const std::string& text) { return storage::fsutil::utf8_to_wide(text); };

        if (view == 0 || view == 1) {
            for (const auto& job : saved_jobs_) {
                std::string summary = build_saved_job_summary(job);
                if (!matches_search(search, std::string("saved"), job.Name, std::string("ready"),
                                    job.Options.SourceRoot, job.Options.DestinationRoot, summary)) {
                    continue;
                }
                Row row;
                row.kind = RowKind::SavedJob;
                row.primary_id = job.JobId;
                row.job_id = job.JobId;
                row.state_utf8 = "Ready";
                row.cells[0] = L"Saved";
                row.cells[1] = wide(job.Name);
                row.cells[2] = L"Ready";
                row.cells[3] = L"-";
                row.cells[4] = L"manual";
                row.cells[5] = format_local_time(job.CreatedUtc);
                row.cells[6] = format_local_time(job.UpdatedUtc);
                row.cells[7] = wide(job.Options.SourceRoot);
                row.cells[8] = wide(job.Options.DestinationRoot);
                row.cells[9] = wide(summary);
                rows_.push_back(std::move(row));
            }
        }

        if (view == 0 || view == 2) {
            for (const auto& entry : queue_entries_) {
                std::string state = entry.AttemptCount > 0 ? "Retry queued" : "Queued";
                std::string summary = build_queue_summary(entry);
                if (!matches_search(search, std::string("queue"), entry.JobName, state,
                                    entry.SourceRoot, entry.DestinationRoot, summary, entry.Trigger)) {
                    continue;
                }
                Row row;
                row.kind = RowKind::QueueEntry;
                row.primary_id = entry.QueueEntryId;
                row.job_id = entry.JobId;
                row.state_utf8 = state;
                row.cells[0] = L"Queue";
                row.cells[1] = wide(entry.JobName);
                row.cells[2] = wide(state);
                row.cells[3] = std::to_wstring(entry.Position);
                row.cells[4] = wide(entry.Trigger);
                row.cells[5] = format_local_time(entry.EnqueuedUtc);
                row.cells[6] = format_local_time(entry.LastUpdatedUtc);
                row.cells[7] = wide(entry.SourceRoot);
                row.cells[8] = wide(entry.DestinationRoot);
                row.cells[9] = wide(summary);
                rows_.push_back(std::move(row));
            }
        }

        if (view == 0 || view == 3) {
            for (const auto& run : runs_) {
                if (status_index > 0 &&
                    run.Status != static_cast<storage::ManagedJobRunStatus>(status_index - 1)) {
                    continue;
                }
                std::string status_text(storage::to_string(run.Status));
                std::string summary = run.Summary.empty() ? std::string("-") : run.Summary;
                if (!matches_search(search, std::string("run"), run.DisplayName, status_text,
                                    run.SourceRoot, run.DestinationRoot, summary, run.Trigger)) {
                    continue;
                }
                Row row;
                row.kind = RowKind::RunHistory;
                row.primary_id = run.RunId;
                row.job_id = run.JobId;
                row.run_status = run.Status;
                row.state_utf8 = status_text;
                row.cells[0] = L"Run";
                row.cells[1] = wide(run.DisplayName);
                row.cells[2] = wide(status_text);
                row.cells[3] = run.QueueAttempt > 0 ? std::to_wstring(run.QueueAttempt) : L"-";
                row.cells[4] = wide(run.Trigger);
                row.cells[5] = format_local_time(run.StartedUtc);
                row.cells[6] = format_local_time(run.LastUpdatedUtc);
                row.cells[7] = wide(run.SourceRoot);
                row.cells[8] = wide(run.DestinationRoot);
                row.cells[9] = wide(summary);
                rows_.push_back(std::move(row));
            }
        }

        populating_ = true;
        ListView_DeleteAllItems(list_);
        for (std::size_t i = 0; i < rows_.size(); ++i) {
            LVITEMW item{};
            item.mask = LVIF_TEXT;
            item.iItem = static_cast<int>(i);
            item.pszText = const_cast<wchar_t*>(rows_[i].cells[0].c_str());
            ListView_InsertItem(list_, &item);
            for (int column = 1; column < 10; ++column) {
                ListView_SetItemText(list_, static_cast<int>(i), column,
                                     const_cast<wchar_t*>(rows_[i].cells[column].c_str()));
            }
        }

        int target_index = -1;
        if (!previous_primary.empty()) {
            for (std::size_t i = 0; i < rows_.size(); ++i) {
                if (rows_[i].kind == previous_kind &&
                    models::detail::equals_ignore_case(rows_[i].primary_id, previous_primary)) {
                    target_index = static_cast<int>(i);
                    break;
                }
            }
        }
        if (target_index < 0 && !rows_.empty()) target_index = 0;
        if (target_index >= 0) {
            ListView_SetItemState(list_, target_index, LVIS_SELECTED | LVIS_FOCUSED,
                                  LVIS_SELECTED | LVIS_FOCUSED);
            ListView_EnsureVisible(list_, target_index, FALSE);
        }
        populating_ = false;

        update_action_states();
        render_selected_details();
    }

    const Row* selected_row() const {
        int index = ListView_GetNextItem(list_, -1, LVNI_SELECTED);
        if (index < 0 || static_cast<std::size_t>(index) >= rows_.size()) return nullptr;
        return &rows_[static_cast<std::size_t>(index)];
    }

    void update_action_states() {
        const Row* row = selected_row();
        auto enable = [](HWND button, bool enabled) {
            if (button != nullptr) EnableWindow(button, enabled ? TRUE : FALSE);
        };
        bool is_saved = row != nullptr && row->kind == RowKind::SavedJob;
        bool is_queue = row != nullptr && row->kind == RowKind::QueueEntry;
        bool is_run = row != nullptr && row->kind == RowKind::RunHistory;
        enable(buttons_[IdRunNow], is_saved || is_queue);
        enable(buttons_[IdQueue], is_saved || (is_run && !row->job_id.empty()));
        enable(buttons_[IdRemoveQueue], is_queue);
        enable(buttons_[IdMoveTop], is_queue);
        enable(buttons_[IdMoveUp], is_queue);
        enable(buttons_[IdMoveDown], is_queue);
        enable(buttons_[IdMoveBottom], is_queue);
        enable(buttons_[IdRename], is_saved);
        enable(buttons_[IdDuplicate], is_saved);
        enable(buttons_[IdDeleteJob], is_saved);
        enable(buttons_[IdDeleteRun], is_run);
        enable(buttons_[IdOpenJournal], is_run);
        enable(buttons_[IdClearQueue], !queue_entries_.empty());
        enable(buttons_[IdClearHistory], !runs_.empty());
    }

    void render_selected_details() {
        // Build the text first and bail out when it is unchanged: the 3 s
        // auto-refresh would otherwise reset the listbox every tick, throwing
        // away the user's scroll position mid-read.
        std::vector<std::wstring> lines;
        auto add = [&lines](const std::wstring& s) { lines.push_back(s); };
        const Row* row = selected_row();
        if (row == nullptr) {
            add(L"No item selected.");
        } else {
            auto line = [&add](const wchar_t* label, const std::wstring& value) {
                add(std::wstring(label) + (value.empty() ? L"-" : value));
            };
            line(L"Type: ", row->cells[0]);
            line(L"Name: ", row->cells[1]);
            line(L"State: ", row->cells[2]);
            line(L"Trigger: ", row->cells[4]);
            line(L"Started/Created: ", row->cells[5]);
            line(L"Updated: ", row->cells[6]);
            line(L"Source: ", row->cells[7]);
            line(L"Destination: ", row->cells[8]);
            line(L"Summary: ", row->cells[9]);
            if (row->kind == RowKind::RunHistory) {
                for (const auto& run : runs_) {
                    if (!models::detail::equals_ignore_case(run.RunId, row->primary_id)) continue;
                    line(L"Journal: ", storage::fsutil::utf8_to_wide(run.JournalPath));
                    if (!run.ErrorMessage.empty()) {
                        line(L"Error: ", storage::fsutil::utf8_to_wide(run.ErrorMessage));
                    }
                    break;
                }
            }
        }
        if (lines == details_lines_) return; // nothing changed; leave the view alone

        const bool same_item = (details_row_id_ == (row != nullptr ? row->primary_id : std::string()));
        const int top_index =
            same_item ? static_cast<int>(SendMessageW(details_list_, LB_GETTOPINDEX, 0, 0)) : 0;

        details_lines_ = lines;
        details_row_id_ = row != nullptr ? row->primary_id : std::string();

        SendMessageW(details_list_, WM_SETREDRAW, FALSE, 0);
        SendMessageW(details_list_, LB_RESETCONTENT, 0, 0);
        for (const auto& text : details_lines_) {
            SendMessageW(details_list_, LB_ADDSTRING, 0, reinterpret_cast<LPARAM>(text.c_str()));
        }
        // Restore the scroll offset when the same item is still selected.
        if (top_index > 0) SendMessageW(details_list_, LB_SETTOPINDEX, top_index, 0);
        SendMessageW(details_list_, WM_SETREDRAW, TRUE, 0);
        InvalidateRect(details_list_, nullptr, TRUE);
    }

    void run_now_selected(HWND hwnd) {
        const Row* row = selected_row();
        if (row == nullptr) return;
        switch (row->kind) {
            case RowKind::SavedJob:
                request_.action = JobManagerRequestAction::RunSavedJob;
                request_.job_id = row->job_id;
                request_.queue_entry_id.clear();
                DestroyWindow(hwnd);
                return;
            case RowKind::QueueEntry:
                request_.action = JobManagerRequestAction::RunQueuedEntry;
                request_.queue_entry_id = row->primary_id;
                request_.job_id = row->job_id;
                DestroyWindow(hwnd);
                return;
            default:
                info_box(hwnd, L"Select a saved job or queue entry to run.");
                return;
        }
    }

    void queue_selected(HWND hwnd) {
        const Row* row = selected_row();
        if (row == nullptr) return;
        bool queued = false;
        switch (row->kind) {
            case RowKind::SavedJob:
                queued = manager_.queue_job(row->job_id, false, "job-manager");
                break;
            case RowKind::RunHistory:
                if (row->job_id.empty()) {
                    info_box(hwnd,
                             L"This run is not linked to a saved job, so it cannot be queued.");
                    return;
                }
                queued = manager_.queue_job(row->job_id, false, "rerun-history");
                break;
            default:
                info_box(hwnd, L"Queue action applies to saved jobs and saved-job run entries.");
                return;
        }
        if (!queued) {
            info_box(hwnd, L"Unable to queue this item. It may already be queued.");
        }
        refresh_data(true);
    }

    void open_journal_selected(HWND hwnd) {
        const Row* row = selected_row();
        if (row == nullptr || row->kind != RowKind::RunHistory) return;
        for (const auto& run : runs_) {
            if (!models::detail::equals_ignore_case(run.RunId, row->primary_id)) continue;
            if (run.JournalPath.empty()) {
                info_box(hwnd, L"This run has no journal path recorded.");
                return;
            }
            std::wstring journal_path = storage::fsutil::utf8_to_wide(run.JournalPath);
            if (storage::fsutil::file_exists(journal_path)) {
                std::wstring arguments = L"/select,\"" + journal_path + L"\"";
                ShellExecuteW(hwnd, L"open", L"explorer.exe", arguments.c_str(), nullptr,
                              SW_SHOWNORMAL);
            } else {
                info_box(hwnd, L"The journal file no longer exists on disk.");
            }
            return;
        }
    }

    void on_command(HWND hwnd, int id, int code) {
        if (id == IdClose || id == IDCANCEL) {
            DestroyWindow(hwnd);
            return;
        }
        if (id == IdRefresh) {
            refresh_data(true);
            return;
        }
        if ((id == IdViewCombo || id == IdStatusCombo) && code == CBN_SELCHANGE) {
            populate_grid(true);
            return;
        }
        if (id == IdSearchEdit && code == EN_CHANGE && !populating_) {
            SetTimer(hwnd, SearchDebounceTimerId, 150, nullptr);
            return;
        }
        if (id == IdRunNow) { run_now_selected(hwnd); return; }
        if (id == IdQueue) { queue_selected(hwnd); return; }
        if (id == IdRemoveQueue) {
            if (const Row* row = selected_row(); row != nullptr && row->kind == RowKind::QueueEntry) {
                manager_.remove_queued_job(row->primary_id);
                refresh_data(true);
            }
            return;
        }
        if (id == IdMoveTop || id == IdMoveUp || id == IdMoveDown || id == IdMoveBottom) {
            if (const Row* row = selected_row(); row != nullptr && row->kind == RowKind::QueueEntry) {
                QueueMoveDirection direction = id == IdMoveTop      ? QueueMoveDirection::Top
                                               : id == IdMoveUp     ? QueueMoveDirection::Up
                                               : id == IdMoveDown   ? QueueMoveDirection::Down
                                                                    : QueueMoveDirection::Bottom;
                manager_.move_queue_entry(row->primary_id, direction);
                refresh_data(true);
            }
            return;
        }
        if (id == IdRename) {
            if (const Row* row = selected_row(); row != nullptr && row->kind == RowKind::SavedJob) {
                std::string current = storage::fsutil::wide_to_utf8(row->cells[1]);
                auto name = TextPromptDialog::show(hwnd, theme_, L"Rename Job",
                                                   L"New name for this job:", current);
                if (name.has_value() && !storage::catalog_detail::is_blank(*name)) {
                    manager_.rename_job(row->job_id, *name);
                    refresh_data(true);
                }
            }
            return;
        }
        if (id == IdDuplicate) {
            if (const Row* row = selected_row(); row != nullptr && row->kind == RowKind::SavedJob) {
                std::string suggested = storage::fsutil::wide_to_utf8(row->cells[1]) + " Copy";
                auto name = TextPromptDialog::show(hwnd, theme_, L"Duplicate Job",
                                                   L"Name for the duplicated job:", suggested);
                if (name.has_value() && !storage::catalog_detail::is_blank(*name)) {
                    manager_.duplicate_job(row->job_id, *name);
                    refresh_data(true);
                }
            }
            return;
        }
        if (id == IdDeleteJob) {
            if (const Row* row = selected_row(); row != nullptr && row->kind == RowKind::SavedJob) {
                std::wstring prompt = L"Delete saved job \"" + row->cells[1] +
                                      L"\"?\n\nQueued entries for this job are removed too.";
                if (confirm_box(hwnd, prompt, MessageIcon::Warning)) {
                    manager_.delete_job(row->job_id);
                    refresh_data(false);
                }
            }
            return;
        }
        if (id == IdDeleteRun) {
            if (const Row* row = selected_row(); row != nullptr && row->kind == RowKind::RunHistory) {
                if (confirm_box(hwnd, L"Delete this run history entry?")) {
                    manager_.delete_run(row->primary_id);
                    refresh_data(false);
                }
            }
            return;
        }
        if (id == IdOpenJournal) { open_journal_selected(hwnd); return; }
        if (id == IdClearQueue) {
            if (confirm_box(hwnd, L"Remove every entry from the queue?")) {
                manager_.clear_queue();
                refresh_data(false);
            }
            return;
        }
        if (id == IdClearHistory) {
            if (confirm_box(hwnd, L"Clear the entire run history?")) {
                manager_.clear_run_history();
                refresh_data(false);
            }
            return;
        }
    }

    void on_create(HWND hwnd) {
        set_dark_title_bar(hwnd, theme_.dark && theme_.themed_chrome);
        window_brush_ = CreateSolidBrush(theme_.window);
        edit_brush_ = CreateSolidBrush(theme_.edit);
        UINT dpi = ui_layout_dpi(GetDpiForWindow(hwnd), theme_.density_percent,
                                 theme_.scale_percent);
        auto scale = [dpi](int value) { return MulDiv(value, static_cast<int>(dpi), 96); };
        font_ = CreateFontW(-MulDiv(9, static_cast<int>(dpi), 72), 0, 0, 0, FW_NORMAL, FALSE,
                            FALSE, FALSE, DEFAULT_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS,
                            CLEARTYPE_QUALITY, DEFAULT_PITCH, L"Segoe UI");
        title_font_ = CreateFontW(-MulDiv(12, static_cast<int>(dpi), 72), 0, 0, 0, FW_SEMIBOLD,
                                  FALSE, FALSE, FALSE, DEFAULT_CHARSET, OUT_DEFAULT_PRECIS,
                                  CLIP_DEFAULT_PRECIS, CLEARTYPE_QUALITY, DEFAULT_PITCH,
                                  L"Segoe UI");

        RECT client;
        GetClientRect(hwnd, &client);
        int margin = scale(12);

        title_label_ = CreateWindowExW(0, L"STATIC", L"Jobs Console",
                                       WS_CHILD | WS_VISIBLE | SS_NOPREFIX, margin, scale(8),
                                       scale(220), scale(22), hwnd, nullptr,
                                       GetModuleHandleW(nullptr), nullptr);
        SendMessageW(title_label_, WM_SETFONT, reinterpret_cast<WPARAM>(title_font_), TRUE);

        summary_label_ = CreateWindowExW(0, L"STATIC", L"Saved: 0 | Queued: 0 | Runs: 0",
                                         WS_CHILD | WS_VISIBLE | SS_NOPREFIX | SS_RIGHT,
                                         client.right - scale(320) - margin, scale(12), scale(320),
                                         scale(18), hwnd, nullptr, GetModuleHandleW(nullptr),
                                         nullptr);
        SendMessageW(summary_label_, WM_SETFONT, reinterpret_cast<WPARAM>(font_), TRUE);

        // Filter row.
        int filter_y = scale(36);
        auto make_label = [&](const wchar_t* text, int x, int width) {
            HWND label = CreateWindowExW(0, L"STATIC", text,
                                         WS_CHILD | WS_VISIBLE | SS_NOPREFIX, x, filter_y + scale(5),
                                         width, scale(18), hwnd, nullptr,
                                         GetModuleHandleW(nullptr), nullptr);
            SendMessageW(label, WM_SETFONT, reinterpret_cast<WPARAM>(font_), TRUE);
        };
        make_label(L"View", margin, scale(34));
        view_combo_ = CreateWindowExW(0, L"COMBOBOX", L"",
                                      WS_CHILD | WS_VISIBLE | WS_TABSTOP | CBS_DROPDOWNLIST |
                                          CBS_OWNERDRAWFIXED | CBS_HASSTRINGS,
                                      margin + scale(38), filter_y, scale(120), scale(200), hwnd,
                                      reinterpret_cast<HMENU>(static_cast<INT_PTR>(IdViewCombo)),
                                      GetModuleHandleW(nullptr), nullptr);
        for (const wchar_t* option : {L"All Items", L"Saved Jobs", L"Queue", L"Run History"}) {
            SendMessageW(view_combo_, CB_ADDSTRING, 0, reinterpret_cast<LPARAM>(option));
        }
        SendMessageW(view_combo_, CB_SETCURSEL, 0, 0);

        make_label(L"Run Status", margin + scale(172), scale(64));
        status_combo_ = CreateWindowExW(0, L"COMBOBOX", L"",
                                        WS_CHILD | WS_VISIBLE | WS_TABSTOP | CBS_DROPDOWNLIST |
                                            CBS_OWNERDRAWFIXED | CBS_HASSTRINGS,
                                        margin + scale(240), filter_y, scale(120), scale(220), hwnd,
                                        reinterpret_cast<HMENU>(static_cast<INT_PTR>(IdStatusCombo)),
                                        GetModuleHandleW(nullptr), nullptr);
        SendMessageW(status_combo_, CB_ADDSTRING, 0, reinterpret_cast<LPARAM>(L"All Statuses"));
        for (const auto& entry : storage::ManagedJobRunStatus_Table) {
            SendMessageW(status_combo_, CB_ADDSTRING, 0,
                         reinterpret_cast<LPARAM>(
                             storage::fsutil::utf8_to_wide(std::string(entry.name)).c_str()));
        }
        SendMessageW(status_combo_, CB_SETCURSEL, 0, 0);

        make_label(L"Search", margin + scale(376), scale(44));
        search_edit_ = CreateWindowExW(0, L"EDIT", L"",
                                       WS_CHILD | WS_VISIBLE | WS_TABSTOP | ES_AUTOHSCROLL |
                                           WS_BORDER,
                                       margin + scale(424), filter_y, scale(220), scale(24), hwnd,
                                       reinterpret_cast<HMENU>(static_cast<INT_PTR>(IdSearchEdit)),
                                       GetModuleHandleW(nullptr), nullptr);

        HWND refresh = CreateWindowExW(0, L"BUTTON", L"Refresh",
                                       WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_OWNERDRAW,
                                       margin + scale(656), filter_y, scale(76), scale(24), hwnd,
                                       reinterpret_cast<HMENU>(static_cast<INT_PTR>(IdRefresh)),
                                       GetModuleHandleW(nullptr), nullptr);
        SendMessageW(refresh, WM_SETFONT, reinterpret_cast<WPARAM>(font_), TRUE);
        buttons_[IdRefresh] = refresh;

        // Grid + details.
        int list_top = filter_y + scale(34);
        int footer_height = scale(78);
        int details_height = scale(110);
        int list_height = client.bottom - list_top - footer_height - details_height - scale(16);
        // LVS_OWNERDRAWFIXED: the control emits WM_DRAWITEM per row and draws no
        // item text itself, so the common-control theme cannot paint black item
        // text on a Light OS (NM_CUSTOMDRAW alone didn't reliably override it
        // on screen). The shared framework keeps the scrollbars dark.
        list_ = CreateWindowExW(0, WC_LISTVIEWW, L"",
                                WS_CHILD | WS_VISIBLE | WS_TABSTOP | LVS_REPORT | LVS_SINGLESEL |
                                    LVS_SHOWSELALWAYS | LVS_OWNERDRAWFIXED,
                                margin, list_top, client.right - margin * 2, list_height, hwnd,
                                reinterpret_cast<HMENU>(static_cast<INT_PTR>(IdList)),
                                GetModuleHandleW(nullptr), nullptr);
        ListView_SetExtendedListViewStyle(list_, LVS_EX_FULLROWSELECT | LVS_EX_DOUBLEBUFFER |
                                                     LVS_EX_HEADERDRAGDROP);
        // Rows/header are still fully owner-drawn on top for per-cell colours.
        apply_control_theme(list_, theme_.dark);
        // Same Segoe UI 9pt as the rest of the dialog (the ListView and its
        // header otherwise fall back to the default system font).
        SendMessageW(list_, WM_SETFONT, reinterpret_cast<WPARAM>(font_), TRUE);
        if (HWND header = ListView_GetHeader(list_)) {
            SendMessageW(header, WM_SETFONT, reinterpret_cast<WPARAM>(font_), TRUE);
            wimukthi::win32_theme::apply_theme_class(header, L"DarkMode_ItemsView");
        }
        ListView_SetBkColor(list_, theme_.edit);
        ListView_SetTextBkColor(list_, theme_.edit);
        ListView_SetTextColor(list_, theme_.text);
        SetWindowSubclass(list_, &list_subclass_proc, 1, reinterpret_cast<DWORD_PTR>(this));

        struct ColumnSpec { const wchar_t* name; int width; };
        const ColumnSpec columns[] = {
            {L"Type", 56},   {L"Name", 150},        {L"State", 88},   {L"Queue", 52},
            {L"Trigger", 92}, {L"Started", 128},    {L"Updated", 128}, {L"Source", 150},
            {L"Destination", 150}, {L"Summary", 210},
        };
        for (int i = 0; i < 10; ++i) {
            LVCOLUMNW column{};
            column.mask = LVCF_TEXT | LVCF_WIDTH;
            column.pszText = const_cast<wchar_t*>(columns[i].name);
            column.cx = scale(columns[i].width);
            ListView_InsertColumn(list_, i, &column);
        }

        details_label_ = CreateWindowExW(0, L"STATIC", L"Selected Item Details",
                                         WS_CHILD | WS_VISIBLE | SS_NOPREFIX, margin,
                                         list_top + list_height + scale(6), scale(220),
                                         scale(16), hwnd, nullptr, GetModuleHandleW(nullptr),
                                         nullptr);
        SendMessageW(details_label_, WM_SETFONT, reinterpret_cast<WPARAM>(font_), TRUE);

        // Owner-drawn listbox (one line per detail field). A listbox's scrollbar
        // darkens under DarkMode_Explorer (an EDIT's does not), so this gives a
        // dark, scrollable details pane.
        details_list_ = CreateWindowExW(
            0, L"LISTBOX", L"",
            WS_CHILD | WS_VISIBLE | WS_VSCROLL | WS_BORDER | LBS_NOINTEGRALHEIGHT |
                LBS_DISABLENOSCROLL | LBS_NOSEL | LBS_OWNERDRAWFIXED | LBS_HASSTRINGS,
            margin, list_top + list_height + scale(24), client.right - margin * 2,
            details_height - scale(24), hwnd,
            reinterpret_cast<HMENU>(static_cast<INT_PTR>(IdDetailsList)), GetModuleHandleW(nullptr),
            nullptr);
        SendMessageW(details_list_, WM_SETFONT, reinterpret_cast<WPARAM>(font_), TRUE);
        SendMessageW(details_list_, LB_SETITEMHEIGHT, 0, scale(18));
        apply_control_theme(details_list_, theme_.dark);

        // Footer actions — positions/wrapping are set by relayout().
        for (const auto& spec : footer_buttons()) {
            HWND button = CreateWindowExW(
                0, L"BUTTON", spec.text, WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_OWNERDRAW, 0, 0, 10,
                10, hwnd, reinterpret_cast<HMENU>(static_cast<INT_PTR>(spec.id)),
                GetModuleHandleW(nullptr), nullptr);
            SendMessageW(button, WM_SETFONT, reinterpret_cast<WPARAM>(font_), TRUE);
            buttons_[spec.id] = button;
        }

        for (HWND control : {view_combo_, status_combo_, search_edit_}) {
            SendMessageW(control, WM_SETFONT, reinterpret_cast<WPARAM>(font_), TRUE);
        }
        apply_combo_theme(view_combo_, theme_.dark);
        apply_combo_theme(status_combo_, theme_.dark);
        apply_control_theme(search_edit_, theme_.dark);
        relayout(hwnd);

        refresh_data(false);
        SetTimer(hwnd, AutoRefreshTimerId, 3000, nullptr);
    }

    struct FooterButton { int id; const wchar_t* text; int width; };
    // Display order; relayout() flows these left-to-right and wraps to as many
    // rows as the window width needs, so they never cluster in a corner.
    static const std::vector<FooterButton>& footer_buttons() {
        static const std::vector<FooterButton> buttons = {
            {IdRunNow, L"Run Now", 84},        {IdQueue, L"Queue", 66},
            {IdRemoveQueue, L"Remove Queue", 108}, {IdMoveTop, L"Top", 48},
            {IdMoveUp, L"Up", 44},             {IdMoveDown, L"Down", 56},
            {IdMoveBottom, L"Bottom", 62},     {IdRename, L"Rename", 70},
            {IdDuplicate, L"Duplicate", 80},   {IdDeleteJob, L"Delete Job", 84},
            {IdDeleteRun, L"Delete Run", 86},  {IdOpenJournal, L"Open Journal", 100},
            {IdClearQueue, L"Clear Queue", 92}, {IdClearHistory, L"Clear History", 96},
            {IdClose, L"Close", 66},
        };
        return buttons;
    }

    // Repositions the resizable regions (list fills the middle, details above
    // the footer, footer pinned to the bottom, summary right-aligned) so the
    // window can be freely resized.
    void relayout(HWND hwnd) {
        if (list_ == nullptr) return;
        UINT dpi = ui_layout_dpi(GetDpiForWindow(hwnd), theme_.density_percent,
                                 theme_.scale_percent);
        auto scale = [dpi](int v) { return MulDiv(v, static_cast<int>(dpi), 96); };
        RECT client;
        GetClientRect(hwnd, &client);
        const int margin = scale(12);
        const int client_right = static_cast<int>(client.right);
        const int client_bottom = static_cast<int>(client.bottom);
        const int content_left = margin;
        const int content_right = client_right - margin;
        const int content_width = std::max(scale(320), content_right - content_left);
        const int button_height = scale(26);
        const int button_gap = scale(8);
        const int row_gap = scale(6);

        MoveWindow(summary_label_, client_right - scale(320) - margin, scale(12), scale(320),
                   scale(18), TRUE);

        // Flow the footer buttons to learn how many rows the current width needs
        // (so the footer grows/shrinks and the list gets the remaining space).
        int footer_rows_needed = 1;
        {
            int x = content_left;
            for (const auto& spec : footer_buttons()) {
                int w = scale(spec.width);
                if (x > content_left && x + w > content_right) {
                    ++footer_rows_needed;
                    x = content_left;
                }
                x += w + button_gap;
            }
        }
        const int footer_height =
            footer_rows_needed * button_height + (footer_rows_needed - 1) * row_gap + scale(12);

        const int list_top = scale(36) + scale(34);
        const int details_height = scale(140);
        const int list_height = std::max(
            scale(80), client_bottom - list_top - footer_height - details_height - scale(16));

        MoveWindow(list_, margin, list_top, content_width, list_height, TRUE);
        MoveWindow(details_label_, margin, list_top + list_height + scale(6), scale(220), scale(16),
                   TRUE);
        MoveWindow(details_list_, margin, list_top + list_height + scale(24), content_width,
                   details_height - scale(24), TRUE);
        stretch_last_column();

        int fx = content_left;
        int fy = client_bottom - footer_height + scale(6);
        for (const auto& spec : footer_buttons()) {
            int w = scale(spec.width);
            if (fx > content_left && fx + w > content_right) {
                fy += button_height + row_gap;
                fx = content_left;
            }
            if (HWND button = buttons_[spec.id]) {
                MoveWindow(button, fx, fy, w, button_height, TRUE);
            }
            fx += w + button_gap;
        }
        InvalidateRect(hwnd, nullptr, FALSE);
    }

    // Grow the Summary column to consume any spare width after a resize.
    void stretch_last_column() {
        if (list_ == nullptr) return;
        RECT lr;
        GetClientRect(list_, &lr);
        int used = 0;
        for (int i = 0; i < 9; ++i) used += ListView_GetColumnWidth(list_, i);
        int last = lr.right - used - GetSystemMetrics(SM_CXVSCROLL) - 4;
        UINT dpi = ui_layout_dpi(GetDpiForWindow(list_), theme_.density_percent,
                                 theme_.scale_percent);
        if (last > MulDiv(120, static_cast<int>(dpi), 96)) {
            ListView_SetColumnWidth(list_, 9, last);
        }
    }

    // Paints one grid row (LVS_OWNERDRAWFIXED WM_DRAWITEM): themed background
    // (selection when selected) and each column's text in the row's state color
    // (State column) or normal text.
    void draw_list_row(const DRAWITEMSTRUCT& draw) {
        const std::size_t index = static_cast<std::size_t>(draw.itemID);
        HDC dc = draw.hDC;
        const bool selected = (draw.itemState & ODS_SELECTED) != 0;

        RECT row_rect = draw.rcItem;
        const COLORREF alternate = blend_color(theme_.edit, theme_.panel, 35);
        const COLORREF bg = selected ? theme_.selection
                                     : (grid_alternating_rows_ && (index & 1) ? alternate
                                                                              : theme_.edit);
        themedraw::fill_rect(dc, row_rect, bg);

        if (index >= rows_.size()) return;
        const Row& row = rows_[index];
        const COLORREF normal = selected ? theme_.selection_text : theme_.text;
        const COLORREF state = selected ? theme_.selection_text : state_color(row);

        HGDIOBJ old_font = SelectObject(dc, font_);
        SetBkMode(dc, TRANSPARENT);
        // LVS_EX_GRIDLINES is ignored for owner-drawn lists, so draw the grid
        // for populated rows explicitly. The empty tail is repainted by the
        // ListView subclass below to prevent native column rules continuing
        // below the last item.
        const COLORREF grid = blend_color(theme_.edit, theme_.border, 70);
        const int column_count = Header_GetItemCount(ListView_GetHeader(list_));
        for (int col = 0; col < column_count && col < 10; ++col) {
            RECT cell;
            if (col == 0) {
                ListView_GetItemRect(list_, static_cast<int>(index), &cell, LVIR_BOUNDS);
                cell.right = cell.left + ListView_GetColumnWidth(list_, 0);
            } else {
                ListView_GetSubItemRect(list_, static_cast<int>(index), col, LVIR_BOUNDS, &cell);
            }
            SetTextColor(dc, col == 2 ? state : normal);
            RECT text_rect = cell;
            text_rect.left += 6;
            text_rect.right -= 4;
            DrawTextW(dc, row.cells[col].c_str(), -1, &text_rect,
                      DT_LEFT | DT_VCENTER | DT_SINGLELINE | DT_END_ELLIPSIS | DT_NOPREFIX);
            if (col > 0) {
                RECT divider{cell.left, row_rect.top, cell.left + 1, row_rect.bottom};
                themedraw::fill_rect(dc, divider, grid);
            }
        }
        RECT underline{row_rect.left, row_rect.bottom - 1, row_rect.right, row_rect.bottom};
        themedraw::fill_rect(dc, underline, grid);
        SelectObject(dc, old_font);
    }

    void draw_details_item(const DRAWITEMSTRUCT& draw) {
        themedraw::fill_rect(draw.hDC, draw.rcItem, theme_.edit);
        if (draw.itemID == static_cast<UINT>(-1)) return;
        const LRESULT length = SendMessageW(details_list_, LB_GETTEXTLEN, draw.itemID, 0);
        if (length < 0) return;
        std::wstring text(static_cast<std::size_t>(length) + 1, L'\0');
        SendMessageW(details_list_, LB_GETTEXT, draw.itemID, reinterpret_cast<LPARAM>(text.data()));
        text.resize(static_cast<std::size_t>(length));
        HGDIOBJ old_font = SelectObject(draw.hDC, font_);
        SetBkMode(draw.hDC, TRANSPARENT);
        SetTextColor(draw.hDC, theme_.text);
        RECT rect = draw.rcItem;
        rect.left += 6;
        rect.right -= 6;
        DrawTextW(draw.hDC, text.c_str(), -1, &rect,
                  DT_LEFT | DT_VCENTER | DT_SINGLELINE | DT_END_ELLIPSIS | DT_NOPREFIX);
        SelectObject(draw.hDC, old_font);
    }

    // Themed message boxes (dark-mode replacements for MessageBoxW).
    void info_box(HWND owner, const std::wstring& text) {
        MessageDialog::show(owner, theme_, L"Job Manager", text, MessageIcon::Information);
    }

    bool confirm_box(HWND owner, const std::wstring& text,
                     MessageIcon icon = MessageIcon::Question) {
        return MessageDialog::show(owner, theme_, L"Job Manager", text, icon,
                                   MessageButtons::YesNo) == IDYES;
    }

    COLORREF state_color(const Row& row) const {
        // Mirrors JobManagerForm.MainGrid_CellFormatting, lightened for dark.
        COLORREF color;
        const std::string& state = row.state_utf8;
        if (state == "Failed") color = RGB(200, 40, 40);
        else if (state == "Running") color = RGB(20, 140, 60);
        else if (state == "Completed") color = RGB(30, 100, 180);
        else if (state == "Cancelled" || state == "Interrupted") color = RGB(160, 120, 20);
        else if (state == "Paused") color = RGB(140, 80, 180);
        else if (state == "Queued" || state == "Retry queued") color = RGB(100, 100, 100);
        else return theme_.text;
        if (theme_.dark) color = blend_color(color, RGB(255, 255, 255), 45);
        return color;
    }

    LRESULT on_notify(HWND hwnd, NMHDR* header) {
        if (header->idFrom == IdList) {
            if (header->code == LVN_ITEMCHANGED && !populating_) {
                update_action_states();
                render_selected_details();
                return 0;
            }
            if (header->code == NM_DBLCLK) {
                const Row* row = selected_row();
                if (row != nullptr) {
                    if (row->kind == RowKind::RunHistory) open_journal_selected(hwnd);
                    else run_now_selected(hwnd);
                }
                return 0;
            }
            if (header->code == LVN_KEYDOWN) {
                const auto* key = reinterpret_cast<NMLVKEYDOWN*>(header);
                if (key->wVKey == VK_DELETE) {
                    const Row* row = selected_row();
                    if (row != nullptr) {
                        if (row->kind == RowKind::QueueEntry) on_command(hwnd, IdRemoveQueue, 0);
                        else if (row->kind == RowKind::SavedJob) on_command(hwnd, IdDeleteJob, 0);
                        else on_command(hwnd, IdDeleteRun, 0);
                    }
                } else if (key->wVKey == VK_F5) {
                    refresh_data(true);
                } else if (key->wVKey == VK_RETURN) {
                    run_now_selected(hwnd);
                }
                return 0;
            }
            // Rows are painted via WM_DRAWITEM (LVS_OWNERDRAWFIXED).
        }
        // The header's NM_CUSTOMDRAW is handled in the ListView subclass (it is
        // sent to the ListView, not forwarded here).
        return 0;
    }

    // Owner-paints one header cell (dark panel + light text). Called from the
    // ListView subclass, which intercepts the header's NM_CUSTOMDRAW.
    LRESULT draw_header(NMCUSTOMDRAW* draw) {
        switch (draw->dwDrawStage) {
            case CDDS_PREPAINT:
                return CDRF_NOTIFYITEMDRAW;
            case CDDS_ITEMPREPAINT: {
                std::string style = grid_header_style_;
                for (char& c : style) {
                    if (c >= 'A' && c <= 'Z') c = static_cast<char>(c - 'A' + 'a');
                }
                COLORREF header_background = theme_.panel;
                COLORREF header_text = theme_.muted_text;
                HFONT header_font = font_;
                if (style == "minimal") {
                    header_background = theme_.edit;
                } else if (style == "prominent") {
                    header_background = blend_color(theme_.panel, theme_.accent, 18);
                    header_text = readable_text_color(header_background);
                    header_font = title_font_;
                }
                themedraw::fill_rect(draw->hdc, draw->rc, header_background);
                RECT border = draw->rc;
                border.left = border.right - 1;
                themedraw::fill_rect(draw->hdc, border, theme_.border);
                wchar_t text[128]{};
                HDITEMW item{};
                item.mask = HDI_TEXT;
                item.pszText = text;
                item.cchTextMax = 127;
                Header_GetItem(ListView_GetHeader(list_), static_cast<int>(draw->dwItemSpec), &item);
                HGDIOBJ old_font = SelectObject(draw->hdc, header_font);
                SetBkMode(draw->hdc, TRANSPARENT);
                SetTextColor(draw->hdc, header_text);
                RECT text_rect = draw->rc;
                text_rect.left += 6;
                text_rect.right -= 6;
                DrawTextW(draw->hdc, text, -1, &text_rect,
                          DT_LEFT | DT_VCENTER | DT_SINGLELINE | DT_END_ELLIPSIS | DT_NOPREFIX);
                SelectObject(draw->hdc, old_font);
                return CDRF_SKIPDEFAULT;
            }
            default:
                return CDRF_DODEFAULT;
        }
    }

    // The native report-view paints column rules through the unused area below
    // the last item. Keep the populated-row grid, but cover only that empty
    // tail after the ListView has painted so the rules stop at the data.
    void paint_list_empty_tail(HWND list) {
        if (list == nullptr) return;
        RECT client{};
        GetClientRect(list, &client);

        const int item_count = ListView_GetItemCount(list);
        if (item_count <= 0) return;
        RECT last_item{};
        if (!ListView_GetItemRect(list, item_count - 1, &last_item, LVIR_BOUNDS)) return;
        if (last_item.bottom >= client.bottom) return;

        RECT empty_tail{client.left, last_item.bottom, client.right, client.bottom};
        HDC dc = GetDC(list);
        if (dc == nullptr) return;
        themedraw::fill_rect(dc, empty_tail, theme_.edit);
        ReleaseDC(list, dc);
    }

    // Intercepts the child header's NM_CUSTOMDRAW (which the ListView would
    // otherwise theme itself, painting black header text on a Light OS).
    static LRESULT CALLBACK list_subclass_proc(HWND hwnd, UINT message, WPARAM wparam,
                                                LPARAM lparam, UINT_PTR, DWORD_PTR ref) {
        auto* self = reinterpret_cast<JobManagerDialog*>(ref);
        if (message == WM_PAINT) {
            LRESULT result = DefSubclassProc(hwnd, message, wparam, lparam);
            if (self != nullptr) self->paint_list_empty_tail(hwnd);
            return result;
        }
        if (message == WM_NOTIFY && self != nullptr) {
            auto* hdr = reinterpret_cast<NMHDR*>(lparam);
            if (hdr->code == NM_CUSTOMDRAW && hdr->hwndFrom == ListView_GetHeader(hwnd)) {
                return self->draw_header(reinterpret_cast<NMCUSTOMDRAW*>(lparam));
            }
        }
        return DefSubclassProc(hwnd, message, wparam, lparam);
    }

    LRESULT proc(HWND hwnd, UINT message, WPARAM wparam, LPARAM lparam, bool& handled) {
        switch (message) {
            case WM_CREATE:
                on_create(hwnd);
                handled = true;
                return 0;
            case WM_SIZE:
                relayout(hwnd);
                handled = true;
                return 0;
            case WM_DPICHANGED: {
                // Moved to a different-DPI monitor: rebuild fonts, re-font every
                // child, adopt the suggested bounds, and relayout.
                UINT dpi = ui_layout_dpi(HIWORD(wparam), theme_.density_percent,
                                         theme_.scale_percent);
                if (font_ != nullptr) DeleteObject(font_);
                if (title_font_ != nullptr) DeleteObject(title_font_);
                font_ = CreateFontW(-MulDiv(9, static_cast<int>(dpi), 72), 0, 0, 0, FW_NORMAL, FALSE,
                                    FALSE, FALSE, DEFAULT_CHARSET, OUT_DEFAULT_PRECIS,
                                    CLIP_DEFAULT_PRECIS, CLEARTYPE_QUALITY, DEFAULT_PITCH,
                                    L"Segoe UI");
                title_font_ = CreateFontW(-MulDiv(12, static_cast<int>(dpi), 72), 0, 0, 0,
                                          FW_SEMIBOLD, FALSE, FALSE, FALSE, DEFAULT_CHARSET,
                                          OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, CLEARTYPE_QUALITY,
                                          DEFAULT_PITCH, L"Segoe UI");
                EnumChildWindows(
                    hwnd,
                    [](HWND child, LPARAM font) -> BOOL {
                        SendMessageW(child, WM_SETFONT, static_cast<WPARAM>(font), TRUE);
                        return TRUE;
                    },
                    reinterpret_cast<LPARAM>(font_));
                if (title_label_ != nullptr) {
                    SendMessageW(title_label_, WM_SETFONT, reinterpret_cast<WPARAM>(title_font_),
                                 TRUE);
                }
                SendMessageW(details_list_, LB_SETITEMHEIGHT, 0, MulDiv(18, static_cast<int>(dpi), 96));
                const RECT* s = reinterpret_cast<const RECT*>(lparam);
                SetWindowPos(hwnd, nullptr, s->left, s->top, s->right - s->left, s->bottom - s->top,
                             SWP_NOZORDER | SWP_NOACTIVATE);
                relayout(hwnd);
                handled = true;
                return 0;
            }
            case WM_GETMINMAXINFO: {
                UINT dpi = ui_layout_dpi(GetDpiForWindow(hwnd), theme_.density_percent,
                                         theme_.scale_percent);
                auto* mmi = reinterpret_cast<MINMAXINFO*>(lparam);
                mmi->ptMinTrackSize.x = MulDiv(760, static_cast<int>(dpi), 96);
                mmi->ptMinTrackSize.y = MulDiv(480, static_cast<int>(dpi), 96);
                handled = true;
                return 0;
            }
            case WM_ERASEBKGND: {
                RECT client;
                GetClientRect(hwnd, &client);
                themedraw::fill_rect(reinterpret_cast<HDC>(wparam), client, theme_.window);
                handled = true;
                return 1;
            }
            case WM_CTLCOLORSTATIC: {
                HDC dc = reinterpret_cast<HDC>(wparam);
                HFONT control_font = reinterpret_cast<HFONT>(
                    SendMessageW(reinterpret_cast<HWND>(lparam), WM_GETFONT, 0, 0));
                SetTextColor(dc, control_font == title_font_ ? theme_.accent : theme_.text);
                SetBkColor(dc, theme_.window);
                handled = true;
                return reinterpret_cast<LRESULT>(window_brush_);
            }
            case WM_CTLCOLOREDIT:
            case WM_CTLCOLORLISTBOX: {
                HDC dc = reinterpret_cast<HDC>(wparam);
                SetTextColor(dc, theme_.text);
                SetBkColor(dc, theme_.edit);
                handled = true;
                return reinterpret_cast<LRESULT>(edit_brush_);
            }
            case WM_DRAWITEM: {
                const auto& draw = *reinterpret_cast<DRAWITEMSTRUCT*>(lparam);
                if (draw.CtlType == ODT_COMBOBOX) {
                    themedraw::draw_combo_item(draw, theme_);
                } else if (draw.CtlType == ODT_LISTBOX &&
                           static_cast<int>(draw.CtlID) == IdDetailsList) {
                    draw_details_item(draw);
                } else if (draw.CtlType == ODT_LISTVIEW &&
                           static_cast<int>(draw.CtlID) == IdList) {
                    draw_list_row(draw);
                } else {
                    themedraw::draw_button(draw, theme_);
                }
                handled = true;
                return TRUE;
            }
            case WM_MEASUREITEM: {
                const auto& measure = *reinterpret_cast<MEASUREITEMSTRUCT*>(lparam);
                if (measure.CtlType == ODT_LISTVIEW &&
                    static_cast<int>(measure.CtlID) == IdList) {
                    auto* mutable_measure = reinterpret_cast<MEASUREITEMSTRUCT*>(lparam);
                    UINT dpi = ui_layout_dpi(GetDpiForWindow(hwnd), theme_.density_percent,
                                             theme_.scale_percent);
                    mutable_measure->itemHeight =
                        MulDiv(grid_row_height_, static_cast<int>(dpi), 96);
                    handled = true;
                    return TRUE;
                }
                return 0;
            }
            case WM_NOTIFY: {
                LRESULT result = on_notify(hwnd, reinterpret_cast<NMHDR*>(lparam));
                handled = true;
                return result;
            }
            case WM_TIMER:
                if (wparam == AutoRefreshTimerId) {
                    // Skip auto refresh while a combo popup is open to avoid
                    // yanking the dropdown shut mid-selection.
                    if (SendMessageW(view_combo_, CB_GETDROPPEDSTATE, 0, 0) == 0 &&
                        SendMessageW(status_combo_, CB_GETDROPPEDSTATE, 0, 0) == 0) {
                        refresh_data(true);
                    }
                } else if (wparam == SearchDebounceTimerId) {
                    KillTimer(hwnd, SearchDebounceTimerId);
                    populate_grid(true);
                }
                handled = true;
                return 0;
            case WM_COMMAND:
                on_command(hwnd, LOWORD(wparam), HIWORD(wparam));
                handled = true;
                return 0;
            case WM_CLOSE:
                save_placement(hwnd);
                DestroyWindow(hwnd);
                handled = true;
                return 0;
            case WM_DESTROY:
                save_placement(hwnd);
                if (list_ != nullptr) RemoveWindowSubclass(list_, &list_subclass_proc, 1);
                KillTimer(hwnd, AutoRefreshTimerId);
                KillTimer(hwnd, SearchDebounceTimerId);
                handled = true;
                return 0;
            default:
                return 0;
        }
    }

    JobManagerService& manager_;
    ThemePalette theme_;
    AppSettings& settings_;
    bool grid_alternating_rows_ = true;
    int grid_row_height_ = 24;
    std::string grid_header_style_ = "default";
    dialog_detail::ModalHost host_;
    JobManagerRequest request_;

    std::vector<storage::ManagedJob> saved_jobs_;
    std::vector<JobQueueEntryView> queue_entries_;
    std::vector<storage::ManagedJobRun> runs_;
    std::vector<Row> rows_;
    // Last rendered details text + the row it belongs to, so the auto-refresh
    // can skip rebuilding (and scrolling) an unchanged pane.
    std::vector<std::wstring> details_lines_;
    std::string details_row_id_;
    bool populating_ = false;

    HWND list_ = nullptr;
    HWND view_combo_ = nullptr;
    HWND status_combo_ = nullptr;
    HWND search_edit_ = nullptr;
    HWND details_list_ = nullptr;
    HWND details_label_ = nullptr;
    HWND title_label_ = nullptr;
    HWND summary_label_ = nullptr;
    std::map<int, HWND> buttons_;
    HFONT font_ = nullptr;
    HFONT title_font_ = nullptr;
    HBRUSH window_brush_ = nullptr;
    HBRUSH edit_brush_ = nullptr;
};

} // namespace xact::ui
