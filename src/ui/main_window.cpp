// -----------------------------------------------------------------------------
// File: src\ui\main_window.cpp
// Purpose: XactCopy native UI. Win32 main window over the ported supervisor/
//          recovery/settings core with the shared dark/light theme layer:
//          owner-drawn buttons/checkboxes/combos, severity-colored log,
//          accent-tinted progress bars, dark title bar + scrollbars, system
//          theme tracking, crash-resume prompt, and single-instance guard.
// -----------------------------------------------------------------------------

#include <deque>
#include <limits>
#include <map>
#include <memory>
#include <optional>
#include <string>
#include <thread>

#ifndef NOMINMAX
#define NOMINMAX
#endif
#ifndef WIN32_LEAN_AND_MEAN
#define WIN32_LEAN_AND_MEAN
#endif
#ifndef _WIN32_WINNT
#define _WIN32_WINNT 0x0A00 // Windows 10+: GetDpiForWindow, DPI awareness APIs
#endif
#ifndef WINVER
#define WINVER 0x0A00 // MinGW gates the per-monitor DPI APIs on WINVER
#endif
#include <windows.h>
#include <commctrl.h>
#include <dwmapi.h>
#include <shellapi.h>
#include <shobjidl.h>

#include "app_icon.h"
#include "dialogs.h"
#include "explorer_integration.h"
#include "job_manager.h"
#include "job_manager_dialog.h"
#include "recovery.h"
#include "selection.h"
#include "settings.h"
#include "supervisor.h"
#include "taskbar_progress.h"
#include "theme.h"
#include "update_dialog.h"
#include "update_service.h"

#if defined(_MSC_VER)
#pragma comment(lib, "comctl32.lib")
#pragma comment(lib, "dwmapi.lib")
#endif

namespace {

using namespace xact;

constexpr UINT WM_APP_LOG = WM_APP + 1;
constexpr UINT WM_APP_PROGRESS = WM_APP + 2;
constexpr UINT WM_APP_RESULT = WM_APP + 3;
constexpr UINT WM_APP_STATE = WM_APP + 4;
constexpr UINT WM_APP_PAUSE = WM_APP + 5;
constexpr UINT WM_APP_START_DONE = WM_APP + 6; // wParam: 1 ok / 0 error; lParam: new std::string*
constexpr UINT WM_APP_UPDATE_DONE = WM_APP + 7; // lParam: new ui::UpdateReleaseInfo*
constexpr UINT WM_APP_APPLY_LAUNCH = WM_APP + 8; // deferred apply_explorer_launch_options()

constexpr int IdSourceEdit = 1001;
constexpr int IdSourceBrowse = 1002;
constexpr int IdDestinationEdit = 1003;
constexpr int IdDestinationBrowse = 1004;
constexpr int IdModeCombo = 1005;
constexpr int IdEngineCombo = 1006;
constexpr int IdOverwriteCombo = 1007;
constexpr int IdVerifyCombo = 1008;
constexpr int IdSalvageCheck = 1009;
constexpr int IdResumeCheck = 1010;
constexpr int IdMapCheck = 1011;
constexpr int IdAdaptiveCheck = 1012;
constexpr int IdStartButton = 1013;
constexpr int IdPauseButton = 1014;
constexpr int IdCancelButton = 1015;
constexpr int IdLogList = 1016;
constexpr int IdSettingsButton = 1017;
constexpr int IdAboutButton = 1018;
constexpr int IdContinueCheck = 1019;
constexpr int IdSkipKnownBadCheck = 1020;
constexpr int IdWaitMediaCheck = 1021;
constexpr int IdFragileCheck = 1022;
constexpr int IdBufferEdit = 1023;
constexpr int IdRetriesEdit = 1024;
constexpr int IdTimeoutEdit = 1025;
constexpr int IdCurrentBar = 1026;
constexpr int IdOverallBar = 1027;
constexpr int IdSaveDefaultsButton = 1028;

// Menu command ids (mirrors the .NET MainForm menu strip).
constexpr int IdMenuStartCopy = 1101;
constexpr int IdMenuPauseCopy = 1102;
constexpr int IdMenuResumeCopy = 1103;
constexpr int IdMenuCancelCopy = 1104;
constexpr int IdMenuOpenJournals = 1105;
constexpr int IdMenuOpenCrash = 1106;
constexpr int IdMenuExit = 1107;
constexpr int IdMenuScanBadBlocks = 1108;
constexpr int IdMenuSettings = 1109;
constexpr int IdMenuCheckUpdates = 1110;
constexpr int IdMenuSaveAsJob = 1111;
constexpr int IdMenuJobManager = 1112;
constexpr int IdMenuResumeInterrupted = 1113;
constexpr int IdMenuRunNextQueued = 1114;
constexpr int IdMenuAbout = 1115;
constexpr int IdMenuSaveDefaults = 1116;
constexpr int IdMenuInspectBadMap = 1117;
constexpr int IdMenuClearBadMap = 1118;

constexpr int IdSourceAddFiles = 1040;
constexpr int IdClearSelection = 1041;

// Explorer invokes a context-menu verb once per selected item, so a multi-select
// arrives as a burst of launches. Selections are accumulated and applied once
// this quiet period elapses, giving one source root and one destination prompt.
constexpr UINT IdSelectionTimer = 1;
constexpr UINT IdDiagnosticsTimer = 2;
constexpr UINT SelectionCoalesceMs = 700;

constexpr const wchar_t* WindowClassName = L"XactCopyNativeMain";

// WM_COPYDATA tag used to hand a second instance's command line to the running
// one (so Explorer verbs still work when XactCopy is already open).
constexpr ULONG_PTR ForwardedLaunchId = 0x58414331; // 'XAC1'

// Defined below; declared here for the WM_COPYDATA handler.
ui::LaunchOptions parse_launch_options(const wchar_t* command_line);

std::string format_bytes_short(std::int64_t value) {
    if (value < 1024) return std::to_string(value) + " B";
    static const char* units[] = {"KB", "MB", "GB", "TB", "PB"};
    double size = static_cast<double>(value);
    int unit = -1;
    do {
        size /= 1024.0;
        ++unit;
    } while (size >= 1024.0 && unit < 4);
    char text[32];
    std::snprintf(text, sizeof(text), "%.2f %s", size, units[unit]);
    return text;
}

class MainWindow {
public:
    int run(HINSTANCE instance, const ui::LaunchOptions& launch) {
        instance_ = instance;
        launch_ = launch;

        INITCOMMONCONTROLSEX icc{
            sizeof(icc), ICC_PROGRESS_CLASS | ICC_BAR_CLASSES | ICC_WIN95_CLASSES};
        InitCommonControlsEx(&icc);

        theme_ = ui::make_theme(settings_.theme(),
                                settings_.get_string("AccentColorMode", "auto"),
                                settings_.get_string("AccentColorHex", "#5A78C8"));
        ui::configure_theme_engine(settings_.theme(), theme_);
        colorize_log_ = settings_.get_bool("UiColorizeLogBySeverity", true);
        ui_scale_percent_ = settings_.get_int("UiScalePercent", 50, 250, 100);
        ui_density_percent_ = ui::ui_density_percent(settings_.get_string("UiDensity", "normal"));
        theme_.themed_chrome = !models::detail::equals_ignore_case(
            settings_.get_string("WindowChromeMode", "themed"), "standard");
        theme_.density_percent = ui_density_percent_;
        theme_.scale_percent = ui_scale_percent_;
        // WM_GETMINMAXINFO can arrive before WM_CREATE. Seed the cached DPI so
        // the first size and minimum bounds are correct on a high-DPI primary
        // monitor instead of assuming 96 DPI until on_create().
        monitor_dpi_ = GetDpiForSystem();

        WNDCLASSW window_class{};
        window_class.lpfnWndProc = &MainWindow::static_window_proc;
        window_class.hInstance = instance;
        window_class.lpszClassName = WindowClassName;
        window_class.hCursor = LoadCursorW(nullptr, IDC_ARROW);
        window_class.hbrBackground = nullptr; // painted per-theme in WM_ERASEBKGND
        ui::assign_window_class_icons(window_class);
        RegisterClassW(&window_class);

        // WS_CLIPCHILDREN stops the parent's WM_ERASEBKGND from painting over the
        // child controls during a resize, which is what smears the owner-drawn
        // progress bars.
        const int initial_dpi = effective_dpi();
        hwnd_ = CreateWindowExW(
            0, WindowClassName, L"XactCopy", WS_OVERLAPPEDWINDOW | WS_CLIPCHILDREN,
            CW_USEDEFAULT, CW_USEDEFAULT, MulDiv(960, initial_dpi, 96),
            MulDiv(800, initial_dpi, 96), nullptr, nullptr, instance, this);
        if (hwnd_ == nullptr) return 1;

        const HWND warning_target = hwnd_;
        auto post_persistence_warning = [warning_target](const std::string& message) {
            PostMessageW(warning_target, WM_APP_LOG, 0,
                         reinterpret_cast<LPARAM>(
                             new std::string("WARNING: " + message)));
        };
        job_manager_.on_persistence_warning = post_persistence_warning;
        recovery_.on_persistence_warning = post_persistence_warning;

        ui::apply_window_icons(hwnd_);
        ui::set_dark_title_bar(hwnd_, theme_.dark && theme_.themed_chrome);
        DragAcceptFiles(hwnd_, TRUE); // dropped items join the selection
        restore_window_placement();
        ShowWindow(hwnd_, restore_maximized_ ? SW_SHOWMAXIMIZED : SW_SHOW);
        UpdateWindow(hwnd_);
        if (!settings_.load_warning().empty()) {
            append_log("WARNING: " + settings_.load_warning());
            warn_box(storage::fsutil::utf8_to_wide(settings_.load_warning()),
                     L"Settings recovery");
        }
        if (!job_manager_.load_warning().empty()) {
            append_log("WARNING: " + job_manager_.load_warning());
            warn_box(storage::fsutil::utf8_to_wide(job_manager_.load_warning()),
                     L"Job catalog recovery");
        }

        // Any managed run still marked Running/Paused is a leftover from a crash.
        job_manager_.mark_any_running_runs_interrupted(
            "XactCopy restarted; the previous run did not finish cleanly.");

        // Retention runs after the interrupted marking above, so a crashed run's
        // journal is already flagged resumable and therefore protected.
        prune_old_journals();

        sync_explorer_integration(false);
        apply_explorer_launch_options();

        maybe_prompt_recovery();

        // Silent update check on launch (CheckUpdatesOnLaunch).
        if (settings_.get_bool("CheckUpdatesOnLaunch", false)) {
            check_for_updates(/*show_dialog*/ false);
        }

        // Startup queue drain (MainForm honors AutoRunQueuedJobsOnStartup).
        if (!supervisor_.is_job_running() &&
            settings_.get_bool("AutoRunQueuedJobsOnStartup", false)) {
            append_log("Startup queue mode enabled. Looking for queued jobs.");
            run_next_queued_job(/*manual*/ false);
        }

        MSG message;
        while (GetMessageW(&message, nullptr, 0, 0) > 0) {
            if (IsDialogMessageW(hwnd_, &message)) continue;
            TranslateMessage(&message);
            DispatchMessageW(&message);
        }
        return static_cast<int>(message.wParam);
    }

private:
    HINSTANCE instance_ = nullptr;
    HWND hwnd_ = nullptr;
    HWND source_label_ = nullptr;
    HWND destination_label_ = nullptr;
    HWND source_edit_ = nullptr;
    HWND destination_edit_ = nullptr;
    HWND source_browse_ = nullptr;
    HWND destination_browse_ = nullptr;
    HWND mode_combo_ = nullptr;
    HWND engine_combo_ = nullptr;
    HWND overwrite_combo_ = nullptr;
    HWND verify_combo_ = nullptr;
    HWND salvage_check_ = nullptr;
    HWND resume_check_ = nullptr;
    HWND map_check_ = nullptr;
    HWND adaptive_check_ = nullptr;
    HWND continue_check_ = nullptr;
    HWND skip_known_bad_check_ = nullptr;
    HWND wait_media_check_ = nullptr;
    HWND fragile_check_ = nullptr;
    HWND buffer_label_ = nullptr;
    HWND buffer_edit_ = nullptr;
    HWND retries_label_ = nullptr;
    HWND retries_edit_ = nullptr;
    HWND timeout_label_ = nullptr;
    HWND timeout_edit_ = nullptr;
    HWND start_button_ = nullptr;
    HWND save_defaults_button_ = nullptr;
    HWND settings_button_ = nullptr;
    HWND about_button_ = nullptr;
    HWND pause_button_ = nullptr;
    HWND cancel_button_ = nullptr;
    HWND current_label_ = nullptr;
    HWND current_bar_ = nullptr;
    HWND overall_bar_ = nullptr;
    int current_permille_ = 0; // owner-drawn flat progress bars (0..1000)
    int overall_permille_ = 0;
    HWND stats_label_ = nullptr;
    HWND throughput_label_ = nullptr;
    HWND eta_label_ = nullptr;
    HWND buffer_usage_label_ = nullptr;
    HWND rescue_label_ = nullptr;
    HWND job_summary_label_ = nullptr;
    HWND journal_label_ = nullptr;
    HWND diagnostics_label_ = nullptr;
    HWND status_label_ = nullptr;
    HWND log_list_ = nullptr;
    int log_horizontal_extent_ = 0;
    HFONT font_ = nullptr;
    HFONT mono_font_ = nullptr;
    HBRUSH window_brush_ = nullptr;
    HBRUSH edit_brush_ = nullptr;
    HMENU menu_bar_ = nullptr;
    UINT monitor_dpi_ = 96;
    ui::TooltipManager tooltips_;

    ui::ThemePalette theme_;
    bool colorize_log_ = true;
    int ui_scale_percent_ = 100; // UiScalePercent setting (global UI scale)
    int ui_density_percent_ = 100; // UiDensity setting (compact/normal/comfortable)
    ui::AppSettings settings_;
    ui::RecoveryService recovery_;
    ui::WorkerSupervisor supervisor_;
    ui::JobManagerService job_manager_;
    ui::TaskbarProgress taskbar_;
    ui::LaunchOptions launch_;
    bool paused_ = false;
    std::string active_run_id_;
    std::string active_managed_run_id_; // JobCatalog run id for the current job
    bool restore_maximized_ = false;    // deferred maximize from saved placement
    bool checking_updates_ = false;
    bool update_check_show_dialog_ = false;
    std::thread update_thread_;
    std::string explorer_selection_root_;    // internal directory root sent to the worker
    std::string explorer_selection_display_; // exact item/common root shown in the Source box
    std::vector<std::string> explorer_selected_paths_; // relative paths for the exact selection

    // Multi-item selection: accumulated across a burst of arrivals, then applied
    // once by the coalescing timer (see queue_selection/commit_selection).
    ui::SelectionModel selection_;
    bool selection_timer_active_ = false;
    bool selection_wants_destination_ = false;
    HWND selection_label_ = nullptr;
    HWND source_add_files_ = nullptr;
    HWND clear_selection_ = nullptr;
    std::map<int, bool> check_states_;
    bool suppress_source_change_ = false;
    int active_mode_index_ = -1;
    int copy_profile_index_ = 2; // Custom preserves advanced saved defaults.
    int copy_overwrite_index_ = 0;
    int copy_verify_index_ = 2;
    bool applying_copy_profile_ = false;
    int scan_profile_index_ = 0;
    int scan_backend_index_ = 0;
    int scan_map_index_ = 0;
    models::TransferEnginePolicy custom_transfer_engine_ =
        models::TransferEnginePolicy::Auto;

    // Async job start (spawn + pipe connect must not block the UI thread).
    std::thread start_thread_;
    bool starting_ = false;
    models::CopyJobOptions pending_start_options_;
    std::string pending_start_job_name_;

    // Transfer telemetry (mirrors MainForm.UpdateTransferTelemetry smoothing).
    double smoothed_bytes_per_second_ = 0.0;
    bool telemetry_started_ = false;
    time::DateTimeOffset telemetry_start_utc_;
    time::DateTimeOffset telemetry_last_sample_utc_;
    std::int64_t telemetry_last_bytes_copied_ = 0;
    std::int64_t buffer_sample_bytes_total_ = 0;
    std::int64_t buffer_samples_count_ = 0;
    bool show_diagnostics_ = false;
    int diagnostics_refresh_ms_ = 250;
    std::optional<models::CopyProgressSnapshot> latest_progress_;
    // Appearance settings consumed by the UI.
    bool show_buffer_row_ = true;
    bool show_rescue_row_ = true;
    bool bar_show_percentage_ = false;
    int progress_bar_height_logical_ = 14;

    static LRESULT CALLBACK static_window_proc(HWND hwnd, UINT message, WPARAM wparam,
                                               LPARAM lparam) {
        MainWindow* self;
        if (message == WM_NCCREATE) {
            self = static_cast<MainWindow*>(
                reinterpret_cast<CREATESTRUCTW*>(lparam)->lpCreateParams);
            SetWindowLongPtrW(hwnd, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(self));
            self->hwnd_ = hwnd;
        } else {
            self = reinterpret_cast<MainWindow*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
        }
        if (self == nullptr) return DefWindowProcW(hwnd, message, wparam, lparam);
        return self->window_proc(message, wparam, lparam);
    }

    LRESULT window_proc(UINT message, WPARAM wparam, LPARAM lparam) {
        switch (message) {
            case WM_CREATE:
                on_create();
                return 0;
            case WM_SIZE:
                layout();
                return 0;
            case WM_DPICHANGED: {
                // Cache the new monitor DPI before touching any metric. During
                // WM_DPICHANGED, querying the window while it still has its old
                // bounds can otherwise mix old item heights with new fonts.
                const RECT* suggested = reinterpret_cast<const RECT*>(lparam);
                monitor_dpi_ = HIWORD(wparam) != 0 ? HIWORD(wparam) : LOWORD(wparam);
                SetWindowPos(hwnd_, nullptr, suggested->left, suggested->top,
                             suggested->right - suggested->left,
                             suggested->bottom - suggested->top, SWP_NOZORDER | SWP_NOACTIVATE);
                rebuild_fonts(monitor_dpi_);
                apply_fonts();
                tooltips_.update_dpi(static_cast<UINT>(effective_dpi()));
                ui::apply_window_icons(hwnd_);
                layout();
                RedrawWindow(hwnd_, nullptr, nullptr,
                             RDW_INVALIDATE | RDW_ERASE | RDW_ALLCHILDREN | RDW_FRAME);
                return 0;
            }
            case WM_GETMINMAXINFO: {
                auto* limits = reinterpret_cast<MINMAXINFO*>(lparam);
                const int dpi = effective_dpi();
                limits->ptMinTrackSize.x = MulDiv(760, dpi, 96);
                limits->ptMinTrackSize.y = MulDiv(620, dpi, 96);
                return 0;
            }
            case WM_ERASEBKGND: {
                RECT client;
                GetClientRect(hwnd_, &client);
                ui::themedraw::fill_rect(reinterpret_cast<HDC>(wparam), client, theme_.window);
                return 1;
            }
            case WM_COMMAND:
                on_command(LOWORD(wparam), HIWORD(wparam));
                return 0;
            case WM_DRAWITEM:
                return on_draw_item(*reinterpret_cast<DRAWITEMSTRUCT*>(lparam)) ? TRUE
                                                                               : DefWindowProcW(
                                                                                     hwnd_, message,
                                                                                     wparam, lparam);
            case WM_CTLCOLORSTATIC: {
                HDC dc = reinterpret_cast<HDC>(wparam);
                SetTextColor(dc, theme_.text);
                SetBkColor(dc, theme_.window);
                return reinterpret_cast<LRESULT>(window_brush_);
            }
            case WM_CTLCOLOREDIT:
            case WM_CTLCOLORLISTBOX: {
                HDC dc = reinterpret_cast<HDC>(wparam);
                SetTextColor(dc, theme_.text);
                SetBkColor(dc, theme_.edit);
                return reinterpret_cast<LRESULT>(edit_brush_);
            }
            case WM_SETTINGCHANGE:
                if (ui::handle_theme_setting_change(lparam)) {
                    apply_theme();
                }
                return 0;
            case WM_APP_LOG: {
                std::unique_ptr<std::string> text(reinterpret_cast<std::string*>(lparam));
                append_log(*text);
                return 0;
            }
            case WM_APP_PROGRESS: {
                std::unique_ptr<models::CopyProgressSnapshot> snapshot(
                    reinterpret_cast<models::CopyProgressSnapshot*>(lparam));
                // Coalesce: the worker emits a snapshot per file transition, which
                // during a whole-drive scan can arrive far faster than the UI can
                // paint. Drain any queued progress and apply only the latest so a
                // burst never saturates the UI thread.
                MSG queued;
                while (PeekMessageW(&queued, hwnd_, WM_APP_PROGRESS, WM_APP_PROGRESS, PM_REMOVE)) {
                    snapshot.reset(reinterpret_cast<models::CopyProgressSnapshot*>(queued.lParam));
                }
                apply_progress(*snapshot);
                return 0;
            }
            case WM_APP_RESULT: {
                std::unique_ptr<models::CopyJobResult> result(
                    reinterpret_cast<models::CopyJobResult*>(lparam));
                apply_result(*result);
                return 0;
            }
            case WM_APP_STATE: {
                std::unique_ptr<std::string> state(reinterpret_cast<std::string*>(lparam));
                SetWindowTextW(status_label_,
                               storage::fsutil::utf8_to_wide("Status: " + *state).c_str());
                return 0;
            }
            case WM_APP_START_DONE: {
                std::unique_ptr<std::string> error(reinterpret_cast<std::string*>(lparam));
                on_start_done(wparam == 1, *error);
                return 0;
            }
            case WM_APP_UPDATE_DONE: {
                std::unique_ptr<ui::UpdateReleaseInfo> info(
                    reinterpret_cast<ui::UpdateReleaseInfo*>(lparam));
                on_update_done(*info);
                return 0;
            }
            case WM_APP_PAUSE:
                paused_ = wparam != 0;
                SetWindowTextW(pause_button_, paused_ ? L"Resume" : L"Pause");
                InvalidateRect(pause_button_, nullptr, FALSE);
                if (!active_managed_run_id_.empty()) {
                    if (paused_) job_manager_.mark_run_paused(active_managed_run_id_);
                    else job_manager_.mark_run_resumed(active_managed_run_id_);
                }
                taskbar_.set_state(hwnd_, paused_ ? ui::TaskbarProgressState::Paused
                                                  : ui::TaskbarProgressState::Normal);
                update_menu_state();
                return 0;
            case WM_COPYDATA: {
                // A second instance handed us its command line: adopt any
                // Explorer selection it carried and apply it here.
                auto* data = reinterpret_cast<COPYDATASTRUCT*>(lparam);
                if (data == nullptr || data->dwData != ForwardedLaunchId ||
                    data->lpData == nullptr || data->cbData < sizeof(wchar_t)) {
                    return 0;
                }
                std::wstring command_line(static_cast<const wchar_t*>(data->lpData),
                                          data->cbData / sizeof(wchar_t));
                while (!command_line.empty() && command_line.back() == L'\0') {
                    command_line.pop_back();
                }
                ui::LaunchOptions forwarded = parse_launch_options(command_line.c_str());
                if (forwarded.explorer_source_paths.empty() &&
                    ui_is_blank(forwarded.explorer_folder_path)) {
                    return TRUE; // nothing actionable; just surfacing the window
                }
                // Queue only. The sender is blocked inside SendMessage until we
                // return, so this must never open a dialog — the destination
                // prompt happens later, from the coalescing timer.
                if (!ui_is_blank(forwarded.explorer_folder_path)) {
                    launch_.explorer_folder_path = forwarded.explorer_folder_path;
                    launch_.explorer_scan_mode = forwarded.explorer_scan_mode;
                    PostMessageW(hwnd_, WM_APP_APPLY_LAUNCH, 0, 0);
                    return TRUE;
                }
                std::vector<std::wstring> paths;
                paths.reserve(forwarded.explorer_source_paths.size());
                for (const auto& raw : forwarded.explorer_source_paths) {
                    paths.push_back(resolve_full_path(raw));
                }
                if (!forwarded.explorer_scan_mode) {
                    SendMessageW(mode_combo_, CB_SETCURSEL, 0, 0);
                    sync_mode_ui();
                }
                queue_selection(paths, forwarded.explorer_scan_mode, /*want_destination*/ true);
                return TRUE;
            }
            case WM_APP_APPLY_LAUNCH:
                apply_explorer_launch_options();
                return 0;
            case WM_TIMER:
                if (wparam == IdSelectionTimer) {
                    commit_selection();
                    return 0;
                }
                if (wparam == IdDiagnosticsTimer) {
                    refresh_diagnostics_label();
                    return 0;
                }
                return DefWindowProcW(hwnd_, message, wparam, lparam);
            case WM_DROPFILES: {
                // Files/folders dropped on the window join the same selection.
                HDROP drop = reinterpret_cast<HDROP>(wparam);
                UINT count = DragQueryFileW(drop, 0xFFFFFFFF, nullptr, 0);
                std::vector<std::wstring> paths;
                for (UINT i = 0; i < count; ++i) {
                    UINT length = DragQueryFileW(drop, i, nullptr, 0);
                    std::wstring path(length + 1, L'\0');
                    if (DragQueryFileW(drop, i, path.data(), length + 1) > 0) {
                        path.resize(length);
                        paths.push_back(std::move(path));
                    }
                }
                DragFinish(drop);
                queue_selection(paths, /*scan_mode*/ false, /*want_destination*/ false);
                return 0;
            }
            case WM_CLOSE:
                if (supervisor_.is_job_running()) {
                    if (!confirm_box(L"A copy job is running. Cancel it and exit?\n\n"
                                     L"The journal keeps progress, so the run can resume next time.",
                                     L"XactCopy", ui::MessageIcon::Warning)) {
                        return 0;
                    }
                    recovery_.mark_run_interrupted("User exited while a run was active.");
                    if (!active_managed_run_id_.empty()) {
                        job_manager_.mark_run_interrupted(active_managed_run_id_,
                                                          "Application closed while the run was active.");
                    }
                }
                if (!save_window_placement()) {
                    warn_box(storage::fsutil::utf8_to_wide(settings_.last_save_error()),
                             L"Window settings could not be saved");
                }
                DestroyWindow(hwnd_);
                return 0;
            case WM_DESTROY:
                KillTimer(hwnd_, IdSelectionTimer);
                KillTimer(hwnd_, IdDiagnosticsTimer);
                if (start_thread_.joinable()) start_thread_.join();
                if (update_thread_.joinable()) update_thread_.join();
                supervisor_.stop();
                if (!supervisor_.is_job_running()) recovery_.mark_clean_shutdown();
                PostQuitMessage(0);
                return 0;
            default:
                return DefWindowProcW(hwnd_, message, wparam, lparam);
        }
    }

    // ---- Construction ------------------------------------------------------

    HWND create(const wchar_t* class_name, const wchar_t* text, DWORD style, int id,
                DWORD ex_style = 0, HFONT font = nullptr) {
        HWND handle = CreateWindowExW(ex_style, class_name, text, WS_CHILD | WS_VISIBLE | style,
                                      0, 0, 10, 10, hwnd_,
                                      reinterpret_cast<HMENU>(static_cast<INT_PTR>(id)),
                                      instance_, nullptr);
        SendMessageW(handle, WM_SETFONT,
                     reinterpret_cast<WPARAM>(font != nullptr ? font : font_), TRUE);
        return handle;
    }

    int effective_scale_percent() const {
        const int base = std::clamp(ui_scale_percent_, 50, 250);
        const int density = std::clamp(ui_density_percent_, 80, 120);
        return std::clamp((base * density + 50) / 100, 50, 250);
    }

    // Monitor DPI folded with both global scale and density, so a single
    // effective DPI drives fonts and geometry consistently.
    int effective_dpi() const {
        return static_cast<int>(monitor_dpi_) * effective_scale_percent() / 100;
    }

    // Per-monitor DPI: fonts are rebuilt for the window's current monitor and
    // re-applied on WM_DPICHANGED (mirrors AxiomCompress MainWindow::update_dpi).
    // `dpi` is the raw monitor DPI; UiScalePercent and UiDensity are folded in
    // here so a density change updates both control geometry and text metrics.
    void rebuild_fonts(UINT dpi) {
        int eff = static_cast<int>(dpi) * effective_scale_percent() / 100;
        if (font_ != nullptr) DeleteObject(font_);
        if (mono_font_ != nullptr) DeleteObject(mono_font_);
        font_ = CreateFontW(-MulDiv(9, eff, 72), 0, 0, 0, FW_NORMAL, FALSE,
                            FALSE, FALSE, DEFAULT_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS,
                            CLEARTYPE_QUALITY, DEFAULT_PITCH, L"Segoe UI");
        std::wstring log_font =
            storage::fsutil::utf8_to_wide(settings_.get_string("LogFontFamily", "Consolas"));
        int log_points = settings_.get_int("LogFontSizePoints", 7, 20, 9);
        mono_font_ = CreateFontW(-MulDiv(log_points, eff, 72), 0, 0, 0,
                                 FW_NORMAL, FALSE, FALSE, FALSE, DEFAULT_CHARSET,
                                 OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, CLEARTYPE_QUALITY,
                                 DEFAULT_PITCH, log_font.c_str());
    }

    void apply_fonts() {
        int log_points = settings_.get_int("LogFontSizePoints", 7, 20, 9);
        // Must cover EVERY control that references font_: rebuild_fonts() deletes
        // the old handle, so any control left on the stale handle repaints with a
        // GDI fallback (looks bold/wrong). Owner-drawn buttons/checks/combos read
        // their font via WM_GETFONT, so they need it set too.
        HWND ui_controls[] = {
            source_label_,     destination_label_,   source_edit_,      destination_edit_,
            source_browse_,    destination_browse_,  source_add_files_, selection_label_,
            clear_selection_,  mode_combo_,          engine_combo_,
            overwrite_combo_,  verify_combo_,        salvage_check_,    resume_check_,
            map_check_,        adaptive_check_,      continue_check_,   skip_known_bad_check_,
            wait_media_check_, fragile_check_,       buffer_label_,     buffer_edit_,
            retries_label_,    retries_edit_,        timeout_label_,    timeout_edit_,
            start_button_,     pause_button_,        cancel_button_,    save_defaults_button_,
            settings_button_,
            about_button_,     current_label_,       stats_label_,      status_label_,
            throughput_label_, eta_label_,           buffer_usage_label_, rescue_label_,
            job_summary_label_, journal_label_,      diagnostics_label_,
        };
        for (HWND control : ui_controls) {
            if (control != nullptr) {
                SendMessageW(control, WM_SETFONT, reinterpret_cast<WPARAM>(font_), TRUE);
            }
        }
        if (log_list_ != nullptr) {
            SendMessageW(log_list_, WM_SETFONT, reinterpret_cast<WPARAM>(mono_font_), TRUE);
            SendMessageW(log_list_, LB_SETITEMHEIGHT, 0,
                         MulDiv(log_points + 6, effective_dpi(), 72));
        }
        const int combo_item_height = MulDiv(22, effective_dpi(), 96);
        for (HWND combo : {mode_combo_, engine_combo_, overwrite_combo_, verify_combo_}) {
            if (combo == nullptr) continue;
            SendMessageW(combo, CB_SETITEMHEIGHT, static_cast<WPARAM>(-1), combo_item_height);
            SendMessageW(combo, CB_SETITEMHEIGHT, 0, combo_item_height);
        }
        rebuild_log_horizontal_extent();
        InvalidateRect(hwnd_, nullptr, TRUE);
    }

    int measure_log_text_width(HDC dc, const std::wstring& text) const {
        if (dc == nullptr || text.empty()) return 0;
        SIZE size{};
        int length = static_cast<int>(std::min<std::size_t>(
            text.size(), static_cast<std::size_t>(std::numeric_limits<int>::max())));
        if (!GetTextExtentPoint32W(dc, text.data(), length, &size)) return 0;
        return size.cx + MulDiv(8, effective_dpi(), 96);
    }

    void update_log_horizontal_extent(const std::wstring& text) {
        if (log_list_ == nullptr || mono_font_ == nullptr) return;
        HDC dc = GetDC(log_list_);
        if (dc == nullptr) return;
        HGDIOBJ old_font = SelectObject(dc, mono_font_);
        int extent = measure_log_text_width(dc, text);
        SelectObject(dc, old_font);
        ReleaseDC(log_list_, dc);
        if (extent <= log_horizontal_extent_) return;
        log_horizontal_extent_ = extent;
        SendMessageW(log_list_, LB_SETHORIZONTALEXTENT,
                     static_cast<WPARAM>(log_horizontal_extent_), 0);
    }

    void rebuild_log_horizontal_extent() {
        if (log_list_ == nullptr || mono_font_ == nullptr) return;
        HDC dc = GetDC(log_list_);
        if (dc == nullptr) return;
        HGDIOBJ old_font = SelectObject(dc, mono_font_);
        int extent = 0;
        int count = static_cast<int>(SendMessageW(log_list_, LB_GETCOUNT, 0, 0));
        for (int index = 0; index < count; ++index) {
            LRESULT length = SendMessageW(log_list_, LB_GETTEXTLEN, index, 0);
            if (length < 0) continue;
            std::wstring text(static_cast<std::size_t>(length) + 1, L'\0');
            SendMessageW(log_list_, LB_GETTEXT, index,
                         reinterpret_cast<LPARAM>(text.data()));
            text.resize(static_cast<std::size_t>(length));
            extent = std::max(extent, measure_log_text_width(dc, text));
        }
        SelectObject(dc, old_font);
        ReleaseDC(log_list_, dc);
        log_horizontal_extent_ = extent;
        SendMessageW(log_list_, LB_SETHORIZONTALEXTENT,
                     static_cast<WPARAM>(log_horizontal_extent_), 0);
    }

    void reset_log_horizontal_extent() {
        log_horizontal_extent_ = 0;
        if (log_list_ != nullptr) {
            SendMessageW(log_list_, LB_SETHORIZONTALEXTENT, 0, 0);
        }
    }

    void on_create() {
        window_brush_ = CreateSolidBrush(theme_.window);
        edit_brush_ = CreateSolidBrush(theme_.edit);

        monitor_dpi_ = GetDpiForWindow(hwnd_);
        rebuild_fonts(monitor_dpi_);
        int log_points = settings_.get_int("LogFontSizePoints", 7, 20, 9);

        source_label_ = create(L"STATIC", L"Source:", 0, 0);
        source_edit_ = create(L"EDIT", L"", WS_TABSTOP | ES_AUTOHSCROLL | WS_BORDER, IdSourceEdit);
        source_browse_ = create(L"BUTTON", L"Browse...", WS_TABSTOP | BS_OWNERDRAW, IdSourceBrowse);
        source_add_files_ =
            create(L"BUTTON", L"Add Files...", WS_TABSTOP | BS_OWNERDRAW, IdSourceAddFiles);
        selection_label_ = create(L"STATIC", L"", SS_LEFT | SS_ENDELLIPSIS | SS_NOPREFIX, 0);
        clear_selection_ =
            create(L"BUTTON", L"Clear", WS_TABSTOP | BS_OWNERDRAW, IdClearSelection);
        ShowWindow(selection_label_, SW_HIDE);
        ShowWindow(clear_selection_, SW_HIDE);
        destination_label_ = create(L"STATIC", L"Destination:", 0, 0);
        destination_edit_ =
            create(L"EDIT", L"", WS_TABSTOP | ES_AUTOHSCROLL | WS_BORDER, IdDestinationEdit);
        destination_browse_ =
            create(L"BUTTON", L"Browse...", WS_TABSTOP | BS_OWNERDRAW, IdDestinationBrowse);

        DWORD combo_style = WS_TABSTOP | CBS_DROPDOWNLIST | CBS_OWNERDRAWFIXED | CBS_HASSTRINGS;
        mode_combo_ = create(L"COMBOBOX", L"", combo_style, IdModeCombo);
        engine_combo_ = create(L"COMBOBOX", L"", combo_style, IdEngineCombo);
        overwrite_combo_ = create(L"COMBOBOX", L"", combo_style, IdOverwriteCombo);
        verify_combo_ = create(L"COMBOBOX", L"", combo_style, IdVerifyCombo);
        for (const wchar_t* item : {L"Verified Copy / Recovery", L"Assess Readable Files"}) {
            SendMessageW(mode_combo_, CB_ADDSTRING, 0, reinterpret_cast<LPARAM>(item));
        }

        salvage_check_ = create(L"BUTTON", L"Salvage unreadable blocks",
                                WS_TABSTOP | BS_OWNERDRAW, IdSalvageCheck);
        resume_check_ =
            create(L"BUTTON", L"Resume from journal", WS_TABSTOP | BS_OWNERDRAW, IdResumeCheck);
        map_check_ = create(L"BUTTON", L"Use bad-range map", WS_TABSTOP | BS_OWNERDRAW, IdMapCheck);
        adaptive_check_ =
            create(L"BUTTON", L"Adaptive buffer", WS_TABSTOP | BS_OWNERDRAW, IdAdaptiveCheck);
        continue_check_ = create(L"BUTTON", L"Continue on error", WS_TABSTOP | BS_OWNERDRAW,
                                 IdContinueCheck);
        skip_known_bad_check_ = create(L"BUTTON", L"Skip known-bad ranges",
                                       WS_TABSTOP | BS_OWNERDRAW, IdSkipKnownBadCheck);
        wait_media_check_ = create(L"BUTTON", L"Wait for media", WS_TABSTOP | BS_OWNERDRAW,
                                   IdWaitMediaCheck);
        fragile_check_ = create(L"BUTTON", L"Fragile-media mode", WS_TABSTOP | BS_OWNERDRAW,
                                IdFragileCheck);

        buffer_label_ = create(L"STATIC", L"Buffer (MB):", SS_NOPREFIX, 0);
        buffer_edit_ = create(L"EDIT", L"", WS_TABSTOP | ES_NUMBER | ES_AUTOHSCROLL | WS_BORDER,
                              IdBufferEdit);
        retries_label_ = create(L"STATIC", L"Retries:", SS_NOPREFIX, 0);
        retries_edit_ = create(L"EDIT", L"", WS_TABSTOP | ES_NUMBER | ES_AUTOHSCROLL | WS_BORDER,
                               IdRetriesEdit);
        timeout_label_ = create(L"STATIC", L"Timeout (s):", SS_NOPREFIX, 0);
        timeout_edit_ = create(L"EDIT", L"", WS_TABSTOP | ES_NUMBER | ES_AUTOHSCROLL | WS_BORDER,
                               IdTimeoutEdit);

        start_button_ = create(L"BUTTON", L"Start", WS_TABSTOP | BS_OWNERDRAW, IdStartButton);
        pause_button_ = create(L"BUTTON", L"Pause", WS_TABSTOP | BS_OWNERDRAW, IdPauseButton);
        cancel_button_ = create(L"BUTTON", L"Cancel", WS_TABSTOP | BS_OWNERDRAW, IdCancelButton);
        save_defaults_button_ =
            create(L"BUTTON", L"Save defaults", WS_TABSTOP | BS_OWNERDRAW, IdSaveDefaultsButton);
        settings_button_ =
            create(L"BUTTON", L"Settings...", WS_TABSTOP | BS_OWNERDRAW, IdSettingsButton);
        about_button_ = create(L"BUTTON", L"About", WS_TABSTOP | BS_OWNERDRAW, IdAboutButton);
        EnableWindow(pause_button_, FALSE);
        EnableWindow(cancel_button_, FALSE);

        current_label_ = create(L"STATIC", L"Current: —", SS_ENDELLIPSIS | SS_NOPREFIX, 0);
        // Owner-drawn flat progress bars: classic msctls_progress32 draws a
        // sunken 3D border (the white edge seen on resize) once its visual
        // styles are stripped to honor a custom bar color, so paint them
        // ourselves for a flat, borderless, consistent look.
        current_bar_ = create(L"STATIC", L"", SS_OWNERDRAW, IdCurrentBar);
        overall_bar_ = create(L"STATIC", L"", SS_OWNERDRAW, IdOverallBar);
        stats_label_ = create(L"STATIC", L"", SS_NOPREFIX, 0);
        throughput_label_ =
            create(L"STATIC", L"Speed: 0 B/s (avg 0 B/s)", SS_ENDELLIPSIS | SS_NOPREFIX, 0);
        eta_label_ = create(L"STATIC", L"ETA: -", SS_NOPREFIX, 0);
        buffer_usage_label_ =
            create(L"STATIC", L"Buffer: -", SS_ENDELLIPSIS | SS_NOPREFIX, 0);
        rescue_label_ = create(L"STATIC", L"Rescue: -", SS_ENDELLIPSIS | SS_NOPREFIX, 0);
        job_summary_label_ = create(L"STATIC", L"Job: -", SS_ENDELLIPSIS | SS_NOPREFIX, 0);
        journal_label_ = create(L"STATIC", L"Journal: -", SS_ENDELLIPSIS | SS_NOPREFIX, 0);
        diagnostics_label_ =
            create(L"STATIC", L"Diagnostics: -", SS_ENDELLIPSIS | SS_NOPREFIX, 0);
        status_label_ = create(L"STATIC", L"Status: Idle", SS_NOPREFIX, 0);
        // WS_BORDER (flat 1px, matches the edits) instead of WS_EX_CLIENTEDGE,
        // which draws a white sunken 3D edge the dark theme can't repaint.
        log_list_ = create(L"LISTBOX", L"",
                           WS_TABSTOP | WS_VSCROLL | WS_HSCROLL | WS_BORDER | LBS_NOINTEGRALHEIGHT |
                               LBS_DISABLENOSCROLL | LBS_NOSEL | LBS_OWNERDRAWFIXED | LBS_HASSTRINGS,
                           IdLogList, 0, mono_font_);
        SendMessageW(log_list_, LB_SETITEMHEIGHT, 0,
                     MulDiv(log_points + 6, effective_dpi(), 72));

        tooltips_.create(hwnd_, static_cast<UINT>(effective_dpi()), theme_.dark);
        configure_tooltips();

        build_menu_bar();
        apply_theme();
        apply_appearance_settings();
        configure_diagnostics_timer();
        apply_settings_defaults();
        wire_supervisor_events();
        update_menu_state();
        layout();
    }

    void configure_tooltips() {
        auto add = [this](HWND control, const wchar_t* text) { tooltips_.add(control, text); };

        add(source_label_, L"The folder or common root XactCopy will read. Drag-and-drop and Explorer selections are also accepted.");
        add(source_edit_, L"Enter the source folder. When individual files are selected, this shows their common root while the selection row lists the subset.");
        add(source_browse_, L"Choose one source folder and replace the current source selection.");
        add(source_add_files_, L"Add one or more files or folders to an exact selection without replacing items already selected.");
        add(selection_label_, L"Summary of the exact files and folders selected beneath the displayed source root.");
        add(clear_selection_, L"Clear the exact-item selection and return to copying or scanning the entire source folder.");
        add(destination_label_, L"The folder that receives copied files. Assess Readable Files mode is read-only and does not require a destination.");
        add(destination_edit_, L"Enter the destination folder. XactCopy stages and verifies each file beside its final path before replacement.");
        add(destination_browse_, L"Choose the destination folder.");

        add(mode_combo_, L"Choose a destination-producing verified copy/recovery workflow, or a read-only assessment of accessible source files.");
        add(engine_combo_, L"In copy mode choose a safe workflow profile. In assessment mode choose Auto, Fast, or Precise scanning. Low-level engine selection remains under Settings.");
        add(overwrite_combo_, L"In copy mode this controls destination conflicts. In assessment mode it selects standard file reads or direct NTFS allocated-extent reads, which require administrator access.");
        add(verify_combo_, L"In copy mode choose destination verification. In assessment mode choose whether findings are saved to the source-specific bad-range map.");

        add(salvage_check_, L"If source blocks remain unreadable, write the configured fill pattern so readable data can still be recovered. The result is reported as recovered/incomplete, never as an exact copy.");
        add(resume_check_, L"Load the matching journal and continue interrupted work. Completed files are reused only after the configured validation allows it.");
        add(map_check_, L"Load remembered bad ranges for the same source. Only fresh, identity-matching entries are trusted as skip hints.");
        add(adaptive_check_, L"Let XactCopy grow or shrink I/O buffers from observed performance and failures instead of using one fixed size.");
        add(continue_check_, L"Continue with later files after an error. Failed or skipped files remain visible and the run cannot be reported as an exact success.");
        add(skip_known_bad_check_, L"Avoid ranges already marked bad in a compatible map. This reduces repeated stress on failing media but does not recover the skipped bytes by itself.");
        add(wait_media_check_, L"Keep waiting if the source or destination disappears, such as a removable drive being reconnected. Cancel remains available.");
        add(fragile_check_, L"Use conservative reads and stop stressing files/media after clustered failures. Recommended for unstable devices; slower on healthy storage.");

        add(buffer_label_, L"Base I/O buffer size in MiB. Smaller buffers isolate bad regions more precisely; larger buffers can improve healthy-device throughput.");
        add(buffer_edit_, L"Base I/O buffer size in MiB. Adaptive mode may adjust it while the job runs.");
        add(retries_label_, L"Additional XactCopy retries after Windows, the filesystem, storage driver, controller, and device have already applied their own retry policies.");
        add(retries_edit_, L"Enter 0-32 additional attempts for retry-capable managed I/O. High values can repeatedly stress failing media; the recommended default is 2.");
        add(timeout_label_, L"Timeout for an individual I/O operation, in seconds.");
        add(timeout_edit_, L"Enter the per-operation timeout. This is not the total job duration; retries and media-wait policies can extend the run.");

        add(start_button_, L"Validate the current options and start the copy or read-only file-readability assessment.");
        add(pause_button_, L"Pause or resume scheduling and managed I/O at a safe cancellation point. An in-flight Windows API call may finish first.");
        add(cancel_button_, L"Request cancellation. If a storage driver does not respond within five seconds, XactCopy force-stops the isolated worker. Click again to force-stop immediately.");
        add(save_defaults_button_, L"Save the visible run controls as defaults for future jobs without saving the current source, destination, or operation mode.");
        add(settings_button_, L"Open advanced appearance, copy, performance, diagnostics, verification, recovery, and Explorer-integration settings.");
        add(about_button_, L"Show version, build, components, and update options.");

        add(current_label_, L"The file or scan item currently being processed.");
        add(current_bar_, L"Progress within the current file or scan item.");
        add(overall_bar_, L"Overall progress across all enumerated source data.");
        add(stats_label_, L"Current processed-file and processed-byte totals for the active job.");
        add(throughput_label_, L"Recent transfer rate and smoothed average rate.");
        add(eta_label_, L"Estimated time remaining. It stabilizes after enough throughput samples are available.");
        add(buffer_usage_label_, L"Current adaptive-buffer size and related I/O buffer telemetry.");
        add(rescue_label_, L"Current rescue pass and unreadable-range status.");
        add(job_summary_label_, L"Active operation name and source summary.");
        add(journal_label_, L"Path of the resumable journal used by this run.");
        add(diagnostics_label_, L"UI and worker diagnostic counters. Enable or configure this strip in Settings.");
        add(log_list_, L"Detailed worker and supervisor log. Severity can be color-coded; use the horizontal scrollbar for long lines.");
        add(status_label_, L"Supervisor state for worker startup, connection, running, pause, cancellation, recovery, and completion.");
    }

    // Mirrors the .NET MainForm menu strip (File / Tools / Jobs / Help).
    void build_menu_bar() {
        HMENU file_menu = CreatePopupMenu();
        AppendMenuW(file_menu, MF_STRING, IdMenuStartCopy, L"&Start Operation");
        AppendMenuW(file_menu, MF_STRING, IdMenuPauseCopy, L"&Pause Operation");
        AppendMenuW(file_menu, MF_STRING, IdMenuResumeCopy, L"&Resume Operation");
        AppendMenuW(file_menu, MF_STRING, IdMenuCancelCopy, L"&Cancel Operation");
        AppendMenuW(file_menu, MF_SEPARATOR, 0, nullptr);
        AppendMenuW(file_menu, MF_STRING, IdMenuOpenJournals, L"Open &Journal Folder");
        AppendMenuW(file_menu, MF_STRING, IdMenuOpenCrash, L"Open &Crash Folder");
        AppendMenuW(file_menu, MF_SEPARATOR, 0, nullptr);
        AppendMenuW(file_menu, MF_STRING, IdMenuExit, L"E&xit");

        HMENU tools_menu = CreatePopupMenu();
        AppendMenuW(tools_menu, MF_STRING, IdMenuScanBadBlocks, L"&Assess Readable Files...");
        AppendMenuW(tools_menu, MF_STRING, IdMenuInspectBadMap, L"Inspect Source &Bad-Range Map...");
        AppendMenuW(tools_menu, MF_STRING, IdMenuClearBadMap, L"Clear Source Bad-Range &Map...");
        AppendMenuW(tools_menu, MF_SEPARATOR, 0, nullptr);
        AppendMenuW(tools_menu, MF_STRING, IdMenuSettings, L"&Settings...");
        AppendMenuW(tools_menu, MF_STRING, IdMenuSaveDefaults,
                    L"Save Current Settings as &Defaults");
        AppendMenuW(tools_menu, MF_STRING, IdMenuCheckUpdates, L"Check for &Updates...");

        HMENU jobs_menu = CreatePopupMenu();
        AppendMenuW(jobs_menu, MF_STRING, IdMenuSaveAsJob, L"&Save Current Options As Job...");
        AppendMenuW(jobs_menu, MF_STRING, IdMenuJobManager, L"&Job Manager...");
        AppendMenuW(jobs_menu, MF_SEPARATOR, 0, nullptr);
        AppendMenuW(jobs_menu, MF_STRING, IdMenuResumeInterrupted, L"&Resume Interrupted Job");
        AppendMenuW(jobs_menu, MF_STRING, IdMenuRunNextQueued, L"Run &Next Queued Job");

        HMENU help_menu = CreatePopupMenu();
        AppendMenuW(help_menu, MF_STRING, IdMenuAbout, L"&About XactCopy");

        menu_bar_ = CreateMenu();
        AppendMenuW(menu_bar_, MF_POPUP, reinterpret_cast<UINT_PTR>(file_menu), L"&File");
        AppendMenuW(menu_bar_, MF_POPUP, reinterpret_cast<UINT_PTR>(tools_menu), L"&Tools");
        AppendMenuW(menu_bar_, MF_POPUP, reinterpret_cast<UINT_PTR>(jobs_menu), L"&Jobs");
        AppendMenuW(menu_bar_, MF_POPUP, reinterpret_cast<UINT_PTR>(help_menu), L"&Help");
        SetMenu(hwnd_, menu_bar_);
    }

    void update_menu_state() {
        if (menu_bar_ == nullptr) return;
        bool running = supervisor_.is_job_running();
        auto set_enabled = [this](int id, bool enabled) {
            EnableMenuItem(menu_bar_, static_cast<UINT>(id),
                           MF_BYCOMMAND | (enabled ? MF_ENABLED : MF_GRAYED));
        };
        set_enabled(IdMenuStartCopy, !running);
        set_enabled(IdMenuScanBadBlocks, !running);
        set_enabled(IdMenuInspectBadMap, !running);
        set_enabled(IdMenuClearBadMap, !running);
        set_enabled(IdMenuPauseCopy, running && !paused_);
        set_enabled(IdMenuResumeCopy, running && paused_);
        set_enabled(IdMenuCancelCopy, running);
        set_enabled(IdMenuResumeInterrupted,
                    !running && recovery_.get_pending_interrupted_run().has_value());
        set_enabled(IdMenuSaveAsJob, !running);
        set_enabled(IdMenuJobManager, true);
        set_enabled(IdMenuRunNextQueued, !running);
        set_enabled(IdMenuCheckUpdates, !checking_updates_);
        DrawMenuBar(hwnd_);
    }

    static void open_folder_in_explorer(const std::wstring& folder) {
        storage::fsutil::create_directories(folder);
        ShellExecuteW(nullptr, L"open", folder.c_str(), nullptr, nullptr, SW_SHOWNORMAL);
    }

    std::optional<std::wstring> selected_source_map_path() {
        std::string source = window_text_utf8(source_edit_);
        if (ui_is_blank(source)) {
            warn_box(L"Choose a source before inspecting its bad-range map.");
            return std::nullopt;
        }
        return storage::BadRangeMapStore::get_default_map_path(
            storage::fsutil::get_full_path(storage::fsutil::utf8_to_wide(source)));
    }

    void inspect_source_bad_range_map() {
        auto path = selected_source_map_path();
        if (!path.has_value()) return;
        storage::BadRangeMapStore store;
        std::optional<storage::BadRangeMap> map;
        try {
            map = store.load(*path);
        } catch (const std::exception& ex) {
            warn_box(L"The bad-range map could not be authenticated or opened.\n\n" +
                         storage::fsutil::utf8_to_wide(ex.what()),
                     L"Bad-range map");
            return;
        }
        std::optional<storage::ManagedJobRun> latest_assessment_run;
        ui::RunJournalInspection latest_assessment;
        const std::wstring selected_source = storage::fsutil::normalize_path(
            storage::fsutil::utf8_to_wide(window_text_utf8(source_edit_)));
        for (const auto& run : job_manager_.get_recent_runs(1000)) {
            const bool scan_mode = run.Options.has_value()
                                       ? run.Options->OperationMode ==
                                             models::JobOperationMode::ScanOnly
                                       : models::detail::equals_ignore_case(run.Trigger, "scan") ||
                                             storage::catalog_detail::is_blank(
                                                 run.DestinationRoot);
            if (!scan_mode || storage::catalog_detail::is_blank(run.SourceRoot) ||
                storage::fsutil::normalize_path(
                    storage::fsutil::utf8_to_wide(run.SourceRoot)) != selected_source) {
                continue;
            }
            ui::RunJournalInspection inspection =
                job_manager_.inspect_run_journal(run.RunId);
            if (!inspection.Loaded) continue;
            latest_assessment_run = run;
            latest_assessment = std::move(inspection);
            break;
        }

        if (!map.has_value()) {
            if (!latest_assessment_run.has_value()) {
                info_box(L"No authenticated bad-range map or readable assessment journal "
                         L"exists for the selected source.\n\n" + *path,
                         L"Bad-range map");
                return;
            }
            storage::BadRangeMap journal_only_map;
            journal_only_map.SourceRoot = window_text_utf8(source_edit_);
            journal_only_map.SourceIdentity = "(no map)";
            map = std::move(journal_only_map);
        }

        std::int64_t bad_bytes = 0;
        std::size_t range_count = 0;
        std::size_t confirmed_files = 0;
        std::size_t finding_files = 0;
        for (const auto& [key, entry] : map->Files.entries) {
            if (entry.BadRanges.empty()) continue;
            ++finding_files;
            if (entry.ConfirmationCount >= 2 && !entry.BadRanges.empty()) ++confirmed_files;
            range_count += entry.BadRanges.size();
            std::int64_t file_bad_bytes = 0;
            for (const auto& range : entry.BadRanges) {
                if (range.Length > 0) {
                    bad_bytes += range.Length;
                    file_bad_bytes += range.Length;
                }
            }
            append_log("Bad-range map finding: " +
                       (entry.RelativePath.empty() ? key : entry.RelativePath) + " — " +
                       std::to_string(entry.BadRanges.size()) + " range(s), " +
                       format_bytes_short(file_bad_bytes) +
                       (entry.ConfirmationCount >= 2 ? " (confirmed)." : " (observed once)."));
        }

        std::size_t journal_finding_files = 0;
        std::size_t journal_localized_files = 0;
        std::size_t journal_unlocalized_files = 0;
        std::int64_t journal_unreadable_bytes = 0;
        if (latest_assessment_run.has_value()) {
            journal_finding_files = latest_assessment.Findings.size();
            for (const auto& finding : latest_assessment.Findings) {
                if (finding.UnreadableRangeCount > 0) {
                    ++journal_localized_files;
                    journal_unreadable_bytes += finding.UnreadableBytes;
                } else {
                    ++journal_unlocalized_files;
                }
                std::string detail = finding.UnreadableRangeCount > 0
                                         ? std::to_string(finding.UnreadableRangeCount) +
                                               " unreadable range(s), " +
                                               format_bytes_short(finding.UnreadableBytes)
                                         : "failure/skip without a localized byte range";
                if (!finding.ErrorMessage.empty()) detail += "; " + finding.ErrorMessage;
                append_log("Latest assessment finding: " + finding.RelativePath + " - " +
                           detail + ".");
            }
        }

        const std::wstring map_updated = map->UpdatedUtc == time::DateTimeOffset::min_value()
                                             ? L"(no map)"
                                             : storage::fsutil::utf8_to_wide(
                                                   map->UpdatedUtc.to_string());
        std::wstring summary =
            L"Source: " + storage::fsutil::utf8_to_wide(map->SourceRoot) +
            L"\nMedia identity: " + storage::fsutil::utf8_to_wide(map->SourceIdentity) +
            L"\nMap updated: " + map_updated +
            L"\n\nBad-range map" +
            L"\nFiles with findings: " + std::to_wstring(finding_files) +
            L"\nFiles with confirmed skip hints: " + std::to_wstring(confirmed_files) +
            L"\nObserved ranges: " + std::to_wstring(range_count) +
            L"\nBytes covered by findings: " +
            storage::fsutil::utf8_to_wide(format_bytes_short(bad_bytes));

        if (latest_assessment_run.has_value()) {
            summary +=
                std::wstring(L"\n\nLatest assessment journal") +
                L"\nRun: " + storage::fsutil::utf8_to_wide(
                                  latest_assessment_run->DisplayName) +
                L"\nStatus: " + storage::fsutil::utf8_to_wide(
                                     std::string(storage::to_string(
                                         latest_assessment_run->Status))) +
                L"\nFiles with findings/failures: " +
                std::to_wstring(journal_finding_files) +
                L"\nFiles with localized ranges: " +
                std::to_wstring(journal_localized_files) +
                L"\nFiles without localized ranges: " +
                std::to_wstring(journal_unlocalized_files) +
                L"\nLocalized unreadable bytes: " +
                storage::fsutil::utf8_to_wide(
                    format_bytes_short(journal_unreadable_bytes));
        }

        summary +=
            L"\n\nAll filenames and details were listed in the main log. A range observed once "
            L"is diagnostic only; it becomes a skip hint after a later matching observation. "
            L"Journal failures without a localized range are not inserted into the range map."
            L"\n\nMap: " + *path;
        info_box(summary, L"Bad-range map");
    }

    void clear_source_bad_range_map() {
        auto path = selected_source_map_path();
        if (!path.has_value()) return;
        if (!confirm_box(
                L"Remove the selected source's authenticated bad-range map, backups, and mirror?\n\n"
                L"This does not modify the source media. Future runs will rediscover unreadable ranges.\n\n" +
                    *path,
                L"Clear bad-range map", ui::MessageIcon::Warning)) {
            return;
        }
        const std::int64_t reclaimed = storage::BadRangeMapStore::remove_map_set(*path);
        if (reclaimed <= 0) {
            info_box(L"No bad-range map artifacts were found for the selected source.",
                     L"Bad-range map");
            return;
        }
        append_log("Cleared source bad-range map and trusted backups (" +
                   format_bytes_short(reclaimed) + ").");
        info_box(L"The map and its trusted backups were removed. The source media was not changed.",
                 L"Bad-range map cleared");
    }

    void apply_theme() {
        theme_ = ui::make_theme(settings_.theme(),
                                settings_.get_string("AccentColorMode", "auto"),
                                settings_.get_string("AccentColorHex", "#5A78C8"));
        theme_.themed_chrome = !models::detail::equals_ignore_case(
            settings_.get_string("WindowChromeMode", "themed"), "standard");
        ui_scale_percent_ = settings_.get_int("UiScalePercent", 50, 250, 100);
        ui_density_percent_ = ui::ui_density_percent(settings_.get_string("UiDensity", "normal"));
        theme_.density_percent = ui_density_percent_;
        theme_.scale_percent = ui_scale_percent_;
        ui::configure_theme_engine(settings_.theme(), theme_);
        ui::set_dark_title_bar(hwnd_, theme_.dark && theme_.themed_chrome);
        tooltips_.update_dpi(static_cast<UINT>(effective_dpi()));
        tooltips_.apply_theme(theme_.dark);
        if (window_brush_ != nullptr) DeleteObject(window_brush_);
        if (edit_brush_ != nullptr) DeleteObject(edit_brush_);
        window_brush_ = CreateSolidBrush(theme_.window);
        edit_brush_ = CreateSolidBrush(theme_.edit);

        ui::attach_theme_window(hwnd_);
        for (HWND control : {log_list_, source_edit_, destination_edit_, buffer_edit_,
                             retries_edit_, timeout_edit_}) {
            ui::apply_control_theme(control, theme_.dark);
        }
        for (HWND combo : {mode_combo_, engine_combo_, overwrite_combo_, verify_combo_}) {
            ui::apply_combo_theme(combo, theme_.dark);
        }
        InvalidateRect(current_bar_, nullptr, FALSE);
        InvalidateRect(overall_bar_, nullptr, FALSE);
        InvalidateRect(hwnd_, nullptr, TRUE);
    }

    static void replace_combo_items(HWND combo,
                                    std::initializer_list<const wchar_t*> items) {
        SendMessageW(combo, CB_RESETCONTENT, 0, 0);
        for (const wchar_t* item : items) {
            SendMessageW(combo, CB_ADDSTRING, 0, reinterpret_cast<LPARAM>(item));
        }
    }

    bool is_scan_mode() const {
        return mode_combo_ != nullptr &&
               SendMessageW(mode_combo_, CB_GETCURSEL, 0, 0) == 1;
    }

    void apply_copy_profile(int profile) {
        copy_profile_index_ = std::clamp(profile, 0, 2);
        if (copy_profile_index_ == 2) return;

        applying_copy_profile_ = true;
        const bool recovery = copy_profile_index_ == 1;
        set_check(IdSalvageCheck, recovery);
        set_check(IdContinueCheck, recovery);
        set_check(IdFragileCheck, recovery);
        set_check(IdAdaptiveCheck, recovery);
        set_check(IdWaitMediaCheck, recovery);
        if (recovery) {
            set_check(IdMapCheck, true);
            set_check(IdSkipKnownBadCheck, false);
        }
        copy_verify_index_ = 2;
        SendMessageW(verify_combo_, CB_SETCURSEL, copy_verify_index_, 0);
        SetWindowTextW(buffer_edit_, recovery ? L"1" : L"4");
        SetWindowTextW(retries_edit_, L"2");
        SetWindowTextW(timeout_edit_, L"10");
        sync_bad_range_controls();
        applying_copy_profile_ = false;
        InvalidateRect(hwnd_, nullptr, FALSE);
    }

    void mark_copy_profile_custom() {
        if (is_scan_mode() || applying_copy_profile_ || copy_profile_index_ == 2) return;
        copy_profile_index_ = 2;
        SendMessageW(engine_combo_, CB_SETCURSEL, copy_profile_index_, 0);
    }

    void sync_mode_ui(bool user_changed = false) {
        int requested = static_cast<int>(SendMessageW(mode_combo_, CB_GETCURSEL, 0, 0));
        if (requested != 1) requested = 0;

        if (active_mode_index_ == 0) {
            copy_profile_index_ = std::max(
                0, static_cast<int>(SendMessageW(engine_combo_, CB_GETCURSEL, 0, 0)));
            copy_overwrite_index_ = std::max(
                0, static_cast<int>(SendMessageW(overwrite_combo_, CB_GETCURSEL, 0, 0)));
            copy_verify_index_ = std::max(
                0, static_cast<int>(SendMessageW(verify_combo_, CB_GETCURSEL, 0, 0)));
        } else if (active_mode_index_ == 1) {
            scan_profile_index_ = std::max(
                0, static_cast<int>(SendMessageW(engine_combo_, CB_GETCURSEL, 0, 0)));
            scan_backend_index_ = std::max(
                0, static_cast<int>(SendMessageW(overwrite_combo_, CB_GETCURSEL, 0, 0)));
            scan_map_index_ = std::max(
                0, static_cast<int>(SendMessageW(verify_combo_, CB_GETCURSEL, 0, 0)));
        }

        active_mode_index_ = requested;
        if (requested == 1) {
            replace_combo_items(engine_combo_, {L"Scan: Auto", L"Scan: Fast", L"Scan: Precise"});
            replace_combo_items(overwrite_combo_,
                                {L"Reads: Standard Files", L"Reads: Direct NTFS (Admin)"});
            replace_combo_items(verify_combo_, {L"Map: Save Findings", L"Map: Do Not Save"});
            SendMessageW(engine_combo_, CB_SETCURSEL, scan_profile_index_, 0);
            SendMessageW(overwrite_combo_, CB_SETCURSEL, scan_backend_index_, 0);
            SendMessageW(verify_combo_, CB_SETCURSEL, scan_map_index_, 0);
            SetWindowTextW(resume_check_, L"Reuse previous scan progress");
            SetWindowTextW(start_button_, L"Start assessment");
        } else {
            replace_combo_items(engine_combo_,
                                {L"Profile: Verified Copy", L"Profile: Recover Media",
                                 L"Profile: Custom"});
            replace_combo_items(overwrite_combo_,
                                {L"Overwrite", L"Skip existing", L"Overwrite if newer",
                                 L"Stop on conflict"});
            replace_combo_items(verify_combo_,
                                {L"Verify: None", L"Verify: Sampled", L"Verify: Full"});
            SendMessageW(engine_combo_, CB_SETCURSEL, copy_profile_index_, 0);
            SendMessageW(overwrite_combo_, CB_SETCURSEL, copy_overwrite_index_, 0);
            SendMessageW(verify_combo_, CB_SETCURSEL, copy_verify_index_, 0);
            SetWindowTextW(resume_check_, L"Reuse verified previous progress");
            SetWindowTextW(start_button_, L"Start copy");
            if (user_changed) launch_.explorer_scan_mode = false;
        }

        const bool scan = requested == 1;
        for (HWND control : {destination_label_, destination_edit_, destination_browse_,
                             salvage_check_}) {
            if (control != nullptr) ShowWindow(control, scan ? SW_HIDE : SW_SHOW);
        }
        layout();
    }

    void apply_settings_defaults() {
        auto defaults = settings_.build_default_options();
        SendMessageW(mode_combo_, CB_SETCURSEL, 0, 0);
        custom_transfer_engine_ = defaults.TransferEnginePolicyValue;
        copy_profile_index_ = 2;
        copy_overwrite_index_ = static_cast<int>(defaults.OverwritePolicyValue);
        copy_verify_index_ = ui::verification_combo_index(defaults);
        scan_profile_index_ = static_cast<int>(defaults.ScanPerformanceProfileValue);
        scan_backend_index_ = defaults.UseExperimentalRawDiskScan ? 1 : 0;
        // A new installation saves scan findings by default. An explicit
        // existing setting still wins, including a deliberate read-only map
        // policy saved from the assessment UI.
        scan_map_index_ = settings_.contains("DefaultUpdateBadRangeMapFromRun")
                              ? (defaults.UpdateBadRangeMapFromRun ? 0 : 1)
                              : 0;
        active_mode_index_ = -1;
        sync_mode_ui();
        set_check(IdSalvageCheck, defaults.SalvageUnreadableBlocks);
        set_check(IdResumeCheck, defaults.ResumeFromJournal);
        set_check(IdMapCheck, defaults.UseBadRangeMap);
        set_check(IdAdaptiveCheck, defaults.UseAdaptiveBufferSizing);
        set_check(IdContinueCheck, defaults.ContinueOnFileError);
        set_check(IdSkipKnownBadCheck, defaults.SkipKnownBadRanges);
        set_check(IdWaitMediaCheck, defaults.WaitForMediaAvailability);
        set_check(IdFragileCheck, defaults.FragileMediaMode);
        sync_bad_range_controls();

        int buffer_mb = std::max(1, defaults.BufferSizeBytes / (1024 * 1024));
        SetWindowTextW(buffer_edit_, std::to_wstring(buffer_mb).c_str());
        SetWindowTextW(retries_edit_, std::to_wstring(std::max(0, defaults.MaxRetries)).c_str());
        std::int64_t timeout_seconds = defaults.OperationTimeout.ticks / 10000000LL;
        if (timeout_seconds <= 0) timeout_seconds = 30;
        SetWindowTextW(timeout_edit_, std::to_wstring(timeout_seconds).c_str());

        show_diagnostics_ = settings_.get_bool("UiShowDiagnostics", false);

        if (!launch_.explorer_folder_path.empty()) {
            SetWindowTextW(source_edit_,
                           storage::fsutil::utf8_to_wide(launch_.explorer_folder_path).c_str());
        }
    }

    void wire_supervisor_events() {
        HWND hwnd = hwnd_;
        supervisor_.on_log = [hwnd](const std::string& message) {
            PostMessageW(hwnd, WM_APP_LOG, 0, reinterpret_cast<LPARAM>(new std::string(message)));
        };
        supervisor_.on_progress = [hwnd](const models::CopyProgressSnapshot& snapshot) {
            PostMessageW(hwnd, WM_APP_PROGRESS, 0,
                         reinterpret_cast<LPARAM>(new models::CopyProgressSnapshot(snapshot)));
        };
        supervisor_.on_job_completed = [hwnd](const models::CopyJobResult& result) {
            PostMessageW(hwnd, WM_APP_RESULT, 0,
                         reinterpret_cast<LPARAM>(new models::CopyJobResult(result)));
        };
        supervisor_.on_worker_state_changed = [hwnd](const std::string& state) {
            PostMessageW(hwnd, WM_APP_STATE, 0, reinterpret_cast<LPARAM>(new std::string(state)));
        };
        supervisor_.on_job_pause_state_changed = [hwnd](bool paused) {
            PostMessageW(hwnd, WM_APP_PAUSE, paused ? 1 : 0, 0);
        };
    }

    // ---- Owner-draw + checkbox state --------------------------------------

    bool get_check(int id) const {
        auto found = check_states_.find(id);
        return found != check_states_.end() && found->second;
    }

    void set_check(int id, bool value) {
        check_states_[id] = value;
        HWND control = GetDlgItem(hwnd_, id);
        if (control != nullptr) InvalidateRect(control, nullptr, FALSE);
    }

    void sync_bad_range_controls() {
        if (skip_known_bad_check_ == nullptr) return;
        bool enabled = get_check(IdMapCheck);
        // A running job disables all parameter inputs. Preserve that state when
        // the dependency is refreshed after a checkbox click or job completion.
        if (map_check_ != nullptr && !IsWindowEnabled(map_check_)) enabled = false;
        EnableWindow(skip_known_bad_check_, enabled ? TRUE : FALSE);
    }

    static bool is_checkbox_id(int id) {
        return id == IdSalvageCheck || id == IdResumeCheck || id == IdMapCheck ||
               id == IdAdaptiveCheck || id == IdContinueCheck || id == IdSkipKnownBadCheck ||
               id == IdWaitMediaCheck || id == IdFragileCheck;
    }

    static bool is_button_id(int id) {
        return id == IdSourceBrowse || id == IdDestinationBrowse || id == IdStartButton ||
               id == IdPauseButton || id == IdCancelButton || id == IdSettingsButton ||
               id == IdAboutButton || id == IdSourceAddFiles || id == IdClearSelection ||
               id == IdSaveDefaultsButton;
    }

    bool on_draw_item(const DRAWITEMSTRUCT& draw) {
        int id = static_cast<int>(draw.CtlID);
        if (draw.CtlType == ODT_BUTTON && is_button_id(id)) {
            ui::themedraw::draw_button(draw, theme_);
            return true;
        }
        if (draw.CtlType == ODT_BUTTON && is_checkbox_id(id)) {
            ui::themedraw::draw_checkbox(draw, theme_, get_check(id));
            return true;
        }
        if (draw.CtlType == ODT_COMBOBOX) {
            ui::themedraw::draw_combo_item(draw, theme_);
            return true;
        }
        if (draw.CtlType == ODT_LISTBOX && id == IdLogList) {
            draw_log_item(draw);
            return true;
        }
        if (draw.CtlType == ODT_STATIC && (id == IdCurrentBar || id == IdOverallBar)) {
            draw_progress_bar(draw, id == IdCurrentBar ? current_permille_ : overall_permille_);
            return true;
        }
        return false;
    }

    // Flat, borderless progress bar: track + accent fill + subtle 1px frame,
    // with an optional centered percentage (ProgressBarShowPercentage setting).
    void draw_progress_bar(const DRAWITEMSTRUCT& draw, int permille) {
        RECT rect = draw.rcItem;
        int clamped = std::clamp(permille, 0, 1000);
        ui::themedraw::fill_rect(draw.hDC, rect, theme_.panel);
        int span = rect.right - rect.left;
        int filled = MulDiv(span, clamped, 1000);
        if (filled > 0) {
            RECT fill = rect;
            fill.right = rect.left + filled;
            ui::themedraw::fill_rect(draw.hDC, fill, theme_.accent);
        }
        ui::themedraw::frame_rect(draw.hDC, rect, theme_.border);
        if (bar_show_percentage_ && (rect.bottom - rect.top) >= 14) {
            wchar_t text[8];
            swprintf(text, 8, L"%d%%", (clamped + 5) / 10);
            HGDIOBJ old_font = SelectObject(draw.hDC, font_);
            SetBkMode(draw.hDC, TRANSPARENT);
            // Readable over both the filled and unfilled halves.
            SetTextColor(draw.hDC, ui::readable_text_color(clamped >= 500 ? theme_.accent
                                                                          : theme_.panel));
            DrawTextW(draw.hDC, text, -1, &rect, DT_CENTER | DT_VCENTER | DT_SINGLELINE);
            SelectObject(draw.hDC, old_font);
        }
    }

    void set_progress_bar(HWND bar, int& store, int permille) {
        int clamped = std::clamp(permille, 0, 1000);
        if (clamped == store) return;
        store = clamped;
        InvalidateRect(bar, nullptr, FALSE);
    }

    // Consume the appearance settings that affect the main window's chrome.
    void apply_appearance_settings() {
        show_diagnostics_ = settings_.get_bool("UiShowDiagnostics", false);
        show_buffer_row_ = settings_.get_bool("ShowBufferStatusRow", true);
        show_rescue_row_ = settings_.get_bool("ShowRescueStatusRow", true);
        bar_show_percentage_ = settings_.get_bool("ProgressBarShowPercentage", false);
        const std::string bar_style = settings_.get_string("ProgressBarStyle", "standard");
        progress_bar_height_logical_ =
            models::detail::equals_ignore_case(bar_style, "thin")
                ? 10
                : models::detail::equals_ignore_case(bar_style, "thick") ? 18 : 14;
    }

    void configure_diagnostics_timer() {
        diagnostics_refresh_ms_ = settings_.get_int("UiDiagnosticsRefreshMs", 100, 5000, 250);
        if (hwnd_ != nullptr) {
            KillTimer(hwnd_, IdDiagnosticsTimer);
            SetTimer(hwnd_, IdDiagnosticsTimer, static_cast<UINT>(diagnostics_refresh_ms_), nullptr);
        }
    }

    // ---- Window placement persistence -------------------------------------

    bool save_window_placement() {
        WINDOWPLACEMENT wp{};
        wp.length = sizeof(wp);
        if (!GetWindowPlacement(hwnd_, &wp)) return true;
        const RECT& r = wp.rcNormalPosition;
        return settings_.update_and_save([&](ui::AppSettings& target) {
            target.set_int("MainWindowX", r.left);
            target.set_int("MainWindowY", r.top);
            target.set_int("MainWindowWidth", r.right - r.left);
            target.set_int("MainWindowHeight", r.bottom - r.top);
            target.set_int("MainWindowDpi", static_cast<int>(monitor_dpi_));
            target.set_bool("MainWindowMaximized", wp.showCmd == SW_SHOWMAXIMIZED);
        });
    }

    void restore_window_placement() {
        int w = settings_.get_int("MainWindowWidth", 0, 20000, 0);
        int h = settings_.get_int("MainWindowHeight", 0, 20000, 0);
        if (w < 400 || h < 300) return; // unset or implausible — keep defaults
        int x = settings_.get_int("MainWindowX", -32000, 32000, 0);
        int y = settings_.get_int("MainWindowY", -32000, 32000, 0);
        const int saved_dpi = settings_.get_int("MainWindowDpi", 0, 960, 0);

        // Clamp onto a visible monitor so a saved position from an unplugged
        // display doesn't strand the window off-screen.
        RECT desired{x, y, x + w, y + h};
        HMONITOR monitor = MonitorFromRect(&desired, MONITOR_DEFAULTTONEAREST);
        MONITORINFO mi{};
        mi.cbSize = sizeof(mi);
        if (GetMonitorInfoW(monitor, &mi)) {
            const RECT& wa = mi.rcWork;
            if (x + w <= wa.left + 40 || x >= wa.right - 40 || y + h <= wa.top + 40 ||
                y >= wa.bottom - 40) {
                x = wa.left + ((wa.right - wa.left) - w) / 2;
                y = wa.top + ((wa.bottom - wa.top) - h) / 2;
            }
        }
        // Crossing monitors while applying a saved size in one SetWindowPos
        // lets WM_DPICHANGED scale that size again. Move first so the window
        // adopts the target monitor's DPI, then apply the saved physical size.
        // New placements carry their source DPI and can also survive a later
        // monitor-scaling change; legacy placements remain byte-for-byte sized.
        SetWindowPos(hwnd_, nullptr, x, y, 0, 0,
                     SWP_NOSIZE | SWP_NOZORDER | SWP_NOACTIVATE);
        const UINT target_dpi = GetDpiForWindow(hwnd_);
        if (saved_dpi > 0 && target_dpi > 0) {
            w = MulDiv(w, static_cast<int>(target_dpi), saved_dpi);
            h = MulDiv(h, static_cast<int>(target_dpi), saved_dpi);
        }
        SetWindowPos(hwnd_, nullptr, x, y, w, h, SWP_NOZORDER | SWP_NOACTIVATE);
        restore_maximized_ = settings_.get_bool("MainWindowMaximized", false);
    }

    // Job-parameter inputs can't change mid-run, so lock them while a scan/copy
    // is active (Pause/Cancel, Settings, About and the log stay usable). Called
    // with false when a job starts and true when it ends.
    void set_inputs_enabled(bool enabled) {
        HWND inputs[] = {
            source_edit_,      destination_edit_,   source_browse_,   destination_browse_,
            source_add_files_, clear_selection_,
            mode_combo_,       engine_combo_,       overwrite_combo_, verify_combo_,
            salvage_check_,    resume_check_,       map_check_,       adaptive_check_,
            continue_check_,   skip_known_bad_check_, wait_media_check_, fragile_check_,
            buffer_edit_,      retries_edit_,       timeout_edit_,
        };
        for (HWND control : inputs) {
            if (control != nullptr) EnableWindow(control, enabled ? TRUE : FALSE);
        }
        sync_bad_range_controls();
    }

    void draw_log_item(const DRAWITEMSTRUCT& draw) {
        ui::themedraw::fill_rect(draw.hDC, draw.rcItem, theme_.edit);
        if (draw.itemID == static_cast<UINT>(-1)) return;

        const LRESULT length = SendMessageW(log_list_, LB_GETTEXTLEN, draw.itemID, 0);
        if (length < 0) return;
        std::wstring text(static_cast<std::size_t>(length) + 1, L'\0');
        SendMessageW(log_list_, LB_GETTEXT, draw.itemID, reinterpret_cast<LPARAM>(text.data()));
        text.resize(static_cast<std::size_t>(length));

        COLORREF color = theme_.text;
        if (colorize_log_) {
            color = ui::severity_color(
                theme_, ui::classify_log_line(storage::fsutil::wide_to_utf8(text)));
        }

        HGDIOBJ old_font = SelectObject(draw.hDC, mono_font_);
        SetBkMode(draw.hDC, TRANSPARENT);
        SetTextColor(draw.hDC, color);
        RECT text_rect = draw.rcItem;
        text_rect.left += 4;
        DrawTextW(draw.hDC, text.c_str(), -1, &text_rect,
                  DT_LEFT | DT_VCENTER | DT_SINGLELINE | DT_NOPREFIX);
        SelectObject(draw.hDC, old_font);
    }

    // ---- Layout ------------------------------------------------------------

    void layout() {
        RECT client;
        GetClientRect(hwnd_, &client);
        int dpi = effective_dpi();
        auto scale = [dpi](int value) { return MulDiv(value, dpi, 96); };

        const int margin = scale(12);
        const int row_height = scale(24);
        const int gap = scale(8);
        int width = client.right - client.left - margin * 2;
        int y = margin;

        auto place = [&](HWND handle, int x, int w, int h = 0) {
            if (handle != nullptr) {
                MoveWindow(handle, margin + x, y, w, h == 0 ? row_height : h, TRUE);
                // MoveWindow only invalidates a resized control's newly exposed
                // strip, so widened statics/bars keep stale pixels. Force a full
                // repaint (WS_CLIPCHILDREN keeps the parent from smearing them).
                InvalidateRect(handle, nullptr, TRUE);
            }
        };

        int label_width = scale(84);
        int browse_width = scale(84);
        int add_width = scale(96);
        // The source row carries an extra "Add Files..." button.
        int source_edit_width = width - label_width - browse_width - add_width - gap * 3;
        int edit_width = width - label_width - browse_width - gap * 2;

        place(source_label_, 0, label_width);
        place(source_edit_, label_width + gap, source_edit_width);
        place(source_add_files_, label_width + gap + source_edit_width + gap, add_width);
        place(source_browse_, label_width + gap + source_edit_width + gap + add_width + gap,
              browse_width);
        y += row_height + gap;

        // Selection summary row, shown only while a subset of the source is queued.
        if (IsWindowVisible(selection_label_)) {
            const int clear_width = scale(64);
            place(selection_label_, label_width + gap,
                  width - label_width - clear_width - gap * 2, scale(20));
            place(clear_selection_, width - clear_width, clear_width, scale(20));
            y += scale(20) + gap;
        }

        if (!is_scan_mode()) {
            place(destination_label_, 0, label_width);
            place(destination_edit_, label_width + gap, edit_width);
            place(destination_browse_, label_width + gap + edit_width + gap, browse_width);
            y += row_height + gap;
        }

        int combo_width = (width - gap * 3) / 4;
        place(mode_combo_, 0, combo_width, row_height * 8);
        place(engine_combo_, combo_width + gap, combo_width, row_height * 8);
        place(overwrite_combo_, (combo_width + gap) * 2, combo_width, row_height * 8);
        place(verify_combo_, (combo_width + gap) * 3, combo_width, row_height * 8);
        y += row_height + gap;

        int check_width = (width - gap * 3) / 4;
        if (is_scan_mode()) {
            place(resume_check_, 0, check_width);
            place(map_check_, check_width + gap, check_width);
            place(adaptive_check_, (check_width + gap) * 2, check_width);
        } else {
            place(salvage_check_, 0, check_width);
            place(resume_check_, check_width + gap, check_width);
            place(map_check_, (check_width + gap) * 2, check_width);
            place(adaptive_check_, (check_width + gap) * 3, check_width);
        }
        y += row_height + gap;

        place(continue_check_, 0, check_width);
        place(skip_known_bad_check_, check_width + gap, check_width);
        place(wait_media_check_, (check_width + gap) * 2, check_width);
        place(fragile_check_, (check_width + gap) * 3, check_width);
        y += row_height + gap;

        // Numeric row: three label+edit pairs.
        int field_label_width = scale(72);
        int field_edit_width = scale(70);
        int field_slot = field_label_width + gap + field_edit_width + scale(16);
        int fx = 0;
        auto place_field = [&](HWND label, HWND edit) {
            place(label, fx, field_label_width);
            place(edit, fx + field_label_width + gap, field_edit_width);
            fx += field_slot;
        };
        place_field(buffer_label_, buffer_edit_);
        place_field(retries_label_, retries_edit_);
        place_field(timeout_label_, timeout_edit_);
        y += row_height + gap;

        int button_width = scale(110);
        place(start_button_, 0, button_width, scale(30));
        place(pause_button_, button_width + gap, button_width, scale(30));
        place(cancel_button_, (button_width + gap) * 2, button_width, scale(30));
        int about_width = scale(80);
        int settings_width = scale(100);
        int save_defaults_width = scale(126);
        place(about_button_, width - about_width, about_width, scale(30));
        place(settings_button_, width - about_width - gap - settings_width, settings_width,
              scale(30));
        place(save_defaults_button_,
              width - about_width - gap - settings_width - gap - save_defaults_width,
              save_defaults_width, scale(30));
        y += scale(30) + gap;

        place(current_label_, 0, width);
        y += row_height;
        const int bar_height = scale(progress_bar_height_logical_);
        place(current_bar_, 0, width, bar_height);
        y += bar_height + gap;
        place(overall_bar_, 0, width, bar_height);
        y += bar_height + scale(2);
        place(stats_label_, 0, width);
        y += row_height;

        // Telemetry strip: two columns of stacked labels. Rows can be hidden via
        // the ShowBufferStatusRow / ShowRescueStatusRow settings.
        int col_width = (width - gap) / 2;
        int label_height = scale(18);
        auto place_two = [&](HWND left, HWND right) {
            place(left, 0, col_width, label_height);
            place(right, col_width + gap, col_width, label_height);
            y += label_height;
        };
        auto show_hide = [](HWND h, bool visible) {
            if (h != nullptr) ShowWindow(h, visible ? SW_SHOW : SW_HIDE);
        };
        place_two(throughput_label_, eta_label_);
        show_hide(buffer_usage_label_, show_buffer_row_);
        show_hide(rescue_label_, show_rescue_row_);
        if (show_buffer_row_ || show_rescue_row_) {
            place_two(show_buffer_row_ ? buffer_usage_label_ : nullptr,
                      show_rescue_row_ ? rescue_label_ : nullptr);
        }
        place_two(job_summary_label_, journal_label_);
        if (show_diagnostics_) {
            if (diagnostics_label_ != nullptr) ShowWindow(diagnostics_label_, SW_SHOW);
            place(diagnostics_label_, 0, width, label_height);
            y += label_height;
        } else if (diagnostics_label_ != nullptr) {
            ShowWindow(diagnostics_label_, SW_HIDE);
        }
        y += gap;

        int status_height = row_height;
        int log_height = client.bottom - y - status_height - margin - gap;
        if (log_height < scale(60)) log_height = scale(60);
        place(log_list_, 0, width, log_height);
        y += log_height + gap;
        place(status_label_, 0, width);
        tooltips_.update_layout();
    }

    // ---- Commands ----------------------------------------------------------

    void on_command(int id, int notification) {
        if (id == IdSourceEdit && notification == EN_CHANGE && !suppress_source_change_ &&
            !explorer_selected_paths_.empty()) {
            clear_selection(/*restore_root*/ false);
            append_log("Exact-item selection cleared because the Source field was edited.");
            return;
        }
        if (id == IdModeCombo && notification == CBN_SELCHANGE) {
            sync_mode_ui(true);
            return;
        }
        if (id == IdEngineCombo && notification == CBN_SELCHANGE && !is_scan_mode()) {
            apply_copy_profile(static_cast<int>(
                SendMessageW(engine_combo_, CB_GETCURSEL, 0, 0)));
            return;
        }
        if (is_checkbox_id(id) && notification == BN_CLICKED) {
            set_check(id, !get_check(id));
            if (id == IdMapCheck) sync_bad_range_controls();
            mark_copy_profile_custom();
            return;
        }
        if (!is_scan_mode() && notification == CBN_SELCHANGE &&
            (id == IdOverwriteCombo || id == IdVerifyCombo)) {
            mark_copy_profile_custom();
            return;
        }
        if (!is_scan_mode() && notification == EN_CHANGE &&
            (id == IdBufferEdit || id == IdRetriesEdit || id == IdTimeoutEdit)) {
            mark_copy_profile_custom();
        }
        switch (id) {
            case IdSourceBrowse:
                // Choosing a source folder replaces any queued item selection.
                if (!selection_.empty()) clear_selection(/*restore_root*/ false);
                browse_folder(source_edit_);
                break;
            case IdSourceAddFiles:
                add_files_to_selection();
                break;
            case IdClearSelection:
                clear_selection();
                break;
            case IdDestinationBrowse:
                browse_folder(destination_edit_);
                break;
            case IdStartButton:
                start_job();
                break;
            case IdPauseButton:
                if (paused_) supervisor_.resume_job("");
                else supervisor_.pause_job("");
                break;
            case IdCancelButton:
                supervisor_.cancel_job("");
                break;
            case IdSettingsButton:
                if (ui::SettingsDialog::show(hwnd_, settings_, theme_)) {
                    colorize_log_ = settings_.get_bool("UiColorizeLogBySeverity", true);
                    show_diagnostics_ = settings_.get_bool("UiShowDiagnostics", false);
                    ui_scale_percent_ = settings_.get_int("UiScalePercent", 50, 250, 100);
                    ui_density_percent_ =
                        ui::ui_density_percent(settings_.get_string("UiDensity", "normal"));
                    apply_appearance_settings();
                    configure_diagnostics_timer();
                    rebuild_fonts(GetDpiForWindow(hwnd_));
                    apply_fonts();
                    apply_theme();
                    layout();
                    sync_explorer_integration(true);
                }
                break;
            case IdSaveDefaultsButton:
            case IdMenuSaveDefaults:
                save_current_as_defaults();
                break;
            case IdAboutButton:
            case IdMenuAbout:
                ui::AboutDialog::show(hwnd_, theme_, settings_, IdMenuCheckUpdates);
                break;
            case IdMenuStartCopy:
                start_job();
                break;
            case IdMenuPauseCopy:
                supervisor_.pause_job("");
                break;
            case IdMenuResumeCopy:
                supervisor_.resume_job("");
                break;
            case IdMenuCancelCopy:
                supervisor_.cancel_job("");
                break;
            case IdMenuOpenJournals:
                open_folder_in_explorer(storage::fsutil::local_app_data() +
                                        L"\\XactCopy\\journals");
                break;
            case IdMenuOpenCrash:
                open_folder_in_explorer(storage::fsutil::local_app_data() +
                                        L"\\XactCopy\\runtime");
                break;
            case IdMenuExit:
                PostMessageW(hwnd_, WM_CLOSE, 0, 0);
                break;
            case IdMenuScanBadBlocks:
                SendMessageW(mode_combo_, CB_SETCURSEL, 1, 0); // Assess Readable Files
                sync_mode_ui();
                start_job();
                break;
            case IdMenuInspectBadMap:
                inspect_source_bad_range_map();
                break;
            case IdMenuClearBadMap:
                clear_source_bad_range_map();
                break;
            case IdMenuSettings:
                on_command(IdSettingsButton, 0);
                break;
            case IdMenuResumeInterrupted:
                if (auto run = recovery_.get_pending_interrupted_run()) {
                    resume_interrupted(*run);
                }
                break;
            case IdMenuSaveAsJob:
                save_current_as_job();
                break;
            case IdMenuJobManager:
                open_job_manager();
                break;
            case IdMenuRunNextQueued:
                run_next_queued_job(/*manual*/ true);
                break;
            case IdMenuCheckUpdates:
                check_for_updates(/*show_dialog*/ true);
                break;
            default:
                break;
        }
    }

    void save_current_as_defaults() {
        models::CopyJobOptions options;
        try {
            options = collect_options();
        } catch (const std::exception& ex) {
            warn_box(storage::fsutil::utf8_to_wide(ex.what()), L"Invalid run setting");
            return;
        }
        if (!settings_.save_run_defaults(options)) {
            const std::wstring error = storage::fsutil::utf8_to_wide(
                settings_.last_save_error().empty()
                    ? "The defaults could not be saved."
                    : settings_.last_save_error());
            warn_box(error, L"Defaults could not be saved");
            SetWindowTextW(status_label_, L"Status: Defaults not saved");
            return;
        }
        append_log("Saved current run settings as defaults.");
        SetWindowTextW(status_label_, L"Status: Defaults saved");
        info_box(L"The current run settings were saved as defaults for future runs.");
    }

    void save_current_as_job() {
        models::CopyJobOptions options;
        try {
            options = collect_options();
        } catch (const std::exception& ex) {
            warn_box(storage::fsutil::utf8_to_wide(ex.what()), L"Invalid run setting");
            return;
        }
        if (ui_is_blank(options.SourceRoot)) {
            info_box(L"Choose a source before saving a job.");
            return;
        }
        bind_media_identities(options);
        std::string suggested = suggest_job_name(options);
        auto name = ui::TextPromptDialog::show(hwnd_, theme_, L"Save Job",
                                               L"Name for this saved job:", suggested);
        if (!name.has_value() || ui_is_blank(*name)) return;
        if (auto saved = job_manager_.save_job(*name, options)) {
            append_log("Saved job '" + saved->Name + "'.");
        }
    }

    void open_job_manager() {
        ui::JobManagerRequest request =
            ui::JobManagerDialog::show(hwnd_, job_manager_, theme_, settings_);
        switch (request.action) {
            case ui::JobManagerRequestAction::RunSavedJob:
                if (!ui_is_blank(request.job_id)) run_saved_job(request.job_id);
                break;
            case ui::JobManagerRequestAction::RunQueuedEntry:
                if (!ui_is_blank(request.queue_entry_id)) {
                    run_queued_entry(request.queue_entry_id, /*manual*/ true);
                }
                break;
            case ui::JobManagerRequestAction::ResumeRun:
                if (!ui_is_blank(request.run_id)) resume_managed_run(request.run_id);
                break;
            default:
                break;
        }
    }

    void resume_managed_run(const std::string& run_id) {
        if (starting_ || supervisor_.is_job_running()) {
            info_box(L"A run is already in progress.");
            return;
        }
        auto run = job_manager_.get_run_by_id(run_id);
        if (!run.has_value()) {
            warn_box(L"Selected run no longer exists.");
            return;
        }

        bool exact = false;
        std::string warning;
        std::optional<models::CopyJobOptions> restored =
            job_manager_.get_resume_options(run_id, &exact, &warning);
        models::CopyJobOptions options;
        if (restored.has_value()) {
            options = *restored;
        } else {
            options = settings_.build_default_options();
            options.SourceRoot = run->SourceRoot;
            options.DestinationRoot = run->DestinationRoot;
            const bool legacy_scan =
                models::detail::equals_ignore_case(run->Trigger, "scan") ||
                run->DisplayName.find("Assessment") != std::string::npos ||
                ui_is_blank(run->DestinationRoot);
            if (legacy_scan) {
                options.OperationMode = models::JobOperationMode::ScanOnly;
                options.DestinationRoot.clear();
            }
        }
        if (!exact &&
            !confirm_box(storage::fsutil::utf8_to_wide(
                             warning +
                             "\n\nThe source, destination, and operation type were recovered. Review the restored controls before continuing. Resume now?"),
                         L"Legacy run settings", ui::MessageIcon::Warning)) {
            apply_run_options_to_ui(options);
            return;
        }

        const std::string journal_path =
            job_manager_.get_resume_journal_path(run_id, /*validate_roots*/ true);
        if (journal_path.empty()) {
            warn_box(L"No authenticated resumable journal could be located for this run.");
            return;
        }
        options.ResumeFromJournal = true;
        options.ResumeJournalPathHint = journal_path;
        apply_run_options_to_ui(options);
        append_log("Resuming run '" + run->DisplayName + "' from " + journal_path + ".");
        start_job_with(options, true, run->RunId, run->DisplayName);
    }

    void run_saved_job(const std::string& job_id) {
        if (supervisor_.is_job_running()) {
            info_box(L"A copy run is already in progress.");
            return;
        }
        auto job = job_manager_.get_job_by_id(job_id);
        if (!job.has_value()) {
            warn_box(L"Selected job no longer exists.");
            return;
        }
        auto run = job_manager_.create_run_for_job(job->JobId, "job-manager");
        std::string managed_run_id = run.has_value() ? run->RunId : std::string();
        start_job_with(job->Options, false, managed_run_id, job->Name);
    }

    void run_next_queued_job(bool manual) { run_queued_entry(std::string(), manual); }

    // Dequeues (a specific entry or the next) and starts it; skips empties and
    // keeps draining like MainForm.RunQueuedEntryByIdAsync.
    void run_queued_entry(const std::string& queue_entry_id, bool manual) {
        if (supervisor_.is_job_running()) return;
        std::string target = queue_entry_id;
        bool show_empty_dialog = manual;
        if (!manual && ui_is_blank(target)) {
            // Do not dequeue attended-only jobs merely to reject them. Leave
            // each one in place with a concrete reason, and continue looking
            // for a later queue entry that is safe to run unattended.
            const auto queued = job_manager_.get_queue_entries();
            target.clear();
            for (const auto& candidate : queued) {
                auto job = job_manager_.get_job_by_id(candidate.JobId);
                if (!job.has_value()) continue;
                const std::string issue = job->Options.unattended_policy_issue();
                if (issue.empty()) {
                    target = candidate.QueueEntryId;
                    break;
                }
                if (job_manager_.mark_queue_entry_blocked(candidate.QueueEntryId, issue)) {
                    append_log("Deferred automatic queued job '" + job->Name + "': " + issue + ".");
                }
            }
            if (target.empty()) return;
        }
        for (;;) {
            ui::QueuedJobWorkItem work_item;
            bool dequeued = ui_is_blank(target)
                                ? job_manager_.try_dequeue_next_job(work_item)
                                : job_manager_.try_dequeue_queued_entry(target, work_item);
            if (!dequeued) {
                if (show_empty_dialog) {
                    const wchar_t* message =
                        ui_is_blank(target)
                            ? L"No queued jobs are waiting to run."
                            : L"Selected queue entry no longer exists.";
                    info_box(message);
                }
                return;
            }

            std::string trigger = manual ? "queued-manual" : "queued-auto";
            auto run = job_manager_.create_run_for_job(work_item.Job.JobId, trigger,
                                                       work_item.QueueEntryId, work_item.Attempt);
            std::string managed_run_id = run.has_value() ? run->RunId : std::string();
            append_log("Dequeued job '" + work_item.Job.Name + "'.");
            start_job_with(work_item.Job.Options, false, managed_run_id, work_item.Job.Name);
            return;
        }
    }

    static std::string display_source(const models::CopyJobOptions& options) {
        if (options.SelectedRelativePaths.size() != 1) return options.SourceRoot;
        std::wstring source =
            storage::fsutil::get_full_path(storage::fsutil::utf8_to_wide(options.SourceRoot));
        if (!source.empty() && source.back() != L'\\' && source.back() != L'/') {
            source.push_back(L'\\');
        }
        source += storage::fsutil::utf8_to_wide(options.SelectedRelativePaths[0]);
        return storage::fsutil::wide_to_utf8(storage::fsutil::get_full_path(source));
    }

    std::string suggest_job_name(const models::CopyJobOptions& options) {
        std::wstring source = storage::fsutil::utf8_to_wide(display_source(options));
        std::wstring leaf = storage::fsutil::get_file_name(source);
        std::string base = leaf.empty() ? std::string("Saved Job") : storage::fsutil::wide_to_utf8(leaf);
        return base;
    }

    // ---- Themed message boxes (dark-mode replacements for MessageBoxW) ------

    void info_box(const std::wstring& text, const wchar_t* title = L"XactCopy") {
        ui::MessageDialog::show(hwnd_, theme_, title, text, ui::MessageIcon::Information);
    }

    void warn_box(const std::wstring& text, const wchar_t* title = L"XactCopy") {
        ui::MessageDialog::show(hwnd_, theme_, title, text, ui::MessageIcon::Warning);
    }

    // Returns true when the user chooses Yes.
    bool confirm_box(const std::wstring& text, const wchar_t* title = L"XactCopy",
                     ui::MessageIcon icon = ui::MessageIcon::Question) {
        return ui::MessageDialog::show(hwnd_, theme_, title, text, icon,
                                       ui::MessageButtons::YesNo) == IDYES;
    }

    // ---- Multi-item selection ----------------------------------------------

    // Single entry point for every source of items: launch arguments, a second
    // instance's forwarded command line, the Add-Files dialog, and drag-and-drop.
    // Nothing is applied immediately — a burst of arrivals (Explorer invoking the
    // verb once per selected item) is coalesced into one selection, so the user
    // gets a single source root and a single destination prompt.
    void queue_selection(const std::vector<std::wstring>& paths, bool scan_mode,
                         bool want_destination) {
        if (paths.empty()) return;
        if (supervisor_.is_job_running()) {
            append_log("Ignored a selection: a job is already running.");
            return;
        }
        int added = 0;
        for (const auto& path : paths) {
            if (selection_.add(path)) ++added;
        }
        if (added == 0 && !selection_timer_active_) return;
        if (scan_mode) launch_.explorer_scan_mode = true;
        if (want_destination) selection_wants_destination_ = true;

        // Restart the quiet period so late arrivals join the same selection.
        SetTimer(hwnd_, IdSelectionTimer, SelectionCoalesceMs, nullptr);
        selection_timer_active_ = true;
    }

    void commit_selection() {
        KillTimer(hwnd_, IdSelectionTimer);
        selection_timer_active_ = false;
        const bool want_destination = selection_wants_destination_;
        selection_wants_destination_ = false;

        explorer_selection_root_.clear();
        explorer_selection_display_.clear();
        explorer_selected_paths_.clear();
        if (selection_.empty()) {
            update_selection_ui();
            return;
        }

        std::wstring root = selection_.common_root();
        if (root.empty() || !directory_exists(root)) {
            append_log("Selected items do not share a common source folder.");
            selection_.clear();
            update_selection_ui();
            return;
        }
        std::wstring display = selection_.display_path();
        suppress_source_change_ = true;
        SetWindowTextW(source_edit_, display.c_str());
        suppress_source_change_ = false;

        std::vector<std::string> relative = selection_.relative_paths(root);
        if (!relative.empty()) {
            explorer_selection_root_ = storage::fsutil::wide_to_utf8(root);
            explorer_selection_display_ = storage::fsutil::wide_to_utf8(display);
            explorer_selected_paths_ = std::move(relative);
        }

        if (launch_.explorer_scan_mode) {
            SendMessageW(mode_combo_, CB_SETCURSEL, 1, 0); // Assess Readable Files
            sync_mode_ui();
        }
        launch_.explorer_scan_mode = false;
        append_log("Source set to " + storage::fsutil::wide_to_utf8(display) + " (" +
                   std::to_string(selection_.size()) + " item(s) selected).");
        update_selection_ui();

        if (want_destination) prompt_destination_from_explorer();
    }

    void clear_selection(bool restore_root = true) {
        if (restore_root && !explorer_selection_root_.empty()) {
            suppress_source_change_ = true;
            SetWindowTextW(source_edit_,
                           storage::fsutil::utf8_to_wide(explorer_selection_root_).c_str());
            suppress_source_change_ = false;
        }
        selection_.clear();
        explorer_selection_root_.clear();
        explorer_selection_display_.clear();
        explorer_selected_paths_.clear();
        update_selection_ui();
        append_log(std::string("Selection cleared; the whole source folder will be ") +
                   (is_scan_mode() ? "assessed." : "copied."));
    }

    // A multi-selection needs a summary because the Source box can only show its
    // common root. Keep the same confirmation for a single exact item.
    void update_selection_ui() {
        const bool active = !explorer_selected_paths_.empty();
        const bool was_visible = selection_label_ != nullptr &&
                                 IsWindowVisible(selection_label_) != FALSE;
        if (selection_label_ != nullptr) {
            std::wstring text =
                active ? selection_.summary() +
                             (selection_.size() == 1 ? L" \x2014 only this will be copied"
                                                     : L" \x2014 only these will be copied")
                       : L"";
            SetWindowTextW(selection_label_, text.c_str());
            ShowWindow(selection_label_, active ? SW_SHOW : SW_HIDE);
        }
        if (clear_selection_ != nullptr) ShowWindow(clear_selection_, active ? SW_SHOW : SW_HIDE);
        // The row only occupies space when shown, so re-run the layout when its
        // visibility changes to open/close the gap beneath the source box.
        if (was_visible != active) layout();
    }

    // Multi-select file picker; folders come in via Browse or drag-and-drop.
    void add_files_to_selection() {
        if (FAILED(CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED | COINIT_DISABLE_OLE1DDE))) {
            return;
        }
        std::vector<std::wstring> picked;
        IFileOpenDialog* dialog = nullptr;
        if (SUCCEEDED(CoCreateInstance(CLSID_FileOpenDialog, nullptr, CLSCTX_INPROC_SERVER,
                                       IID_PPV_ARGS(&dialog)))) {
            DWORD options = 0;
            dialog->GetOptions(&options);
            dialog->SetOptions(options | FOS_ALLOWMULTISELECT | FOS_FORCEFILESYSTEM |
                               FOS_FILEMUSTEXIST);
            dialog->SetTitle(L"Add files to copy");
            if (SUCCEEDED(dialog->Show(hwnd_))) {
                IShellItemArray* items = nullptr;
                if (SUCCEEDED(dialog->GetResults(&items))) {
                    DWORD count = 0;
                    items->GetCount(&count);
                    for (DWORD i = 0; i < count; ++i) {
                        IShellItem* item = nullptr;
                        if (SUCCEEDED(items->GetItemAt(i, &item))) {
                            PWSTR path = nullptr;
                            if (SUCCEEDED(item->GetDisplayName(SIGDN_FILESYSPATH, &path))) {
                                picked.emplace_back(path);
                                CoTaskMemFree(path);
                            }
                            item->Release();
                        }
                    }
                    items->Release();
                }
            }
            dialog->Release();
        }
        CoUninitialize();
        // Already-chosen items stay; the picker adds to the set.
        queue_selection(picked, /*scan_mode*/ false, /*want_destination*/ false);
    }

    // ---- Explorer launch application (--from-explorer[-folder]) ------------

    static std::wstring resolve_full_path(const std::string& raw) {
        std::wstring wide = storage::fsutil::trim(storage::fsutil::utf8_to_wide(raw));
        if (wide.empty()) return std::wstring();
        return storage::fsutil::get_full_path(wide);
    }

    static bool directory_exists(const std::wstring& path) {
        DWORD attributes = GetFileAttributesW(path.c_str());
        return attributes != INVALID_FILE_ATTRIBUTES &&
               (attributes & FILE_ATTRIBUTE_DIRECTORY) != 0;
    }

    // Root/relative-path resolution now lives in ui::SelectionModel (selection.h).

    // Feeds launch arguments into the selection model. A folder-background
    // launch is a whole-folder copy and applies immediately; item selections go
    // through the coalescing queue so Explorer's one-launch-per-item burst
    // becomes a single job.
    void apply_explorer_launch_options() {
        if (!ui_is_blank(launch_.explorer_folder_path)) {
            std::wstring folder = resolve_full_path(launch_.explorer_folder_path);
            if (folder.empty() || !directory_exists(folder)) {
                append_log("Explorer folder path was not found: " + launch_.explorer_folder_path);
                return;
            }
            selection_.clear();
            explorer_selection_root_.clear();
            explorer_selection_display_.clear();
            explorer_selected_paths_.clear();
            update_selection_ui();
            SetWindowTextW(source_edit_, folder.c_str());
            if (launch_.explorer_scan_mode) {
                SendMessageW(mode_combo_, CB_SETCURSEL, 1, 0);
                sync_mode_ui();
            } else {
                SendMessageW(mode_combo_, CB_SETCURSEL, 0, 0);
                sync_mode_ui();
            }
            launch_.explorer_scan_mode = false;
            append_log("Source set from Explorer background folder: " +
                       storage::fsutil::wide_to_utf8(folder));
            prompt_destination_from_explorer();
            return;
        }

        if (launch_.explorer_source_paths.empty()) return;
        std::vector<std::wstring> paths;
        paths.reserve(launch_.explorer_source_paths.size());
        for (const auto& raw : launch_.explorer_source_paths) {
            paths.push_back(resolve_full_path(raw));
        }
        queue_selection(paths, launch_.explorer_scan_mode, /*want_destination*/ true);
    }

    void prompt_destination_from_explorer() {
        if (is_scan_mode()) return; // assessments write no destination
        if (supervisor_.is_job_running()) return;
        std::string current = window_text_utf8(destination_edit_);
        if (!ui_is_blank(current)) return;
        append_log("Explorer launch detected. Choose a destination folder.");
        browse_folder(destination_edit_);
    }

    // ---- Update check (fetch + compare; download handled via the browser) --

    void check_for_updates(bool show_dialog) {
        if (checking_updates_ || supervisor_.is_job_running()) return;
        std::string url = settings_.get_string("UpdateReleaseUrl", ui::kDefaultUpdateReleaseUrl);
        if (ui_is_blank(url)) {
            if (show_dialog) {
                info_box(L"No update source is configured (UpdateReleaseUrl is empty).",
                         L"XactCopy Updates");
            }
            return;
        }
        checking_updates_ = true;
        update_check_show_dialog_ = show_dialog;
        update_menu_state();
        append_log("Checking for updates...");

        if (update_thread_.joinable()) update_thread_.join();
        HWND hwnd = hwnd_;
        update_thread_ = std::thread([hwnd, url]() {
            auto* info = new ui::UpdateReleaseInfo;
            try {
                *info = ui::UpdateService::get_latest_release(url);
            } catch (const std::exception& ex) {
                info->error = std::string("Update check failed: ") + ex.what();
            } catch (...) {
                info->error = "Update check failed unexpectedly.";
            }
            if (!PostMessageW(hwnd, WM_APP_UPDATE_DONE, 0,
                              reinterpret_cast<LPARAM>(info))) {
                delete info;
            }
        });
    }

    void on_update_done(ui::UpdateReleaseInfo& info) {
        checking_updates_ = false;
        update_menu_state();
        if (!info.ok) {
            append_log("Update check failed: " + info.error);
            if (update_check_show_dialog_) {
                warn_box(storage::fsutil::utf8_to_wide(info.error), L"XactCopy Updates");
            }
            return;
        }
        if (!ui::UpdateService::is_update_available(ui::kNativeVersion, info.version)) {
            append_log("Update check complete: already up to date.");
            if (update_check_show_dialog_) {
                info_box(L"You're already running the latest version.", L"XactCopy Updates");
            }
            return;
        }
        append_log("Update available: " + info.tag_name);
        // Show the themed update dialog (download -> verify -> apply). When it
        // launches the updater it returns true and we close so the update runs.
        bool should_exit = ui::UpdateDialog::show(hwnd_, info, ui::kNativeVersion, settings_, theme_);
        if (should_exit) {
            append_log("Applying update; XactCopy will close and restart.");
            PostMessageW(hwnd_, WM_CLOSE, 0, 0);
        }
    }

    // Registers/removes the shell context-menu verbs to match the setting; only
    // touches the registry when the desired state differs from the current one.
    void sync_explorer_integration(bool announce) {
        bool desired = settings_.get_bool("EnableExplorerContextMenu", false);
        bool current = ui::ExplorerIntegrationService::is_registered();
        if (current == desired) return;
        wchar_t exe_path[MAX_PATH]{};
        GetModuleFileNameW(nullptr, exe_path, MAX_PATH);
        bool ok = ui::ExplorerIntegrationService::sync(desired, exe_path);
        if (announce) {
            if (ok) {
                append_log(std::string("Explorer context menu integration ") +
                           (desired ? "enabled." : "disabled."));
            } else {
                append_log("Explorer integration update failed.");
            }
        }
    }

    void browse_folder(HWND target_edit) {
        if (FAILED(CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED | COINIT_DISABLE_OLE1DDE))) {
            return;
        }
        IFileOpenDialog* dialog = nullptr;
        if (SUCCEEDED(CoCreateInstance(CLSID_FileOpenDialog, nullptr, CLSCTX_INPROC_SERVER,
                                       IID_PPV_ARGS(&dialog)))) {
            DWORD options = 0;
            dialog->GetOptions(&options);
            dialog->SetOptions(options | FOS_PICKFOLDERS | FOS_FORCEFILESYSTEM);
            if (SUCCEEDED(dialog->Show(hwnd_))) {
                IShellItem* item = nullptr;
                if (SUCCEEDED(dialog->GetResult(&item))) {
                    PWSTR path = nullptr;
                    if (SUCCEEDED(item->GetDisplayName(SIGDN_FILESYSPATH, &path))) {
                        SetWindowTextW(target_edit, path);
                        CoTaskMemFree(path);
                    }
                    item->Release();
                }
            }
            dialog->Release();
        }
        CoUninitialize();
    }

    std::string window_text_utf8(HWND handle) {
        wchar_t buffer[2048];
        int length = GetWindowTextW(handle, buffer, 2048);
        return storage::fsutil::wide_to_utf8(std::wstring(buffer, length));
    }

    int read_int_field(HWND edit, int min_value, int max_value, const char* label) {
        std::string text = window_text_utf8(edit);
        try {
            std::size_t consumed = 0;
            long long value = std::stoll(text, &consumed, 10);
            if (consumed != text.size() || value < min_value || value > max_value) {
                throw std::out_of_range("value outside range");
            }
            return static_cast<int>(value);
        } catch (...) {
            throw std::runtime_error(std::string(label) + " must be a whole number from " +
                                     std::to_string(min_value) + " to " +
                                     std::to_string(max_value) + ".");
        }
    }

    void apply_run_options_to_ui(const models::CopyJobOptions& options) {
        explorer_selection_root_ = options.SourceRoot;
        explorer_selected_paths_ = options.SelectedRelativePaths;
        explorer_selection_display_ = display_source(options);
        SetWindowTextW(source_edit_,
                       storage::fsutil::utf8_to_wide(explorer_selection_display_).c_str());
        SetWindowTextW(destination_edit_,
                       storage::fsutil::utf8_to_wide(options.DestinationRoot).c_str());

        custom_transfer_engine_ = options.TransferEnginePolicyValue;
        copy_profile_index_ = 2; // exact persisted options, not a mutable preset
        copy_overwrite_index_ = std::clamp(static_cast<int>(options.OverwritePolicyValue), 0, 3);
        copy_verify_index_ = ui::verification_combo_index(options);
        scan_profile_index_ =
            std::clamp(static_cast<int>(options.ScanPerformanceProfileValue), 0, 2);
        scan_backend_index_ = options.UseExperimentalRawDiskScan ? 1 : 0;
        scan_map_index_ = options.UpdateBadRangeMapFromRun ? 0 : 1;

        active_mode_index_ = -1;
        SendMessageW(mode_combo_, CB_SETCURSEL,
                     options.OperationMode == models::JobOperationMode::ScanOnly ? 1 : 0, 0);
        sync_mode_ui();

        set_check(IdSalvageCheck, options.SalvageUnreadableBlocks);
        set_check(IdResumeCheck, true);
        set_check(IdMapCheck, options.UseBadRangeMap);
        set_check(IdAdaptiveCheck, options.UseAdaptiveBufferSizing);
        set_check(IdContinueCheck, options.ContinueOnFileError);
        set_check(IdSkipKnownBadCheck,
                  options.UseBadRangeMap && options.SkipKnownBadRanges);
        set_check(IdWaitMediaCheck, options.WaitForMediaAvailability);
        set_check(IdFragileCheck, options.FragileMediaMode);
        sync_bad_range_controls();

        const int buffer_mb = std::clamp(
            std::max(1, options.BufferSizeBytes / (1024 * 1024)), 1, 256);
        SetWindowTextW(buffer_edit_, std::to_wstring(buffer_mb).c_str());
        SetWindowTextW(retries_edit_,
                       std::to_wstring(std::clamp(options.MaxRetries, 0, 32)).c_str());
        const std::int64_t timeout_seconds = std::clamp<std::int64_t>(
            options.OperationTimeout.ticks / time::TicksPerSecond, 1, 3600);
        SetWindowTextW(timeout_edit_, std::to_wstring(timeout_seconds).c_str());
    }

    models::CopyJobOptions collect_options() {
        models::CopyJobOptions options = settings_.build_default_options();
        options.SourceRoot = window_text_utf8(source_edit_);
        options.DestinationRoot = window_text_utf8(destination_edit_);

        int mode = static_cast<int>(SendMessageW(mode_combo_, CB_GETCURSEL, 0, 0));
        options.OperationMode =
            mode == 1 ? models::JobOperationMode::ScanOnly : models::JobOperationMode::Copy;
        if (options.OperationMode == models::JobOperationMode::ScanOnly) {
            options.DestinationRoot.clear();
            int profile = static_cast<int>(SendMessageW(engine_combo_, CB_GETCURSEL, 0, 0));
            options.ScanPerformanceProfileValue =
                profile == 1   ? models::ScanPerformanceProfile::Fast
                : profile == 2 ? models::ScanPerformanceProfile::Precise
                               : models::ScanPerformanceProfile::Auto;
            options.UseExperimentalRawDiskScan =
                SendMessageW(overwrite_combo_, CB_GETCURSEL, 0, 0) == 1;
            options.UpdateBadRangeMapFromRun =
                SendMessageW(verify_combo_, CB_GETCURSEL, 0, 0) != 1;
            options.TransferEnginePolicyValue = models::TransferEnginePolicy::ManagedRescue;
            options.VerifyAfterCopy = false;
            options.VerificationModeValue = models::VerificationMode::None;
            options.SalvageUnreadableBlocks = false;
        } else {
            int profile = static_cast<int>(SendMessageW(engine_combo_, CB_GETCURSEL, 0, 0));
            options.TransferEnginePolicyValue =
                profile == 0   ? models::TransferEnginePolicy::Auto
                : profile == 1 ? models::TransferEnginePolicy::ManagedRescue
                               : custom_transfer_engine_;
            if (profile == 1) {
                options.UpdateBadRangeMapFromRun = true;
                // Access-denied-as-contention is an expert lock policy that is
                // intentionally incompatible with salvage. The Recover Media
                // profile must remain runnable even if an older custom default
                // left that hidden setting enabled.
                options.TreatAccessDeniedAsContention = false;
            }
            int overwrite = static_cast<int>(
                SendMessageW(overwrite_combo_, CB_GETCURSEL, 0, 0));
            options.OverwritePolicyValue = static_cast<models::OverwritePolicy>(
                overwrite < 0 || overwrite > 3 ? 0 : overwrite);
            int verify = static_cast<int>(SendMessageW(verify_combo_, CB_GETCURSEL, 0, 0));
            options.VerificationModeValue =
                verify == 1   ? models::VerificationMode::Sampled
                : verify == 2 ? models::VerificationMode::Full
                              : models::VerificationMode::None;
            options.VerifyAfterCopy =
                options.VerificationModeValue != models::VerificationMode::None;
            options.SalvageUnreadableBlocks = get_check(IdSalvageCheck);
        }

        options.ResumeFromJournal = get_check(IdResumeCheck);
        bool use_map = get_check(IdMapCheck);
        options.UseBadRangeMap = use_map;
        options.SkipKnownBadRanges = use_map && get_check(IdSkipKnownBadCheck);
        options.UseAdaptiveBufferSizing = get_check(IdAdaptiveCheck);
        options.ContinueOnFileError = get_check(IdContinueCheck);
        options.WaitForMediaAvailability = get_check(IdWaitMediaCheck);
        options.FragileMediaMode = get_check(IdFragileCheck);

        int buffer_mb = read_int_field(buffer_edit_, 1, 256, "Buffer size");
        options.BufferSizeBytes = static_cast<std::int32_t>(
            static_cast<std::int64_t>(buffer_mb) * 1024 * 1024);
        options.MaxRetries = read_int_field(retries_edit_, 0, 32, "Retry count");
        int timeout_seconds = read_int_field(timeout_edit_, 1, 3600, "Operation timeout");
        options.OperationTimeout = time::TimeSpan::from_seconds(timeout_seconds);

        // The Source box shows an exact single item, while the worker accepts a
        // directory root plus relative paths. Translate back only while the box
        // still contains the captured display path (manual edits opt out).
        if (!explorer_selected_paths_.empty() &&
            models::detail::equals_ignore_case(
                storage::fsutil::wide_to_utf8(storage::fsutil::get_full_path(
                    storage::fsutil::utf8_to_wide(options.SourceRoot))),
                explorer_selection_display_)) {
            options.SourceRoot = explorer_selection_root_;
            options.SelectedRelativePaths = explorer_selected_paths_;
        }
        return options;
    }

    void start_job() {
        try {
            start_job_with(collect_options(), false);
        } catch (const std::exception& ex) {
            warn_box(storage::fsutil::utf8_to_wide(ex.what()), L"Invalid run setting");
        }
    }

    static void bind_media_identities(models::CopyJobOptions& options) {
        if (options.ExpectedSourceIdentity.empty()) {
            options.ExpectedSourceIdentity = storage::fsutil::resolve_media_identity(
                storage::fsutil::utf8_to_wide(options.SourceRoot));
        }
        if (options.OperationMode == models::JobOperationMode::ScanOnly) return;
        if (options.ExpectedDestinationIdentity.empty()) {
            options.ExpectedDestinationIdentity = storage::fsutil::resolve_media_identity(
                storage::fsutil::utf8_to_wide(options.DestinationRoot));
        }
    }

    void start_job_with(models::CopyJobOptions options, bool from_recovery,
                        std::string managed_run_id = std::string(),
                        const std::string& managed_display_name = std::string()) {
        if (starting_ || supervisor_.is_job_running()) return;
        if (ui_is_blank(options.SourceRoot)) {
            info_box(L"Choose a source first.");
            return;
        }
        if (options.OperationMode == models::JobOperationMode::Copy &&
            ui_is_blank(options.DestinationRoot)) {
            info_box(L"Choose a destination folder first.");
            return;
        }

        if (options.OperationMode == models::JobOperationMode::Copy) {
            std::wstring risks;
            auto add_risk = [&risks](const wchar_t* risk) {
                risks += L"\n  - ";
                risks += risk;
            };
            if (!options.VerifyAfterCopy ||
                options.VerificationModeValue == models::VerificationMode::None) {
                add_risk(L"Destination bytes will not be verified after copying.");
            } else if (options.VerificationModeValue != models::VerificationMode::Full) {
                add_risk(L"Sampled verification can miss corruption outside the sampled ranges.");
            }
            if (options.ContinueOnFileError) {
                add_risk(L"Continuing after an error can leave a mixed, incomplete destination set.");
            }
            if (options.SalvageUnreadableBlocks &&
                options.AllowRecoveredOverwriteExisting) {
                add_risk(L"Synthetic salvage bytes may replace an existing destination file.");
            }
            if (options.AllowDecryptedDestination) {
                add_risk(L"EFS-encrypted source data may be published as plaintext.");
            }
            if (!risks.empty() &&
                !confirm_box(L"This run uses attended-only integrity settings:" + risks +
                                 L"\n\nProceed with these risks?",
                             L"XactCopy - Confirm integrity risks",
                             ui::MessageIcon::Warning)) {
                return;
            }
        }
        bind_media_identities(options);

        SendMessageW(log_list_, LB_RESETCONTENT, 0, 0);
        reset_log_horizontal_extent();
        reset_telemetry();
        SetWindowTextW(journal_label_, storage::fsutil::utf8_to_wide(
                                           "Journal: " + compute_journal_path(options)).c_str());
        SetWindowTextW(job_summary_label_,
                       storage::fsutil::utf8_to_wide(
                           "Job: " + (options.OperationMode == models::JobOperationMode::ScanOnly
                                          ? std::string("Readability Assessment")
                                          : std::string("Copy")) +
                           " — " + display_source(options)).c_str());

        // Record a managed run in the catalog so the Job Manager history tracks
        // every copy/scan, matching MainForm's CreateAdHocRun/CreateRunForJob.
        if (ui_is_blank(managed_run_id)) {
            std::string display_name = managed_display_name;
            std::string trigger = "manual";
            if (display_name.empty()) {
                if (from_recovery) {
                    display_name = options.OperationMode == models::JobOperationMode::ScanOnly
                                       ? "Recovered Readability Assessment"
                                       : "Recovered Copy Session";
                    trigger = "resume-interrupted";
                }
                else if (options.OperationMode == models::JobOperationMode::ScanOnly) {
                    display_name = "Readability Assessment"; trigger = "scan";
                } else { display_name = "Manual Copy"; }
            }
            managed_run_id = job_manager_.create_ad_hoc_run(options, display_name, trigger).RunId;
        }
        active_managed_run_id_ = managed_run_id;
        pending_start_options_ = options;
        pending_start_job_name_ =
            !managed_display_name.empty()
                ? managed_display_name
                : (from_recovery ? std::string("Recovered run") : std::string("Manual run"));

        // Spawn + pipe-connect can take a moment (worker cold start); run it on a
        // background thread so the UI stays responsive, then finish in
        // WM_APP_START_DONE. Without this the Start click blocks the message loop.
        starting_ = true;
        EnableWindow(start_button_, FALSE);
        set_inputs_enabled(false);
        SetWindowTextW(status_label_, L"Status: Starting worker...");
        append_log("Starting worker...");

        if (start_thread_.joinable()) start_thread_.join();
        HWND hwnd = hwnd_;
        models::CopyJobOptions options_copy = options;
        start_thread_ = std::thread([this, hwnd, options_copy]() {
            try {
                supervisor_.start_job(options_copy);
                PostMessageW(hwnd, WM_APP_START_DONE, 1,
                             reinterpret_cast<LPARAM>(new std::string()));
            } catch (const std::exception& ex) {
                PostMessageW(hwnd, WM_APP_START_DONE, 0,
                             reinterpret_cast<LPARAM>(new std::string(ex.what())));
            }
        });
    }

    void on_start_done(bool ok, const std::string& error) {
        starting_ = false;
        if (ok) {
            active_run_id_ = supervisor_.current_job_id();
            recovery_.mark_job_started(active_run_id_, pending_start_job_name_,
                                       pending_start_options_,
                                       compute_journal_path(pending_start_options_), settings_,
                                       active_managed_run_id_);
            if (latest_progress_.has_value()) {
                recovery_.update_media_identities(
                    active_run_id_, latest_progress_->SourceMediaIdentity,
                    latest_progress_->DestinationMediaIdentity);
            }
            if (!active_managed_run_id_.empty()) {
                job_manager_.mark_run_running(active_managed_run_id_,
                                              compute_journal_path(pending_start_options_),
                                              &pending_start_options_);
            }
            EnableWindow(start_button_, FALSE);
            EnableWindow(pause_button_, TRUE);
            EnableWindow(cancel_button_, TRUE);
            update_menu_state();
            append_log("Job started (" + active_run_id_ + ").");
        } else {
            if (!active_managed_run_id_.empty()) {
                models::CopyJobResult failure;
                failure.ErrorMessage = error;
                job_manager_.mark_run_completed(active_managed_run_id_, failure);
                active_managed_run_id_.clear();
            }
            EnableWindow(start_button_, TRUE);
            set_inputs_enabled(true);
            SetWindowTextW(status_label_, L"Status: Idle");
            update_menu_state();
            ui::MessageDialog::show(hwnd_, theme_, L"XactCopy",
                                    storage::fsutil::utf8_to_wide(error), ui::MessageIcon::Error);
        }
    }

    // Mirrors MainForm.ComputeJournalPath: derive the deterministic journal path
    // from source|destination so the Job Manager can open it later.
    std::string compute_journal_path(const models::CopyJobOptions& options) {
        std::string destination = options.DestinationRoot;
        if (options.OperationMode == models::JobOperationMode::ScanOnly && ui_is_blank(destination)) {
            destination = options.SourceRoot;
        }
        if (ui_is_blank(options.SourceRoot) || ui_is_blank(destination)) return std::string();
        std::string job_id = storage::JobJournalStore::build_job_id(
            storage::fsutil::utf8_to_wide(options.SourceRoot),
            storage::fsutil::utf8_to_wide(destination));
        return storage::fsutil::wide_to_utf8(storage::JobJournalStore::get_default_journal_path(job_id));
    }

    static bool ui_is_blank(const std::string& text) {
        for (char c : text) {
            if (c != ' ' && c != '\t') return false;
        }
        return true;
    }

    // ---- Event application -------------------------------------------------

    void append_log(const std::string& message) {
        std::wstring wide = storage::fsutil::utf8_to_wide(message);
        int max_lines = settings_.get_int("UiMaxLogLines", 1000, 1000000, 50000);
        int cap = std::min(max_lines, 20000); // classic listbox practical cap
        int count = static_cast<int>(SendMessageW(log_list_, LB_GETCOUNT, 0, 0));
        while (count >= cap) {
            SendMessageW(log_list_, LB_DELETESTRING, 0, 0);
            --count;
        }
        int index = static_cast<int>(
            SendMessageW(log_list_, LB_ADDSTRING, 0, reinterpret_cast<LPARAM>(wide.c_str())));
        update_log_horizontal_extent(wide);
        SendMessageW(log_list_, LB_SETTOPINDEX, index, 0);
    }

    void apply_progress(const models::CopyProgressSnapshot& snapshot) {
        recovery_.touch_active_run(active_run_id_);
        recovery_.update_media_identities(active_run_id_, snapshot.SourceMediaIdentity,
                                          snapshot.DestinationMediaIdentity);
        latest_progress_ = snapshot;

        SetWindowTextW(current_label_,
                       storage::fsutil::utf8_to_wide(
                           "Current: " + (snapshot.CurrentFile.empty() ? std::string("—")
                                                                       : snapshot.CurrentFile))
                           .c_str());
        set_progress_bar(current_bar_, current_permille_,
                         static_cast<int>(snapshot.current_file_progress() * 1000.0));
        set_progress_bar(overall_bar_, overall_permille_,
                         static_cast<int>(snapshot.overall_progress() * 1000.0));
        taskbar_.set_progress(hwnd_, static_cast<unsigned>(overall_permille_), 1000,
                              paused_ ? ui::TaskbarProgressState::Paused
                                      : ui::TaskbarProgressState::Normal);

        const std::int64_t work_completed = snapshot.WorkBytesCompleted > 0
                                                ? snapshot.WorkBytesCompleted
                                                : snapshot.TotalBytesCopied;
        std::string stats =
            std::to_string(snapshot.CompletedFiles) + "/" + std::to_string(snapshot.TotalFiles) +
            " files  •  " + format_bytes_short(work_completed) + " / " +
            format_bytes_short(snapshot.TotalBytes);
        stats += pending_start_options_.OperationMode == models::JobOperationMode::ScanOnly
                     ? " assessed"
                     : " processed";
        if (snapshot.BytesSkipped > 0) {
            stats += "  •  " + format_bytes_short(snapshot.BytesSkipped) + " skipped";
        }
        if (snapshot.BytesReused > 0) {
            stats += "  •  " + format_bytes_short(snapshot.BytesReused) + " reused";
        }
        if (snapshot.FailedFiles > 0) {
            stats += "  •  " + std::to_string(snapshot.FailedFiles) + " failed";
        }
        if (snapshot.RecoveredFiles > 0) {
            stats += "  •  " + std::to_string(snapshot.RecoveredFiles) +
                     (pending_start_options_.OperationMode == models::JobOperationMode::ScanOnly
                          ? " with unreadable ranges"
                          : " recovered");
        }
        if (!snapshot.RescuePass.empty()) stats += "  •  pass: " + snapshot.RescuePass;
        if (snapshot.ScanWorkerCount > 1) {
            stats += "  •  workers: " + std::to_string(snapshot.ScanWorkerCount);
        }
        SetWindowTextW(stats_label_, storage::fsutil::utf8_to_wide(stats).c_str());

        update_transfer_telemetry(snapshot);
    }

    void reset_telemetry() {
        smoothed_bytes_per_second_ = 0.0;
        telemetry_started_ = false;
        latest_progress_.reset();
        telemetry_last_bytes_copied_ = 0;
        buffer_sample_bytes_total_ = 0;
        buffer_samples_count_ = 0;
        set_progress_bar(current_bar_, current_permille_, 0);
        set_progress_bar(overall_bar_, overall_permille_, 0);
        taskbar_.clear(hwnd_);
        SetWindowTextW(throughput_label_, L"Speed: 0 B/s (avg 0 B/s)");
        SetWindowTextW(eta_label_, L"ETA: -");
        SetWindowTextW(buffer_usage_label_, L"Buffer: -");
        SetWindowTextW(rescue_label_, L"Rescue: -");
        SetWindowTextW(diagnostics_label_, L"Diagnostics: -");
    }

    void refresh_diagnostics_label() {
        if (!show_diagnostics_ || !latest_progress_.has_value()) {
            SetWindowTextW(diagnostics_label_, L"Diagnostics: -");
            return;
        }
        const auto& snapshot = *latest_progress_;
        std::string diag = "Diagnostics: active " + std::to_string(snapshot.ActiveFileCount) +
                           " | scan workers " + std::to_string(snapshot.ScanWorkerCount) +
                           " | recovered " + std::to_string(snapshot.RecoveredFiles) +
                           " | skipped " + std::to_string(snapshot.SkippedFiles);
        SetWindowTextW(diagnostics_label_, storage::fsutil::utf8_to_wide(diag).c_str());
    }

    static std::string format_eta(double seconds) {
        if (seconds < 0) seconds = 0;
        long total = static_cast<long>(seconds);
        long hours = total / 3600;
        long minutes = (total % 3600) / 60;
        long secs = total % 60;
        char buffer[32];
        std::snprintf(buffer, sizeof(buffer), "%02ld:%02ld:%02ld", hours, minutes, secs);
        return buffer;
    }

    // Mirrors MainForm.UpdateTransferTelemetry: EWMA speed (0.65/0.35), running
    // average, ETA from smoothed speed, and average buffer utilization.
    void update_transfer_telemetry(const models::CopyProgressSnapshot& snapshot) {
        const bool scan =
            pending_start_options_.OperationMode == models::JobOperationMode::ScanOnly;
        std::int64_t activity_bytes = scan ? snapshot.BytesRead : snapshot.BytesWritten;
        // Compatibility with an older worker during an in-place update.
        if (activity_bytes == 0 && snapshot.BytesRead == 0 && snapshot.BytesWritten == 0 &&
            snapshot.WorkBytesCompleted == 0) {
            activity_bytes = snapshot.TotalBytesCopied;
        }
        const std::int64_t work_completed = snapshot.WorkBytesCompleted > 0
                                                ? snapshot.WorkBytesCompleted
                                                : snapshot.TotalBytesCopied;
        time::DateTimeOffset now_utc = time::DateTimeOffset::now_utc();
        if (!telemetry_started_) {
            telemetry_started_ = true;
            telemetry_start_utc_ = now_utc;
            telemetry_last_sample_utc_ = now_utc;
            telemetry_last_bytes_copied_ = activity_bytes;
        }

        double interval_seconds =
            static_cast<double>(now_utc.utc_ticks() - telemetry_last_sample_utc_.utc_ticks()) /
            10000000.0;
        std::int64_t delta_bytes = activity_bytes - telemetry_last_bytes_copied_;
        if (interval_seconds > 0 && delta_bytes >= 0) {
            double instant = static_cast<double>(delta_bytes) / interval_seconds;
            smoothed_bytes_per_second_ = smoothed_bytes_per_second_ <= 0
                                             ? instant
                                             : smoothed_bytes_per_second_ * 0.65 + instant * 0.35;
        }
        telemetry_last_sample_utc_ = now_utc;
        telemetry_last_bytes_copied_ = activity_bytes;

        double elapsed_seconds =
            static_cast<double>(now_utc.utc_ticks() - telemetry_start_utc_.utc_ticks()) /
            10000000.0;
        double average_bps =
            elapsed_seconds > 0 ? static_cast<double>(activity_bytes) / elapsed_seconds : 0.0;

        SetWindowTextW(throughput_label_,
                       storage::fsutil::utf8_to_wide(
                           std::string(scan ? "Read: " : "Write: ") +
                           format_bytes_short(static_cast<std::int64_t>(smoothed_bytes_per_second_ + 0.5)) +
                           "/s (avg " +
                           format_bytes_short(static_cast<std::int64_t>(average_bps + 0.5)) + "/s)")
                           .c_str());

        std::int64_t remaining_bytes =
            std::max<std::int64_t>(0, snapshot.TotalBytes - work_completed);
        double reference_speed = smoothed_bytes_per_second_ > 1.0 ? smoothed_bytes_per_second_ : average_bps;
        if (remaining_bytes <= 0) {
            SetWindowTextW(eta_label_, L"ETA: 00:00:00");
        } else if (reference_speed > 1.0) {
            SetWindowTextW(eta_label_,
                           storage::fsutil::utf8_to_wide(
                               "ETA: " + format_eta(static_cast<double>(remaining_bytes) / reference_speed))
                               .c_str());
        } else {
            SetWindowTextW(eta_label_, L"ETA: Calculating...");
        }

        if (snapshot.BufferSizeBytes > 0 && snapshot.LastChunkBytesTransferred > 0) {
            ++buffer_samples_count_;
            buffer_sample_bytes_total_ += snapshot.LastChunkBytesTransferred;
        }
        if (snapshot.BufferSizeBytes > 0 && buffer_samples_count_ > 0) {
            double average_util = (static_cast<double>(buffer_sample_bytes_total_) /
                                   static_cast<double>(buffer_samples_count_)) /
                                  static_cast<double>(snapshot.BufferSizeBytes);
            average_util = std::clamp(average_util, 0.0, 1.0);
            char buffer[96];
            std::snprintf(buffer, sizeof(buffer), "Buffer: %s (avg %.0f%%)",
                          format_bytes_short(snapshot.BufferSizeBytes).c_str(), average_util * 100.0);
            SetWindowTextW(buffer_usage_label_, storage::fsutil::utf8_to_wide(buffer).c_str());
        }

        if (!snapshot.RescuePass.empty() || snapshot.RescueBadRegionCount > 0) {
            std::string rescue = "Rescue: " +
                                 (snapshot.RescuePass.empty() ? std::string("-") : snapshot.RescuePass) +
                                 ", regions " + std::to_string(snapshot.RescueBadRegionCount) +
                                 ", remaining " + format_bytes_short(snapshot.RescueRemainingBytes);
            SetWindowTextW(rescue_label_, storage::fsutil::utf8_to_wide(rescue).c_str());
        }

        // The diagnostics strip is refreshed by UiDiagnosticsRefreshMs rather
        // than by every worker snapshot, keeping high-frequency progress events
        // from becoming a second UI update stream.
    }

    void apply_result(const models::CopyJobResult& result) {
        recovery_.mark_job_ended(active_run_id_);
        active_run_id_.clear();
        if (!active_managed_run_id_.empty()) {
            job_manager_.mark_run_completed(active_managed_run_id_, result);
            active_managed_run_id_.clear();
        }

        EnableWindow(start_button_, TRUE);
        EnableWindow(pause_button_, FALSE);
        EnableWindow(cancel_button_, FALSE);
        set_inputs_enabled(true);
        SetWindowTextW(pause_button_, L"Pause");
        paused_ = false;
        update_menu_state();
        InvalidateRect(start_button_, nullptr, FALSE);
        InvalidateRect(pause_button_, nullptr, FALSE);
        InvalidateRect(cancel_button_, nullptr, FALSE);

        const bool scan =
            pending_start_options_.OperationMode == models::JobOperationMode::ScanOnly;
        const bool assessment_findings = scan && result.RecoveredFiles > 0;
        std::string summary =
            std::string(result.Succeeded ? (assessment_findings ? "Completed with findings"
                                                                : "Completed")
                                         : (result.Cancelled ? "Cancelled"
                                                              : (result.is_incomplete() ? "Incomplete" : "Failed"))) +
            ": " + std::to_string(result.CompletedFiles) + "/" +
            std::to_string(result.TotalFiles) + " files, " +
            (scan ? format_bytes_short(result.BytesRead) + " read"
                  : format_bytes_short(result.BytesWritten > 0 ? result.BytesWritten
                                                               : result.CopiedBytes) +
                        " written");
        if (!scan && result.BytesVerified > 0) {
            summary += ", " + format_bytes_short(result.BytesVerified) + " verification I/O";
        }
        if (result.FailedFiles > 0) summary += ", " + std::to_string(result.FailedFiles) + " failed";
        if (result.RecoveredFiles > 0) {
            summary += ", " + std::to_string(result.RecoveredFiles) +
                       (scan ? " file(s) with unreadable ranges" : " recovered");
        }
        if (result.SkippedFiles > 0) summary += ", " + std::to_string(result.SkippedFiles) + " skipped";
        if (!result.ErrorMessage.empty()) summary += " — " + result.ErrorMessage;
        append_log(summary);
        if (!result.IntegrityNotice.empty()) {
            append_log("Integrity notice: " + result.IntegrityNotice);
        }
        if (!result.MetadataNotice.empty()) {
            append_log("Metadata notice: " + result.MetadataNotice);
        }
        set_progress_bar(overall_bar_, overall_permille_, result.Succeeded ? 1000 : 0);
        if (result.Succeeded) {
            taskbar_.clear(hwnd_);
        } else {
            taskbar_.set_progress(hwnd_, 1000, 1000, ui::TaskbarProgressState::Error);
        }
        SetWindowTextW(job_summary_label_,
                       storage::fsutil::utf8_to_wide("Job: " + summary).c_str());
        if (!result.JournalPath.empty()) {
            SetWindowTextW(journal_label_,
                           storage::fsutil::utf8_to_wide("Journal: " + result.JournalPath).c_str());
        }

        // A run that finished cleanly no longer needs its rotating backups —
        // those only guard a journal that is still being rewritten. The snapshot,
        // ledger, and anchor stay, so the run remains inspectable and verifiable.
        if (result.Succeeded && !result.JournalPath.empty() &&
            settings_.get_bool("CompactJournalsOnCompletion", true)) {
            std::int64_t reclaimed = storage::JobJournalStore::compact_completed(
                storage::fsutil::utf8_to_wide(result.JournalPath));
            if (reclaimed > 0) {
                append_log("Reclaimed " + format_bytes_short(reclaimed) +
                           " of journal backups for the completed run.");
            }
        }

        // Drain the queue after each run finishes (MainForm.QueueAutoQueuedRun).
        run_next_queued_job(/*manual*/ false);
    }

    // ---- Journal retention -------------------------------------------------

    // Removes old completed journals at startup. Anything still resumable — the
    // active run, or a run left Running/Paused/Interrupted — is protected, since
    // resume reads straight from the journal.
    void prune_old_journals() {
        storage::JobJournalStore::PruneOptions options;
        options.retention_days = settings_.get_int("JournalRetentionDays", 0, 3650, 30);
        options.keep_minimum = settings_.get_int("JournalKeepMinimum", 0, 1000, 10);
        const bool compact = settings_.get_bool("CompactJournalsOnCompletion", true);
        if (options.retention_days <= 0 && !compact) return;

        for (const auto& run : job_manager_.get_recent_runs(1000)) {
            if (run.JournalPath.empty()) continue;
            switch (run.Status) {
                case storage::ManagedJobRunStatus::Running:
                case storage::ManagedJobRunStatus::Paused:
                case storage::ManagedJobRunStatus::Queued:
                case storage::ManagedJobRunStatus::Interrupted:
                    options.protected_paths.push_back(
                        storage::fsutil::utf8_to_wide(run.JournalPath));
                    break;
                default:
                    break;
            }
        }

        if (options.retention_days > 0) {
            storage::JobJournalStore::PruneResult result =
                storage::JobJournalStore::prune(options);
            if (result.journals_removed > 0) {
                append_log("Journal retention: removed " + std::to_string(result.journals_removed) +
                           " journal(s) older than " + std::to_string(options.retention_days) +
                           " days, reclaiming " + format_bytes_short(result.bytes_reclaimed) + ".");
            }
        }

        // Reclaim rotations left behind by runs that finished earlier; only the
        // still-resumable journals keep their backups.
        if (compact) {
            std::int64_t reclaimed =
                storage::JobJournalStore::compact_all(options.protected_paths);
            if (reclaimed > 0) {
                append_log("Journal maintenance: reclaimed " + format_bytes_short(reclaimed) +
                           " of backup rotations.");
            }
        }
    }

    // ---- Recovery prompt ---------------------------------------------------

    void maybe_prompt_recovery() {
        ui::RecoveryStartupInfo info = recovery_.initialize_session(settings_, launch_);
        if (!info.has_interrupted_run()) return;

        const auto& run = *info.interrupted_run;
        if (info.should_auto_resume) {
            append_log("Auto-resuming interrupted run: " + run.JobName);
            resume_interrupted(run);
            return;
        }
        if (!info.should_prompt) return;

        std::wstring text = storage::fsutil::utf8_to_wide(
            info.interruption_reason + "\n\nSource: " + run.Options.SourceRoot +
            "\nDestination: " + run.Options.DestinationRoot + "\n\nResume this run now?");
        if (confirm_box(text, L"XactCopy — Resume interrupted run")) {
            resume_interrupted(run);
        } else {
            recovery_.mark_resume_prompt_deferred(settings_, false);
        }
    }

    void resume_interrupted(const ui::RecoveryActiveRun& run) {
        models::CopyJobOptions options = run.Options;
        options.ResumeFromJournal = true;
        if (!run.JournalPath.empty()) options.ResumeJournalPathHint = run.JournalPath;
        apply_run_options_to_ui(options);
        append_log("Restored exact " +
                   std::string(options.OperationMode == models::JobOperationMode::ScanOnly
                                   ? "assessment"
                                   : "copy") +
                   " options from the interrupted run.");
        start_job_with(options, true, run.ManagedRunId, run.JobName);
    }
};

void add_explorer_source_path(ui::LaunchOptions& launch, const std::wstring& value) {
    std::string resolved = storage::fsutil::wide_to_utf8(value);
    if (resolved.empty()) return;
    // Skip unexpanded shell tokens (registration passes %1/%V/%*).
    if (resolved == "%1" || resolved == "%V" || resolved == "%*") return;
    for (const auto& existing : launch.explorer_source_paths) {
        if (models::detail::equals_ignore_case(existing, resolved)) return;
    }
    launch.explorer_source_paths.push_back(resolved);
}

// Takes the FULL command line (GetCommandLineW), which includes the executable
// path as argv[0]. Note wWinMain's lpCmdLine does NOT include it — passing that
// here would silently drop the first switch.
ui::LaunchOptions parse_launch_options(const wchar_t* command_line) {
    ui::LaunchOptions launch;
    int argc = 0;
    wchar_t** argv = CommandLineToArgvW(command_line, &argc);
    if (argv == nullptr) return launch;
    for (int i = 1; i < argc; ++i) { // argv[0] is the executable path
        const wchar_t* arg = argv[i];
        if (_wcsicmp(arg, L"--recovery-autostart") == 0) {
            launch.is_recovery_autostart = true;
        } else if (_wcsicmp(arg, L"--force-resume-prompt") == 0 ||
                   _wcsicmp(arg, L"--resume-interrupted") == 0) {
            launch.force_resume_prompt = true;
        } else if ((_wcsicmp(arg, L"--source") == 0 ||
                    _wcsicmp(arg, L"--from-explorer-folder") == 0) &&
                   i + 1 < argc) {
            launch.explorer_folder_path = storage::fsutil::wide_to_utf8(argv[++i]);
        } else if (_wcsnicmp(arg, L"--from-explorer-folder=", 23) == 0) {
            launch.explorer_folder_path = storage::fsutil::wide_to_utf8(arg + 23);
        } else if (_wcsicmp(arg, L"--from-explorer") == 0) {
            // Consume following non-flag arguments as the selection set.
            while (i + 1 < argc && _wcsnicmp(argv[i + 1], L"--", 2) != 0) {
                add_explorer_source_path(launch, argv[++i]);
            }
        } else if (_wcsnicmp(arg, L"--from-explorer=", 16) == 0) {
            add_explorer_source_path(launch, arg + 16);
        } else if (_wcsicmp(arg, L"--scan-from-explorer") == 0) {
            launch.explorer_scan_mode = true;
            while (i + 1 < argc && _wcsnicmp(argv[i + 1], L"--", 2) != 0) {
                add_explorer_source_path(launch, argv[++i]);
            }
        } else if (_wcsnicmp(arg, L"--scan-from-explorer=", 21) == 0) {
            launch.explorer_scan_mode = true;
            add_explorer_source_path(launch, arg + 21);
        }
    }
    LocalFree(argv);
    return launch;
}

} // namespace

int WINAPI wWinMain(HINSTANCE instance, HINSTANCE, PWSTR, int) {
    SetProcessDpiAwarenessContext(DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2);

    // GetCommandLineW (not lpCmdLine) so argv[0] is the executable, matching
    // parse_launch_options' expectations.
    const wchar_t* full_command_line = GetCommandLineW();

    // Single instance: hand our launch arguments to the running copy (so an
    // Explorer verb still lands on a selection) and surface its window.
    HANDLE mutex = CreateMutexW(nullptr, TRUE, L"Local\\XactCopyNative.SingleInstance");
    if (mutex != nullptr && GetLastError() == ERROR_ALREADY_EXISTS) {
        HWND existing = nullptr;
        // The first process owns the mutex before it registers its window
        // class. Wait for that startup gap instead of silently dropping an
        // Explorer selection launched a few milliseconds later.
        for (int attempt = 0; attempt < 200 && existing == nullptr; ++attempt) {
            existing = FindWindowW(WindowClassName, nullptr);
            if (existing == nullptr) Sleep(50);
        }
        if (existing != nullptr) {
            std::wstring payload(full_command_line);
            COPYDATASTRUCT data{};
            data.dwData = ForwardedLaunchId;
            data.cbData = static_cast<DWORD>((payload.size() + 1) * sizeof(wchar_t));
            data.lpData = const_cast<wchar_t*>(payload.c_str());
            DWORD_PTR forwarded = 0;
            if (!SendMessageTimeoutW(existing, WM_COPYDATA, 0,
                                     reinterpret_cast<LPARAM>(&data),
                                     SMTO_ABORTIFHUNG | SMTO_BLOCK, 5000, &forwarded)) {
                MessageBoxW(nullptr,
                            L"The running XactCopy window did not accept the Explorer selection. "
                            L"Please retry after it becomes responsive.",
                            L"XactCopy", MB_OK | MB_ICONWARNING);
            }
            if (IsIconic(existing)) ShowWindow(existing, SW_RESTORE);
            SetForegroundWindow(existing);
            CloseHandle(mutex);
            return 0;
        }
        // If the original process exited during startup, take ownership and
        // continue as the primary instance. Otherwise report the handoff
        // failure rather than pretending the requested selection was queued.
        if (WaitForSingleObject(mutex, 0) != WAIT_OBJECT_0) {
            MessageBoxW(nullptr,
                        L"XactCopy is starting, but its window was not available to receive this "
                        L"selection. Please try the Explorer command again.",
                        L"XactCopy", MB_OK | MB_ICONWARNING);
            CloseHandle(mutex);
            return 1;
        }
    }

    MainWindow window;
    int result = window.run(instance, parse_launch_options(full_command_line));
    if (mutex != nullptr) CloseHandle(mutex);
    return result;
}
