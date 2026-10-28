// -----------------------------------------------------------------------------
// File: src\ui\main_window.cpp
// Purpose: XactCopy native UI. Win32 main window over the ported supervisor/
//          recovery/settings core with the shared dark/light theme layer:
//          owner-drawn buttons/checkboxes/combos, severity-colored log,
//          accent-tinted progress bars, dark title bar + scrollbars, system
//          theme tracking, crash-resume prompt, and single-instance guard.
// -----------------------------------------------------------------------------

#include <deque>
#include <atomic>
#include <cstring>
#include <filesystem>
#include <fstream>
#include <limits>
#include <map>
#include <memory>
#include <mutex>
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
#include <windowsx.h>
#include <commctrl.h>
#include <dwmapi.h>
#include <shellapi.h>
#include <shobjidl.h>

#include "app_icon.h"
#include "dialogs.h"
#include "elevation.h"
#include "task_model.h"
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
constexpr UINT WM_APP_LOG_TASK = WM_APP + 9;
constexpr UINT WM_APP_SOURCE_SIZE = WM_APP + 10;   // lParam: new SourceSizeResult*
constexpr UINT WM_APP_RUN_NEXT_TASK = WM_APP + 11; // start the next queued session task

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
constexpr int IdContinueCheck = 1019;
constexpr int IdSkipKnownBadCheck = 1020;
constexpr int IdWaitMediaCheck = 1021;
constexpr int IdFragileCheck = 1022;
constexpr int IdBufferEdit = 1023;
constexpr int IdRetriesEdit = 1024;
constexpr int IdTimeoutEdit = 1025;
constexpr int IdCurrentBar = 1026;
constexpr int IdOverallBar = 1027;
constexpr int IdLogFilter = 1029;
constexpr int IdLogFind = 1030;
constexpr int IdLogFindNext = 1031;
constexpr int IdLogCopySelected = 1032;
constexpr int IdLogCopyAll = 1033;
constexpr int IdLogSave = 1034;

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
constexpr int IdSourcePaste = 1042;
constexpr int IdSourcesList = 1043;
constexpr int IdTasksList = 1044;
constexpr int IdAddTaskButton = 1045;
constexpr int IdOptionsToggle = 1046;
constexpr int IdRecentDestinations = 1047;
constexpr int IdPauseListCheck = 1048;
constexpr int IdClearFinished = 1049;
constexpr int IdModeCopyTab = 1050;
constexpr int IdModeScanTab = 1051;
// Context-menu commands for the sources and tasks lists.
constexpr int IdSourceRemove = 1060;
constexpr int IdSourceContentsOnly = 1061;
constexpr int IdSourceOpenLocation = 1062;
constexpr int IdSourceCopyPath = 1063;
constexpr int IdTaskStartNow = 1070;
constexpr int IdTaskMoveUp = 1071;
constexpr int IdTaskMoveDown = 1072;
constexpr int IdTaskRemove = 1073;
constexpr int IdTaskRetry = 1074;
constexpr int IdTaskOpenDestination = 1075;
constexpr int IdTaskSaveAsJob = 1076;
constexpr int IdTaskEdit = 1077;
constexpr int IdRecentDestinationBase = 1200; // 1200..1209
constexpr int MaxRecentDestinations = 10;

// Explorer invokes a context-menu verb once per selected item, so a multi-select
// arrives as a burst of launches. Selections are accumulated and applied once
// this quiet period elapses, giving one source root and one destination prompt.
constexpr UINT IdSelectionTimer = 1;
constexpr UINT IdDiagnosticsTimer = 2;
// Polls an elevated relaunch until it accepts the handoff or exits.
constexpr UINT IdElevationTimer = 3;
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
    // Set by wWinMain when this process is the elevated relaunch of a window
    // that asked for a direct-NTFS assessment; run() starts it after startup.
    void set_incoming_handoff(ui::elevation::HandoffRequest request) {
        incoming_handoff_ = std::move(request);
    }

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

        auto post_persistence_warning = [this](const std::string& message) {
            queue_log("WARNING: " + message);
        };
        job_manager_.on_persistence_warning = post_persistence_warning;
        recovery_.on_persistence_warning = post_persistence_warning;

        ui::apply_window_icons(hwnd_);
        ui::set_dark_title_bar(hwnd_, theme_.dark && theme_.themed_chrome);
        if (process_elevated_) {
            // Explorer verbs and drag-and-drop come from normal-integrity
            // processes; let them reach this High-integrity window.
            ui::elevation::allow_unelevated_messages(hwnd_);
            SetWindowTextW(hwnd_, L"XactCopy (Administrator)");
        }
        ui::elevation::cleanup_stale_handoffs();
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

        sources_changed();
        refresh_tasks_view();
        restore_pending_tasks();

        // An elevated relaunch exists to run one specific request; start it
        // before any startup queue drain could take the worker first.
        if (incoming_handoff_.has_value()) {
            ui::elevation::HandoffRequest request = std::move(*incoming_handoff_);
            incoming_handoff_.reset();
            append_log("Running with administrator rights for direct NTFS volume reads.");
            apply_run_options_to_ui(request.Options);
            start_job_with(request.Options, request.FromRecovery, request.ManagedRunId,
                           request.DisplayName);
        }

        // Startup queue drain (MainForm honors AutoRunQueuedJobsOnStartup).
        if (!supervisor_.is_job_running() && !starting_ &&
            settings_.get_bool("AutoRunQueuedJobsOnStartup", false)) {
            append_log("Startup queue mode enabled. Looking for queued jobs.");
            run_next_queued_job(/*manual*/ false);
        }

        MSG message;
        while (GetMessageW(&message, nullptr, 0, 0) > 0) {
            if (handle_log_shortcuts(message)) continue;
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
    HWND destination_edit_ = nullptr;
    HWND source_browse_ = nullptr; // "+ Folder"
    HWND destination_browse_ = nullptr;
    HWND recent_destinations_button_ = nullptr;

    // "What to copy" panel: one row per source item (see task_model.h).
    HWND sources_list_ = nullptr;
    HWND sources_footer_ = nullptr;
    HWND source_paste_ = nullptr;
    std::vector<ui::SourceItem> sources_;
    struct SourceRowInfo {
        int icon = -1;               // system image list index
        std::int64_t bytes = -1;     // -1 while counting
        std::int64_t files = 0;
        bool size_incomplete = false; // some entries could not be read
    };
    std::vector<SourceRowInfo> source_rows_;
    HIMAGELIST system_small_icons_ = nullptr;
    struct SourceSizeResult {
        std::uint64_t generation = 0;
        std::wstring path;
        std::int64_t bytes = 0;
        std::int64_t files = 0;
        bool incomplete = false;
    };
    // Background folder sizing: one long-lived thread, newest request wins.
    std::mutex sizer_lock_;
    std::condition_variable sizer_ready_;
    std::vector<std::wstring> sizer_request_;
    std::atomic<std::uint64_t> sizer_generation_{0};
    bool sizer_stopping_ = false;
    std::thread sizer_thread_;

    // Options section collapses to one summary line.
    HWND options_toggle_ = nullptr;
    bool options_expanded_ = false;
    HWND add_task_button_ = nullptr;

    // "Tasks" panel: the session task list (task_model.h).
    HWND tasks_label_ = nullptr;
    HWND tasks_list_ = nullptr;
    HWND pause_list_check_ = nullptr;
    HWND clear_finished_button_ = nullptr;
    HWND detail_label_ = nullptr;
    ui::TaskList tasks_;
    std::uint64_t running_task_id_ = 0; // task bound to the job currently starting/running
    bool tasks_running_ = false;        // user started the list; advance automatically
    bool closing_for_handoff_ = false;
    std::unique_ptr<ui::PendingTaskWriter> pending_task_writer_;
    // A start's result can be processed before its WM_APP_START_DONE (a run
    // that ends at once); the late start notification must not revive it.
    bool result_applied_since_start_ = false;

    // Copy / Assess segmented switch. mode_combo_ stays as the hidden source
    // of truth so every existing mode read/write keeps working.
    HWND mode_copy_tab_ = nullptr;
    HWND mode_scan_tab_ = nullptr;
    // Captions for the three mode-dependent combos.
    HWND profile_label_ = nullptr;
    HWND policy_label_ = nullptr;
    HWND verify_label_ = nullptr;
    // Fills the destination row's slot in Assess mode so both modes keep the
    // same geometry.
    HWND scan_info_label_ = nullptr;
    // Client height the current mode/options need; feeds WM_GETMINMAXINFO.
    int required_client_height_ = 0;
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
    HWND settings_button_ = nullptr;
    HWND pause_button_ = nullptr;
    HWND cancel_button_ = nullptr;
    HWND current_label_ = nullptr;
    HWND current_bar_ = nullptr;
    HWND overall_bar_ = nullptr;
    HWND file_progress_label_ = nullptr;
    HWND overall_progress_label_ = nullptr;
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
    HWND log_filter_ = nullptr;
    HWND log_find_ = nullptr;
    HWND log_find_button_ = nullptr;
    ui::LogModel log_model_;
    ui::LogFilter log_filter_mode_ = ui::LogFilter::All;
    struct QueuedLog {
        std::string text;
        std::chrono::system_clock::time_point timestamp;
    };
    std::mutex log_queue_lock_;
    std::deque<QueuedLog> log_queue_;
    bool log_wake_pending_ = false;
    std::size_t log_max_characters_ = 0;
    std::thread log_save_thread_;
    std::atomic<bool> log_save_running_{false};
    enum class LogTaskKind { Find, Copy, Save };
    struct LogTaskResult {
        LogTaskKind kind = LogTaskKind::Find;
        ui::LogFilter filter = ui::LogFilter::All;
        bool found = false;
        std::uint64_t sequence = 0;
        std::wstring text;
        HGLOBAL clipboard = nullptr;
        std::string error;
        ~LogTaskResult() { if (clipboard != nullptr) GlobalFree(clipboard); }
    };
    int log_horizontal_extent_ = 0;
    HFONT font_ = nullptr;
    HFONT mono_font_ = nullptr;
    HFONT header_font_ = nullptr; // semibold section headings
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

    // Direct-NTFS elevation (see elevation.h). `outgoing_handoff_` is a start
    // this window passed to an elevated relaunch and is waiting on;
    // `incoming_handoff_` is a start this (elevated) window received.
    struct OutgoingHandoff {
        ui::elevation::HandoffRequest request;
        std::string token;
        HANDLE process = nullptr;
        ULONGLONG started_tick = 0;
    };
    const bool process_elevated_ = ui::elevation::is_process_elevated();
    std::optional<OutgoingHandoff> outgoing_handoff_;
    std::optional<ui::elevation::HandoffRequest> incoming_handoff_;
    std::string pending_recovery_run_id_; // crash-recovery run being resumed by start_job_with
    bool restore_maximized_ = false;    // deferred maximize from saved placement
    bool checking_updates_ = false;
    bool update_check_show_dialog_ = false;
    std::thread update_thread_;

    // Explorer verbs arrive as a burst of launches, one per item. Items join
    // the sources list at once; the coalescing timer only decides, once the
    // burst ends, whether to switch mode and ask for a destination.
    bool selection_timer_active_ = false;
    bool selection_wants_destination_ = false;
    HWND source_add_files_ = nullptr;
    HWND clear_selection_ = nullptr;
    std::map<int, bool> check_states_;
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
    bool bar_show_percentage_ = true;
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
                limits->ptMinTrackSize.x = MulDiv(860, dpi, 96);
                limits->ptMinTrackSize.y = MulDiv(740, dpi, 96); // two columns + log
                // Never below what the current mode and options need.
                RECT window{}, client{};
                if (required_client_height_ > 0 && GetWindowRect(hwnd_, &window) &&
                    GetClientRect(hwnd_, &client)) {
                    const int frame = (window.bottom - window.top) - client.bottom;
                    limits->ptMinTrackSize.y = std::max<LONG>(limits->ptMinTrackSize.y,
                                                              required_client_height_ + frame);
                }
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
                const HWND control = reinterpret_cast<HWND>(lparam);
                // Captions and secondary lines recede; values stay full contrast.
                const bool muted = control == profile_label_ || control == policy_label_ ||
                                   control == verify_label_ || control == sources_footer_ ||
                                   control == scan_info_label_ ||
                                   control == file_progress_label_ ||
                                   control == overall_progress_label_ ||
                                   control == journal_label_ || control == diagnostics_label_;
                SetTextColor(dc, muted ? theme_.muted_text : theme_.text);
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
                drain_log_queue();
                return 0;
            }
            case WM_APP_LOG_TASK: {
                std::unique_ptr<LogTaskResult> result(reinterpret_cast<LogTaskResult*>(lparam));
                finish_log_task(*result);
                return 0;
            }
            case WM_CONTEXTMENU:
                if (reinterpret_cast<HWND>(wparam) == log_list_) {
                    show_log_context_menu(lparam);
                    return 0;
                }
                return DefWindowProcW(hwnd_, message, wparam, lparam);
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
                if (ui::SessionTask* task = tasks_.find(running_task_id_)) {
                    if (ui::is_active(task->state)) {
                        task->state = paused_ ? ui::TaskState::Paused : ui::TaskState::Running;
                        invalidate_task_row(task->id);
                    }
                }
                update_menu_state();
                return 0;
            case WM_APP_SOURCE_SIZE: {
                std::unique_ptr<SourceSizeResult> result(reinterpret_cast<SourceSizeResult*>(lparam));
                on_source_size(*result);
                return 0;
            }
            case WM_APP_RUN_NEXT_TASK:
                on_run_next_task();
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
            case WM_POWERBROADCAST:
                // Keep the worker watchdog from mistaking a sleep for a hang.
                if (wparam == PBT_APMSUSPEND) {
                    supervisor_.notify_system_suspending();
                } else if (wparam == PBT_APMRESUMEAUTOMATIC || wparam == PBT_APMRESUMESUSPEND) {
                    supervisor_.notify_system_resumed();
                }
                return TRUE;
            case WM_TIMER:
                if (wparam == IdSelectionTimer) {
                    commit_selection();
                    return 0;
                }
                if (wparam == IdDiagnosticsTimer) {
                    refresh_diagnostics_label();
                    return 0;
                }
                if (wparam == IdElevationTimer) {
                    poll_elevated_handoff();
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
                if (!closing_for_handoff_ && !supervisor_.is_job_running()) {
                    const auto counts = tasks_.counts();
                    if (counts.queued > 0 &&
                        !confirm_box(std::to_wstring(counts.queued) +
                                         (counts.queued == 1 ? L" task has" : L" tasks have") +
                                         L" not started yet. Exit anyway?\n\nThey will be offered "
                                         L"again the next time XactCopy starts.",
                                     L"XactCopy", ui::MessageIcon::Question)) {
                        return 0;
                    }
                }
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
                KillTimer(hwnd_, IdElevationTimer);
                if (outgoing_handoff_.has_value() && outgoing_handoff_->process != nullptr) {
                    CloseHandle(outgoing_handoff_->process);
                    outgoing_handoff_.reset();
                }
                if (start_thread_.joinable()) start_thread_.join();
                if (update_thread_.joinable()) update_thread_.join();
                stop_sizer();
                if (pending_task_writer_ != nullptr) {
                    // Final synchronous write; the window is going away. The
                    // running task (if any) is not "pending": crash recovery
                    // owns it.
                    pending_task_writer_->flush(
                        ui::pending_tasks_to_json(tasks_.queued(), tasks_running_));
                }
                supervisor_.stop();
                if (log_save_thread_.joinable()) log_save_thread_.join();
                {
                    MSG pending{};
                    while (PeekMessageW(&pending, hwnd_, WM_APP_LOG_TASK, WM_APP_LOG_TASK, PM_REMOVE))
                        delete reinterpret_cast<LogTaskResult*>(pending.lParam);
                }
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
        if (header_font_ != nullptr) DeleteObject(header_font_);
        font_ = CreateFontW(-MulDiv(9, eff, 72), 0, 0, 0, FW_NORMAL, FALSE,
                            FALSE, FALSE, DEFAULT_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS,
                            CLEARTYPE_QUALITY, DEFAULT_PITCH, L"Segoe UI");
        header_font_ = CreateFontW(-MulDiv(10, eff, 72), 0, 0, 0, FW_SEMIBOLD, FALSE,
                                   FALSE, FALSE, DEFAULT_CHARSET, OUT_DEFAULT_PRECIS,
                                   CLIP_DEFAULT_PRECIS, CLEARTYPE_QUALITY, DEFAULT_PITCH,
                                   L"Segoe UI");
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
            source_label_,     destination_label_,   destination_edit_,  sources_list_,
            source_browse_,    destination_browse_,  source_add_files_,  source_paste_,
            sources_footer_,   recent_destinations_button_, options_toggle_, add_task_button_,
            tasks_label_,      tasks_list_,          pause_list_check_,  clear_finished_button_,
            detail_label_,     mode_copy_tab_,       mode_scan_tab_,     profile_label_,
            policy_label_,     verify_label_,        scan_info_label_,
            clear_selection_,  mode_combo_,          engine_combo_,
            overwrite_combo_,  verify_combo_,        salvage_check_,    resume_check_,
            map_check_,        adaptive_check_,      continue_check_,   skip_known_bad_check_,
            wait_media_check_, fragile_check_,       buffer_label_,     buffer_edit_,
            retries_label_,    retries_edit_,        timeout_label_,    timeout_edit_,
            start_button_,     pause_button_,        cancel_button_,    settings_button_,
            log_filter_,       log_find_,            log_find_button_,
            current_label_,    file_progress_label_, overall_progress_label_,
            stats_label_,      status_label_,
            throughput_label_, eta_label_,           buffer_usage_label_, rescue_label_,
            job_summary_label_, journal_label_,      diagnostics_label_,
        };
        for (HWND control : ui_controls) {
            if (control != nullptr) {
                SendMessageW(control, WM_SETFONT, reinterpret_cast<WPARAM>(font_), TRUE);
            }
        }
        for (HWND heading : {source_label_, tasks_label_, detail_label_}) {
            if (heading != nullptr) {
                SendMessageW(heading, WM_SETFONT, reinterpret_cast<WPARAM>(header_font_), TRUE);
            }
        }
        if (log_list_ != nullptr) {
            SendMessageW(log_list_, WM_SETFONT, reinterpret_cast<WPARAM>(mono_font_), TRUE);
            SendMessageW(log_list_, LB_SETITEMHEIGHT, 0,
                         MulDiv(log_points + 6, effective_dpi(), 72));
        }
        // Sources: one line (icon, name, folder, size). Tasks: two lines plus a
        // thin inline progress bar.
        if (sources_list_ != nullptr) {
            SendMessageW(sources_list_, LB_SETITEMHEIGHT, 0, MulDiv(24, effective_dpi(), 96));
        }
        if (tasks_list_ != nullptr) {
            SendMessageW(tasks_list_, LB_SETITEMHEIGHT, 0, MulDiv(46, effective_dpi(), 96));
        }
        const int combo_item_height = MulDiv(22, effective_dpi(), 96);
        for (HWND combo : {mode_combo_, engine_combo_, overwrite_combo_, verify_combo_, log_filter_}) {
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

    void rebuild_log_horizontal_extent() {
        if (log_list_ == nullptr || mono_font_ == nullptr) return;
        HDC dc = GetDC(log_list_);
        if (dc == nullptr) return;
        HGDIOBJ old_font = SelectObject(dc, mono_font_);
        TEXTMETRICW metrics{};
        GetTextMetricsW(dc, &metrics);
        const auto safe_width = std::min<std::size_t>(
            log_max_characters_ * static_cast<std::size_t>(std::max<LONG>(1, metrics.tmMaxCharWidth)),
            static_cast<std::size_t>(std::numeric_limits<int>::max() - 16));
        int extent = static_cast<int>(safe_width) + MulDiv(8, effective_dpi(), 96);
        SelectObject(dc, old_font);
        ReleaseDC(log_list_, dc);
        log_horizontal_extent_ = extent;
        SendMessageW(log_list_, LB_SETHORIZONTALEXTENT,
                     static_cast<WPARAM>(log_horizontal_extent_), 0);
    }

    void reset_log_horizontal_extent() {
        log_horizontal_extent_ = 0;
        log_max_characters_ = 0;
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

        // ---- What to copy ----
        source_label_ = create(L"STATIC", L"What to copy", SS_NOPREFIX | SS_CENTERIMAGE, 0);
        source_add_files_ =
            create(L"BUTTON", L"+ Files", WS_TABSTOP | BS_OWNERDRAW, IdSourceAddFiles);
        source_browse_ = create(L"BUTTON", L"+ Folder", WS_TABSTOP | BS_OWNERDRAW, IdSourceBrowse);
        source_paste_ = create(L"BUTTON", L"Paste", WS_TABSTOP | BS_OWNERDRAW, IdSourcePaste);
        clear_selection_ =
            create(L"BUTTON", L"Clear", WS_TABSTOP | BS_OWNERDRAW, IdClearSelection);
        // Owner-data list (LBS_NODATA): rows are painted from sources_.
        sources_list_ = create(L"LISTBOX", L"",
                               WS_TABSTOP | WS_VSCROLL | WS_BORDER | LBS_NOINTEGRALHEIGHT |
                                   LBS_NODATA | LBS_OWNERDRAWFIXED | LBS_EXTENDEDSEL |
                                   LBS_NOTIFY,
                               IdSourcesList);
        SetWindowSubclass(sources_list_, &panel_list_subclass, 2, reinterpret_cast<DWORD_PTR>(this));
        sources_footer_ = create(L"STATIC", L"", SS_ENDELLIPSIS | SS_NOPREFIX, 0);
        {
            SHFILEINFOW info{};
            system_small_icons_ = reinterpret_cast<HIMAGELIST>(SHGetFileInfoW(
                L"C:\\", FILE_ATTRIBUTE_DIRECTORY, &info, sizeof(info),
                SHGFI_SYSICONINDEX | SHGFI_SMALLICON | SHGFI_USEFILEATTRIBUTES));
        }
        sizer_thread_ = std::thread([this] { sizer_loop(); });

        destination_label_ = create(L"STATIC", L"Destination", SS_CENTERIMAGE, 0);
        destination_edit_ =
            create(L"EDIT", L"", WS_TABSTOP | ES_AUTOHSCROLL | WS_BORDER, IdDestinationEdit);
        destination_browse_ =
            create(L"BUTTON", L"Browse...", WS_TABSTOP | BS_OWNERDRAW, IdDestinationBrowse);
        recent_destinations_button_ = create(L"BUTTON", L"\x25BE", WS_TABSTOP | BS_OWNERDRAW,
                                             IdRecentDestinations);

        DWORD combo_style = WS_TABSTOP | CBS_DROPDOWNLIST | CBS_OWNERDRAWFIXED | CBS_HASSTRINGS;
        mode_combo_ = create(L"COMBOBOX", L"", combo_style, IdModeCombo);
        engine_combo_ = create(L"COMBOBOX", L"", combo_style, IdEngineCombo);
        overwrite_combo_ = create(L"COMBOBOX", L"", combo_style, IdOverwriteCombo);
        verify_combo_ = create(L"COMBOBOX", L"", combo_style, IdVerifyCombo);
        for (const wchar_t* item : {L"Verified Copy / Recovery", L"Assess Readable Files"}) {
            SendMessageW(mode_combo_, CB_ADDSTRING, 0, reinterpret_cast<LPARAM>(item));
        }
        ShowWindow(mode_combo_, SW_HIDE);
        mode_copy_tab_ = create(L"BUTTON", L"Copy", WS_TABSTOP | BS_OWNERDRAW, IdModeCopyTab);
        mode_scan_tab_ =
            create(L"BUTTON", L"Assess readability", WS_TABSTOP | BS_OWNERDRAW, IdModeScanTab);
        profile_label_ = create(L"STATIC", L"Profile", SS_NOPREFIX | SS_CENTERIMAGE, 0);
        policy_label_ = create(L"STATIC", L"If file exists", SS_NOPREFIX | SS_CENTERIMAGE, 0);
        verify_label_ = create(L"STATIC", L"Verify", SS_NOPREFIX | SS_CENTERIMAGE, 0);
        scan_info_label_ = create(L"STATIC",
                                  L"Read-only: nothing is written except the optional bad-range map.",
                                  SS_NOPREFIX | SS_CENTERIMAGE | SS_ENDELLIPSIS, 0);

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

        buffer_label_ = create(L"STATIC", L"Buffer (MB)", SS_NOPREFIX | SS_CENTERIMAGE, 0);
        buffer_edit_ = create(L"EDIT", L"", WS_TABSTOP | ES_NUMBER | ES_AUTOHSCROLL | WS_BORDER,
                              IdBufferEdit);
        retries_label_ = create(L"STATIC", L"Retries", SS_NOPREFIX | SS_CENTERIMAGE, 0);
        retries_edit_ = create(L"EDIT", L"", WS_TABSTOP | ES_NUMBER | ES_AUTOHSCROLL | WS_BORDER,
                               IdRetriesEdit);
        timeout_label_ = create(L"STATIC", L"Timeout (s)", SS_NOPREFIX | SS_CENTERIMAGE, 0);
        timeout_edit_ = create(L"EDIT", L"", WS_TABSTOP | ES_NUMBER | ES_AUTOHSCROLL | WS_BORDER,
                               IdTimeoutEdit);

        options_toggle_ = create(L"BUTTON", L"", WS_TABSTOP | BS_OWNERDRAW, IdOptionsToggle);
        options_expanded_ = settings_.get_bool("UiOptionsExpanded", false);

        start_button_ = create(L"BUTTON", L"Start \x2192", WS_TABSTOP | BS_OWNERDRAW, IdStartButton);
        add_task_button_ =
            create(L"BUTTON", L"Add to tasks", WS_TABSTOP | BS_OWNERDRAW, IdAddTaskButton);
        pause_button_ = create(L"BUTTON", L"Pause", WS_TABSTOP | BS_OWNERDRAW, IdPauseButton);
        cancel_button_ = create(L"BUTTON", L"Cancel", WS_TABSTOP | BS_OWNERDRAW, IdCancelButton);
        settings_button_ =
            create(L"BUTTON", L"Settings...", WS_TABSTOP | BS_OWNERDRAW, IdSettingsButton);
        EnableWindow(pause_button_, FALSE);
        EnableWindow(cancel_button_, FALSE);

        // ---- Tasks ----
        tasks_label_ = create(L"STATIC", L"Tasks", SS_NOPREFIX | SS_CENTERIMAGE | SS_ENDELLIPSIS, 0);
        pause_list_check_ = create(L"BUTTON", L"Pause after current task",
                                   WS_TABSTOP | BS_OWNERDRAW, IdPauseListCheck);
        clear_finished_button_ =
            create(L"BUTTON", L"Clear finished", WS_TABSTOP | BS_OWNERDRAW, IdClearFinished);
        tasks_list_ = create(L"LISTBOX", L"",
                             WS_TABSTOP | WS_VSCROLL | WS_BORDER | LBS_NOINTEGRALHEIGHT |
                                 LBS_NODATA | LBS_OWNERDRAWFIXED | LBS_NOTIFY,
                             IdTasksList);
        SetWindowSubclass(tasks_list_, &panel_list_subclass, 3, reinterpret_cast<DWORD_PTR>(this));
        detail_label_ =
            create(L"STATIC", L"No task running", SS_NOPREFIX | SS_ENDELLIPSIS | SS_CENTERIMAGE, 0);
        pending_task_writer_ = std::make_unique<ui::PendingTaskWriter>();

        current_label_ = create(L"STATIC", L"Current: —", SS_ENDELLIPSIS | SS_NOPREFIX, 0);
        // Owner-drawn flat progress bars: classic msctls_progress32 draws a
        // sunken 3D border (the white edge seen on resize) once its visual
        // styles are stripped to honor a custom bar color, so paint them
        // ourselves for a flat, borderless, consistent look.
        current_bar_ = create(L"STATIC", L"", SS_OWNERDRAW, IdCurrentBar);
        overall_bar_ = create(L"STATIC", L"", SS_OWNERDRAW, IdOverallBar);
        file_progress_label_ = create(L"STATIC", L"File", SS_CENTERIMAGE | SS_NOPREFIX, 0);
        overall_progress_label_ = create(L"STATIC", L"Overall", SS_CENTERIMAGE | SS_NOPREFIX, 0);
        stats_label_ = create(L"STATIC", L"\x2014", SS_NOPREFIX | SS_ENDELLIPSIS, 0);
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
        status_label_ = create(L"STATIC", L"Status: Idle", SS_NOPREFIX | SS_CENTERIMAGE, 0);
        // WS_BORDER (flat 1px, matches the edits) instead of WS_EX_CLIENTEDGE,
        // which draws a white sunken 3D edge the dark theme can't repaint.
        log_filter_ = create(L"COMBOBOX", L"", combo_style, IdLogFilter);
        for (const wchar_t* item : {L"All", L"Warnings and errors", L"Errors only"}) {
            SendMessageW(log_filter_, CB_ADDSTRING, 0, reinterpret_cast<LPARAM>(item));
        }
        SendMessageW(log_filter_, CB_SETCURSEL, 0, 0);
        log_find_ = create(L"EDIT", L"", WS_TABSTOP | ES_AUTOHSCROLL | WS_BORDER, IdLogFind);
        log_find_button_ = create(L"BUTTON", L"Find next", WS_TABSTOP | BS_OWNERDRAW, IdLogFindNext);
        log_model_.set_capacity(static_cast<std::size_t>(
            settings_.get_int("UiMaxLogLines", 1000, 1000000, 50000)));
        log_list_ = create(L"LISTBOX", L"",
                           WS_TABSTOP | WS_VSCROLL | WS_HSCROLL | WS_BORDER | LBS_NOINTEGRALHEIGHT |
                               LBS_DISABLENOSCROLL | LBS_NODATA | LBS_OWNERDRAWFIXED,
                           IdLogList, 0, mono_font_);
        SetWindowSubclass(log_list_, &log_list_subclass, 1, reinterpret_cast<DWORD_PTR>(this));
        SendMessageW(log_list_, LB_SETITEMHEIGHT, 0,
                     MulDiv(log_points + 6, effective_dpi(), 72));

        tooltips_.create(hwnd_, static_cast<UINT>(effective_dpi()), theme_.dark);
        configure_tooltips();

        build_menu_bar();
        apply_fonts(); // also sets the owner-draw list row heights
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

        add(source_label_, L"The files and folders to copy or assess. Each item lands in the destination under its own name; items from different folders become separate tasks.");
        add(sources_list_, L"Files and folders to copy. Drag items here, paste paths (Ctrl+V), press Del to remove, or right-click a folder to copy only its contents.");
        add(source_browse_, L"Add one or more folders. Each folder is copied as itself unless you choose \"Copy folder contents only\" from its right-click menu.");
        add(source_add_files_, L"Add one or more files.");
        add(source_paste_, L"Add files copied in Explorer, or paths copied as text (one per line).");
        add(clear_selection_, L"Remove every item from the list.");
        add(sources_footer_, L"Item count and total size. Sizes are counted in the background.");
        add(recent_destinations_button_, L"Pick a recently used destination.");
        add(options_toggle_, L"Show or hide the detailed options. The summary shows what the next task will use.");
        add(add_task_button_, L"Add tasks for these sources to the task list without starting them.");
        add(tasks_label_, L"Tasks run one at a time, in order. Right-click a task for more actions.");
        add(tasks_list_, L"Queued, running, and finished tasks. Enter starts the selected queued task; Del removes it; right-click for retry, move, open destination, and save as job.");
        add(pause_list_check_, L"When checked, the next queued task waits after the current one finishes.");
        add(clear_finished_button_, L"Remove finished tasks from the list. Run history stays in the Job Manager.");
        add(detail_label_, L"Live detail for the running task.");
        add(destination_label_, L"The folder that receives copied files. Assess Readable Files mode is read-only and does not require a destination.");
        add(destination_edit_, L"Enter the destination folder. XactCopy stages and verifies each file beside its final path before replacement.");
        add(destination_browse_, L"Choose the destination folder.");

        add(mode_copy_tab_, L"Copy or recover the listed items into a destination, with verification.");
        add(mode_scan_tab_, L"Read the listed items without copying them and record which ranges are unreadable. Nothing is written except the optional bad-range map.");
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

        add(start_button_, L"Add tasks for these sources and start running the task list (Ctrl+Enter). While a task runs, this queues them instead.");
        add(pause_button_, L"Pause or resume scheduling and managed I/O at a safe cancellation point. An in-flight Windows API call may finish first.");
        add(cancel_button_, L"Request cancellation. If a storage driver does not respond within five seconds, XactCopy force-stops the isolated worker. Click again to force-stop immediately.");
        add(settings_button_, L"Open advanced appearance, copy, performance, diagnostics, verification, recovery, and Explorer-integration settings.");

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
        add(log_filter_, L"Show all log lines, warnings and errors, or errors only.");
        add(log_find_, L"Find text in the visible log. Ctrl+F focuses this box; F3 finds the next match.");
        add(log_find_button_, L"Find the next match in the visible log, wrapping at the end.");
        add(status_label_, L"Supervisor state for worker startup, connection, running, pause, cancellation, recovery, and completion.");
    }

    // Mirrors the .NET MainForm menu strip (File / Tools / Jobs / Help).
    void build_menu_bar() {
        HMENU file_menu = CreatePopupMenu();
        AppendMenuW(file_menu, MF_STRING, IdMenuStartCopy, L"&Start\tCtrl+Enter");
        AppendMenuW(file_menu, MF_STRING, IdAddTaskButton, L"&Add to tasks");
        AppendMenuW(file_menu, MF_STRING, IdMenuPauseCopy, L"&Pause Operation");
        AppendMenuW(file_menu, MF_STRING, IdMenuResumeCopy, L"&Resume Operation");
        AppendMenuW(file_menu, MF_STRING, IdMenuCancelCopy, L"&Cancel Operation");
        AppendMenuW(file_menu, MF_SEPARATOR, 0, nullptr);
        AppendMenuW(file_menu, MF_STRING, IdMenuSaveDefaults,
                    L"Save current settings as &defaults");
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
        set_enabled(IdMenuStartCopy, !outgoing_handoff_.has_value()); // queues while a task runs
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
        ui::elevation::open_as_user(nullptr, folder);
    }

    // The bad-range map is keyed by the root a job reads from: the first
    // source group's root (the folder holding the first item, or a
    // contents-only folder itself).
    std::string primary_source_root() const {
        const auto groups = ui::group_sources(sources_);
        return groups.empty() ? std::string() : storage::fsutil::wide_to_utf8(groups.front().root);
    }

    std::optional<std::wstring> selected_source_map_path() {
        std::string source = primary_source_root();
        if (ui_is_blank(source)) {
            warn_box(L"Add a source before inspecting its bad-range map.");
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
            storage::fsutil::utf8_to_wide(primary_source_root()));
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
            journal_only_map.SourceRoot = primary_source_root();
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
        for (HWND control : {log_list_, log_find_, sources_list_, tasks_list_, destination_edit_,
                             buffer_edit_, retries_edit_, timeout_edit_}) {
            ui::apply_control_theme(control, theme_.dark);
        }
        for (HWND combo : {mode_combo_, engine_combo_, overwrite_combo_, verify_combo_, log_filter_}) {
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
            replace_combo_items(engine_combo_, {L"Auto", L"Fast", L"Precise"});
            replace_combo_items(overwrite_combo_,
                                {L"Standard file reads", L"Direct NTFS (admin)"});
            replace_combo_items(verify_combo_, {L"Save to bad-range map", L"Do not save"});
            SendMessageW(engine_combo_, CB_SETCURSEL, scan_profile_index_, 0);
            SendMessageW(overwrite_combo_, CB_SETCURSEL, scan_backend_index_, 0);
            SendMessageW(verify_combo_, CB_SETCURSEL, scan_map_index_, 0);
            SetWindowTextW(resume_check_, L"Resume previous scan");
        } else {
            replace_combo_items(engine_combo_,
                                {L"Verified Copy", L"Recover Media", L"Custom"});
            replace_combo_items(overwrite_combo_,
                                {L"Overwrite", L"Skip existing", L"Overwrite if newer",
                                 L"Stop on conflict"});
            replace_combo_items(verify_combo_,
                                {L"None", L"Sampled", L"Full (SHA)"});
            SendMessageW(engine_combo_, CB_SETCURSEL, copy_profile_index_, 0);
            SendMessageW(overwrite_combo_, CB_SETCURSEL, copy_overwrite_index_, 0);
            SendMessageW(verify_combo_, CB_SETCURSEL, copy_verify_index_, 0);
            SetWindowTextW(resume_check_, L"Resume verified progress");
            if (user_changed) launch_.explorer_scan_mode = false;
        }

        const bool scan = requested == 1;
        for (HWND control : {destination_label_, destination_edit_, destination_browse_,
                             recent_destinations_button_}) {
            if (control != nullptr) ShowWindow(control, scan ? SW_HIDE : SW_SHOW);
        }
        if (scan_info_label_ != nullptr) ShowWindow(scan_info_label_, scan ? SW_SHOW : SW_HIDE);
        SetWindowTextW(source_label_, scan ? L"What to assess" : L"What to copy");
        SetWindowTextW(profile_label_, scan ? L"Scan" : L"Profile");
        SetWindowTextW(policy_label_, scan ? L"Read method" : L"If file exists");
        SetWindowTextW(verify_label_, scan ? L"Findings" : L"Verify");
        InvalidateRect(mode_copy_tab_, nullptr, FALSE);
        InvalidateRect(mode_scan_tab_, nullptr, FALSE);
        layout();
        ensure_window_fits_layout(); // the Options section shows/hides salvage per mode
        if (options_toggle_ != nullptr) InvalidateRect(options_toggle_, nullptr, TRUE);
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
        // Explorer launch paths join the sources list in apply_explorer_launch_options.
    }

    void wire_supervisor_events() {
        HWND hwnd = hwnd_;
        supervisor_.on_log = [this](const std::string& message) {
            queue_log(message);
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
               id == IdWaitMediaCheck || id == IdFragileCheck || id == IdPauseListCheck;
    }

    static bool is_button_id(int id) {
        return id == IdSourceBrowse || id == IdDestinationBrowse || id == IdStartButton ||
               id == IdPauseButton || id == IdCancelButton || id == IdSettingsButton ||
               id == IdSourceAddFiles || id == IdClearSelection || id == IdLogFindNext ||
               id == IdSourcePaste || id == IdAddTaskButton || id == IdRecentDestinations ||
               id == IdClearFinished;
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
        if (draw.CtlType == ODT_LISTBOX && id == IdSourcesList) {
            draw_source_item(draw);
            return true;
        }
        if (draw.CtlType == ODT_LISTBOX && id == IdTasksList) {
            draw_task_item(draw);
            return true;
        }
        if (draw.CtlType == ODT_BUTTON && id == IdOptionsToggle) {
            draw_options_toggle(draw);
            return true;
        }
        if (draw.CtlType == ODT_BUTTON && (id == IdModeCopyTab || id == IdModeScanTab)) {
            draw_mode_tab(draw, (id == IdModeScanTab) == is_scan_mode());
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
        if (bar_show_percentage_) {
            wchar_t text[8];
            swprintf(text, 8, L"%d%%", (clamped + 5) / 10);
            HGDIOBJ old_font = SelectObject(draw.hDC, font_);
            SetBkMode(draw.hDC, TRANSPARENT);
            // A percentage may straddle the fill edge; paint each half with
            // contrast against the color actually underneath it.
            for (bool over_fill : {false, true}) {
                const int saved_dc = SaveDC(draw.hDC);
                if (saved_dc == 0) continue;
                if (over_fill) IntersectClipRect(draw.hDC, rect.left, rect.top,
                                                  rect.left + filled, rect.bottom);
                else IntersectClipRect(draw.hDC, rect.left + filled, rect.top,
                                       rect.right, rect.bottom);
                SetTextColor(draw.hDC, ui::readable_text_color(over_fill ? theme_.accent
                                                                         : theme_.panel));
                DrawTextW(draw.hDC, text, -1, &rect, DT_CENTER | DT_VCENTER | DT_SINGLELINE);
                RestoreDC(draw.hDC, saved_dc);
            }
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
        bar_show_percentage_ = settings_.get_bool("ProgressBarShowPercentage", true);
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
    // The task builder stays editable while a job runs, so more tasks can be
    // queued. Only an elevation handoff in progress locks Start/Add, because
    // this window is about to close.
    void set_inputs_enabled(bool /*enabled*/) {
        update_action_buttons();
        sync_bad_range_controls();
    }

    void update_action_buttons() {
        const BOOL can_queue = outgoing_handoff_.has_value() ? FALSE : TRUE;
        if (start_button_ != nullptr) EnableWindow(start_button_, can_queue);
        if (add_task_button_ != nullptr) EnableWindow(add_task_button_, can_queue);
        if (start_button_ != nullptr) {
            const bool busy = supervisor_.is_job_running() || starting_;
            SetWindowTextW(start_button_, busy ? L"Queue \x2192" : L"Start \x2192");
            InvalidateRect(start_button_, nullptr, FALSE);
        }
    }

    // ---- Sources panel -----------------------------------------------------

    void draw_source_item(const DRAWITEMSTRUCT& draw) {
        const bool selected = (draw.itemState & ODS_SELECTED) != 0;
        ui::themedraw::fill_rect(draw.hDC, draw.rcItem, selected ? theme_.selection : theme_.edit);
        if (draw.itemID == static_cast<UINT>(-1) || draw.itemID >= sources_.size()) return;
        const ui::SourceItem& item = sources_[draw.itemID];
        const SourceRowInfo& info = source_rows_[draw.itemID];
        const int dpi = effective_dpi();
        auto scale = [dpi](int value) { return MulDiv(value, dpi, 96); };
        RECT rect = draw.rcItem;
        rect.left += scale(6);
        rect.right -= scale(6);

        const int icon_size = GetSystemMetricsForDpi(SM_CXSMICON, static_cast<UINT>(dpi));
        if (system_small_icons_ != nullptr && info.icon >= 0) {
            ImageList_DrawEx(system_small_icons_, info.icon, draw.hDC, rect.left,
                             rect.top + (rect.bottom - rect.top - icon_size) / 2, icon_size,
                             icon_size, CLR_NONE, CLR_NONE, ILD_TRANSPARENT);
        }
        rect.left += icon_size + scale(8);

        HGDIOBJ old_font = SelectObject(draw.hDC, font_);
        SetBkMode(draw.hDC, TRANSPARENT);
        const COLORREF text = selected ? theme_.selection_text : theme_.text;
        const COLORREF muted = selected ? theme_.selection_text : theme_.muted_text;

        // Right: size or item count.
        std::wstring size_text;
        if (info.bytes < 0) {
            size_text = L"counting\x2026";
        } else {
            size_text = storage::fsutil::utf8_to_wide(format_bytes_short(info.bytes));
            if (item.is_directory) {
                size_text = std::to_wstring(info.files) + (info.files == 1 ? L" file, " : L" files, ") +
                            size_text;
            }
            if (info.size_incomplete) size_text += L"+";
        }
        RECT size_rect = rect;
        SetTextColor(draw.hDC, muted);
        DrawTextW(draw.hDC, size_text.c_str(), -1, &size_rect,
                  DT_RIGHT | DT_VCENTER | DT_SINGLELINE | DT_NOPREFIX);
        SIZE size_extent{};
        GetTextExtentPoint32W(draw.hDC, size_text.c_str(), static_cast<int>(size_text.size()),
                              &size_extent);
        rect.right -= size_extent.cx + scale(12);

        // Name, then the containing folder (muted, path-ellipsized).
        std::wstring name = storage::fsutil::get_file_name(item.path);
        if (name.empty()) name = item.path;
        if (item.contents_only) name += L"\\*  (contents)";
        SetTextColor(draw.hDC, text);
        RECT name_rect = rect;
        DrawTextW(draw.hDC, name.c_str(), -1, &name_rect,
                  DT_LEFT | DT_VCENTER | DT_SINGLELINE | DT_END_ELLIPSIS | DT_NOPREFIX);
        SIZE name_extent{};
        GetTextExtentPoint32W(draw.hDC, name.c_str(), static_cast<int>(name.size()), &name_extent);
        RECT folder_rect = rect;
        folder_rect.left = std::min<LONG>(rect.right, rect.left + name_extent.cx + scale(10));
        std::wstring folder = item.contents_only ? item.path
                                                 : storage::fsutil::get_directory_name(item.path);
        SetTextColor(draw.hDC, muted);
        DrawTextW(draw.hDC, folder.c_str(), -1, &folder_rect,
                  DT_LEFT | DT_VCENTER | DT_SINGLELINE | DT_PATH_ELLIPSIS | DT_NOPREFIX);
        SelectObject(draw.hDC, old_font);
        if ((draw.itemState & ODS_FOCUS) != 0) DrawFocusRect(draw.hDC, &draw.rcItem);
    }

    static ui::SourceItem make_source_item(const std::wstring& raw) {
        ui::SourceItem item;
        std::wstring full = storage::fsutil::get_full_path(storage::fsutil::trim(raw));
        while (full.size() > 3 && (full.back() == L'\\' || full.back() == L'/')) full.pop_back();
        item.path = full;
        DWORD attributes = GetFileAttributesW(storage::fsutil::long_path(full).c_str());
        item.is_directory = attributes != INVALID_FILE_ATTRIBUTES &&
                            (attributes & FILE_ATTRIBUTE_DIRECTORY) != 0;
        // A drive root has no parent to be copied "as itself" into.
        if (item.is_directory && full.size() <= 3) item.contents_only = true;
        return item;
    }

    // Adds existing paths that are not already listed. Returns how many were added.
    int add_sources(const std::vector<std::wstring>& paths, bool contents_only = false) {
        int added = 0;
        for (const auto& raw : paths) {
            if (storage::fsutil::trim(raw).empty()) continue;
            ui::SourceItem item = make_source_item(raw);
            if (item.path.empty() ||
                GetFileAttributesW(storage::fsutil::long_path(item.path).c_str()) ==
                    INVALID_FILE_ATTRIBUTES) {
                append_log("Not added (not found): " + storage::fsutil::wide_to_utf8(raw));
                continue;
            }
            if (contents_only && item.is_directory) item.contents_only = true;
            const bool duplicate = std::any_of(sources_.begin(), sources_.end(),
                                               [&](const ui::SourceItem& existing) {
                                                   return _wcsicmp(existing.path.c_str(),
                                                                   item.path.c_str()) == 0;
                                               });
            if (duplicate) continue;
            sources_.push_back(std::move(item));
            ++added;
        }
        if (added > 0) sources_changed();
        return added;
    }

    void set_sources(std::vector<ui::SourceItem> items) {
        sources_ = std::move(items);
        sources_changed();
    }

    // Rebuilds row metadata (icons from the extension only, so slow media is
    // never touched on this thread) and asks the sizer for fresh totals.
    void sources_changed() {
        source_rows_.assign(sources_.size(), SourceRowInfo{});
        std::vector<std::wstring> paths;
        paths.reserve(sources_.size());
        for (std::size_t i = 0; i < sources_.size(); ++i) {
            SHFILEINFOW info{};
            const DWORD attributes =
                sources_[i].is_directory ? FILE_ATTRIBUTE_DIRECTORY : FILE_ATTRIBUTE_NORMAL;
            if (SHGetFileInfoW(sources_[i].path.c_str(), attributes, &info, sizeof(info),
                               SHGFI_SYSICONINDEX | SHGFI_SMALLICON | SHGFI_USEFILEATTRIBUTES) != 0) {
                source_rows_[i].icon = info.iIcon;
            }
            paths.push_back(sources_[i].path);
        }
        {
            std::lock_guard<std::mutex> guard(sizer_lock_);
            sizer_request_ = std::move(paths);
            sizer_generation_.fetch_add(1);
        }
        sizer_ready_.notify_one();
        SendMessageW(sources_list_, LB_SETCOUNT, sources_.size(), 0);
        InvalidateRect(sources_list_, nullptr, TRUE);
        update_sources_footer();
    }

    void update_sources_footer() {
        if (sources_footer_ == nullptr) return;
        if (sources_.empty()) {
            SetWindowTextW(sources_footer_,
                           L"Add files or folders: + Files, + Folder, Paste, or drag them here.");
            return;
        }
        int folders = 0;
        std::int64_t bytes = 0;
        bool counting = false;
        for (std::size_t i = 0; i < sources_.size(); ++i) {
            if (sources_[i].is_directory) ++folders;
            if (source_rows_[i].bytes < 0) counting = true;
            else bytes += source_rows_[i].bytes;
        }
        const int files = static_cast<int>(sources_.size()) - folders;
        std::string text = std::to_string(sources_.size()) +
                           (sources_.size() == 1 ? " item" : " items");
        text += " \xC2\xB7 " + std::to_string(folders) + (folders == 1 ? " folder, " : " folders, ") +
                std::to_string(files) + (files == 1 ? " file" : " files");
        text += " \xC2\xB7 " + (counting ? std::string("counting\xE2\x80\xA6")
                                         : format_bytes_short(bytes));
        const auto groups = ui::group_sources(sources_);
        if (groups.size() > 1) {
            text += " \xC2\xB7 " + std::to_string(groups.size()) + " tasks (one per source folder)";
        }
        SetWindowTextW(sources_footer_, storage::fsutil::utf8_to_wide(text).c_str());
    }

    void on_source_size(const SourceSizeResult& result) {
        if (result.generation != sizer_generation_.load()) return;
        for (std::size_t i = 0; i < sources_.size(); ++i) {
            if (_wcsicmp(sources_[i].path.c_str(), result.path.c_str()) != 0) continue;
            source_rows_[i].bytes = result.bytes;
            source_rows_[i].files = result.files;
            source_rows_[i].size_incomplete = result.incomplete;
            RECT row{};
            if (SendMessageW(sources_list_, LB_GETITEMRECT, i, reinterpret_cast<LPARAM>(&row)) !=
                LB_ERR) {
                InvalidateRect(sources_list_, &row, FALSE);
            }
        }
        update_sources_footer();
    }

    // Sizes each requested path on this background thread. A newer request
    // (generation bump) abandons the current walk at the next entry.
    void sizer_loop() {
        for (;;) {
            std::vector<std::wstring> request;
            std::uint64_t generation = 0;
            {
                std::unique_lock<std::mutex> guard(sizer_lock_);
                sizer_ready_.wait(guard, [this] { return sizer_stopping_ || !sizer_request_.empty(); });
                if (sizer_stopping_) return;
                request = std::move(sizer_request_);
                sizer_request_.clear();
                generation = sizer_generation_.load();
            }
            for (const auto& path : request) {
                if (generation != sizer_generation_.load() || sizer_stopping_) break;
                auto* result = new SourceSizeResult;
                result->generation = generation;
                result->path = path;
                size_path(path, generation, *result);
                if (generation != sizer_generation_.load() ||
                    !PostMessageW(hwnd_, WM_APP_SOURCE_SIZE, 0, reinterpret_cast<LPARAM>(result))) {
                    delete result;
                }
            }
        }
    }

    void size_path(const std::wstring& path, std::uint64_t generation, SourceSizeResult& result) {
        WIN32_FILE_ATTRIBUTE_DATA data{};
        if (!GetFileAttributesExW(storage::fsutil::long_path(path).c_str(), GetFileExInfoStandard,
                                  &data)) {
            result.incomplete = true;
            return;
        }
        if ((data.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) == 0) {
            result.bytes = (static_cast<std::int64_t>(data.nFileSizeHigh) << 32) | data.nFileSizeLow;
            result.files = 1;
            return;
        }
        std::vector<std::wstring> pending{path};
        while (!pending.empty()) {
            if (generation != sizer_generation_.load() || sizer_stopping_) return;
            std::wstring directory = std::move(pending.back());
            pending.pop_back();
            WIN32_FIND_DATAW find{};
            HANDLE handle = FindFirstFileExW(
                storage::fsutil::long_path(directory + L"\\*").c_str(), FindExInfoBasic, &find,
                FindExSearchNameMatch, nullptr, FIND_FIRST_EX_LARGE_FETCH);
            if (handle == INVALID_HANDLE_VALUE) {
                result.incomplete = true;
                continue;
            }
            do {
                const std::wstring name = find.cFileName;
                if (name == L"." || name == L"..") continue;
                if ((find.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) != 0) {
                    // Do not follow junctions/symlinks; the copy does not either by default.
                    if ((find.dwFileAttributes & FILE_ATTRIBUTE_REPARSE_POINT) == 0) {
                        pending.push_back(directory + L"\\" + name);
                    }
                    continue;
                }
                result.bytes += (static_cast<std::int64_t>(find.nFileSizeHigh) << 32) | find.nFileSizeLow;
                result.files += 1;
            } while (FindNextFileW(handle, &find));
            FindClose(handle);
        }
    }

    void stop_sizer() {
        {
            std::lock_guard<std::mutex> guard(sizer_lock_);
            sizer_stopping_ = true;
        }
        sizer_generation_.fetch_add(1);
        sizer_ready_.notify_one();
        if (sizer_thread_.joinable()) sizer_thread_.join();
    }

    std::vector<int> selected_list_indices(HWND list) const {
        const int count = static_cast<int>(SendMessageW(list, LB_GETSELCOUNT, 0, 0));
        std::vector<int> indices;
        if (count > 0) {
            indices.resize(static_cast<std::size_t>(count));
            SendMessageW(list, LB_GETSELITEMS, count, reinterpret_cast<LPARAM>(indices.data()));
        } else if (count == LB_ERR) { // single-selection list
            const int current = static_cast<int>(SendMessageW(list, LB_GETCURSEL, 0, 0));
            if (current >= 0) indices.push_back(current);
        }
        return indices;
    }

    void remove_selected_sources() {
        auto indices = selected_list_indices(sources_list_);
        if (indices.empty()) return;
        std::sort(indices.rbegin(), indices.rend());
        for (int index : indices) {
            if (index >= 0 && static_cast<std::size_t>(index) < sources_.size()) {
                sources_.erase(sources_.begin() + index);
            }
        }
        SendMessageW(sources_list_, LB_SETSEL, FALSE, -1);
        sources_changed();
    }

    void paste_sources() {
        std::vector<std::wstring> paths;
        if (!OpenClipboard(hwnd_)) return;
        if (HANDLE drop = GetClipboardData(CF_HDROP)) {
            HDROP files = static_cast<HDROP>(drop);
            UINT count = DragQueryFileW(files, 0xFFFFFFFF, nullptr, 0);
            for (UINT i = 0; i < count; ++i) {
                UINT length = DragQueryFileW(files, i, nullptr, 0);
                std::wstring path(length + 1, L'\0');
                DragQueryFileW(files, i, path.data(), length + 1);
                path.resize(length);
                paths.push_back(std::move(path));
            }
        } else if (HANDLE text = GetClipboardData(CF_UNICODETEXT)) {
            if (const auto* chars = static_cast<const wchar_t*>(GlobalLock(text))) {
                std::wstring all(chars);
                GlobalUnlock(text);
                std::size_t start = 0;
                while (start <= all.size()) {
                    std::size_t end = all.find_first_of(L"\r\n", start);
                    std::wstring line = storage::fsutil::trim(
                        all.substr(start, end == std::wstring::npos ? std::wstring::npos : end - start));
                    if (line.size() >= 2 && line.front() == L'"' && line.back() == L'"') {
                        line = line.substr(1, line.size() - 2);
                    }
                    if (!line.empty()) paths.push_back(std::move(line));
                    if (end == std::wstring::npos) break;
                    start = end + 1;
                }
            }
        }
        CloseClipboard();
        if (paths.empty()) {
            info_box(L"The clipboard has no file paths. Copy files in Explorer, or copy one path per line.");
            return;
        }
        const int added = add_sources(paths);
        append_log("Pasted " + std::to_string(added) + " source item(s).");
    }

    void show_sources_context_menu(LPARAM screen_point) {
        POINT point{GET_X_LPARAM(screen_point), GET_Y_LPARAM(screen_point)};
        if (point.x == -1 && point.y == -1) GetCursorPos(&point);
        const auto indices = selected_list_indices(sources_list_);
        const bool has_selection = !indices.empty();
        const bool single_folder = indices.size() == 1 &&
                                   static_cast<std::size_t>(indices[0]) < sources_.size() &&
                                   sources_[static_cast<std::size_t>(indices[0])].is_directory;
        HMENU menu = CreatePopupMenu();
        AppendMenuW(menu, MF_STRING | (has_selection ? 0 : MF_GRAYED), IdSourceRemove,
                    L"&Remove\tDel");
        if (single_folder) {
            const auto& item = sources_[static_cast<std::size_t>(indices[0])];
            AppendMenuW(menu, MF_STRING | (item.contents_only ? MF_CHECKED : 0) |
                                  (item.path.size() <= 3 ? MF_GRAYED : 0),
                        IdSourceContentsOnly, L"Copy folder &contents only");
        }
        AppendMenuW(menu, MF_SEPARATOR, 0, nullptr);
        AppendMenuW(menu, MF_STRING | (indices.size() == 1 ? 0 : MF_GRAYED), IdSourceOpenLocation,
                    L"&Open location");
        AppendMenuW(menu, MF_STRING | (has_selection ? 0 : MF_GRAYED), IdSourceCopyPath,
                    L"Copy &path");
        AppendMenuW(menu, MF_SEPARATOR, 0, nullptr);
        AppendMenuW(menu, MF_STRING, IdSourcePaste, L"P&aste paths\tCtrl+V");
        AppendMenuW(menu, MF_STRING | (sources_.empty() ? MF_GRAYED : 0), IdClearSelection,
                    L"C&lear all");
        TrackPopupMenu(menu, TPM_RIGHTBUTTON, point.x, point.y, 0, hwnd_, nullptr);
        DestroyMenu(menu);
    }

    void copy_text_to_clipboard(const std::wstring& text) {
        if (!OpenClipboard(hwnd_)) return;
        EmptyClipboard();
        const std::size_t bytes = (text.size() + 1) * sizeof(wchar_t);
        if (HGLOBAL memory = GlobalAlloc(GMEM_MOVEABLE, bytes)) {
            if (void* target = GlobalLock(memory)) {
                memcpy(target, text.c_str(), bytes);
                GlobalUnlock(memory);
                if (SetClipboardData(CF_UNICODETEXT, memory) == nullptr) GlobalFree(memory);
            } else {
                GlobalFree(memory);
            }
        }
        CloseClipboard();
    }

    void on_source_command(int id) {
        const auto indices = selected_list_indices(sources_list_);
        switch (id) {
            case IdSourceRemove:
                remove_selected_sources();
                break;
            case IdSourceContentsOnly:
                if (indices.size() == 1 && static_cast<std::size_t>(indices[0]) < sources_.size()) {
                    auto& item = sources_[static_cast<std::size_t>(indices[0])];
                    if (item.is_directory && item.path.size() > 3) {
                        item.contents_only = !item.contents_only;
                        sources_changed();
                    }
                }
                break;
            case IdSourceOpenLocation:
                if (indices.size() == 1 && static_cast<std::size_t>(indices[0]) < sources_.size()) {
                    const auto& item = sources_[static_cast<std::size_t>(indices[0])];
                    ui::elevation::open_as_user(hwnd_, L"explorer.exe",
                                                L"/select,\"" + item.path + L"\"");
                }
                break;
            case IdSourceCopyPath: {
                std::wstring text;
                for (int index : indices) {
                    if (static_cast<std::size_t>(index) >= sources_.size()) continue;
                    if (!text.empty()) text += L"\r\n";
                    text += sources_[static_cast<std::size_t>(index)].path;
                }
                if (!text.empty()) copy_text_to_clipboard(text);
                break;
            }
            default:
                break;
        }
    }

    // ---- Mode switch --------------------------------------------------------

    // One half of the Copy / Assess segmented control: the active half is
    // filled with the accent, the other reads as a flat button.
    void draw_mode_tab(const DRAWITEMSTRUCT& draw, bool active) {
        const bool hot = (draw.itemState & ODS_HOTLIGHT) != 0;
        const bool focused = (draw.itemState & ODS_FOCUS) != 0;
        const COLORREF fill = active ? theme_.accent : hot ? theme_.button_hot : theme_.button;
        ui::themedraw::fill_rect(draw.hDC, draw.rcItem, fill);
        ui::themedraw::frame_rect(draw.hDC, draw.rcItem, active ? theme_.accent : theme_.border);
        wchar_t text[64]{};
        GetWindowTextW(draw.hwndItem, text, 64);
        HGDIOBJ old_font = SelectObject(draw.hDC, active ? header_font_ : font_);
        SetBkMode(draw.hDC, TRANSPARENT);
        SetTextColor(draw.hDC, active ? ui::readable_text_color(theme_.accent) : theme_.text);
        RECT rect = draw.rcItem;
        DrawTextW(draw.hDC, text, -1, &rect,
                  DT_CENTER | DT_VCENTER | DT_SINGLELINE | DT_END_ELLIPSIS | DT_NOPREFIX);
        SelectObject(draw.hDC, old_font);
        if (focused) {
            RECT focus = draw.rcItem;
            InflateRect(&focus, -MulDiv(3, effective_dpi(), 96), -MulDiv(3, effective_dpi(), 96));
            DrawFocusRect(draw.hDC, &focus);
        }
    }

    void select_mode(bool scan) {
        if (scan == is_scan_mode()) return;
        SendMessageW(mode_combo_, CB_SETCURSEL, scan ? 1 : 0, 0);
        sync_mode_ui(true);
    }

    // ---- Options section ----------------------------------------------------

    void draw_options_toggle(const DRAWITEMSTRUCT& draw) {
        const bool focused = (draw.itemState & ODS_FOCUS) != 0;
        ui::themedraw::fill_rect(draw.hDC, draw.rcItem, theme_.window);
        RECT rect = draw.rcItem;
        rect.left += MulDiv(2, effective_dpi(), 96);
        HGDIOBJ old_font = SelectObject(draw.hDC, font_);
        SetBkMode(draw.hDC, TRANSPARENT);
        std::wstring heading = options_expanded_ ? L"\x25BE  Options" : L"\x25B8  Options";
        SetTextColor(draw.hDC, theme_.accent);
        DrawTextW(draw.hDC, heading.c_str(), -1, &rect,
                  DT_LEFT | DT_VCENTER | DT_SINGLELINE | DT_NOPREFIX);
        if (!options_expanded_) {
            SIZE extent{};
            GetTextExtentPoint32W(draw.hDC, heading.c_str(), static_cast<int>(heading.size()), &extent);
            RECT summary_rect = rect;
            summary_rect.left += extent.cx + MulDiv(12, effective_dpi(), 96);
            std::wstring summary;
            try {
                summary = storage::fsutil::utf8_to_wide(ui::summarize_options(collect_options()));
            } catch (const std::exception&) {
                summary = L"(check the option values)";
            }
            SetTextColor(draw.hDC, theme_.muted_text);
            DrawTextW(draw.hDC, summary.c_str(), -1, &summary_rect,
                      DT_LEFT | DT_VCENTER | DT_SINGLELINE | DT_END_ELLIPSIS | DT_NOPREFIX);
        }
        SelectObject(draw.hDC, old_font);
        if (focused) DrawFocusRect(draw.hDC, &draw.rcItem);
    }

    void set_options_expanded(bool expanded) {
        options_expanded_ = expanded;
        settings_.set_bool("UiOptionsExpanded", expanded);
        layout();
        ensure_window_fits_layout();
        InvalidateRect(options_toggle_, nullptr, TRUE);
    }

    // ---- Recent destinations ------------------------------------------------

    std::vector<std::wstring> recent_destinations() {
        std::vector<std::wstring> result;
        const std::wstring all =
            storage::fsutil::utf8_to_wide(settings_.get_string("RecentDestinations", ""));
        std::size_t start = 0;
        while (start < all.size()) {
            std::size_t end = all.find(L'|', start);
            std::wstring entry = all.substr(start, end == std::wstring::npos ? std::wstring::npos
                                                                              : end - start);
            if (!entry.empty()) result.push_back(std::move(entry));
            if (end == std::wstring::npos) break;
            start = end + 1;
        }
        return result;
    }

    void remember_destination(const std::string& destination) {
        if (ui_is_blank(destination)) return;
        const std::wstring wide = storage::fsutil::utf8_to_wide(destination);
        auto list = recent_destinations();
        list.erase(std::remove_if(list.begin(), list.end(),
                                  [&](const std::wstring& entry) {
                                      return _wcsicmp(entry.c_str(), wide.c_str()) == 0;
                                  }),
                   list.end());
        list.insert(list.begin(), wide);
        if (list.size() > static_cast<std::size_t>(MaxRecentDestinations)) {
            list.resize(MaxRecentDestinations);
        }
        std::wstring joined;
        for (const auto& entry : list) {
            if (!joined.empty()) joined += L"|";
            joined += entry;
        }
        settings_.set_string("RecentDestinations", storage::fsutil::wide_to_utf8(joined));
    }

    void show_recent_destinations() {
        const auto list = recent_destinations();
        HMENU menu = CreatePopupMenu();
        if (list.empty()) {
            AppendMenuW(menu, MF_STRING | MF_GRAYED, 0, L"No recent destinations");
        }
        for (std::size_t i = 0; i < list.size(); ++i) {
            AppendMenuW(menu, MF_STRING, static_cast<UINT_PTR>(IdRecentDestinationBase + i),
                        list[i].c_str());
        }
        RECT button{};
        GetWindowRect(recent_destinations_button_, &button);
        const int chosen = static_cast<int>(TrackPopupMenu(
            menu, TPM_RIGHTALIGN | TPM_TOPALIGN | TPM_RETURNCMD | TPM_NONOTIFY, button.right,
            button.bottom, 0, hwnd_, nullptr));
        DestroyMenu(menu);
        const int index = chosen - IdRecentDestinationBase;
        if (chosen != 0 && index >= 0 && static_cast<std::size_t>(index) < list.size()) {
            SetWindowTextW(destination_edit_, list[static_cast<std::size_t>(index)].c_str());
        }
    }

    // ---- Tasks panel ---------------------------------------------------------

    COLORREF task_state_color(ui::TaskState state) const {
        COLORREF color;
        switch (state) {
            case ui::TaskState::Running: color = RGB(20, 140, 60); break;
            case ui::TaskState::Starting:
            case ui::TaskState::WaitingForElevation: color = RGB(40, 120, 190); break;
            case ui::TaskState::Paused: color = RGB(140, 80, 180); break;
            case ui::TaskState::Completed: color = RGB(30, 100, 180); break;
            case ui::TaskState::CompletedWithRecovery:
            case ui::TaskState::Incomplete: color = RGB(210, 125, 30); break;
            case ui::TaskState::Failed: color = RGB(200, 40, 40); break;
            case ui::TaskState::Cancelled: color = RGB(160, 120, 20); break;
            default: return theme_.muted_text;
        }
        if (theme_.dark) color = ui::blend_color(color, RGB(255, 255, 255), 45);
        return color;
    }

    void draw_task_item(const DRAWITEMSTRUCT& draw) {
        const bool selected = (draw.itemState & ODS_SELECTED) != 0;
        ui::themedraw::fill_rect(draw.hDC, draw.rcItem, selected ? theme_.selection : theme_.edit);
        if (draw.itemID == static_cast<UINT>(-1) || draw.itemID >= tasks_.size()) return;
        const ui::SessionTask& task = tasks_.at(draw.itemID);
        const int dpi = effective_dpi();
        auto scale = [dpi](int value) { return MulDiv(value, dpi, 96); };

        RECT rect = draw.rcItem;
        rect.left += scale(10);
        rect.right -= scale(8);
        const int line = (rect.bottom - rect.top - scale(8)) / 2;
        RECT title_rect{rect.left + scale(14), rect.top + scale(3), rect.right,
                        rect.top + scale(3) + line};
        RECT status_rect{rect.left + scale(14), title_rect.bottom, rect.right,
                         title_rect.bottom + line};

        // State dot.
        const COLORREF state_color = task_state_color(task.state);
        const int dot = scale(8);
        RECT dot_rect{rect.left, title_rect.top + (line - dot) / 2, rect.left + dot,
                      title_rect.top + (line - dot) / 2 + dot};
        ui::themedraw::fill_rect(draw.hDC, dot_rect, state_color);

        HGDIOBJ old_font = SelectObject(draw.hDC, font_);
        SetBkMode(draw.hDC, TRANSPARENT);
        SetTextColor(draw.hDC, selected ? theme_.selection_text : theme_.text);
        const std::wstring title = storage::fsutil::utf8_to_wide(ui::task_title(task.options));
        DrawTextW(draw.hDC, title.c_str(), -1, &title_rect,
                  DT_LEFT | DT_VCENTER | DT_SINGLELINE | DT_END_ELLIPSIS | DT_NOPREFIX);

        std::string status = ui::to_display(task.state);
        if (task.state == ui::TaskState::Running || task.state == ui::TaskState::Paused) {
            status += " \xC2\xB7 " + std::to_string(static_cast<int>(task.progress * 100.0 + 0.5)) + "%";
            if (task.bytes_per_second > 0 && task.state == ui::TaskState::Running) {
                status += " \xC2\xB7 " + format_bytes_short(
                                              static_cast<std::int64_t>(task.bytes_per_second)) + "/s";
            }
            if (task.eta_seconds >= 0 && task.state == ui::TaskState::Running) {
                status += " \xC2\xB7 " + format_eta(static_cast<double>(task.eta_seconds)) + " left";
            }
        } else if (!task.outcome.empty()) {
            status += " \xC2\xB7 " + task.outcome;
        }
        SetTextColor(draw.hDC, selected ? theme_.selection_text : state_color);
        const std::wstring status_wide = storage::fsutil::utf8_to_wide(status);
        DrawTextW(draw.hDC, status_wide.c_str(), -1, &status_rect,
                  DT_LEFT | DT_VCENTER | DT_SINGLELINE | DT_END_ELLIPSIS | DT_NOPREFIX);
        SelectObject(draw.hDC, old_font);

        // Thin inline progress bar along the bottom edge for an active task.
        if (ui::is_active(task.state)) {
            RECT track{rect.left, draw.rcItem.bottom - scale(5), rect.right, draw.rcItem.bottom - scale(2)};
            ui::themedraw::fill_rect(draw.hDC, track, theme_.panel);
            RECT fill = track;
            fill.right = track.left + static_cast<LONG>((track.right - track.left) *
                                                        std::clamp(task.progress, 0.0, 1.0));
            if (fill.right > fill.left) ui::themedraw::fill_rect(draw.hDC, fill, theme_.accent);
        }
        RECT divider{draw.rcItem.left, draw.rcItem.bottom - 1, draw.rcItem.right, draw.rcItem.bottom};
        ui::themedraw::fill_rect(draw.hDC, divider, ui::blend_color(theme_.edit, theme_.border, 60));
        if ((draw.itemState & ODS_FOCUS) != 0) DrawFocusRect(draw.hDC, &draw.rcItem);
    }

    void refresh_tasks_view() {
        const int selected = static_cast<int>(SendMessageW(tasks_list_, LB_GETCURSEL, 0, 0));
        SendMessageW(tasks_list_, LB_SETCOUNT, tasks_.size(), 0);
        if (selected >= 0 && static_cast<std::size_t>(selected) < tasks_.size()) {
            SendMessageW(tasks_list_, LB_SETCURSEL, selected, 0);
        }
        InvalidateRect(tasks_list_, nullptr, TRUE);
        const auto counts = tasks_.counts();
        std::wstring label = L"Tasks";
        if (tasks_.size() > 0) {
            label += L"  \x00B7  " + std::to_wstring(counts.active) + L" running, " +
                     std::to_wstring(counts.queued) + L" queued, " +
                     std::to_wstring(counts.finished) + L" finished";
        }
        SetWindowTextW(tasks_label_, label.c_str());
        EnableWindow(clear_finished_button_, counts.finished > 0 ? TRUE : FALSE);
    }

    void invalidate_task_row(std::uint64_t id) {
        auto index = tasks_.index_of(id);
        if (!index.has_value()) return;
        RECT row{};
        if (SendMessageW(tasks_list_, LB_GETITEMRECT, *index, reinterpret_cast<LPARAM>(&row)) != LB_ERR) {
            InvalidateRect(tasks_list_, &row, FALSE);
        }
    }

    void persist_pending_tasks() {
        if (pending_task_writer_ == nullptr) return;
        pending_task_writer_->post(ui::pending_tasks_to_json(tasks_.queued(), tasks_running_));
    }

    void tasks_changed() {
        refresh_tasks_view();
        persist_pending_tasks();
    }

    // Turns the builder (sources + destination + options) into tasks: one per
    // source folder group. Returns how many tasks were added.
    int add_tasks_from_builder() {
        if (sources_.empty()) {
            info_box(L"Add files or folders to copy first: + Files, + Folder, Paste, or drag them "
                     L"onto the list.");
            return 0;
        }
        models::CopyJobOptions base;
        try {
            base = collect_options();
        } catch (const std::exception& ex) {
            warn_box(storage::fsutil::utf8_to_wide(ex.what()), L"Invalid run setting");
            return 0;
        }
        const bool scan = base.OperationMode == models::JobOperationMode::ScanOnly;
        if (!scan && ui_is_blank(base.DestinationRoot)) {
            info_box(L"Choose a destination folder first.");
            SetFocus(destination_edit_);
            return 0;
        }
        const auto groups = ui::group_sources(sources_);
        if (groups.empty()) return 0;
        if (!scan) {
            const auto collisions = ui::find_destination_collisions(groups);
            if (!collisions.empty()) {
                std::wstring names;
                for (const auto& name : collisions) names += L"\n  \x2022 " + name;
                if (!confirm_box(L"These names come from more than one source folder and would be "
                                 L"written to the same place in the destination:" + names +
                                     L"\n\nThe later task would overwrite or merge into the earlier "
                                     L"one, as the overwrite policy decides. Add the tasks anyway?",
                                 L"XactCopy - Name collision", ui::MessageIcon::Warning)) {
                    return 0;
                }
            }
            remember_destination(base.DestinationRoot);
        }
        for (const auto& group : groups) {
            ui::SessionTask task;
            task.options = base;
            task.options.SourceRoot = storage::fsutil::wide_to_utf8(group.root);
            task.options.SelectedRelativePaths = group.relative_paths;
            tasks_.add(std::move(task));
        }
        append_log("Added " + std::to_string(groups.size()) +
                   (groups.size() == 1 ? " task." : " tasks (one per source folder)."));
        tasks_changed();
        return static_cast<int>(groups.size());
    }

    // Starts the next queued task when nothing else is running.
    void run_next_task() {
        if (supervisor_.is_job_running() || starting_ || outgoing_handoff_.has_value()) return;
        if (running_task_id_ != 0) return;
        ui::SessionTask* task = tasks_.next_queued();
        if (task == nullptr) {
            if (tasks_running_) append_log("All queued tasks have been started.");
            tasks_running_ = false;
            persist_pending_tasks();
            return;
        }
        running_task_id_ = task->id;
        task->state = ui::TaskState::Starting;
        const std::uint64_t id = task->id;
        models::CopyJobOptions options = task->options;
        const bool from_recovery = task->from_recovery;
        const std::string managed_run_id = task->managed_run_id;
        const std::string display_name = task->display_name;
        tasks_changed();
        start_job_with(options, from_recovery, managed_run_id, display_name);
        settle_unstarted_task(id, "Not started");
    }

    // start_job_with can decline (validation, a risk prompt answered No,
    // elevation declined without fallback). Leave the task visibly not run and
    // stop advancing instead of looping over the queue.
    void settle_unstarted_task(std::uint64_t id, const std::string& outcome) {
        if (running_task_id_ != id) return;
        ui::SessionTask* task = tasks_.find(id);
        if (task == nullptr) {
            running_task_id_ = 0;
            return;
        }
        if (outgoing_handoff_.has_value()) {
            task->state = ui::TaskState::WaitingForElevation;
        } else if (!starting_ && !supervisor_.is_job_running()) {
            task->state = ui::TaskState::Cancelled;
            task->outcome = outcome;
            running_task_id_ = 0;
            tasks_running_ = false;
        }
        tasks_changed();
    }

    // Every start path funnels through start_job_with; a job started outside
    // the task list (Job Manager, crash recovery, an elevation handoff) gets a
    // task row too, so the panel always shows what is running.
    void bind_started_job_to_task() {
        ui::SessionTask* task = running_task_id_ != 0 ? tasks_.find(running_task_id_) : nullptr;
        if (task == nullptr) {
            ui::SessionTask external;
            external.options = pending_start_options_;
            external.display_name = pending_start_job_name_;
            external.managed_run_id = active_managed_run_id_;
            running_task_id_ = tasks_.add(std::move(external));
            task = tasks_.find(running_task_id_);
        }
        task->state = paused_ ? ui::TaskState::Paused : ui::TaskState::Running;
        task->progress = 0.0;
        set_detail_heading(L"Running: " +
                           storage::fsutil::utf8_to_wide(ui::task_title(task->options)));
        tasks_changed();
    }

    void set_detail_heading(const std::wstring& text) {
        if (detail_label_ != nullptr) SetWindowTextW(detail_label_, text.c_str());
    }

    void finish_running_task(const models::CopyJobResult& result) {
        ui::SessionTask* task = running_task_id_ != 0 ? tasks_.find(running_task_id_) : nullptr;
        // Only one job runs at a time, so a row still marked active belongs
        // to this result even if its binding was lost.
        if (task == nullptr) task = tasks_.active();
        running_task_id_ = 0;
        if (task == nullptr) {
            set_detail_heading(L"Last task");
            return;
        }
        task->state = ui::state_from_result(result);
        task->progress = task->state == ui::TaskState::Completed ? 1.0 : task->progress;
        task->bytes_per_second = 0;
        task->eta_seconds = -1;
        std::string outcome = std::to_string(result.CompletedFiles) + "/" +
                              std::to_string(result.TotalFiles) + " files";
        if (result.FailedFiles > 0) outcome += ", " + std::to_string(result.FailedFiles) + " failed";
        if (result.RecoveredFiles > 0) outcome += ", " + std::to_string(result.RecoveredFiles) + " recovered";
        if (result.SkippedFiles > 0) outcome += ", " + std::to_string(result.SkippedFiles) + " skipped";
        if (!result.ErrorMessage.empty() && task->state == ui::TaskState::Failed) {
            outcome += " \xE2\x80\x94 " + result.ErrorMessage;
        }
        task->outcome = std::move(outcome);
        set_detail_heading(L"Last task: " +
                           storage::fsutil::utf8_to_wide(ui::to_display(task->state)));
        tasks_changed();
    }

    void show_tasks_context_menu(LPARAM screen_point) {
        POINT point{GET_X_LPARAM(screen_point), GET_Y_LPARAM(screen_point)};
        if (point.x == -1 && point.y == -1) GetCursorPos(&point);
        const ui::SessionTask* task = selected_task();
        const bool queued = task != nullptr && task->state == ui::TaskState::Queued;
        const bool finished = task != nullptr && ui::is_finished(task->state);
        const bool copy = task != nullptr &&
                          task->options.OperationMode == models::JobOperationMode::Copy;
        auto flag = [](bool enabled) { return MF_STRING | (enabled ? 0u : static_cast<UINT>(MF_GRAYED)); };
        HMENU menu = CreatePopupMenu();
        AppendMenuW(menu, flag(queued), IdTaskStartNow, L"&Start now");
        AppendMenuW(menu, flag(queued), IdTaskMoveUp, L"Move &up");
        AppendMenuW(menu, flag(queued), IdTaskMoveDown, L"Move &down");
        AppendMenuW(menu, MF_SEPARATOR, 0, nullptr);
        AppendMenuW(menu, flag(finished && task->state != ui::TaskState::Completed), IdTaskRetry,
                    L"&Retry (resume from journal)");
        AppendMenuW(menu, flag(copy), IdTaskOpenDestination, L"&Open destination");
        AppendMenuW(menu, flag(task != nullptr), IdTaskEdit, L"&Edit in builder");
        AppendMenuW(menu, flag(task != nullptr), IdTaskSaveAsJob, L"Save as &job...");
        AppendMenuW(menu, MF_SEPARATOR, 0, nullptr);
        AppendMenuW(menu, flag(task != nullptr && !ui::is_active(task->state)), IdTaskRemove,
                    L"Re&move\tDel");
        TrackPopupMenu(menu, TPM_RIGHTBUTTON, point.x, point.y, 0, hwnd_, nullptr);
        DestroyMenu(menu);
    }

    ui::SessionTask* selected_task() {
        const int index = static_cast<int>(SendMessageW(tasks_list_, LB_GETCURSEL, 0, 0));
        if (index < 0 || static_cast<std::size_t>(index) >= tasks_.size()) return nullptr;
        return &tasks_.at(static_cast<std::size_t>(index));
    }

    void select_task(std::uint64_t id) {
        if (auto index = tasks_.index_of(id)) {
            SendMessageW(tasks_list_, LB_SETCURSEL, *index, 0);
        }
    }

    void on_task_command(int id) {
        ui::SessionTask* task = selected_task();
        if (task == nullptr) return;
        const std::uint64_t task_id = task->id;
        switch (id) {
            case IdTaskStartNow:
                tasks_.move_to_front(task_id);
                tasks_running_ = true;
                tasks_changed();
                select_task(task_id);
                run_next_task();
                break;
            case IdTaskMoveUp:
            case IdTaskMoveDown:
                if (tasks_.move(task_id, id == IdTaskMoveUp ? -1 : 1)) {
                    tasks_changed();
                    select_task(task_id);
                }
                break;
            case IdTaskRemove:
                if (tasks_.remove(task_id)) tasks_changed();
                break;
            case IdTaskRetry: {
                ui::SessionTask retry;
                retry.options = task->options;
                retry.options.ResumeFromJournal = true;
                retry.display_name = task->display_name;
                const std::uint64_t retry_id = tasks_.add(std::move(retry));
                tasks_changed();
                select_task(retry_id);
                if (!supervisor_.is_job_running() && !starting_) {
                    tasks_.move_to_front(retry_id);
                    tasks_running_ = true;
                    run_next_task();
                }
                break;
            }
            case IdTaskOpenDestination:
                if (!ui_is_blank(task->options.DestinationRoot)) {
                    open_folder_in_explorer(storage::fsutil::utf8_to_wide(task->options.DestinationRoot));
                }
                break;
            case IdTaskEdit:
                apply_run_options_to_ui(task->options);
                append_log("Loaded the task's settings into the builder.");
                break;
            case IdTaskSaveAsJob: {
                models::CopyJobOptions options = task->options;
                bind_media_identities(options);
                auto name = ui::TextPromptDialog::show(hwnd_, theme_, L"Save Job",
                                                       L"Name for this saved job:",
                                                       suggest_job_name(options));
                if (name.has_value() && !ui_is_blank(*name)) {
                    if (auto saved = job_manager_.save_job(*name, options)) {
                        append_log("Saved job '" + saved->Name + "'.");
                    }
                }
                break;
            }
            default:
                break;
        }
    }

    // Offers tasks that never started in the previous session. An elevated
    // relaunch restores them silently: the unelevated window wrote them a
    // moment ago as part of the same session.
    void restore_pending_tasks() {
        ui::PendingTasks pending = ui::load_pending_tasks();
        if (pending.tasks.empty()) return;
        const bool silent = incoming_handoff_.has_value();
        if (!silent) {
            const std::wstring text = std::to_wstring(pending.tasks.size()) +
                                      (pending.tasks.size() == 1
                                           ? L" task from your last session did not run. Restore it?"
                                           : L" tasks from your last session did not run. Restore them?");
            if (!confirm_box(text, L"XactCopy - Unstarted tasks")) {
                persist_pending_tasks(); // forget them
                return;
            }
        }
        for (auto& task : pending.tasks) {
            task.state = ui::TaskState::Queued;
            tasks_.add(std::move(task));
        }
        tasks_running_ = silent && pending.list_running;
        append_log("Restored " + std::to_string(pending.tasks.size()) + " unstarted task(s)." +
                   (tasks_running_ ? "" : " Press Start to run them."));
        tasks_changed();
    }

    // Keyboard handling for the two panel lists (Del removes; Enter starts a
    // queued task), and right-click context menus.
    static LRESULT CALLBACK panel_list_subclass(HWND control, UINT message, WPARAM wparam,
                                                LPARAM lparam, UINT_PTR, DWORD_PTR data) {
        auto* self = reinterpret_cast<MainWindow*>(data);
        if (message == WM_KEYDOWN) {
            if (control == self->sources_list_) {
                if (wparam == VK_DELETE) { self->remove_selected_sources(); return 0; }
                if (wparam == 'V' && (GetKeyState(VK_CONTROL) & 0x8000) != 0) {
                    self->paste_sources();
                    return 0;
                }
                if (wparam == 'A' && (GetKeyState(VK_CONTROL) & 0x8000) != 0) {
                    SendMessageW(control, LB_SETSEL, TRUE, -1);
                    return 0;
                }
            } else if (control == self->tasks_list_) {
                if (wparam == VK_DELETE) { self->on_task_command(IdTaskRemove); return 0; }
                if (wparam == VK_RETURN) { self->on_task_command(IdTaskStartNow); return 0; }
            }
        }
        if (message == WM_GETDLGCODE && (wparam == VK_RETURN || wparam == VK_DELETE)) {
            return DLGC_WANTALLKEYS;
        }
        if (message == WM_CONTEXTMENU) {
            if (control == self->sources_list_) self->show_sources_context_menu(lparam);
            else self->show_tasks_context_menu(lparam);
            return 0;
        }
        if (message == WM_RBUTTONDOWN) {
            // Select the row under the cursor before the context menu opens.
            const LRESULT hit = SendMessageW(control, LB_ITEMFROMPOINT, 0, lparam);
            if (HIWORD(hit) == 0) {
                const int index = LOWORD(hit);
                if (control == self->tasks_list_) {
                    SendMessageW(control, LB_SETCURSEL, index, 0);
                } else if (SendMessageW(control, LB_GETSEL, index, 0) <= 0) {
                    SendMessageW(control, LB_SETSEL, FALSE, -1);
                    SendMessageW(control, LB_SETSEL, TRUE, index);
                }
            }
        }
        if (message == WM_NCDESTROY) {
            RemoveWindowSubclass(control, &panel_list_subclass, control == self->sources_list_ ? 2 : 3);
        }
        return DefSubclassProc(control, message, wparam, lparam);
    }

    void draw_log_item(const DRAWITEMSTRUCT& draw) {
        const bool selected = (draw.itemState & ODS_SELECTED) != 0;
        ui::themedraw::fill_rect(draw.hDC, draw.rcItem, selected ? theme_.selection : theme_.edit);
        if (draw.itemID == static_cast<UINT>(-1)) return;
        const auto entry = log_model_.at(log_filter_mode_, draw.itemID);
        if (entry == nullptr) return;
        COLORREF color = selected ? theme_.selection_text
                                  : colorize_log_ ? ui::severity_color(theme_, entry->severity)
                                                  : theme_.text;

        HGDIOBJ old_font = SelectObject(draw.hDC, mono_font_);
        SetBkMode(draw.hDC, TRANSPARENT);
        SetTextColor(draw.hDC, color);
        RECT text_rect = draw.rcItem;
        text_rect.left += 4;
        DrawTextW(draw.hDC, entry->text.c_str(), -1, &text_rect,
                  DT_LEFT | DT_VCENTER | DT_SINGLELINE | DT_NOPREFIX);
        if ((draw.itemState & ODS_FOCUS) != 0) DrawFocusRect(draw.hDC, &draw.rcItem);
        SelectObject(draw.hDC, old_font);
    }

    // ---- Layout ------------------------------------------------------------

    // Two columns above a full-width log. Left, the task builder: the Copy /
    // Assess switch, the item list, destination (copy only), the three
    // mode-specific choices with captions, the collapsible options, and the
    // actions. Right, the task list and the running task's detail. Each
    // column is laid out by one function that can measure without moving
    // anything, so the fixed parts are known exactly and only the two lists
    // and the log flex; required_client_height_ keeps the window from going
    // below what the current mode and options need.
    void layout() {
        RECT client;
        GetClientRect(hwnd_, &client);
        const int dpi = effective_dpi();
        auto scale = [dpi](int value) { return MulDiv(value, dpi, 96); };

        const int margin = scale(12);
        const int row_height = scale(24);
        const int gap = scale(8);
        const int section_gap = scale(14);
        const int column_gap = scale(20);
        const int button_height = scale(30);
        const int label_height = scale(18);
        const int tab_height = scale(30);
        const int list_floor = scale(72);
        const int client_width = static_cast<int>(client.right - client.left) - margin * 2;
        const int client_height = static_cast<int>(client.bottom - client.top);
        const bool scan = is_scan_mode();

        const int left_width = std::max(scale(400), (client_width - column_gap) * 46 / 100);
        const int right_width = std::max(scale(200), client_width - left_width - column_gap);

        int origin_x = margin;
        int y = margin;
        bool commit = false;
        auto place = [&](HWND handle, int x, int w, int h = 0) {
            if (!commit || handle == nullptr) return;
            MoveWindow(handle, origin_x + x, y, std::max(0, w), h == 0 ? row_height : h, TRUE);
            // MoveWindow only invalidates a resized control's newly exposed
            // strip, so widened statics/bars keep stale pixels.
            InvalidateRect(handle, nullptr, TRUE);
        };
        auto show = [&](HWND handle, bool visible) {
            if (commit && handle != nullptr) ShowWindow(handle, visible ? SW_SHOW : SW_HIDE);
        };

        // Both lists start at the same height: the add buttons share a row
        // with the "What to copy" heading when the column is wide enough, and
        // the task list's toolbar row mirrors that.
        const int files_width = scale(66);
        const int folder_width = scale(78);
        const int paste_width = scale(58);
        const int clear_width = scale(56);
        const int buttons_width = files_width + folder_width + paste_width + clear_width + gap * 3;
        const bool one_row = left_width >= buttons_width + scale(104) + gap;
        const int header_height = one_row ? row_height : row_height * 2 + scale(4);

        // ---- Left column: returns its height for a given list height ----
        HWND option_checks[] = {salvage_check_,    resume_check_,   map_check_,
                                adaptive_check_,   continue_check_, skip_known_bad_check_,
                                wait_media_check_, fragile_check_};
        HWND option_fields[] = {buffer_label_,  buffer_edit_,   retries_label_,
                                retries_edit_,  timeout_label_, timeout_edit_};
        auto left_column = [&](int sources_height) {
            origin_x = margin;
            y = margin;
            const int width = left_width;

            // Copy / Assess segmented switch.
            const int tab_width = std::min(width / 2, scale(180));
            place(mode_copy_tab_, 0, tab_width, tab_height);
            place(mode_scan_tab_, tab_width - 1, tab_width + 1, tab_height);
            y += tab_height + section_gap;

            // Heading + add buttons; the buttons drop to their own row when
            // the column is too narrow for both.
            place(source_label_, 0, one_row ? width - buttons_width - gap : width, row_height);
            if (!one_row) y += row_height + scale(4);
            int x = width;
            x -= clear_width;          place(clear_selection_, x, clear_width);
            x -= gap + paste_width;    place(source_paste_, x, paste_width);
            x -= gap + folder_width;   place(source_browse_, x, folder_width);
            x -= gap + files_width;    place(source_add_files_, x, files_width);
            y += row_height + scale(6);

            place(sources_list_, 0, width, sources_height);
            y += sources_height + scale(3);
            place(sources_footer_, 0, width, label_height);
            y += label_height + section_gap;

            // The destination row's slot is kept in Assess mode (a note
            // fills it) so switching modes never moves the lists.
            place(scan_info_label_, 0, width);
            {
                const int label_width = scale(78);
                const int browse_width = scale(84);
                const int recent_width = scale(26);
                const int edit_width = width - label_width - browse_width - recent_width - gap;
                place(destination_label_, 0, label_width);
                place(destination_edit_, label_width, edit_width);
                place(recent_destinations_button_, label_width + edit_width - 1, recent_width);
                place(destination_browse_, width - browse_width, browse_width);
                y += row_height + gap + scale(4);
            }

            // The three mode-specific choices, captioned: three across when
            // there is room (caption above), otherwise one per row.
            HWND captions[] = {profile_label_, policy_label_, verify_label_};
            HWND combos[] = {engine_combo_, overwrite_combo_, verify_combo_};
            // A closed owner-draw combo is its item height plus its frame.
            const int combo_height =
                engine_combo_ != nullptr
                    ? static_cast<int>(SendMessageW(engine_combo_, CB_GETITEMHEIGHT,
                                                    static_cast<WPARAM>(-1), 0)) + scale(6)
                    : row_height;
            if (width >= scale(500)) {
                const int third = (width - gap * 2) / 3;
                for (int i = 0; i < 3; ++i) {
                    place(captions[i], i * (third + gap), third, label_height);
                }
                y += label_height + scale(2);
                for (int i = 0; i < 3; ++i) {
                    place(combos[i], i * (third + gap), third, row_height * 8);
                }
                y += combo_height + gap;
            } else {
                const int caption_width = scale(92);
                for (int i = 0; i < 3; ++i) {
                    place(captions[i], 0, caption_width, combo_height);
                    place(combos[i], caption_width, width - caption_width, row_height * 8);
                    y += combo_height + scale(6);
                }
            }

            y += scale(4);
            place(options_toggle_, 0, width);
            y += row_height + gap;

            for (HWND control : option_checks) show(control, false);
            for (HWND control : option_fields) show(control, options_expanded_);
            if (options_expanded_) {
                std::vector<HWND> visible;
                for (HWND control : option_checks) {
                    if (control == salvage_check_ && scan) continue;
                    visible.push_back(control);
                }
                const int half = (width - gap) / 2;
                for (std::size_t i = 0; i < visible.size(); i += 2) {
                    show(visible[i], true);
                    place(visible[i], scale(4), half - scale(4));
                    if (i + 1 < visible.size()) {
                        show(visible[i + 1], true);
                        place(visible[i + 1], half + gap, half);
                    }
                    y += row_height + scale(4);
                }
                y += scale(4);
                const int slot = (width - gap * 2) / 3;
                const int edit_width = std::min(scale(50), slot / 2);
                int fx = 0;
                auto place_field = [&](HWND label, HWND edit) {
                    place(label, fx + scale(4), slot - edit_width - scale(8));
                    place(edit, fx + slot - edit_width, edit_width);
                    fx += slot + gap;
                };
                place_field(buffer_label_, buffer_edit_);
                place_field(retries_label_, retries_edit_);
                place_field(timeout_label_, timeout_edit_);
                y += row_height + gap;
            }

            y += scale(6);
            // Settings on the left, Add/Start on the right; narrow columns
            // shrink the two action buttons rather than overlapping Settings.
            const int settings_width = scale(100);
            const int action_width =
                std::min(scale(124), (width - settings_width - gap * 2) / 2);
            place(settings_button_, 0, settings_width, button_height);
            place(start_button_, width - action_width, action_width, button_height);
            place(add_task_button_, width - action_width * 2 - gap, action_width, button_height);
            y += button_height;
            return y - margin;
        };

        // ---- Right column: returns its height for a given tasks height ----
        auto right_column = [&](int tasks_height) {
            origin_x = margin + left_width + column_gap;
            y = margin;
            const int width = right_width;

            place(tasks_label_, 0, width, tab_height);
            y += tab_height + section_gap;
            // Toolbar row level with the add buttons opposite.
            y += header_height - row_height;
            const int finished_width = scale(118);
            place(pause_list_check_, 0, std::max(0, width - finished_width - gap), row_height);
            place(clear_finished_button_, width - finished_width, finished_width);
            y += row_height + scale(6);
            place(tasks_list_, 0, width, tasks_height);
            y += tasks_height + scale(3) + label_height + section_gap;

            // Running task: heading with Pause/Cancel beside it, then detail.
            const int control_width = scale(92);
            place(detail_label_, 0, std::max(0, width - control_width * 2 - gap * 2),
                  button_height);
            place(cancel_button_, width - control_width, control_width, button_height);
            place(pause_button_, width - control_width * 2 - gap, control_width, button_height);
            y += button_height + scale(4);
            place(current_label_, 0, width);
            y += row_height;
            const int bar_height = scale(bar_show_percentage_
                                             ? std::max(progress_bar_height_logical_, 14)
                                             : progress_bar_height_logical_);
            const int progress_label_width = scale(56);
            const int progress_row_height = std::max(row_height, bar_height);
            for (auto [label, bar] : {std::pair<HWND, HWND>{file_progress_label_, current_bar_},
                                      std::pair<HWND, HWND>{overall_progress_label_, overall_bar_}}) {
                place(label, 0, progress_label_width, progress_row_height);
                const int offset = (progress_row_height - bar_height) / 2;
                y += offset;
                place(bar, progress_label_width, width - progress_label_width, bar_height);
                y += progress_row_height - offset + scale(4);
            }
            place(stats_label_, 0, width, label_height);
            y += label_height + scale(2);
            const int col_width = (width - gap) / 2;
            auto place_two = [&](HWND left, HWND right) {
                place(left, 0, col_width, label_height);
                place(right, col_width + gap, col_width, label_height);
                y += label_height;
            };
            place_two(throughput_label_, eta_label_);
            show(buffer_usage_label_, show_buffer_row_);
            show(rescue_label_, show_rescue_row_);
            if (show_buffer_row_ || show_rescue_row_) {
                place_two(show_buffer_row_ ? buffer_usage_label_ : nullptr,
                          show_rescue_row_ ? rescue_label_ : nullptr);
            }
            place(job_summary_label_, 0, width, label_height);
            y += label_height;
            place(journal_label_, 0, width, label_height);
            y += label_height;
            show(diagnostics_label_, show_diagnostics_);
            if (show_diagnostics_) {
                place(diagnostics_label_, 0, width, label_height);
                y += label_height;
            }
            return y - margin;
        };

        // ---- Measure, then split the height ----
        const int left_fixed = left_column(0);
        const int right_fixed = right_column(0);
        const int status_height = row_height;
        const int log_toolbar_height = row_height;
        const int min_log = scale(100);
        const int bottom_fixed = gap + log_toolbar_height + gap + gap + status_height;
        const int top_min = std::max(left_fixed, right_fixed) + list_floor;
        required_client_height_ = margin * 2 + top_min + bottom_fixed + min_log;

        const int body = client_height - margin * 2;
        // The lists get a useful height (list_target) before the log grows
        // past its minimum; beyond that the split is two-thirds / one-third.
        const int list_target = scale(160);
        int top_height = std::max({top_min, (body - bottom_fixed) * 66 / 100,
                                   std::max(left_fixed, right_fixed) + list_target});
        top_height = std::max(top_min, std::min(top_height, body - bottom_fixed - min_log));

        // One height for both lists, so they line up top and bottom.
        commit = true;
        const int list_height = std::max(list_floor, top_height - std::max(left_fixed, right_fixed));
        left_column(list_height);
        right_column(list_height);
        show(mode_combo_, false);

        // ---- Bottom: log ----
        origin_x = margin;
        y = margin + top_height + gap;
        const int filter_width = scale(190);
        const int find_button_width = scale(95);
        place(log_filter_, 0, filter_width, scale(160));
        place(log_find_, filter_width + gap,
              std::max(scale(80), client_width - filter_width - find_button_width - 2 * gap));
        place(log_find_button_, client_width - find_button_width, find_button_width);
        y += log_toolbar_height + gap;
        const int log_height =
            std::max(scale(40), client_height - margin - status_height - gap - y);
        place(log_list_, 0, client_width, log_height);
        y += log_height + gap;
        place(status_label_, 0, client_width);
        tooltips_.update_layout();
    }

    // Grows the window (within its monitor's work area) when the current mode
    // and options need more height than it has, e.g. after expanding Options.
    void ensure_window_fits_layout() {
        if (hwnd_ == nullptr || IsZoomed(hwnd_) || IsIconic(hwnd_)) return;
        RECT client{};
        GetClientRect(hwnd_, &client);
        const int shortfall = required_client_height_ - static_cast<int>(client.bottom);
        if (shortfall <= 0) return;
        RECT window{};
        GetWindowRect(hwnd_, &window);
        MONITORINFO monitor{};
        monitor.cbSize = sizeof(monitor);
        if (!GetMonitorInfoW(MonitorFromWindow(hwnd_, MONITOR_DEFAULTTONEAREST), &monitor)) return;
        const RECT& work = monitor.rcWork;
        const int height = std::min<int>(window.bottom - window.top + shortfall,
                                         work.bottom - work.top);
        int top = window.top;
        if (top + height > work.bottom) top = std::max<int>(work.top, work.bottom - height);
        SetWindowPos(hwnd_, nullptr, window.left, top, window.right - window.left, height,
                     SWP_NOZORDER | SWP_NOACTIVATE);
    }

    // ---- Commands ----------------------------------------------------------

    void on_command(int id, int notification) {
        if (id == IdLogFilter && notification == CBN_SELCHANGE) {
            const int selection = static_cast<int>(SendMessageW(log_filter_, CB_GETCURSEL, 0, 0));
            log_filter_mode_ = static_cast<ui::LogFilter>(std::clamp(selection, 0, 2));
            SendMessageW(log_list_, LB_SETCOUNT, log_model_.size(log_filter_mode_), 0);
            InvalidateRect(log_list_, nullptr, TRUE);
            return;
        }
        if (id == IdLogFindNext) { find_log_next(); return; }
        if (id == IdLogCopySelected) { copy_log(false); return; }
        if (id == IdLogCopyAll) { copy_log(true); return; }
        if (id == IdLogSave) { save_log(); return; }
        // The collapsed Options header summarises the current option values.
        if (options_toggle_ != nullptr && !options_expanded_ && id != IdOptionsToggle) {
            InvalidateRect(options_toggle_, nullptr, FALSE);
        }
        if (id == IdPauseListCheck && notification == BN_CLICKED) {
            set_check(id, !get_check(id));
            InvalidateRect(pause_list_check_, nullptr, FALSE);
            return;
        }
        if (id == IdTasksList && notification == LBN_DBLCLK) {
            if (const ui::SessionTask* task = selected_task()) {
                if (task->state == ui::TaskState::Queued) on_task_command(IdTaskStartNow);
                else if (ui::is_finished(task->state)) on_task_command(IdTaskEdit);
            }
            return;
        }
        if (id >= IdSourceRemove && id <= IdSourceCopyPath) {
            on_source_command(id);
            return;
        }
        if (id >= IdTaskStartNow && id <= IdTaskEdit) {
            on_task_command(id);
            return;
        }
        if (id == IdModeCombo && notification == CBN_SELCHANGE) {
            sync_mode_ui(true);
            return;
        }
        if ((id == IdModeCopyTab || id == IdModeScanTab) && notification == BN_CLICKED) {
            select_mode(id == IdModeScanTab);
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
                add_folders_to_sources();
                break;
            case IdSourceAddFiles:
                add_files_to_selection();
                break;
            case IdSourcePaste:
                paste_sources();
                break;
            case IdClearSelection:
                set_sources({});
                break;
            case IdDestinationBrowse:
                browse_folder(destination_edit_);
                break;
            case IdRecentDestinations:
                show_recent_destinations();
                break;
            case IdOptionsToggle:
                set_options_expanded(!options_expanded_);
                break;
            case IdAddTaskButton:
                add_tasks_from_builder();
                break;
            case IdClearFinished:
                tasks_.clear_finished();
                tasks_changed();
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
                    log_model_.set_capacity(static_cast<std::size_t>(
                        settings_.get_int("UiMaxLogLines", 1000, 1000000, 50000)));
                    SendMessageW(log_list_, LB_SETCOUNT, log_model_.size(log_filter_mode_), 0);
                    configure_diagnostics_timer();
                    rebuild_fonts(GetDpiForWindow(hwnd_));
                    apply_fonts();
                    apply_theme();
                    layout();
                    sync_explorer_integration(true);
                }
                break;
            case IdMenuSaveDefaults:
                save_current_as_defaults();
                break;
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
            info_box(L"Add a source before saving a job.");
            return;
        }
        const std::size_t groups = ui::group_sources(sources_).size();
        if (groups > 1) {
            info_box(L"A saved job reads from one folder, but these sources span " +
                     std::to_wstring(groups) +
                     L" folders. Add them as tasks, then use Save as job on each task.");
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

    // ---- Incoming items ------------------------------------------------------

    // Single entry point for launch arguments, a second instance's forwarded
    // command line, the pickers, and drag-and-drop. Items join the sources list
    // immediately; a burst of Explorer launches (one per selected item) is
    // coalesced only for the follow-up: switching to assessment mode and
    // asking for a destination once.
    void queue_selection(const std::vector<std::wstring>& paths, bool scan_mode,
                         bool want_destination) {
        if (paths.empty()) return;
        add_sources(paths);
        if (scan_mode) launch_.explorer_scan_mode = true;
        if (want_destination) selection_wants_destination_ = true;
        SetTimer(hwnd_, IdSelectionTimer, SelectionCoalesceMs, nullptr);
        selection_timer_active_ = true;
    }

    void commit_selection() {
        KillTimer(hwnd_, IdSelectionTimer);
        selection_timer_active_ = false;
        const bool want_destination = selection_wants_destination_;
        selection_wants_destination_ = false;
        // An Explorer verb names its mode ("Copy with" / "Assess with"), so it
        // wins over whatever mode the window was in; drops keep the mode.
        if (want_destination || launch_.explorer_scan_mode) {
            select_mode(launch_.explorer_scan_mode);
        }
        launch_.explorer_scan_mode = false;
        if (IsIconic(hwnd_)) ShowWindow(hwnd_, SW_RESTORE);
        SetForegroundWindow(hwnd_);
        if (want_destination) prompt_destination_from_explorer();
    }

    // Folder picker with multi-select; each folder is added as an item.
    void add_folders_to_sources() {
        auto picked = pick_paths(/*folders*/ true, L"Add folders to copy");
        add_sources(picked);
    }

    // Multi-select shell picker for files or folders.
    std::vector<std::wstring> pick_paths(bool folders, const wchar_t* title) {
        std::vector<std::wstring> picked;
        if (FAILED(CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED | COINIT_DISABLE_OLE1DDE))) {
            return picked;
        }
        IFileOpenDialog* dialog = nullptr;
        if (SUCCEEDED(CoCreateInstance(CLSID_FileOpenDialog, nullptr, CLSCTX_INPROC_SERVER,
                                       IID_PPV_ARGS(&dialog)))) {
            DWORD options = 0;
            dialog->GetOptions(&options);
            dialog->SetOptions(options | FOS_ALLOWMULTISELECT | FOS_FORCEFILESYSTEM |
                               (folders ? FOS_PICKFOLDERS : FOS_FILEMUSTEXIST));
            dialog->SetTitle(title);
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
        return picked;
    }

    void add_files_to_selection() {
        add_sources(pick_paths(/*folders*/ false, L"Add files to copy"));
    }

    // ---- Explorer launch application (--from-explorer[-folder]) ------------

    static std::wstring resolve_full_path(const std::string& raw) {
        std::wstring wide = storage::fsutil::trim(storage::fsutil::utf8_to_wide(raw));
        if (wide.empty()) return std::wstring();
        return storage::fsutil::get_full_path(wide);
    }

    static bool directory_exists(const std::wstring& path) {
        DWORD attributes = GetFileAttributesW(storage::fsutil::long_path(path).c_str());
        return attributes != INVALID_FILE_ATTRIBUTES &&
               (attributes & FILE_ATTRIBUTE_DIRECTORY) != 0;
    }


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
            // The folder-background verb means "this folder's contents".
            add_sources({folder}, /*contents_only*/ true);
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
        set_sources(ui::items_from_options(options));
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
        options.SourceRoot.clear(); // set from the sources list below
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

        // Single-job callers (save as job, defaults) get the first source
        // group; task creation splits every group into its own task.
        const auto groups = ui::group_sources(sources_);
        if (!groups.empty()) {
            options.SourceRoot = storage::fsutil::wide_to_utf8(groups.front().root);
            options.SelectedRelativePaths = groups.front().relative_paths;
        }
        return options;
    }

    // Start: add tasks for the builder's sources and run the list. While a
    // task is running the new tasks simply queue behind it.
    void start_job() {
        if (add_tasks_from_builder() <= 0) return;
        tasks_running_ = true;
        persist_pending_tasks();
        if (supervisor_.is_job_running() || starting_) {
            append_log("Queued; the new task(s) start after the current one.");
            return;
        }
        run_next_task();
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
        if (outgoing_handoff_.has_value()) return; // an elevated relaunch is taking over

        // Direct NTFS reads need administrator rights. Every start path (manual,
        // Job Manager, queue, crash recovery, Explorer verb) passes through
        // here, so this is where an unelevated window hands the run to an
        // elevated relaunch instead of silently falling back to file reads.
        std::string raw_unsupported;
        if (ui::elevation::raw_scan_requires_elevation(options, process_elevated_,
                                                       raw_unsupported)) {
            if (!begin_elevated_handoff(options, from_recovery, managed_run_id,
                                        managed_display_name)) {
                return;
            }
            // Declined or failed, and the user chose standard file reads.
        } else if (!raw_unsupported.empty() &&
                   options.OperationMode == models::JobOperationMode::ScanOnly &&
                   options.UseExperimentalRawDiskScan) {
            append_log("Direct NTFS reads are not available: " + raw_unsupported +
                       ". Standard file reads will be used.");
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
        result_applied_since_start_ = false;
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

    // ---- Direct-NTFS elevation handoff ------------------------------------

    // Returns true when the caller should continue starting the run here
    // (with direct reads switched off because elevation was declined or
    // failed); false when the run was handed off or abandoned.
    bool begin_elevated_handoff(models::CopyJobOptions& options, bool from_recovery,
                                const std::string& managed_run_id,
                                const std::string& managed_display_name) {
        ui::elevation::HandoffRequest request;
        request.Options = options;
        request.FromRecovery = from_recovery;
        request.ManagedRunId = managed_run_id;
        request.DisplayName = managed_display_name;
        request.RecoveryRunId = from_recovery ? pending_recovery_run_id_ : std::string();
        request.UserSid = ui::elevation::current_user_sid();
        request.CreatedUnixMs = ui::elevation::unix_now_ms();

        append_log("Direct NTFS volume reads need administrator rights. Asking Windows for approval...");
        const std::string token = ui::elevation::write_handoff(request);
        if (token.empty()) {
            return handle_elevation_unavailable(
                request, options, "The elevation request could not be saved.");
        }
        HANDLE process = nullptr;
        DWORD error = 0;
        const auto outcome = ui::elevation::relaunch_elevated(hwnd_, token, process, error);
        if (outcome != ui::elevation::LaunchOutcome::Started) {
            ui::elevation::discard_handoff(token);
            return handle_elevation_unavailable(
                request, options,
                outcome == ui::elevation::LaunchOutcome::Declined
                    ? std::string("Administrator approval was declined.")
                    : "XactCopy could not be restarted as administrator (Win32 " +
                          std::to_string(error) + ").");
        }

        outgoing_handoff_ = OutgoingHandoff{std::move(request), token, process, GetTickCount64()};
        EnableWindow(start_button_, FALSE);
        set_inputs_enabled(false);
        SetWindowTextW(status_label_, L"Status: Handing the run to the administrator window...");
        SetTimer(hwnd_, IdElevationTimer, 250, nullptr);
        return false;
    }

    // Elevation is not possible for this attempt. Offer the same run with
    // standard file reads, or leave every record as it was before the click.
    bool handle_elevation_unavailable(const ui::elevation::HandoffRequest& request,
                                      models::CopyJobOptions& options,
                                      const std::string& reason) {
        append_log("WARNING: " + reason);
        if (confirm_box(storage::fsutil::utf8_to_wide(
                            reason + "\n\nRun this assessment with standard file reads instead? "
                                     "Results are equivalent for readable data, but unreadable "
                                     "regions are located through the filesystem rather than "
                                     "the raw volume."),
                        L"XactCopy - Direct NTFS reads unavailable", ui::MessageIcon::Question)) {
            options.UseExperimentalRawDiskScan = false;
            append_log("Continuing with standard file reads for this run.");
            return true;
        }
        abandon_handoff_request(request, reason);
        return false;
    }

    void abandon_handoff_request(const ui::elevation::HandoffRequest& request,
                                 const std::string& reason) {
        // A queued or Job Manager run was already recorded; close it out so it
        // does not sit in the catalog as created-but-never-started.
        if (!request.ManagedRunId.empty()) {
            models::CopyJobResult failure;
            failure.ErrorMessage = "Not started: " + reason;
            job_manager_.mark_run_completed(request.ManagedRunId, failure);
        }
        // A declined crash-recovery resume stays pending for the next launch.
        if (request.FromRecovery) recovery_.mark_resume_prompt_deferred(settings_, false);
        append_log("Run not started: " + reason);
    }

    void poll_elevated_handoff() {
        if (!outgoing_handoff_.has_value()) {
            KillTimer(hwnd_, IdElevationTimer);
            return;
        }
        OutgoingHandoff& handoff = *outgoing_handoff_;
        const bool exited = WaitForSingleObject(handoff.process, 0) == WAIT_OBJECT_0;
        // The elevated instance deletes the request file the moment it has
        // read and validated it; that is its acceptance.
        if (!exited && !ui::elevation::handoff_pending(handoff.token)) {
            KillTimer(hwnd_, IdElevationTimer);
            CloseHandle(handoff.process);
            outgoing_handoff_.reset();
            append_log("The administrator window accepted the run. Closing this window.");
            // Queued tasks are already in the pending-task file; the elevated
            // window restores them and continues the list.
            closing_for_handoff_ = true;
            ShowWindow(hwnd_, SW_HIDE);
            // Clean shutdown releases the single-instance mutex the elevated
            // window is waiting for. No job is running in this window.
            PostMessageW(hwnd_, WM_CLOSE, 0, 0);
            return;
        }
        const bool timed_out = GetTickCount64() - handoff.started_tick > 60'000;
        if (!exited && !timed_out) return;

        KillTimer(hwnd_, IdElevationTimer);
        DWORD exit_code = 0;
        if (exited) GetExitCodeProcess(handoff.process, &exit_code);
        CloseHandle(handoff.process);
        ui::elevation::discard_handoff(handoff.token);
        ui::elevation::HandoffRequest request = std::move(handoff.request);
        outgoing_handoff_.reset();
        EnableWindow(start_button_, TRUE);
        set_inputs_enabled(true);
        SetWindowTextW(status_label_, L"Status: Idle");

        std::string reason;
        if (!exited) {
            reason = "The administrator window did not respond.";
        } else if (exit_code == ui::elevation::ExitHandoffWrongUser) {
            reason = "Administrator approval was given for a different Windows account. XactCopy "
                     "keeps journals and keys per account, so the run cannot continue there.";
        } else {
            reason = "The administrator window could not take over the run (exit code " +
                     std::to_string(exit_code) + ").";
        }
        models::CopyJobOptions options = request.Options;
        if (handle_elevation_unavailable(request, options, reason)) {
            if (request.FromRecovery) pending_recovery_run_id_ = request.RecoveryRunId;
            start_job_with(options, request.FromRecovery, request.ManagedRunId,
                           request.DisplayName);
            pending_recovery_run_id_.clear();
        }
        settle_unstarted_task(running_task_id_, "Not started: administrator approval unavailable");
    }

    void on_start_done(bool ok, const std::string& error) {
        starting_ = false;
        if (ok && result_applied_since_start_) {
            // The worker finished (or failed) before this notification was
            // processed; apply_result already settled the task and run.
            update_action_buttons();
            update_menu_state();
            return;
        }
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
            EnableWindow(pause_button_, TRUE);
            EnableWindow(cancel_button_, TRUE);
            update_action_buttons();
            update_menu_state();
            bind_started_job_to_task();
            append_log("Job started (" + active_run_id_ + ").");
        } else {
            if (!active_managed_run_id_.empty()) {
                models::CopyJobResult failure;
                failure.ErrorMessage = error;
                job_manager_.mark_run_completed(active_managed_run_id_, failure);
                active_managed_run_id_.clear();
            }
            if (ui::SessionTask* task = tasks_.find(running_task_id_)) {
                task->state = ui::TaskState::Failed;
                task->outcome = "Could not start: " + error;
            }
            running_task_id_ = 0;
            tasks_running_ = false; // a start failure would repeat for every queued task
            tasks_changed();
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

    static LRESULT CALLBACK log_list_subclass(HWND control, UINT message, WPARAM wparam,
                                              LPARAM lparam, UINT_PTR, DWORD_PTR data) {
        auto* self = reinterpret_cast<MainWindow*>(data);
        if (message == LB_RESETCONTENT) self->log_model_.clear();
        return DefSubclassProc(control, message, wparam, lparam);
    }

    void queue_log(const std::string& text) {
        bool wake = false;
        {
            std::lock_guard<std::mutex> guard(log_queue_lock_);
            log_queue_.push_back({text, std::chrono::system_clock::now()});
            if (!log_wake_pending_) {
                log_wake_pending_ = true;
                wake = true;
            }
        }
        if (wake && !PostMessageW(hwnd_, WM_APP_LOG, 0, 0)) {
            std::lock_guard<std::mutex> guard(log_queue_lock_);
            log_wake_pending_ = false;
        }
    }

    void drain_log_queue() {
        std::deque<QueuedLog> batch;
        bool repost = false;
        {
            std::lock_guard<std::mutex> guard(log_queue_lock_);
            // Bound one wake's work so a producer burst cannot starve input.
            for (std::size_t count = 0; count < 512 && !log_queue_.empty(); ++count) {
                batch.push_back(std::move(log_queue_.front()));
                log_queue_.pop_front();
            }
            repost = !log_queue_.empty();
            log_wake_pending_ = repost;
        }
        if (repost && !PostMessageW(hwnd_, WM_APP_LOG, 0, 0)) {
            std::lock_guard<std::mutex> guard(log_queue_lock_);
            log_wake_pending_ = false;
        }
        append_log_batch(batch);
    }

    void append_log(const std::string& message) {
        std::deque<QueuedLog> batch;
        batch.push_back({message, std::chrono::system_clock::now()});
        append_log_batch(batch);
    }

    void append_log_batch(const std::deque<QueuedLog>& batch) {
        if (batch.empty() || log_list_ == nullptr) return;
        const int old_count = static_cast<int>(log_model_.size(log_filter_mode_));
        int top = std::max(0, static_cast<int>(SendMessageW(log_list_, LB_GETTOPINDEX, 0, 0)));
        int selected = static_cast<int>(SendMessageW(log_list_, LB_GETCURSEL, 0, 0));
        RECT client{};
        GetClientRect(log_list_, &client);
        const int item_height = std::max(1, static_cast<int>(SendMessageW(log_list_, LB_GETITEMHEIGHT, 0, 0)));
        const int visible_rows = (client.bottom - client.top + item_height - 1) / item_height;
        const bool follow_tail = old_count == 0 || top + visible_rows >= old_count;
        int added_visible = 0;
        HDC dc = GetDC(log_list_);
        HGDIOBJ old_font = dc != nullptr ? SelectObject(dc, mono_font_) : nullptr;
        int extent = log_horizontal_extent_;
        for (const auto& line : batch) {
            const auto severity = ui::classify_log_line(line.text);
            if (ui::LogModel::matches_filter(severity, log_filter_mode_)) {
                ++added_visible;
            }
            std::wstring wide = storage::fsutil::utf8_to_wide(line.text);
            log_max_characters_ = std::max(log_max_characters_, wide.size());
            if (dc != nullptr) extent = std::max(extent, measure_log_text_width(dc, wide));
            log_model_.append(std::move(wide), severity, line.timestamp);
        }
        if (dc != nullptr) {
            if (old_font != nullptr) SelectObject(dc, old_font);
            ReleaseDC(log_list_, dc);
        }
        const int count = static_cast<int>(log_model_.size(log_filter_mode_));
        const int removed = std::max(0, old_count + added_visible - count);
        SendMessageW(log_list_, WM_SETREDRAW, FALSE, 0);
        SendMessageW(log_list_, LB_SETCOUNT, count, 0);
        if (selected >= 0) {
            selected -= removed;
            SendMessageW(log_list_, LB_SETCURSEL, selected >= 0 ? selected : -1, 0);
        }
        if (count > 0) {
            const int target = follow_tail ? std::max(0, count - visible_rows)
                                           : std::max(0, top - removed);
            SendMessageW(log_list_, LB_SETTOPINDEX, target, 0);
        }
        if (extent > log_horizontal_extent_) {
            log_horizontal_extent_ = extent;
            SendMessageW(log_list_, LB_SETHORIZONTALEXTENT, extent, 0);
        }
        SendMessageW(log_list_, WM_SETREDRAW, TRUE, 0);
        InvalidateRect(log_list_, nullptr, TRUE);
    }

    bool handle_log_shortcuts(const MSG& message) {
        if (message.message != WM_KEYDOWN) return false;
        const bool control = (GetKeyState(VK_CONTROL) & 0x8000) != 0;
        const bool shift = (GetKeyState(VK_SHIFT) & 0x8000) != 0;
        if (control && message.wParam == VK_RETURN) { start_job(); return true; }
        if (control && message.wParam == 'O') {
            if (shift) add_folders_to_sources();
            else add_files_to_selection();
            return true;
        }
        if (message.wParam == 'F' && (GetKeyState(VK_CONTROL) & 0x8000) != 0) {
            SetFocus(log_find_);
            SendMessageW(log_find_, EM_SETSEL, 0, -1);
            return true;
        }
        if (message.wParam == VK_F3) { find_log_next(); return true; }
        return false;
    }

    void find_log_next() {
        const int length = GetWindowTextLengthW(log_find_);
        if (length <= 0) { SetFocus(log_find_); return; }
        std::wstring query(static_cast<std::size_t>(length) + 1, L'\0');
        GetWindowTextW(log_find_, query.data(), length + 1);
        query.resize(static_cast<std::size_t>(length));
        const int selected = static_cast<int>(SendMessageW(log_list_, LB_GETCURSEL, 0, 0));
        const auto entry = selected < 0 ? nullptr :
            log_model_.at(log_filter_mode_, static_cast<std::size_t>(selected));
        start_log_task(LogTaskKind::Find, {}, std::move(query), entry != nullptr,
                       entry != nullptr ? entry->sequence : 0);
    }

    void copy_log(bool all) {
        if (all) { start_log_task(LogTaskKind::Copy); return; }
        const int selected = static_cast<int>(SendMessageW(log_list_, LB_GETCURSEL, 0, 0));
        const auto entry = selected < 0 ? nullptr :
            log_model_.at(log_filter_mode_, static_cast<std::size_t>(selected));
        if (entry != nullptr) copy_log_text(entry->text + L"\r\n");
    }

    void copy_log_text(const std::wstring& text) {
        if (text.empty() || !OpenClipboard(hwnd_)) return;
        HGLOBAL memory = GlobalAlloc(GMEM_MOVEABLE, (text.size() + 1) * sizeof(wchar_t));
        if (memory != nullptr) {
            void* buffer = GlobalLock(memory);
            if (buffer != nullptr) {
                std::memcpy(buffer, text.c_str(), (text.size() + 1) * sizeof(wchar_t));
                GlobalUnlock(memory);
                EmptyClipboard();
                if (SetClipboardData(CF_UNICODETEXT, memory) != nullptr) memory = nullptr;
            }
            if (memory != nullptr) GlobalFree(memory);
        }
        CloseClipboard();
    }

    void show_log_context_menu(LPARAM point) {
        HMENU menu = CreatePopupMenu();
        AppendMenuW(menu, MF_STRING, IdLogCopySelected, L"Copy selected");
        AppendMenuW(menu, MF_STRING, IdLogCopyAll, L"Copy all visible");
        AppendMenuW(menu, MF_SEPARATOR, 0, nullptr);
        AppendMenuW(menu, MF_STRING, IdLogSave, L"Save log as...");
        POINT cursor{GET_X_LPARAM(point), GET_Y_LPARAM(point)};
        if (cursor.x == -1 && cursor.y == -1) {
            RECT rect{}; GetWindowRect(log_list_, &rect); cursor = {rect.left, rect.top};
        }
        TrackPopupMenu(menu, TPM_RIGHTBUTTON, cursor.x, cursor.y, 0, hwnd_, nullptr);
        DestroyMenu(menu);
    }

    void save_log() {
        if (log_save_running_.load()) { MessageBeep(MB_OK); return; }
        if (log_save_thread_.joinable()) log_save_thread_.join();
        IFileSaveDialog* dialog = nullptr;
        if (FAILED(CoCreateInstance(CLSID_FileSaveDialog, nullptr, CLSCTX_INPROC_SERVER,
                                    IID_PPV_ARGS(&dialog)))) return;
        const COMDLG_FILTERSPEC filters[] = {{L"Text files", L"*.txt"}, {L"All files", L"*.*"}};
        dialog->SetFileTypes(2, filters);
        dialog->SetDefaultExtension(L"txt");
        dialog->SetFileName(L"XactCopy-log.txt");
        std::wstring path;
        if (SUCCEEDED(dialog->Show(hwnd_))) {
            IShellItem* item = nullptr;
            if (SUCCEEDED(dialog->GetResult(&item))) {
                PWSTR name = nullptr;
                if (SUCCEEDED(item->GetDisplayName(SIGDN_FILESYSPATH, &name))) {
                    path = name;
                    CoTaskMemFree(name);
                }
                item->Release();
            }
        }
        dialog->Release();
        if (path.empty()) return;
        start_log_task(LogTaskKind::Save, std::move(path));
    }

    void start_log_task(LogTaskKind kind, std::wstring path = {}, std::wstring query = {},
                        bool has_selection = false, std::uint64_t selected_sequence = 0) {
        if (log_save_running_.exchange(true)) { MessageBeep(MB_OK); return; }
        if (log_save_thread_.joinable()) log_save_thread_.join();
        auto snapshot = log_model_.begin_snapshot();
        if (!snapshot) { log_save_running_ = false; MessageBeep(MB_OK); return; }
        const auto filter = log_filter_mode_;
        const HWND target = hwnd_;
        log_save_thread_ = std::thread([this, target, snapshot, kind, filter,
                                        path = std::move(path), query = std::move(query),
                                        has_selection, selected_sequence]() mutable {
            auto result = std::make_unique<LogTaskResult>();
            result->kind = kind;
            result->filter = filter;
            try {
                std::ofstream stream;
                if (kind == LogTaskKind::Save) {
                    stream.open(std::filesystem::path(path), std::ios::binary | std::ios::trunc);
                    if (!stream) throw std::runtime_error("Could not open the log file.");
                }
                for (auto& c : query) c = static_cast<wchar_t>(std::towlower(c));
                bool first_found = false;
                std::uint64_t first_sequence = 0;
                while (auto entry = log_model_.next_snapshot(snapshot)) {
                    if (!ui::LogModel::matches_filter(entry->severity, filter)) continue;
                    if (kind == LogTaskKind::Find) {
                        const auto match = std::search(entry->text.begin(), entry->text.end(),
                            query.begin(), query.end(), [](wchar_t left, wchar_t right) {
                                return std::towlower(left) == right;
                            });
                        if (match == entry->text.end()) continue;
                        if (!first_found) { first_found = true; first_sequence = entry->sequence; }
                        if (!has_selection || entry->sequence > selected_sequence) {
                            result->found = true; result->sequence = entry->sequence; break;
                        }
                    } else if (kind == LogTaskKind::Copy) {
                        result->text += entry->text;
                        result->text += L"\r\n";
                    } else {
                        const std::string bytes = storage::fsutil::wide_to_utf8(entry->text);
                        stream.write(bytes.data(), static_cast<std::streamsize>(bytes.size()));
                        stream.write("\r\n", 2);
                        if (!stream) throw std::runtime_error("Could not write the log file.");
                    }
                }
                if (kind == LogTaskKind::Find && !result->found && first_found) {
                    result->found = true; result->sequence = first_sequence;
                }
                if (kind == LogTaskKind::Save) {
                    stream.close();
                    if (!stream) throw std::runtime_error("Could not finish writing the log file.");
                } else if (kind == LogTaskKind::Copy && !result->text.empty()) {
                    const std::size_t bytes = (result->text.size() + 1) * sizeof(wchar_t);
                    result->clipboard = GlobalAlloc(GMEM_MOVEABLE, bytes);
                    if (result->clipboard == nullptr) throw std::runtime_error("Could not allocate clipboard data.");
                    void* buffer = GlobalLock(result->clipboard);
                    if (buffer == nullptr) throw std::runtime_error("Could not prepare clipboard data.");
                    std::memcpy(buffer, result->text.c_str(), bytes);
                    GlobalUnlock(result->clipboard);
                    std::wstring().swap(result->text);
                }
            } catch (const std::exception& error) {
                result->error = error.what();
            }
            log_model_.end_snapshot(snapshot);
            if (PostMessageW(target, WM_APP_LOG_TASK, 0,
                             reinterpret_cast<LPARAM>(result.get()))) result.release();
            log_save_running_ = false;
        });
    }

    void finish_log_task(LogTaskResult& result) {
        if (!result.error.empty()) { append_log("ERROR: " + result.error); return; }
        if (result.kind == LogTaskKind::Copy) {
            if (result.clipboard != nullptr && OpenClipboard(hwnd_)) {
                EmptyClipboard();
                if (SetClipboardData(CF_UNICODETEXT, result.clipboard) != nullptr)
                    result.clipboard = nullptr;
                CloseClipboard();
            }
            return;
        }
        if (result.kind == LogTaskKind::Find) {
            const auto index = result.found && result.filter == log_filter_mode_
                                   ? log_model_.index_of_sequence(log_filter_mode_, result.sequence)
                                   : log_model_.size(log_filter_mode_);
            if (index < log_model_.size(log_filter_mode_)) {
                SendMessageW(log_list_, LB_SETCURSEL, index, 0);
                SendMessageW(log_list_, LB_SETTOPINDEX, index, 0);
                SetFocus(log_list_);
            } else MessageBeep(MB_OK);
        }
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

        // Mirror the headline numbers onto the running task's row.
        if (ui::SessionTask* task = tasks_.find(running_task_id_)) {
            task->progress = snapshot.overall_progress();
            task->bytes_per_second = smoothed_bytes_per_second_;
            task->eta_seconds = remaining_bytes <= 0 ? 0
                                : reference_speed > 1.0
                                    ? static_cast<std::int64_t>(static_cast<double>(remaining_bytes) /
                                                                reference_speed)
                                    : -1;
            invalidate_task_row(task->id);
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
        result_applied_since_start_ = true;
        recovery_.mark_job_ended(active_run_id_);
        active_run_id_.clear();
        if (!active_managed_run_id_.empty()) {
            job_manager_.mark_run_completed(active_managed_run_id_, result);
            active_managed_run_id_.clear();
        }

        EnableWindow(pause_button_, FALSE);
        EnableWindow(cancel_button_, FALSE);
        SetWindowTextW(pause_button_, L"Pause");
        paused_ = false;
        set_inputs_enabled(true);
        finish_running_task(result);
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

        // Continue the session task list, or, when it is not running, drain
        // the Job Manager queue as before (MainForm.QueueAutoQueuedRun). Posted
        // so the supervisor finishes its completion bookkeeping first.
        PostMessageW(hwnd_, WM_APP_RUN_NEXT_TASK, 0, 0);
    }

    void on_run_next_task() {
        if (supervisor_.is_job_running() || starting_ || outgoing_handoff_.has_value()) return;
        if (tasks_running_ && tasks_.next_queued() != nullptr) {
            if (get_check(IdPauseListCheck)) {
                tasks_running_ = false;
                persist_pending_tasks();
                append_log("Task list paused after the current task. Press Start or use Start now to continue.");
                return;
            }
            run_next_task();
            return;
        }
        if (tasks_running_) {
            tasks_running_ = false;
            persist_pending_tasks();
        }
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
        // An elevated relaunch carries the run the user already chose in the
        // unelevated window, which had this same recovery state and already
        // asked about it. Asking again here would offer the run twice; any
        // other pending run stays pending for the next normal launch.
        if (incoming_handoff_.has_value()) return;

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
        pending_recovery_run_id_ = run.RunId;
        start_job_with(options, true, run.ManagedRunId, run.JobName);
        pending_recovery_run_id_.clear();
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
        } else if (_wcsicmp(arg, L"--scan-from-explorer-folder") == 0 && i + 1 < argc) {
            launch.explorer_scan_mode = true;
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

    // Elevated relaunch for a direct-NTFS assessment (see elevation.h). The
    // unelevated window that asked for it still holds the single-instance
    // mutex and is waiting for this process to accept the request, so this
    // path must not forward to it: validate, accept, then wait for the mutex.
    std::string handoff_token;
    std::string origin_sid;
    {
        int argc = 0;
        wchar_t** argv = CommandLineToArgvW(full_command_line, &argc);
        for (int i = 1; argv != nullptr && i + 1 < argc; ++i) {
            if (_wcsicmp(argv[i], ui::elevation::HandoffSwitch) == 0) {
                handoff_token = storage::fsutil::wide_to_utf8(argv[++i]);
            } else if (_wcsicmp(argv[i], ui::elevation::OriginSidSwitch) == 0) {
                origin_sid = storage::fsutil::wide_to_utf8(argv[++i]);
            }
        }
        if (argv != nullptr) LocalFree(argv);
    }
    std::optional<ui::elevation::HandoffRequest> handoff;
    if (!handoff_token.empty()) {
        if (origin_sid.empty() || origin_sid != ui::elevation::current_user_sid()) {
            // Over-the-shoulder UAC: an administrator approved with their own
            // account. This process's journals, keys, and settings belong to
            // that account, and it cannot read the requester's handoff file.
            MessageBoxW(nullptr,
                        L"Administrator approval was given for a different Windows account.\n\n"
                        L"XactCopy keeps its journals and keys per account, so direct NTFS reads "
                        L"must be approved by an administrator account that is also the signed-in "
                        L"user. The original window will offer standard file reads instead.",
                        L"XactCopy", MB_OK | MB_ICONWARNING);
            return static_cast<int>(ui::elevation::ExitHandoffWrongUser);
        }
        ui::elevation::HandoffRequest request;
        const auto status = ui::elevation::consume_handoff(handoff_token, request);
        if (status != ui::elevation::HandoffStatus::Ok) {
            const std::wstring text =
                L"XactCopy was started as administrator, but the run request was refused: " +
                storage::fsutil::utf8_to_wide(ui::elevation::describe(status)) + L".";
            MessageBoxW(nullptr, text.c_str(), L"XactCopy", MB_OK | MB_ICONWARNING);
            return static_cast<int>(ui::elevation::ExitHandoffRejected);
        }
        handoff = std::move(request);
    }

    if (handoff.has_value()) {
        HANDLE mutex = CreateMutexW(nullptr, TRUE, L"Local\\XactCopyNative.SingleInstance");
        if (mutex != nullptr && GetLastError() == ERROR_ALREADY_EXISTS) {
            const DWORD wait = WaitForSingleObject(mutex, 60'000);
            if (wait != WAIT_OBJECT_0 && wait != WAIT_ABANDONED) {
                MessageBoxW(nullptr,
                            L"The previous XactCopy window did not close, so the administrator "
                            L"window cannot take over. Close XactCopy and start the assessment again.",
                            L"XactCopy", MB_OK | MB_ICONWARNING);
                CloseHandle(mutex);
                return static_cast<int>(ui::elevation::ExitHandoffMutexTimeout);
            }
        }
        MainWindow window;
        window.set_incoming_handoff(std::move(*handoff));
        int result = window.run(instance, ui::LaunchOptions{});
        if (mutex != nullptr) {
            ReleaseMutex(mutex);
            CloseHandle(mutex);
        }
        return result;
    }

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
