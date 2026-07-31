// -----------------------------------------------------------------------------
// File: src\ui\dialogs.h
// Purpose: Themed modal dialogs for the native UI. The Settings dialog has seven
//          navigation pages (Appearance, Copy Defaults, Performance,
//          Diagnostics, Verification, Recovery & Startup, Explorer Integration)
//          built from a declarative field table bound to the DOM-preserving
//          settings store. Update settings live in the About dialog. Both sit on
//          a shared ModalHost.
// -----------------------------------------------------------------------------

#pragma once

#include <algorithm>
#include <functional>
#include <map>
#include <string>
#include <vector>

#include "../version.h"
#include "app_icon.h"
#include "settings.h"
#include "theme.h"

namespace xact::ui {

inline constexpr const char* NativeAppVersion = XACTCOPY_VERSION_STRING;

namespace dialog_detail {

// Minimal modal host: registers a class once, disables the owner, and pumps
// messages until the dialog window is destroyed.
class ModalHost {
public:
    using MessageHandler =
        std::function<LRESULT(HWND, UINT, WPARAM, LPARAM, bool& handled)>;

    HWND create(HWND owner, const wchar_t* class_name, const wchar_t* title, int width,
                int height, MessageHandler handler, bool resizable = false) {
        handler_ = std::move(handler);
        owner_ = owner;

        // WNDCLASSEXW (not WNDCLASSW) so the class can carry a small icon —
        // that is what the title bar draws.
        WNDCLASSEXW window_class{};
        window_class.cbSize = sizeof(window_class);
        window_class.lpfnWndProc = &ModalHost::static_proc;
        window_class.hInstance = GetModuleHandleW(nullptr);
        window_class.lpszClassName = class_name;
        window_class.hCursor = LoadCursorW(nullptr, IDC_ARROW);
        window_class.hIcon = load_app_icon(GetSystemMetrics(SM_CXICON), GetSystemMetrics(SM_CYICON));
        window_class.hIconSm =
            load_app_icon(GetSystemMetrics(SM_CXSMICON), GetSystemMetrics(SM_CYSMICON));
        RegisterClassExW(&window_class); // idempotent; re-register fails harmlessly

        UINT dpi = owner != nullptr ? GetDpiForWindow(owner) : 96;
        int scaled_width = MulDiv(width, static_cast<int>(dpi), 96);
        int scaled_height = MulDiv(height, static_cast<int>(dpi), 96);

        RECT owner_rect{};
        if (owner != nullptr) GetWindowRect(owner, &owner_rect);
        int x = owner_rect.left + ((owner_rect.right - owner_rect.left) - scaled_width) / 2;
        int y = owner_rect.top + ((owner_rect.bottom - owner_rect.top) - scaled_height) / 2;

        DWORD style = WS_POPUP | WS_CAPTION | WS_SYSMENU;
        DWORD ex_style = WS_EX_DLGMODALFRAME;
        if (resizable) {
            style |= WS_THICKFRAME | WS_MAXIMIZEBOX;
            ex_style = 0; // the modal frame forces a fixed border; drop it to size
        }
        hwnd_ = CreateWindowExW(ex_style, class_name, title, style, x, y, scaled_width,
                                scaled_height, owner, nullptr, GetModuleHandleW(nullptr), this);
        return hwnd_;
    }

    void run_modal() {
        if (hwnd_ == nullptr) return;
        bool owner_was_enabled = owner_ != nullptr && IsWindowEnabled(owner_);
        if (owner_was_enabled) EnableWindow(owner_, FALSE);
        ShowWindow(hwnd_, SW_SHOW);

        MSG message;
        while (IsWindow(hwnd_) && GetMessageW(&message, nullptr, 0, 0) > 0) {
            if (IsDialogMessageW(hwnd_, &message)) continue;
            TranslateMessage(&message);
            DispatchMessageW(&message);
        }
        if (owner_was_enabled && owner_ != nullptr) {
            EnableWindow(owner_, TRUE);
            SetForegroundWindow(owner_);
        }
    }

    HWND hwnd() const noexcept { return hwnd_; }

private:
    HWND hwnd_ = nullptr;
    HWND owner_ = nullptr;
    MessageHandler handler_;

    static LRESULT CALLBACK static_proc(HWND hwnd, UINT message, WPARAM wparam, LPARAM lparam) {
        ModalHost* self;
        if (message == WM_NCCREATE) {
            self = static_cast<ModalHost*>(reinterpret_cast<CREATESTRUCTW*>(lparam)->lpCreateParams);
            SetWindowLongPtrW(hwnd, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(self));
            self->hwnd_ = hwnd;
        } else {
            self = reinterpret_cast<ModalHost*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
        }
        if (self != nullptr && self->handler_) {
            bool handled = false;
            LRESULT result = self->handler_(hwnd, message, wparam, lparam, handled);
            if (handled) return result;
        }
        return DefWindowProcW(hwnd, message, wparam, lparam);
    }
};

} // namespace dialog_detail

// ---------------------------------------------------------------------------
// Themed message dialog — the dark-mode replacement for MessageBoxW. Adapted
// from the AxiomCompress message dialog: auto-sized wrapped text, a system
// icon, and owner-drawn buttons on the shared ModalHost.
// ---------------------------------------------------------------------------

enum class MessageIcon { None, Information, Warning, Error, Question };
enum class MessageButtons { Ok, OkCancel, YesNo, YesNoCancel };

class MessageDialog {
public:
    // Returns IDOK / IDCANCEL / IDYES / IDNO, matching MessageBoxW.
    static int show(HWND owner, const ThemePalette& theme, const std::wstring& title,
                    const std::wstring& message, MessageIcon icon = MessageIcon::Information,
                    MessageButtons buttons = MessageButtons::Ok) {
        MessageDialog dialog(theme, message, icon, buttons);
        return dialog.run(owner, title);
    }

private:
    static constexpr int IdFirstButton = 2400;

    ThemePalette theme_;
    std::wstring message_;
    MessageIcon icon_;
    MessageButtons buttons_;
    dialog_detail::ModalHost host_;
    HFONT font_ = nullptr;
    HBRUSH window_brush_ = nullptr;
    int result_ = IDCANCEL;
    UINT dpi_ = 96;
    RECT text_rect_{};
    struct ButtonSpec { const wchar_t* text; int id; };
    std::vector<ButtonSpec> button_specs_;
    std::vector<HWND> button_windows_;

    MessageDialog(const ThemePalette& theme, const std::wstring& message, MessageIcon icon,
                  MessageButtons buttons)
        : theme_(theme), message_(message), icon_(icon), buttons_(buttons) {
        switch (buttons_) {
            case MessageButtons::Ok:
                button_specs_ = {{L"OK", IDOK}};
                break;
            case MessageButtons::OkCancel:
                button_specs_ = {{L"OK", IDOK}, {L"Cancel", IDCANCEL}};
                break;
            case MessageButtons::YesNo:
                button_specs_ = {{L"Yes", IDYES}, {L"No", IDNO}};
                break;
            case MessageButtons::YesNoCancel:
                button_specs_ = {{L"Yes", IDYES}, {L"No", IDNO}, {L"Cancel", IDCANCEL}};
                break;
        }
        // Dismissing via Esc / the close box maps to the least destructive choice.
        result_ = (buttons_ == MessageButtons::Ok) ? IDOK
                  : (buttons_ == MessageButtons::YesNo) ? IDNO
                                                        : IDCANCEL;
    }

    ~MessageDialog() {
        if (font_ != nullptr) DeleteObject(font_);
        if (window_brush_ != nullptr) DeleteObject(window_brush_);
    }

    HICON system_icon() const {
        switch (icon_) {
            case MessageIcon::Information: return LoadIconW(nullptr, IDI_INFORMATION);
            case MessageIcon::Warning: return LoadIconW(nullptr, IDI_WARNING);
            case MessageIcon::Error: return LoadIconW(nullptr, IDI_ERROR);
            case MessageIcon::Question: return LoadIconW(nullptr, IDI_QUESTION);
            case MessageIcon::None: break;
        }
        return nullptr;
    }

    int run(HWND owner, const std::wstring& title) {
        // Measure the message at the owner's DPI so the dialog is sized to fit.
        dpi_ = owner != nullptr ? GetDpiForWindow(owner) : 96;
        auto s = [this](int v) { return MulDiv(v, static_cast<int>(dpi_), 96); };
        HDC screen = GetDC(nullptr);
        HFONT measure_font = CreateFontW(-MulDiv(9, static_cast<int>(dpi_), 72), 0, 0, 0, FW_NORMAL,
                                         FALSE, FALSE, FALSE, DEFAULT_CHARSET, OUT_DEFAULT_PRECIS,
                                         CLIP_DEFAULT_PRECIS, CLEARTYPE_QUALITY, DEFAULT_PITCH,
                                         L"Segoe UI");
        HGDIOBJ old = SelectObject(screen, measure_font);
        RECT calc{0, 0, s(360), 0};
        DrawTextW(screen, message_.c_str(), -1, &calc, DT_CALCRECT | DT_WORDBREAK | DT_NOPREFIX);
        SelectObject(screen, old);
        DeleteObject(measure_font);
        ReleaseDC(nullptr, screen);

        const int icon_span = (icon_ == MessageIcon::None) ? 0 : 32 + 16;
        int text_w = MulDiv(calc.right - calc.left, 96, static_cast<int>(dpi_));
        int text_h = MulDiv(calc.bottom - calc.top, 96, static_cast<int>(dpi_));
        int content_w = icon_span + std::max(text_w, 180);
        int width = std::max(300, std::min(460, content_w + 40));
        // Buttons need room too (max 3 at 96 + gaps).
        width = std::max(width, 40 + static_cast<int>(button_specs_.size()) * 104);
        int height = 24 + std::max(text_h, (icon_ == MessageIcon::None) ? 0 : 32) + 24 + 30 + 20;
        height = std::max(height, 150);

        HWND hwnd = host_.create(owner, L"XactCopyMessageDlg", title.c_str(), width, height,
                                 [this](HWND h, UINT m, WPARAM w, LPARAM l, bool& handled) {
                                     return proc(h, m, w, l, handled);
                                 });
        if (hwnd == nullptr) return result_;
        apply_window_icons(hwnd);
        set_dark_title_bar(hwnd, theme_.dark);
        host_.run_modal();
        return result_;
    }

    LRESULT proc(HWND hwnd, UINT message, WPARAM wparam, LPARAM lparam, bool& handled) {
        switch (message) {
            case WM_CREATE: {
                window_brush_ = CreateSolidBrush(theme_.window);
                dpi_ = GetDpiForWindow(hwnd);
                auto s = [this](int v) { return MulDiv(v, static_cast<int>(dpi_), 96); };
                font_ = CreateFontW(-MulDiv(9, static_cast<int>(dpi_), 72), 0, 0, 0, FW_NORMAL,
                                    FALSE, FALSE, FALSE, DEFAULT_CHARSET, OUT_DEFAULT_PRECIS,
                                    CLIP_DEFAULT_PRECIS, CLEARTYPE_QUALITY, DEFAULT_PITCH,
                                    L"Segoe UI");
                RECT client;
                GetClientRect(hwnd, &client);
                const int margin = s(16);
                const int btn_w = s(96);
                const int btn_h = s(30);
                const int gap = s(8);
                int x = client.right - margin - btn_w;
                const int y = client.bottom - margin - btn_h;
                // Lay buttons out right-to-left so the first spec ends up leftmost
                // of the group (matching MessageBoxW's ordering).
                for (auto it = button_specs_.rbegin(); it != button_specs_.rend(); ++it) {
                    HWND b = CreateWindowExW(
                        0, L"BUTTON", it->text, WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_OWNERDRAW,
                        x, y, btn_w, btn_h, hwnd,
                        reinterpret_cast<HMENU>(static_cast<INT_PTR>(it->id)),
                        GetModuleHandleW(nullptr), nullptr);
                    SendMessageW(b, WM_SETFONT, reinterpret_cast<WPARAM>(font_), TRUE);
                    button_windows_.push_back(b);
                    x -= btn_w + gap;
                }
                if (!button_windows_.empty()) SetFocus(button_windows_.back());
                text_rect_ = {margin + (icon_ == MessageIcon::None ? 0 : s(48)), margin,
                              client.right - margin, y - s(12)};
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
            case WM_PAINT: {
                PAINTSTRUCT paint{};
                HDC dc = BeginPaint(hwnd, &paint);
                auto s = [this](int v) { return MulDiv(v, static_cast<int>(dpi_), 96); };
                HICON icon = system_icon();
                if (icon != nullptr) {
                    DrawIconEx(dc, s(16), s(16), icon, s(32), s(32), 0, nullptr, DI_NORMAL);
                }
                HGDIOBJ old_font = SelectObject(dc, font_);
                SetBkMode(dc, TRANSPARENT);
                SetTextColor(dc, theme_.text);
                RECT r = text_rect_;
                DrawTextW(dc, message_.c_str(), -1, &r, DT_LEFT | DT_WORDBREAK | DT_NOPREFIX);
                SelectObject(dc, old_font);
                EndPaint(hwnd, &paint);
                handled = true;
                return 0;
            }
            case WM_DRAWITEM:
                themedraw::draw_button(*reinterpret_cast<DRAWITEMSTRUCT*>(lparam), theme_);
                handled = true;
                return TRUE;
            case WM_COMMAND: {
                int id = LOWORD(wparam);
                for (const auto& spec : button_specs_) {
                    if (spec.id == id) {
                        result_ = id;
                        DestroyWindow(hwnd);
                        break;
                    }
                }
                handled = true;
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
    }
};

// Convenience wrapper mirroring the MessageBoxW call shape.
inline int themed_message_box(HWND owner, const ThemePalette& theme, const std::wstring& message,
                              const std::wstring& title,
                              MessageIcon icon = MessageIcon::Information,
                              MessageButtons buttons = MessageButtons::Ok) {
    return MessageDialog::show(owner, theme, title, message, icon, buttons);
}

// ---------------------------------------------------------------------------
// Settings dialog — declarative field table over the full settings surface.
// ---------------------------------------------------------------------------

class SettingsDialog {
public:
    static bool show(HWND owner, AppSettings& settings, const ThemePalette& theme) {
        SettingsDialog dialog(settings, theme);
        return dialog.run(owner);
    }

private:
    enum class FieldKind { Section, Check, Combo, ComboInt, EditInt, EditText };

    struct FieldSpec {
        int page;
        FieldKind kind;
        const char* label;      // section title / field label / checkbox text
        const char* key;        // settings key (nullptr for Section)
        std::vector<const wchar_t*> options;
        std::vector<const char*> values;      // kebab values for Combo
        std::vector<int> int_values;          // for ComboInt
        int min = 0, max = 0, fallback_int = 0;
        const char* fallback_text = "";
        bool fallback_bool = false;
    };

    static const std::vector<const wchar_t*>& page_names() {
        static const std::vector<const wchar_t*> names = {
            L"Appearance",    L"Copy Defaults", L"Performance",
            L"Diagnostics",   L"Verification",  L"Recovery & Startup",
            L"Explorer Integration"};
        return names;
    }

    static const std::vector<FieldSpec>& fields() {
        static const std::vector<FieldSpec> table = [] {
            std::vector<FieldSpec> f;
            auto section = [&f](int page, const char* title) {
                f.push_back({page, FieldKind::Section, title, nullptr, {}, {}, {}});
            };
            auto check = [&f](int page, const char* label, const char* key, bool fallback) {
                FieldSpec spec{page, FieldKind::Check, label, key, {}, {}, {}};
                spec.fallback_bool = fallback;
                f.push_back(spec);
            };
            auto combo = [&f](int page, const char* label, const char* key,
                              std::vector<const wchar_t*> options,
                              std::vector<const char*> values, const char* fallback) {
                FieldSpec spec{page, FieldKind::Combo, label, key, std::move(options),
                               std::move(values), {}};
                spec.fallback_text = fallback;
                f.push_back(spec);
            };
            auto combo_int = [&f](int page, const char* label, const char* key,
                                  std::vector<const wchar_t*> options, std::vector<int> values,
                                  int fallback) {
                FieldSpec spec{page, FieldKind::ComboInt, label, key, std::move(options), {},
                               std::move(values)};
                spec.fallback_int = fallback;
                f.push_back(spec);
            };
            auto edit_int = [&f](int page, const char* label, const char* key, int min, int max,
                                 int fallback) {
                FieldSpec spec{page, FieldKind::EditInt, label, key, {}, {}, {}};
                spec.min = min;
                spec.max = max;
                spec.fallback_int = fallback;
                f.push_back(spec);
            };
            auto edit_text = [&f](int page, const char* label, const char* key,
                                  const char* fallback) {
                FieldSpec spec{page, FieldKind::EditText, label, key, {}, {}, {}};
                spec.fallback_text = fallback;
                f.push_back(spec);
            };

            // --- Page 0: Appearance -------------------------------------
            section(0, "Theme & Accent");
            combo(0, "Mode", "Theme", {L"Dark", L"System", L"Classic"},
                  {"dark", "system", "classic"}, "dark");
            combo(0, "Accent source", "AccentColorMode", {L"Auto", L"System", L"Custom"},
                  {"auto", "system", "custom"}, "auto");
            edit_text(0, "Custom accent color", "AccentColorHex", "#5A78C8");
            combo(0, "Window chrome", "WindowChromeMode",
                  {L"Themed title bar", L"Standard title bar"}, {"themed", "standard"}, "themed");
            section(0, "Layout");
            combo(0, "UI density", "UiDensity", {L"Compact", L"Normal", L"Comfortable"},
                  {"compact", "normal", "comfortable"}, "normal");
            combo_int(0, "UI scale", "UiScalePercent", {L"90%", L"100%", L"110%", L"125%"},
                      {90, 100, 110, 125}, 100);
            section(0, "Operations Log");
            edit_text(0, "Log font", "LogFontFamily", "Consolas");
            edit_int(0, "Log size (pt)", "LogFontSizePoints", 7, 20, 9);
            check(0, "Color-code log by severity", "UiColorizeLogBySeverity", true);
            section(0, "Grid Appearance");
            check(0, "Use alternating grid rows", "GridAlternatingRows", true);
            edit_int(0, "Grid row height", "GridRowHeight", 18, 48, 24);
            combo(0, "Grid header style", "GridHeaderStyle",
                  {L"Default", L"Minimal", L"Prominent"}, {"default", "minimal", "prominent"},
                  "default");
            section(0, "Status & Progress");
            check(0, "Show buffer status row", "ShowBufferStatusRow", true);
            check(0, "Show rescue status row", "ShowRescueStatusRow", true);
            check(0, "Show percentage on progress bars", "ProgressBarShowPercentage", false);
            combo(0, "Progress bar style", "ProgressBarStyle", {L"Thin", L"Standard", L"Thick"},
                  {"thin", "standard", "thick"}, "standard");

            // --- Page 1: Copy Defaults ----------------------------------
            section(1, "Default Run Behavior");
            check(1, "Resume from journal", "DefaultResumeFromJournal", true);
            check(1, "Salvage unreadable blocks", "DefaultSalvageUnreadableBlocks", true);
            check(1, "Continue on file errors", "DefaultContinueOnFileError", true);
            check(1, "Preserve source timestamps", "DefaultPreserveTimestamps", true);
            check(1, "Copy empty directories", "DefaultCopyEmptyDirectories", true);
            check(1, "Wait forever for source/destination", "DefaultWaitForMediaAvailability",
                  false);
            check(1, "Wait for lock/contention release", "DefaultWaitForFileLockRelease", false);
            check(1, "Treat Access Denied as contention", "DefaultTreatAccessDeniedAsContention",
                  false);
            section(1, "Policies");
            combo(1, "Overwrite policy", "DefaultOverwritePolicy",
                  {L"Overwrite existing", L"Skip existing", L"Overwrite if newer", L"Always ask"},
                  {"overwrite", "skip-existing", "overwrite-if-newer", "ask"}, "overwrite");
            combo(1, "Symlink handling", "DefaultSymlinkHandling",
                  {L"Skip symbolic links", L"Follow symbolic links"}, {"skip", "follow"}, "skip");
            combo(1, "Salvage fill pattern", "DefaultSalvageFillPattern",
                  {L"Zero-fill", L"0xFF-fill", L"Random-fill"}, {"zero", "ones", "random"},
                  "zero");
            section(1, "Bad Range Map");
            check(1, "Use bad-range map when available", "DefaultUseBadRangeMap", true);
            check(1, "Skip known bad ranges during copy", "DefaultSkipKnownBadRanges", true);
            check(1, "Update map from scan/copy runs", "DefaultUpdateBadRangeMapFromRun", true);
            check(1, "Experimental raw disk scan backend", "DefaultUseExperimentalRawDiskScan",
                  false);
            edit_int(1, "Map max age (days, 0=never)", "DefaultBadRangeMapMaxAgeDays", 0, 3650,
                     30);

            // --- Page 2: Performance ------------------------------------
            section(2, "Transfer Tuning");
            check(2, "Enable adaptive buffer by default", "DefaultUseAdaptiveBuffer", false);
            combo(2, "Transfer engine", "DefaultTransferEnginePolicy",
                  {L"Auto", L"Managed rescue", L"Native fast"}, {"auto", "managed", "native"},
                  "auto");
            combo(2, "Scan profile", "DefaultScanPerformanceProfile",
                  {L"Auto fast scan", L"Fast health scan", L"Precise bad-range scan"},
                  {"auto", "fast", "precise"}, "auto");
            combo(2, "Worker priority", "DefaultWorkerProcessPriorityClass",
                  {L"Idle", L"Below normal", L"Normal", L"Above normal", L"High"},
                  {"Idle", "BelowNormal", "Normal", "AboveNormal", "High"}, "Normal");
            edit_int(2, "Manual buffer MB", "DefaultBufferSizeMb", 1, 256, 4);
            edit_int(2, "Max retries", "DefaultMaxRetries", 0, 1000, 12);
            edit_int(2, "Operation timeout (sec)", "DefaultOperationTimeoutSeconds", 1, 3600, 10);
            edit_int(2, "Per-file timeout (sec, 0=off)", "DefaultPerFileTimeoutSeconds", 0, 86400,
                     0);
            edit_int(2, "Max throughput (MB/s, 0=off)", "DefaultMaxThroughputMbPerSecond", 0,
                     4096, 0);
            edit_int(2, "Small-file workers (0=auto)", "DefaultParallelSmallFileWorkers", 0, 64,
                     0);
            edit_int(2, "Scan workers (0=auto)", "DefaultParallelScanWorkers", 0, 64, 0);
            edit_int(2, "Small-file threshold KB", "DefaultSmallFileThresholdKb", 4, 1048576,
                     256);
            section(2, "Fragile Media Guard");
            check(2, "Enable fragile media mode by default", "DefaultFragileMediaMode", false);
            check(2, "Skip file on first read error", "DefaultSkipFileOnFirstReadError", true);
            check(2, "Persist skips across resume", "DefaultPersistFragileSkipsAcrossResume",
                  true);
            edit_int(2, "Failure window (sec)", "DefaultFragileFailureWindowSeconds", 1, 3600,
                     20);
            edit_int(2, "Failure threshold (files)", "DefaultFragileFailureThreshold", 1, 1000, 3);
            edit_int(2, "Cooldown (sec, 0=off)", "DefaultFragileCooldownSeconds", 0, 600, 6);
            section(2, "Contention & Source Mutation");
            edit_int(2, "Lock probe interval (ms)", "DefaultLockContentionProbeIntervalMs", 100,
                     10000, 500);
            combo(2, "Source mutation policy", "DefaultSourceMutationPolicy",
                  {L"Fail file", L"Skip file", L"Wait for reappearance"},
                  {"fail-file", "skip-file", "wait-for-reappearance"}, "fail-file");
            section(2, "Rescue Engine Tuning");
            edit_int(2, "FastScan chunk KB (0=auto)", "DefaultRescueFastScanChunkKb", 0, 262144,
                     0);
            edit_int(2, "TrimSweep chunk KB (0=auto)", "DefaultRescueTrimChunkKb", 0, 262144, 0);
            edit_int(2, "Scrape chunk KB (0=auto)", "DefaultRescueScrapeChunkKb", 0, 262144, 0);
            edit_int(2, "RetryBad chunk KB (0=auto)", "DefaultRescueRetryChunkKb", 0, 262144, 0);
            edit_int(2, "Split minimum KB (0=auto)", "DefaultRescueSplitMinimumKb", 0, 65536, 0);
            edit_int(2, "FastScan retries", "DefaultRescueFastScanRetries", 0, 1000, 0);
            edit_int(2, "TrimSweep retries", "DefaultRescueTrimRetries", 0, 1000, 1);
            edit_int(2, "Scrape retries", "DefaultRescueScrapeRetries", 0, 1000, 2);

            // --- Page 3: Diagnostics ------------------------------------
            section(3, "Worker Telemetry");
            combo(3, "Worker telemetry profile", "WorkerTelemetryProfile",
                  {L"Normal", L"Verbose", L"Debug"}, {"normal", "verbose", "debug"}, "normal");
            edit_int(3, "Progress interval (ms)", "WorkerProgressIntervalMs", 20, 1000, 75);
            edit_int(3, "Log rate cap (/sec, 0=off)", "WorkerMaxLogsPerSecond", 0, 5000, 100);
            section(3, "UI Diagnostics");
            check(3, "Show UI diagnostics strip", "UiShowDiagnostics", true);
            edit_int(3, "Diagnostics refresh (ms)", "UiDiagnosticsRefreshMs", 100, 5000, 250);
            edit_int(3, "Virtual log max lines", "UiMaxLogLines", 1000, 1000000, 50000);
            section(3, "Journal Storage");
            check(3, "Discard backups after a run completes", "CompactJournalsOnCompletion", true);
            edit_int(3, "Delete journals after (days, 0=keep)", "JournalRetentionDays", 0, 3650, 30);
            edit_int(3, "Always keep newest journals", "JournalKeepMinimum", 0, 1000, 10);

            // --- Page 4: Verification -----------------------------------
            section(4, "Verification Defaults");
            check(4, "Enable verification by default", "DefaultVerifyAfterCopy", false);
            combo(4, "Verification mode", "DefaultVerificationMode",
                  {L"Full hash", L"Sampled hash"}, {"full", "sampled"}, "full");
            combo(4, "Hash algorithm", "DefaultVerificationHashAlgorithm",
                  {L"SHA-256", L"SHA-512"}, {"sha256", "sha512"}, "sha256");
            section(4, "Sampled Mode");
            edit_int(4, "Sample chunk size (KB)", "DefaultSampleVerificationChunkKb", 32, 4096,
                     128);
            edit_int(4, "Sample count", "DefaultSampleVerificationChunkCount", 1, 64, 3);

            // Update settings live in the About dialog, not here.

            // --- Page 5: Recovery & Startup -----------------------------
            section(5, "Startup");
            check(5, "Auto-start on next logon after interruption", "EnableRecoveryAutostart",
                  true);
            check(5, "Automatically run queued jobs on startup", "AutoRunQueuedJobsOnStartup",
                  false);
            section(5, "Interrupted Run Handling");
            check(5, "Prompt to resume interrupted runs", "PromptResumeAfterCrash", true);
            check(5, "Auto-resume interrupted runs", "AutoResumeAfterCrash", false);
            check(5, "Keep prompting until resolved", "KeepResumePromptUntilResolved", true);
            edit_int(5, "Heartbeat write interval (sec)", "RecoveryTouchIntervalSeconds", 1, 60,
                     2);

            // --- Page 6: Explorer Integration ---------------------------
            section(6, "Context Menu");
            check(6, "Enable Explorer context menu integration", "EnableExplorerContextMenu",
                  false);
            return f;
        }();
        return table;
    }

    static constexpr int IdNavList = 2001;
    static constexpr int IdSaveButton = 2002;
    static constexpr int IdCancelButton = 2003;
    static constexpr int IdFieldBase = 2100;

    AppSettings& settings_;
    ThemePalette theme_;
    dialog_detail::ModalHost host_;
    HFONT font_ = nullptr;
    HFONT section_font_ = nullptr;
    HBRUSH window_brush_ = nullptr;
    HBRUSH edit_brush_ = nullptr;
    HBRUSH panel_brush_ = nullptr;
    HWND nav_list_ = nullptr;
    int current_page_ = 0;
    bool saved_ = false;

    // Staged values (all pages), keyed by settings key.
    std::map<std::string, bool> staged_bools_;
    std::map<std::string, int> staged_ints_;
    std::map<std::string, std::string> staged_strings_;

    // Live controls for the current page: field-table index -> hwnd.
    std::map<int, HWND> page_controls_;
    std::vector<HWND> page_labels_;
    std::map<int, bool> page_checks_; // field index -> checked (owner-draw state)

    SettingsDialog(AppSettings& settings, const ThemePalette& theme)
        : settings_(settings), theme_(theme) {}

    ~SettingsDialog() {
        if (font_ != nullptr) DeleteObject(font_);
        if (section_font_ != nullptr) DeleteObject(section_font_);
        if (window_brush_ != nullptr) DeleteObject(window_brush_);
        if (edit_brush_ != nullptr) DeleteObject(edit_brush_);
        if (panel_brush_ != nullptr) DeleteObject(panel_brush_);
    }

    bool run(HWND owner) {
        stage_from_settings();
        HWND hwnd = host_.create(
            owner, L"XactCopySettingsDlg", L"Settings", 900, 720,
            [this](HWND h, UINT m, WPARAM w, LPARAM l, bool& handled) {
                return proc(h, m, w, l, handled);
            });
        if (hwnd == nullptr) return false;
        apply_window_icons(hwnd);
        set_dark_title_bar(hwnd, theme_.dark);
        host_.run_modal();
        return saved_;
    }

    void stage_from_settings() {
        for (const auto& field : fields()) {
            if (field.key == nullptr) continue;
            switch (field.kind) {
                case FieldKind::Check:
                    staged_bools_[field.key] = settings_.get_bool(field.key, field.fallback_bool);
                    break;
                case FieldKind::Combo:
                    staged_strings_[field.key] =
                        settings_.get_string(field.key, field.fallback_text);
                    break;
                case FieldKind::ComboInt:
                    staged_ints_[field.key] = settings_.get_int(field.key, INT32_MIN / 2,
                                                               INT32_MAX / 2, field.fallback_int);
                    break;
                case FieldKind::EditInt:
                    staged_ints_[field.key] =
                        settings_.get_int(field.key, field.min, field.max, field.fallback_int);
                    break;
                case FieldKind::EditText:
                    staged_strings_[field.key] =
                        settings_.get_string(field.key, field.fallback_text);
                    break;
                default:
                    break;
            }
        }
    }

    void save_all() {
        harvest_current_page();
        for (const auto& field : fields()) {
            if (field.key == nullptr) continue;
            switch (field.kind) {
                case FieldKind::Check:
                    settings_.set_bool(field.key, staged_bools_[field.key]);
                    break;
                case FieldKind::Combo:
                case FieldKind::EditText:
                    settings_.set_string(field.key, staged_strings_[field.key]);
                    break;
                case FieldKind::ComboInt:
                case FieldKind::EditInt:
                    settings_.set_int(field.key, staged_ints_[field.key]);
                    break;
                default:
                    break;
            }
        }
        settings_.save();
        saved_ = true;
    }

    // ---- Page building -----------------------------------------------------

    void destroy_page_controls() {
        for (auto& [index, control] : page_controls_) DestroyWindow(control);
        for (HWND label : page_labels_) DestroyWindow(label);
        page_controls_.clear();
        page_labels_.clear();
        page_checks_.clear();
    }

    void harvest_current_page() {
        const auto& table = fields();
        for (const auto& [index, control] : page_controls_) {
            const FieldSpec& field = table[static_cast<std::size_t>(index)];
            switch (field.kind) {
                case FieldKind::Check:
                    staged_bools_[field.key] = page_checks_[index];
                    break;
                case FieldKind::Combo: {
                    int selection = static_cast<int>(SendMessageW(control, CB_GETCURSEL, 0, 0));
                    if (selection >= 0 && selection < static_cast<int>(field.values.size())) {
                        staged_strings_[field.key] =
                            field.values[static_cast<std::size_t>(selection)];
                    }
                    break;
                }
                case FieldKind::ComboInt: {
                    int selection = static_cast<int>(SendMessageW(control, CB_GETCURSEL, 0, 0));
                    if (selection >= 0 && selection < static_cast<int>(field.int_values.size())) {
                        staged_ints_[field.key] =
                            field.int_values[static_cast<std::size_t>(selection)];
                    }
                    break;
                }
                case FieldKind::EditInt: {
                    wchar_t buffer[32];
                    GetWindowTextW(control, buffer, 32);
                    int value = _wtoi(buffer);
                    staged_ints_[field.key] = std::clamp(value, field.min, field.max);
                    break;
                }
                case FieldKind::EditText: {
                    wchar_t buffer[512];
                    int length = GetWindowTextW(control, buffer, 512);
                    staged_strings_[field.key] =
                        storage::fsutil::wide_to_utf8(std::wstring(buffer, length));
                    break;
                }
                default:
                    break;
            }
        }
    }

    void build_page(int page) {
        destroy_page_controls();
        current_page_ = page;

        HWND hwnd = host_.hwnd();
        UINT dpi = GetDpiForWindow(hwnd);
        auto scale = [dpi](int value) { return MulDiv(value, static_cast<int>(dpi), 96); };

        RECT client;
        GetClientRect(hwnd, &client);
        const int nav_width = scale(180);
        const int margin = scale(16);
        const int row_height = scale(24);
        const int row_gap = scale(5);
        const int section_gap = scale(10); // breathing room above a section header
        const int footer_height = scale(48);
        const int column_gap = scale(28);
        const int label_width = scale(176);
        const int control_gap = scale(8);
        const int top = margin + scale(2);
        const int content_left = nav_width + margin;
        const int content_width = client.right - content_left - margin;
        const int column_width = (content_width - column_gap) / 2;
        const int usable_bottom = client.bottom - footer_height;
        auto column_x = [&](int column) {
            return content_left + column * (column_width + column_gap);
        };

        const auto& table = fields();

        // First pass: pack whole sections into two columns so a section's fields
        // never get orphaned from their header across the column break.
        std::map<int, POINT> pos;
        int column = 0;
        int y = top;
        for (int index = 0; index < static_cast<int>(table.size()); ++index) {
            const FieldSpec& field = table[static_cast<std::size_t>(index)];
            if (field.page != page) continue;

            if (field.kind == FieldKind::Section) {
                int rows = 1; // header
                for (int j = index + 1; j < static_cast<int>(table.size()); ++j) {
                    if (table[static_cast<std::size_t>(j)].page != page) continue;
                    if (table[static_cast<std::size_t>(j)].kind == FieldKind::Section) break;
                    ++rows;
                }
                const bool at_column_top = (y == top);
                const int block_height = rows * (row_height + row_gap) + (at_column_top ? 0 : section_gap);
                if (!at_column_top && column == 0 && y + block_height > usable_bottom) {
                    column = 1;
                    y = top;
                } else if (!at_column_top) {
                    y += section_gap;
                }
            }

            pos[index] = POINT{column_x(column), y};
            y += row_height + row_gap;
        }

        // Second pass: create the controls at the packed positions.
        for (int index = 0; index < static_cast<int>(table.size()); ++index) {
            const FieldSpec& field = table[static_cast<std::size_t>(index)];
            if (field.page != page) continue;

            const int x = pos[index].x;
            y = pos[index].y;
            int control_id = IdFieldBase + index;

            if (field.kind == FieldKind::Section) {
                HWND label = CreateWindowExW(
                    0, L"STATIC", storage::fsutil::utf8_to_wide(field.label).c_str(),
                    WS_CHILD | WS_VISIBLE | SS_NOPREFIX | SS_CENTERIMAGE, x, y, column_width,
                    row_height, hwnd, nullptr, GetModuleHandleW(nullptr), nullptr);
                SendMessageW(label, WM_SETFONT, reinterpret_cast<WPARAM>(section_font_), TRUE);
                page_labels_.push_back(label);
                continue;
            }

            if (field.kind == FieldKind::Check) {
                HWND check = CreateWindowExW(
                    0, L"BUTTON", storage::fsutil::utf8_to_wide(field.label).c_str(),
                    WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_OWNERDRAW, x, y, column_width, row_height,
                    hwnd, reinterpret_cast<HMENU>(static_cast<INT_PTR>(control_id)),
                    GetModuleHandleW(nullptr), nullptr);
                SendMessageW(check, WM_SETFONT, reinterpret_cast<WPARAM>(font_), TRUE);
                page_controls_[index] = check;
                page_checks_[index] = staged_bools_[field.key];
                continue;
            }

            // Labeled field: label + control on a shared baseline.
            HWND label = CreateWindowExW(
                0, L"STATIC", storage::fsutil::utf8_to_wide(field.label).c_str(),
                WS_CHILD | WS_VISIBLE | SS_ENDELLIPSIS | SS_NOPREFIX | SS_CENTERIMAGE, x, y,
                label_width, row_height, hwnd, nullptr, GetModuleHandleW(nullptr), nullptr);
            SendMessageW(label, WM_SETFONT, reinterpret_cast<WPARAM>(font_), TRUE);
            page_labels_.push_back(label);

            int control_x = x + label_width + control_gap;
            int control_width = column_width - label_width - control_gap;
            int control_height = row_height - scale(4);
            int control_y = y + (row_height - control_height) / 2;

            if (field.kind == FieldKind::Combo || field.kind == FieldKind::ComboInt) {
                HWND combo = CreateWindowExW(
                    0, L"COMBOBOX", L"",
                    WS_CHILD | WS_VISIBLE | WS_TABSTOP | CBS_DROPDOWNLIST | CBS_OWNERDRAWFIXED |
                        CBS_HASSTRINGS,
                    control_x, control_y, control_width, row_height * 10, hwnd,
                    reinterpret_cast<HMENU>(static_cast<INT_PTR>(control_id)),
                    GetModuleHandleW(nullptr), nullptr);
                SendMessageW(combo, WM_SETFONT, reinterpret_cast<WPARAM>(font_), TRUE);
                for (const wchar_t* option : field.options) {
                    SendMessageW(combo, CB_ADDSTRING, 0, reinterpret_cast<LPARAM>(option));
                }
                int selection = 0;
                if (field.kind == FieldKind::Combo) {
                    const std::string& current = staged_strings_[field.key];
                    for (std::size_t i = 0; i < field.values.size(); ++i) {
                        if (models::detail::equals_ignore_case(field.values[i], current)) {
                            selection = static_cast<int>(i);
                            break;
                        }
                    }
                } else {
                    int current = staged_ints_[field.key];
                    int best_distance = INT32_MAX;
                    for (std::size_t i = 0; i < field.int_values.size(); ++i) {
                        int distance = std::abs(field.int_values[i] - current);
                        if (distance < best_distance) {
                            best_distance = distance;
                            selection = static_cast<int>(i);
                        }
                    }
                }
                SendMessageW(combo, CB_SETCURSEL, selection, 0);
                apply_combo_theme(combo, theme_.dark);
                page_controls_[index] = combo;
                continue;
            }

            std::wstring initial =
                field.kind == FieldKind::EditInt
                    ? std::to_wstring(staged_ints_[field.key])
                    : storage::fsutil::utf8_to_wide(staged_strings_[field.key]);
            HWND edit = CreateWindowExW(
                0, L"EDIT", initial.c_str(),
                WS_CHILD | WS_VISIBLE | WS_TABSTOP | ES_AUTOHSCROLL | WS_BORDER |
                    (field.kind == FieldKind::EditInt ? ES_NUMBER : 0),
                control_x, control_y, control_width, control_height, hwnd,
                reinterpret_cast<HMENU>(static_cast<INT_PTR>(control_id)),
                GetModuleHandleW(nullptr), nullptr);
            SendMessageW(edit, WM_SETFONT, reinterpret_cast<WPARAM>(font_), TRUE);
            apply_control_theme(edit, theme_.dark);
            page_controls_[index] = edit;
        }
        InvalidateRect(hwnd, nullptr, TRUE);
    }

    void on_create() {
        HWND hwnd = host_.hwnd();
        window_brush_ = CreateSolidBrush(theme_.window);
        edit_brush_ = CreateSolidBrush(theme_.edit);
        panel_brush_ = CreateSolidBrush(theme_.panel);
        UINT dpi = GetDpiForWindow(hwnd);
        auto scale = [dpi](int value) { return MulDiv(value, static_cast<int>(dpi), 96); };
        font_ = CreateFontW(-MulDiv(9, static_cast<int>(dpi), 72), 0, 0, 0, FW_NORMAL, FALSE,
                            FALSE, FALSE, DEFAULT_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS,
                            CLEARTYPE_QUALITY, DEFAULT_PITCH, L"Segoe UI");
        section_font_ = CreateFontW(-MulDiv(10, static_cast<int>(dpi), 72), 0, 0, 0, FW_SEMIBOLD,
                                    FALSE, FALSE, FALSE, DEFAULT_CHARSET, OUT_DEFAULT_PRECIS,
                                    CLIP_DEFAULT_PRECIS, CLEARTYPE_QUALITY, DEFAULT_PITCH,
                                    L"Segoe UI");

        RECT client;
        GetClientRect(hwnd, &client);
        nav_list_ = CreateWindowExW(
            0, L"LISTBOX", L"",
            WS_CHILD | WS_VISIBLE | WS_TABSTOP | LBS_NOTIFY | LBS_OWNERDRAWFIXED |
                LBS_HASSTRINGS | LBS_NOINTEGRALHEIGHT,
            scale(8), scale(12), scale(172), client.bottom - scale(24), hwnd,
            reinterpret_cast<HMENU>(static_cast<INT_PTR>(IdNavList)), GetModuleHandleW(nullptr),
            nullptr);
        SendMessageW(nav_list_, WM_SETFONT, reinterpret_cast<WPARAM>(font_), TRUE);
        SendMessageW(nav_list_, LB_SETITEMHEIGHT, 0, scale(32));
        for (const wchar_t* name : page_names()) {
            SendMessageW(nav_list_, LB_ADDSTRING, 0, reinterpret_cast<LPARAM>(name));
        }
        SendMessageW(nav_list_, LB_SETCURSEL, 0, 0);
        apply_control_theme(nav_list_, theme_.dark);

        auto make_button = [&](int id, const wchar_t* text, int x) {
            HWND button = CreateWindowExW(
                0, L"BUTTON", text, WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_OWNERDRAW, x,
                client.bottom - scale(38), scale(90), scale(28), hwnd,
                reinterpret_cast<HMENU>(static_cast<INT_PTR>(id)), GetModuleHandleW(nullptr),
                nullptr);
            SendMessageW(button, WM_SETFONT, reinterpret_cast<WPARAM>(font_), TRUE);
        };
        make_button(IdSaveButton, L"Save", client.right - scale(200));
        make_button(IdCancelButton, L"Cancel", client.right - scale(100));

        build_page(0);
    }

    void draw_nav_item(const DRAWITEMSTRUCT& draw) {
        const bool selected = (draw.itemState & ODS_SELECTED) != 0;
        UINT dpi = GetDpiForWindow(nav_list_);
        auto s = [dpi](int v) { return MulDiv(v, static_cast<int>(dpi), 96); };
        RECT r = draw.rcItem;
        // Selected rows get an accent-tinted fill; others sit flush on the rail.
        const COLORREF fill = selected ? blend_color(theme_.panel, theme_.accent, 26) : theme_.panel;
        themedraw::fill_rect(draw.hDC, r, fill);
        if (selected) {
            RECT bar{r.left, r.top + s(4), r.left + s(3), r.bottom - s(4)};
            themedraw::fill_rect(draw.hDC, bar, theme_.accent);
        }
        if (draw.itemID == static_cast<UINT>(-1)) return;
        wchar_t text[64]{};
        SendMessageW(nav_list_, LB_GETTEXT, draw.itemID, reinterpret_cast<LPARAM>(text));
        HGDIOBJ old_font = SelectObject(draw.hDC, selected ? section_font_ : font_);
        SetBkMode(draw.hDC, TRANSPARENT);
        SetTextColor(draw.hDC, selected ? theme_.text : theme_.muted_text);
        RECT text_rect = r;
        text_rect.left += s(16);
        DrawTextW(draw.hDC, text, -1, &text_rect,
                  DT_LEFT | DT_VCENTER | DT_SINGLELINE | DT_NOPREFIX);
        SelectObject(draw.hDC, old_font);
    }

    LRESULT proc(HWND hwnd, UINT message, WPARAM wparam, LPARAM lparam, bool& handled) {
        switch (message) {
            case WM_CREATE:
                on_create();
                handled = true;
                return 0;
            case WM_ERASEBKGND: {
                HDC dc = reinterpret_cast<HDC>(wparam);
                RECT client;
                GetClientRect(hwnd, &client);
                themedraw::fill_rect(dc, client, theme_.window);
                // Left navigation rail: a panel-colored sidebar with a 1px edge.
                UINT dpi = GetDpiForWindow(hwnd);
                auto s = [dpi](int v) { return MulDiv(v, static_cast<int>(dpi), 96); };
                RECT rail{0, 0, s(188), client.bottom};
                themedraw::fill_rect(dc, rail, theme_.panel);
                RECT sep{s(188) - 1, 0, s(188), client.bottom};
                themedraw::fill_rect(dc, sep, theme_.border);
                handled = true;
                return 1;
            }
            case WM_CTLCOLORSTATIC: {
                HDC dc = reinterpret_cast<HDC>(wparam);
                // Section headers use the accent color; regular labels the text color.
                HFONT control_font = reinterpret_cast<HFONT>(
                    SendMessageW(reinterpret_cast<HWND>(lparam), WM_GETFONT, 0, 0));
                SetTextColor(dc, control_font == section_font_ ? theme_.accent : theme_.text);
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
            case WM_CTLCOLORLISTBOX: {
                // The only listbox is the navigation rail — paint it panel-colored.
                HDC dc = reinterpret_cast<HDC>(wparam);
                SetTextColor(dc, theme_.text);
                SetBkColor(dc, theme_.panel);
                handled = true;
                return reinterpret_cast<LRESULT>(panel_brush_);
            }
            case WM_DRAWITEM: {
                const auto& draw = *reinterpret_cast<DRAWITEMSTRUCT*>(lparam);
                int id = static_cast<int>(draw.CtlID);
                if (id == IdNavList) {
                    draw_nav_item(draw);
                } else if (draw.CtlType == ODT_COMBOBOX) {
                    themedraw::draw_combo_item(draw, theme_);
                } else if (id == IdSaveButton || id == IdCancelButton) {
                    themedraw::draw_button(draw, theme_);
                } else if (id >= IdFieldBase) {
                    int index = id - IdFieldBase;
                    themedraw::draw_checkbox(draw, theme_, page_checks_[index]);
                }
                handled = true;
                return TRUE;
            }
            case WM_COMMAND: {
                int id = LOWORD(wparam);
                if (id == IdNavList && HIWORD(wparam) == LBN_SELCHANGE) {
                    int page = static_cast<int>(SendMessageW(nav_list_, LB_GETCURSEL, 0, 0));
                    if (page >= 0 && page != current_page_) {
                        harvest_current_page();
                        build_page(page);
                    }
                    handled = true;
                    return 0;
                }
                if (id >= IdFieldBase && HIWORD(wparam) == BN_CLICKED) {
                    int index = id - IdFieldBase;
                    if (page_checks_.count(index) != 0) {
                        page_checks_[index] = !page_checks_[index];
                        InvalidateRect(page_controls_[index], nullptr, FALSE);
                        handled = true;
                        return 0;
                    }
                }
                if (id == IdSaveButton) {
                    save_all();
                    DestroyWindow(hwnd);
                    handled = true;
                    return 0;
                }
                if (id == IdCancelButton) {
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
                // No PostQuitMessage: the modal loop exits on IsWindow();
                // posting WM_QUIT here would leak into the owner's loop.
                handled = true;
                return 0;
            default:
                return 0;
        }
    }
};

// ---------------------------------------------------------------------------
// About dialog
// ---------------------------------------------------------------------------

class AboutDialog {
public:
    // `check_updates_command` (if non-zero) is posted to `owner` when the user
    // clicks "Check for Updates...", so About can trigger the main window's check.
    // `settings` backs the "check automatically" toggle (update settings live
    // here rather than in the Settings dialog).
    static void show(HWND owner, const ThemePalette& theme, AppSettings& settings,
                     UINT check_updates_command = 0) {
        AboutDialog dialog(theme, settings, check_updates_command);
        dialog.run(owner);
    }

private:
    static constexpr int IdOkButton = 2101;
    static constexpr int IdCheckButton = 2102;
    static constexpr int IdAutoUpdateCheck = 2103;

    ThemePalette theme_;
    AppSettings& settings_;
    bool auto_update_ = false;
    UINT check_updates_command_ = 0;
    HWND owner_ = nullptr;
    dialog_detail::ModalHost host_;
    HFONT font_ = nullptr;
    HFONT title_font_ = nullptr;
    HBRUSH window_brush_ = nullptr;

    HWND icon_ = nullptr;
    HWND title_ = nullptr;
    HWND description_ = nullptr;
    HWND metadata_[4] = {};
    HWND components_title_ = nullptr;
    HWND components_ = nullptr;
    HWND auto_update_check_ = nullptr;
    HWND check_btn_ = nullptr;
    HWND ok_btn_ = nullptr;

    AboutDialog(const ThemePalette& theme, AppSettings& settings, UINT check_updates_command)
        : theme_(theme), settings_(settings), check_updates_command_(check_updates_command) {
        auto_update_ = settings_.get_bool("CheckUpdatesOnLaunch", false);
    }

    ~AboutDialog() {
        if (font_ != nullptr) DeleteObject(font_);
        if (title_font_ != nullptr) DeleteObject(title_font_);
        if (window_brush_ != nullptr) DeleteObject(window_brush_);
    }

    void run(HWND owner) {
        owner_ = owner;
        HWND hwnd = host_.create(
            owner, L"XactCopyAboutDlg", L"About XactCopy", 480, 400,
            [this](HWND h, UINT m, WPARAM w, LPARAM l, bool& handled) {
                return proc(h, m, w, l, handled);
            });
        if (hwnd == nullptr) return;
        apply_window_icons(hwnd);
        set_dark_title_bar(hwnd, theme_.dark);
        host_.run_modal();
    }

    static std::wstring build_stamp() {
        // __DATE__ __TIME__ captured at compile time (mirrors Axiom's About).
        return storage::fsutil::utf8_to_wide(std::string(__DATE__) + " " + __TIME__);
    }

    LRESULT proc(HWND hwnd, UINT message, WPARAM wparam, LPARAM lparam, bool& handled) {
        switch (message) {
            case WM_CREATE: {
                window_brush_ = CreateSolidBrush(theme_.window);
                UINT dpi = GetDpiForWindow(hwnd);
                font_ = CreateFontW(-MulDiv(9, static_cast<int>(dpi), 72), 0, 0, 0, FW_NORMAL,
                                    FALSE, FALSE, FALSE, DEFAULT_CHARSET, OUT_DEFAULT_PRECIS,
                                    CLIP_DEFAULT_PRECIS, CLEARTYPE_QUALITY, DEFAULT_PITCH,
                                    L"Segoe UI");
                title_font_ = CreateFontW(-MulDiv(22, static_cast<int>(dpi), 72), 0, 0, 0,
                                          FW_SEMIBOLD, FALSE, FALSE, FALSE, DEFAULT_CHARSET,
                                          OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, CLEARTYPE_QUALITY,
                                          DEFAULT_PITCH, L"Segoe UI");
                HINSTANCE inst = GetModuleHandleW(nullptr);
                auto make_static = [&](const wchar_t* text, DWORD extra) {
                    return CreateWindowExW(0, L"STATIC", text,
                                           WS_CHILD | WS_VISIBLE | WS_CLIPSIBLINGS | extra, 0, 0, 0,
                                           0, hwnd, nullptr, inst, nullptr);
                };
                auto make_button = [&](const wchar_t* text, int id) {
                    return CreateWindowExW(0, L"BUTTON", text,
                                           WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_OWNERDRAW, 0, 0, 0,
                                           0, hwnd, reinterpret_cast<HMENU>(static_cast<INT_PTR>(id)),
                                           inst, nullptr);
                };

                icon_ = CreateWindowExW(0, L"STATIC", nullptr,
                                        WS_CHILD | WS_VISIBLE | WS_CLIPSIBLINGS | SS_ICON |
                                            SS_CENTERIMAGE,
                                        0, 0, 0, 0, hwnd, nullptr, inst, nullptr);
                SendMessageW(icon_, STM_SETICON,
                             reinterpret_cast<WPARAM>(load_app_icon(MulDiv(48, static_cast<int>(dpi),
                                                                           96),
                                                                    MulDiv(48, static_cast<int>(dpi),
                                                                           96))),
                             0);
                title_ = make_static(L"XactCopy", SS_NOPREFIX);
                description_ = make_static(
                    L"Resilient file mover and bad-block scanner for unstable media.",
                    SS_NOPREFIX);
                const std::wstring meta[4] = {
                    L"Version:  " + storage::fsutil::utf8_to_wide(NativeAppVersion),
                    L"Build:  " + build_stamp(),
                    L"Author:  Wimukthi Bandara",
                    L"Licence:  GNU GPL v3.0",
                };
                for (int i = 0; i < 4; ++i) metadata_[i] = make_static(meta[i].c_str(), SS_NOPREFIX);
                components_title_ = make_static(L"Components", SS_NOPREFIX);
                components_ = make_static(
                    L"Windows CNG (BCrypt) \x2014 SHA-256 / HMAC integrity\r\n"
                    L"WinHTTP \x2014 update checks and downloads\r\n"
                    L"Common Controls v6 \x2014 native themed UI\r\n"
                    L"Named-pipe IPC \x2014 supervised copy worker",
                    SS_NOPREFIX);

                auto_update_check_ = CreateWindowExW(
                    0, L"BUTTON", L"Check automatically",
                    WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_OWNERDRAW, 0, 0, 0, 0, hwnd,
                    reinterpret_cast<HMENU>(static_cast<INT_PTR>(IdAutoUpdateCheck)), inst,
                    nullptr);
                check_btn_ = make_button(L"Check for Updates...", IdCheckButton);
                ok_btn_ = make_button(L"OK", IdOkButton);
                if (check_updates_command_ == 0) EnableWindow(check_btn_, FALSE);

                for (HWND h : {title_, description_, metadata_[0], metadata_[1], metadata_[2],
                               metadata_[3], components_title_, components_, auto_update_check_,
                               check_btn_, ok_btn_}) {
                    SendMessageW(h, WM_SETFONT,
                                 reinterpret_cast<WPARAM>(h == title_ ? title_font_ : font_), TRUE);
                }
                layout(hwnd);
                SetFocus(ok_btn_);
                handled = true;
                return 0;
            }
            case WM_SIZE:
                if (ok_btn_ != nullptr) layout(hwnd);
                handled = true;
                return 0;
            case WM_ERASEBKGND: {
                RECT client;
                GetClientRect(hwnd, &client);
                themedraw::fill_rect(reinterpret_cast<HDC>(wparam), client, theme_.window);
                handled = true;
                return 1;
            }
            case WM_PAINT: {
                PAINTSTRUCT paint{};
                HDC dc = BeginPaint(hwnd, &paint);
                RECT client;
                GetClientRect(hwnd, &client);
                UINT dpi = GetDpiForWindow(hwnd);
                auto scale = [dpi](int v) { return MulDiv(v, static_cast<int>(dpi), 96); };
                int y = scale(96);
                HPEN pen = CreatePen(PS_SOLID, 1, theme_.border);
                HGDIOBJ old = SelectObject(dc, pen);
                MoveToEx(dc, scale(20), y, nullptr);
                LineTo(dc, client.right - scale(20), y);
                SelectObject(dc, old);
                DeleteObject(pen);
                EndPaint(hwnd, &paint);
                handled = true;
                return 0;
            }
            case WM_CTLCOLORSTATIC: {
                HDC dc = reinterpret_cast<HDC>(wparam);
                HWND ctl = reinterpret_cast<HWND>(lparam);
                SetBkColor(dc, theme_.window);
                SetTextColor(dc, ctl == title_ ? theme_.accent
                                 : ctl == components_ ? theme_.muted_text
                                                      : theme_.text);
                handled = true;
                return reinterpret_cast<LRESULT>(window_brush_);
            }
            case WM_DRAWITEM: {
                const auto& draw = *reinterpret_cast<DRAWITEMSTRUCT*>(lparam);
                if (draw.CtlID == IdAutoUpdateCheck) {
                    themedraw::draw_checkbox(draw, theme_, auto_update_);
                } else {
                    themedraw::draw_button(draw, theme_);
                }
                handled = true;
                return TRUE;
            }
            case WM_COMMAND: {
                int id = LOWORD(wparam);
                if (id == IdAutoUpdateCheck) {
                    auto_update_ = !auto_update_;
                    settings_.set_bool("CheckUpdatesOnLaunch", auto_update_);
                    settings_.save();
                    InvalidateRect(auto_update_check_, nullptr, FALSE);
                    handled = true;
                    return 0;
                }
                if (id == IdOkButton) {
                    DestroyWindow(hwnd);
                } else if (id == IdCheckButton && check_updates_command_ != 0) {
                    HWND owner = owner_;
                    UINT cmd = check_updates_command_;
                    DestroyWindow(hwnd);
                    PostMessageW(owner, WM_COMMAND, MAKEWPARAM(cmd, 0), 0);
                }
                handled = true;
                return 0;
            }
            case WM_CLOSE:
                DestroyWindow(hwnd);
                handled = true;
                return 0;
            case WM_DESTROY:
                // No PostQuitMessage: the modal loop exits on IsWindow();
                // posting WM_QUIT here would leak into the owner's loop.
                handled = true;
                return 0;
            default:
                return 0;
        }
    }

    void layout(HWND hwnd) {
        RECT client;
        GetClientRect(hwnd, &client);
        UINT dpi = GetDpiForWindow(hwnd);
        auto s = [dpi](int v) { return MulDiv(v, static_cast<int>(dpi), 96); };
        const int margin = s(20);
        const int icon = s(48);
        const int text_left = margin + icon + s(16);
        const int text_w = client.right - text_left - margin;

        // Size the title from the font's real line height (ascent + descent +
        // leading); a fixed height clips descenders at some DPIs/scales.
        int title_h = s(34);
        if (title_font_ != nullptr) {
            HDC dc = GetDC(hwnd);
            HGDIOBJ old = SelectObject(dc, title_font_);
            TEXTMETRICW metrics{};
            if (GetTextMetricsW(dc, &metrics)) {
                title_h = metrics.tmHeight + metrics.tmExternalLeading + s(2);
            }
            SelectObject(dc, old);
            ReleaseDC(hwnd, dc);
        }

        MoveWindow(icon_, margin, margin + s(2), icon, icon, TRUE);
        MoveWindow(title_, text_left, margin - s(4), text_w, title_h, TRUE);
        MoveWindow(description_, text_left, margin - s(4) + title_h, text_w, s(34), TRUE);

        int y = s(108); // below the separator at y=96
        const int row = s(22);
        for (int i = 0; i < 4; ++i) {
            MoveWindow(metadata_[i], margin, y + i * row, client.right - margin * 2, row, TRUE);
        }
        int cy = y + 4 * row + s(10);
        MoveWindow(components_title_, margin, cy, client.right - margin * 2, row, TRUE);
        MoveWindow(components_, margin, cy + s(24), client.right - margin * 2, s(80), TRUE);

        const int btn_h = s(30);
        const int btn_w = s(96);
        const int check_w = s(150);
        const int btn_y = client.bottom - margin - btn_h;
        MoveWindow(ok_btn_, client.right - margin - btn_w, btn_y, btn_w, btn_h, TRUE);
        MoveWindow(check_btn_, client.right - margin - btn_w - s(10) - check_w, btn_y, check_w,
                   btn_h, TRUE);
        // Auto-update toggle sits bottom-left, vertically centred on the buttons
        // and clamped so it can never run into the Check-for-Updates button.
        const int toggle_right = client.right - margin - btn_w - s(10) - check_w - s(12);
        const int toggle_w = std::max(s(120), std::min(s(170), toggle_right - margin));
        MoveWindow(auto_update_check_, margin, btn_y + (btn_h - s(20)) / 2, toggle_w, s(20), TRUE);
        InvalidateRect(hwnd, nullptr, TRUE);
    }
};

} // namespace xact::ui
