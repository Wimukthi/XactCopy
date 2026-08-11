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
#include <cerrno>
#include <cstdlib>
#include <cwchar>
#include <functional>
#include <map>
#include <string>
#include <string_view>
#include <vector>

#include "../version.h"
#include "app_icon.h"
#include "settings.h"
#include "theme.h"

namespace xact::ui {

inline constexpr const char* NativeAppVersion = XACTCOPY_VERSION_STRING;

namespace dialog_detail {

// Modal host: registers a class once, disables the owner, and pumps messages
// until the dialog window is destroyed. The owner is restored while the child
// still exists so USER32 can transfer activation directly within XactCopy
// instead of briefly activating and repainting an unrelated window.
class ModalHost {
public:
    using MessageHandler =
        std::function<LRESULT(HWND, UINT, WPARAM, LPARAM, bool& handled)>;

    void set_density_percent(int density_percent) {
        density_percent_ = std::clamp(density_percent, 80, 120);
    }

    void set_scale_percent(int scale_percent) {
        scale_percent_ = std::clamp(scale_percent, 50, 250);
    }

    HWND create(HWND owner, const wchar_t* class_name, const wchar_t* title, int width,
                int height, MessageHandler handler, bool resizable = false) {
        handler_ = std::move(handler);
        owner_ = owner;
        hwnd_ = nullptr;
        owner_was_enabled_ = false;
        owner_restored_ = false;
        restore_activation_ = false;
        activation_owned_ = false;
        close_requested_ = false;
        command_dispatch_ = false;

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
        UINT layout_dpi = ui_layout_dpi(dpi, density_percent_, scale_percent_);
        const int scaled_client_width = MulDiv(width, static_cast<int>(layout_dpi), 96);
        const int scaled_client_height = MulDiv(height, static_cast<int>(layout_dpi), 96);

        DWORD style = WS_POPUP | WS_CAPTION | WS_SYSMENU | WS_CLIPCHILDREN;
        DWORD ex_style = WS_EX_DLGMODALFRAME;
        if (resizable) {
            style |= WS_THICKFRAME | WS_MAXIMIZEBOX;
            ex_style = 0; // the modal frame forces a fixed border; drop it to size
        }

        // Width/height describe the usable client area. Passing those values
        // directly to CreateWindowEx made borders/title bars consume a growing
        // fraction of the content on high-DPI monitors and changed the layout
        // again after a monitor transition.
        RECT window_rect{0, 0, scaled_client_width, scaled_client_height};
        if (!AdjustWindowRectExForDpi(&window_rect, style, FALSE, ex_style, dpi)) {
            AdjustWindowRectEx(&window_rect, style, FALSE, ex_style);
        }
        const int window_width = window_rect.right - window_rect.left;
        const int window_height = window_rect.bottom - window_rect.top;

        RECT owner_rect{};
        if (owner != nullptr) GetWindowRect(owner, &owner_rect);
        int x = owner_rect.left + ((owner_rect.right - owner_rect.left) - window_width) / 2;
        int y = owner_rect.top + ((owner_rect.bottom - owner_rect.top) - window_height) / 2;

        hwnd_ = CreateWindowExW(ex_style, class_name, title, style, x, y, window_width,
                                window_height, owner, nullptr, GetModuleHandleW(nullptr), this);
        return hwnd_;
    }

    void run_modal() {
        if (hwnd_ == nullptr) return;
        owner_was_enabled_ = owner_ != nullptr && IsWindow(owner_) &&
                             IsWindowEnabled(owner_);
        if (owner_was_enabled_) {
            activation_owned_ = modal_owner_has_foreground(owner_, hwnd_);
            EnableWindow(owner_, FALSE);
        }
        ShowWindow(hwnd_, SW_SHOW);

        MSG message{};
        while (IsWindow(hwnd_) && GetMessageW(&message, nullptr, 0, 0) > 0) {
            if (IsDialogMessageW(hwnd_, &message)) continue;
            TranslateMessage(&message);
            DispatchMessageW(&message);
        }
        restore_owner_after_modal();
    }

    // Close a modal child through the same activation-safe path used by the
    // AxiomCompress dialogs. Callers use this instead of DestroyWindow so the
    // owner is enabled before USER32 processes the child destruction.
    void destroy() {
        if (hwnd_ == nullptr || !IsWindow(hwnd_)) return;
        close_requested_ = true;
        activation_owned_ = activation_owned_ ||
                            modal_owner_has_foreground(owner_, hwnd_) ||
                            GetActiveWindow() == hwnd_;
        restore_owner_before_destroy();
        DestroyWindow(hwnd_);
    }

    HWND hwnd() const noexcept { return hwnd_; }

private:
    HWND hwnd_ = nullptr;
    HWND owner_ = nullptr;
    int density_percent_ = 100;
    int scale_percent_ = 100;
    MessageHandler handler_;
    bool owner_was_enabled_ = false;
    bool owner_restored_ = false;
    bool restore_activation_ = false;
    bool activation_owned_ = false;
    bool close_requested_ = false;
    bool command_dispatch_ = false;

    static bool modal_owner_has_foreground(HWND owner, HWND dialog) {
        if (owner == nullptr || !IsWindow(owner)) return false;
        const HWND foreground = GetForegroundWindow();
        if (foreground == nullptr) {
            const HWND active = GetActiveWindow();
            return active == owner || (dialog != nullptr && active == dialog);
        }
        if (foreground == owner || foreground == dialog) return true;

        // Nested XactCopy dialogs share the same root owner. Preserve
        // activation within that chain, but do not pull XactCopy in front of
        // an application the user switched to while the dialog was open.
        const HWND owner_root = GetAncestor(owner, GA_ROOTOWNER);
        const HWND foreground_root = GetAncestor(foreground, GA_ROOTOWNER);
        return owner_root != nullptr && foreground_root == owner_root;
    }

    static bool window_belongs_to_modal_chain(HWND owner, HWND window) {
        if (owner == nullptr || window == nullptr) return false;
        if (window == owner) return true;
        const HWND owner_root = GetAncestor(owner, GA_ROOTOWNER);
        return owner_root != nullptr &&
               GetAncestor(window, GA_ROOTOWNER) == owner_root;
    }

    void observe_message(HWND hwnd, UINT message, WPARAM wparam, LPARAM lparam) {
        const bool command_can_destroy = message == WM_COMMAND;
        if (message == WM_CLOSE ||
            (message == WM_SYSCOMMAND && (wparam & 0xFFF0u) == SC_CLOSE)) {
            close_requested_ = true;
            activation_owned_ = activation_owned_ ||
                                modal_owner_has_foreground(owner_, hwnd) ||
                                GetActiveWindow() == hwnd;
        }
        if (command_can_destroy) {
            command_dispatch_ = true;
            if (lparam != 0 || GetActiveWindow() == hwnd ||
                modal_owner_has_foreground(owner_, hwnd)) {
                activation_owned_ = true;
            }
        }
        if (message == WM_ACTIVATE) {
            if (LOWORD(wparam) != WA_INACTIVE) {
                activation_owned_ = true;
            } else {
                const HWND next_active = reinterpret_cast<HWND>(lparam);
                if (!close_requested_ && !command_dispatch_ && next_active != nullptr &&
                    !window_belongs_to_modal_chain(owner_, next_active)) {
                    activation_owned_ = false;
                }
            }
        } else if (message == WM_ACTIVATEAPP && wparam == FALSE &&
                   !close_requested_ && !command_dispatch_) {
            activation_owned_ = false;
        }
    }

    void restore_owner_before_destroy() {
        if (!owner_was_enabled_ || owner_restored_ || owner_ == nullptr ||
            !IsWindow(owner_)) {
            return;
        }
        owner_restored_ = true;
        EnableWindow(owner_, TRUE);
        restore_activation_ = activation_owned_;
    }

    void restore_owner_after_modal() {
        if (!owner_was_enabled_ || owner_ == nullptr || !IsWindow(owner_)) return;

        bool restore_activation = restore_activation_;
        // Fallback for an unexpected close path that did not pass through
        // ModalHost::destroy or the host window procedure's WM_DESTROY.
        if (!IsWindowEnabled(owner_)) {
            restore_activation = modal_owner_has_foreground(owner_, nullptr);
            EnableWindow(owner_, TRUE);
        }
        if (!restore_activation || !IsWindowVisible(owner_) || IsIconic(owner_)) return;

        // Activation is deliberately deferred until the child is completely
        // gone. Usually USER32 has already selected the owner, so these calls
        // are no-ops. Do not steal focus from another app or a nested dialog.
        const HWND foreground = GetForegroundWindow();
        const HWND owner_root = GetAncestor(owner_, GA_ROOTOWNER);
        if (foreground != nullptr && foreground != owner_ && owner_root != nullptr &&
            GetAncestor(foreground, GA_ROOTOWNER) == owner_root) {
            return;
        }
        if (GetActiveWindow() != owner_) SetActiveWindow(owner_);
        if (GetForegroundWindow() != owner_) {
            BringWindowToTop(owner_);
            SetForegroundWindow(owner_);
        }
    }

    static LRESULT CALLBACK static_proc(HWND hwnd, UINT message, WPARAM wparam, LPARAM lparam) {
        ModalHost* self;
        if (message == WM_NCCREATE) {
            self = static_cast<ModalHost*>(reinterpret_cast<CREATESTRUCTW*>(lparam)->lpCreateParams);
            SetWindowLongPtrW(hwnd, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(self));
            self->hwnd_ = hwnd;
        } else {
            self = reinterpret_cast<ModalHost*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
        }
        const bool command_can_destroy = message == WM_COMMAND;
        if (self != nullptr) {
            self->observe_message(hwnd, message, wparam, lparam);
            if (message == WM_DESTROY || message == WM_NCDESTROY) {
                self->restore_owner_before_destroy();
            }
        }
        if (self != nullptr && self->handler_) {
            bool handled = false;
            LRESULT result = self->handler_(hwnd, message, wparam, lparam, handled);
            if (handled) {
                if (command_can_destroy && IsWindow(hwnd)) self->command_dispatch_ = false;
                if (message == WM_NCDESTROY) self->hwnd_ = nullptr;
                return result;
            }
        }
        const LRESULT result = DefWindowProcW(hwnd, message, wparam, lparam);
        if (command_can_destroy && self != nullptr && IsWindow(hwnd)) {
            self->command_dispatch_ = false;
        }
        if (message == WM_NCDESTROY && self != nullptr) self->hwnd_ = nullptr;
        return result;
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
    TooltipManager tooltips_;

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

    void rebuild_font(UINT monitor_dpi) {
        dpi_ = ui_layout_dpi(monitor_dpi, theme_.density_percent, theme_.scale_percent);
        if (font_ != nullptr) DeleteObject(font_);
        font_ = CreateFontW(-MulDiv(9, static_cast<int>(dpi_), 72), 0, 0, 0, FW_NORMAL,
                            FALSE, FALSE, FALSE, DEFAULT_CHARSET, OUT_DEFAULT_PRECIS,
                            CLIP_DEFAULT_PRECIS, CLEARTYPE_QUALITY, DEFAULT_PITCH, L"Segoe UI");
    }

    void relayout(HWND hwnd) {
        if (button_windows_.empty()) return;
        auto s = [this](int v) { return MulDiv(v, static_cast<int>(dpi_), 96); };
        RECT client{};
        GetClientRect(hwnd, &client);
        const int margin = s(16);
        const int button_width = s(96);
        const int button_height = s(30);
        const int gap = s(8);
        int x = client.right - margin - button_width;
        const int y = client.bottom - margin - button_height;
        for (HWND button : button_windows_) {
            MoveWindow(button, x, y, button_width, button_height, TRUE);
            x -= button_width + gap;
        }
        text_rect_ = {margin + (icon_ == MessageIcon::None ? 0 : s(48)), margin,
                      client.right - margin, y - s(12)};
        tooltips_.update_layout();
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
        host_.set_density_percent(theme_.density_percent);
        host_.set_scale_percent(theme_.scale_percent);
        dpi_ = ui_layout_dpi(owner != nullptr ? GetDpiForWindow(owner) : 96,
                             theme_.density_percent, theme_.scale_percent);
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
        set_dark_title_bar(hwnd, theme_.dark && theme_.themed_chrome);
        host_.run_modal();
        return result_;
    }

    LRESULT proc(HWND hwnd, UINT message, WPARAM wparam, LPARAM lparam, bool& handled) {
        switch (message) {
            case WM_CREATE: {
                window_brush_ = CreateSolidBrush(theme_.window);
                rebuild_font(GetDpiForWindow(hwnd));
                // Lay buttons out right-to-left so the first spec ends up leftmost
                // of the group (matching MessageBoxW's ordering).
                for (auto it = button_specs_.rbegin(); it != button_specs_.rend(); ++it) {
                    HWND b = CreateWindowExW(
                        0, L"BUTTON", it->text, WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_OWNERDRAW,
                        0, 0, 0, 0, hwnd,
                        reinterpret_cast<HMENU>(static_cast<INT_PTR>(it->id)),
                        GetModuleHandleW(nullptr), nullptr);
                    SendMessageW(b, WM_SETFONT, reinterpret_cast<WPARAM>(font_), TRUE);
                    button_windows_.push_back(b);
                }
                if (!button_windows_.empty()) SetFocus(button_windows_.back());
                tooltips_.create(hwnd, dpi_, theme_.dark);
                for (HWND button : button_windows_) {
                    switch (GetDlgCtrlID(button)) {
                        case IDOK:
                            tooltips_.add(button, L"Acknowledge this message and close the dialog.");
                            break;
                        case IDYES:
                            tooltips_.add(button, L"Confirm the requested action.");
                            break;
                        case IDNO:
                            tooltips_.add(button, L"Decline the requested action.");
                            break;
                        case IDCANCEL:
                            tooltips_.add(button, L"Cancel the requested action and close the dialog.");
                            break;
                    }
                }
                relayout(hwnd);
                handled = true;
                return 0;
            }
            case WM_SIZE:
                relayout(hwnd);
                InvalidateRect(hwnd, nullptr, TRUE);
                handled = true;
                return 0;
            case WM_DPICHANGED: {
                rebuild_font(HIWORD(wparam) != 0 ? HIWORD(wparam) : LOWORD(wparam));
                const RECT* suggested = reinterpret_cast<const RECT*>(lparam);
                SetWindowPos(hwnd, nullptr, suggested->left, suggested->top,
                             suggested->right - suggested->left,
                             suggested->bottom - suggested->top,
                             SWP_NOZORDER | SWP_NOACTIVATE);
                for (HWND button : button_windows_) {
                    SendMessageW(button, WM_SETFONT, reinterpret_cast<WPARAM>(font_), TRUE);
                }
                tooltips_.update_dpi(dpi_);
                apply_window_icons(hwnd);
                relayout(hwnd);
                RedrawWindow(hwnd, nullptr, nullptr,
                             RDW_INVALIDATE | RDW_ERASE | RDW_ALLCHILDREN);
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
                        host_.destroy();
                        break;
                    }
                }
                handled = true;
                return 0;
            }
            case WM_CLOSE:
                host_.destroy();
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
            combo(0, "Mode", "Theme", {L"Dark", L"Light", L"System", L"Classic"},
                  {"dark", "light", "system", "classic"}, "dark");
            combo(0, "Accent source", "AccentColorMode", {L"Auto", L"System", L"Custom"},
                  {"auto", "system", "custom"}, "auto");
            edit_text(0, "Custom accent color", "AccentColorHex", "#5A78C8");
            combo(0, "Window chrome", "WindowChromeMode",
                  {L"Themed title bar", L"Standard title bar"}, {"themed", "standard"}, "themed");
            section(0, "Layout");
            combo(0, "UI density", "UiDensity", {L"Compact", L"Normal", L"Comfortable"},
                  {"compact", "normal", "comfortable"}, "normal");
            combo_int(0, "UI scale", "UiScalePercent",
                      {L"50%", L"75%", L"90%", L"100%", L"110%", L"125%",
                       L"150%", L"175%", L"200%", L"225%", L"250%"},
                      {50, 75, 90, 100, 110, 125, 150, 175, 200, 225, 250}, 100);
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
            check(1, "Salvage unreadable blocks", "DefaultSalvageUnreadableBlocks", false);
            check(1, "Continue on file errors", "DefaultContinueOnFileError", false);
            check(1, "Preserve source timestamps", "DefaultPreserveTimestamps", true);
            check(1, "Copy empty directories", "DefaultCopyEmptyDirectories", true);
            check(1, "Wait forever for source/destination", "DefaultWaitForMediaAvailability",
                  false);
            check(1, "Wait for lock/contention release", "DefaultWaitForFileLockRelease", false);
            check(1, "Treat Access Denied as contention", "DefaultTreatAccessDeniedAsContention",
                  false);
            section(1, "Policies");
            combo(1, "Overwrite policy", "DefaultOverwritePolicy",
                  {L"Overwrite existing", L"Skip existing", L"Overwrite if newer", L"Stop on conflict"},
                  {"overwrite", "skip-existing", "overwrite-if-newer", "ask"}, "overwrite");
            combo(1, "Symlink handling", "DefaultSymlinkHandling",
                  {L"Skip symbolic links", L"Follow targets inside source only",
                   L"Follow external targets (expert)"},
                  {"skip", "follow-internal", "follow"}, "skip");
            check(1, "Allow salvaged files to replace existing files",
                  "DefaultAllowRecoveredOverwriteExisting", false);
            check(1, "Allow encrypted sources to become plaintext",
                  "DefaultAllowDecryptedDestination", false);
            combo(1, "Salvage fill pattern", "DefaultSalvageFillPattern",
                  {L"Zero-fill", L"0xFF-fill"}, {"zero", "ones"}, "zero");
            section(1, "Bad Range Map");
            check(1, "Use bad-range map when available", "DefaultUseBadRangeMap", false);
            check(1, "Skip known bad ranges during copy", "DefaultSkipKnownBadRanges", false);
            check(1, "Update map from scan/copy runs", "DefaultUpdateBadRangeMapFromRun", false);
            check(1, "Raw volume scan (local NTFS; Administrator)",
                  "DefaultUseExperimentalRawDiskScan", false);
            edit_int(1, "Map max age (days, 0=never)", "DefaultBadRangeMapMaxAgeDays", 0, 3650,
                     30);

            // --- Page 2: Performance ------------------------------------
            section(2, "Transfer Tuning");
            check(2, "Enable adaptive buffer by default", "DefaultUseAdaptiveBuffer", false);
            combo(2, "Transfer engine", "DefaultTransferEnginePolicy",
                  {L"Auto", L"Managed rescue", L"Native fast"}, {"auto", "managed", "native"},
                  "auto");
            combo(2, "Scan profile", "DefaultScanPerformanceProfile",
                  {L"Automatic assessment", L"Fast file assessment", L"Precise range assessment"},
                  {"auto", "fast", "precise"}, "auto");
            combo(2, "Worker priority", "DefaultWorkerProcessPriorityClass",
                  {L"Idle", L"Below normal", L"Normal", L"Above normal", L"High"},
                  {"Idle", "BelowNormal", "Normal", "AboveNormal", "High"}, "Normal");
            edit_int(2, "Manual buffer MB", "DefaultBufferSizeMb", 1, 256, 4);
            edit_int(2, "Max retries", "DefaultMaxRetries", 0, 32, 2);
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
            edit_int(2, "FastScan retries", "DefaultRescueFastScanRetries", 0, 32, 0);
            edit_int(2, "TrimSweep retries", "DefaultRescueTrimRetries", 0, 32, 1);
            edit_int(2, "Scrape retries", "DefaultRescueScrapeRetries", 0, 32, 2);

            // --- Page 3: Diagnostics ------------------------------------
            section(3, "Worker Telemetry");
            combo(3, "Worker telemetry profile", "WorkerTelemetryProfile",
                  {L"Normal", L"Verbose", L"Debug"}, {"normal", "verbose", "debug"}, "normal");
            edit_int(3, "Progress interval (ms)", "WorkerProgressIntervalMs", 20, 1000, 75);
            edit_int(3, "Log rate cap (/sec, 0=off)", "WorkerMaxLogsPerSecond", 0, 5000, 100);
            section(3, "UI Diagnostics");
            check(3, "Show UI diagnostics strip", "UiShowDiagnostics", false);
            edit_int(3, "Diagnostics refresh (ms)", "UiDiagnosticsRefreshMs", 100, 5000, 250);
            edit_int(3, "Virtual log max lines", "UiMaxLogLines", 1000, 1000000, 50000);
            section(3, "Journal Storage");
            check(3, "Discard backups after a run completes", "CompactJournalsOnCompletion", true);
            edit_int(3, "Delete journals after (days, 0=keep)", "JournalRetentionDays", 0, 3650, 30);
            edit_int(3, "Always keep newest journals", "JournalKeepMinimum", 0, 1000, 10);

            // --- Page 4: Verification -----------------------------------
            section(4, "Verification Defaults");
            check(4, "Enable verification by default", "DefaultVerifyAfterCopy", true);
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

    static const wchar_t* field_help(const char* key) {
        if (key == nullptr) return L"";
        const std::string_view setting(key);

        if (setting == "Theme") return L"Choose XactCopy's application palette. Light keeps the custom light palette, System follows Windows, and Classic uses standard Windows controls.";
        if (setting == "AccentColorMode") return L"Choose whether the accent is selected automatically, inherited from Windows, or read from the custom color field.";
        if (setting == "AccentColorHex") return L"Enter a custom accent as #RRGGBB. It is used only when Accent source is Custom.";
        if (setting == "WindowChromeMode") return L"Themed applies XactCopy's dark/light title bar. Standard leaves title-bar rendering to Windows.";
        if (setting == "UiDensity") return L"Adjust spacing and control heights without changing monitor DPI. Compact fits more content; Comfortable adds breathing room.";
        if (setting == "UiScalePercent") return L"Apply an additional application-only scale on top of each monitor's Windows DPI setting.";
        if (setting == "LogFontFamily") return L"Font family used by the operations log. A monospaced font such as Consolas keeps columns aligned.";
        if (setting == "LogFontSizePoints") return L"Point size used by the operations log on every monitor.";
        if (setting == "UiColorizeLogBySeverity") return L"Color errors, warnings, successful recovery, and supervisor messages so important events are easier to find.";
        if (setting == "GridAlternatingRows") return L"Use alternating row backgrounds in Job Manager to make wide records easier to follow.";
        if (setting == "GridRowHeight") return L"Logical Job Manager row height before monitor DPI and UI scale are applied.";
        if (setting == "GridHeaderStyle") return L"Choose the visual emphasis used for Job Manager column headers.";
        if (setting == "ShowBufferStatusRow") return L"Show the active I/O buffer and adaptive-buffer status beneath the progress bars.";
        if (setting == "ShowRescueStatusRow") return L"Show rescue pass, unreadable-region, and recovery information beneath the progress bars.";
        if (setting == "ProgressBarShowPercentage") return L"Draw a numeric percentage inside progress bars when the selected bar height has enough room.";
        if (setting == "ProgressBarStyle") return L"Choose the visual height of current-file and overall progress bars.";

        if (setting == "DefaultResumeFromJournal") return L"Load matching journal state by default. Completed files still require the configured validation before reuse.";
        if (setting == "DefaultSalvageUnreadableBlocks") return L"When bytes remain unreadable, write the selected fill pattern and report recovery instead of abandoning all readable data.";
        if (setting == "DefaultContinueOnFileError") return L"Continue processing later files after an error. The run remains failed or incomplete when any file is lost or skipped.";
        if (setting == "DefaultPreserveTimestamps") return L"Apply source creation, access, and last-write times to the staged destination before publication.";
        if (setting == "DefaultCopyEmptyDirectories") return L"Create source directories that contain no copied files. Disable when only file content matters.";
        if (setting == "DefaultWaitForMediaAvailability") return L"Wait indefinitely for a missing source or destination volume to return instead of failing immediately.";
        if (setting == "DefaultWaitForFileLockRelease") return L"Probe and wait for sharing violations or locks to clear. Cancel remains available while waiting.";
        if (setting == "DefaultTreatAccessDeniedAsContention") return L"Treat Access Denied like a temporary lock and retry it. Use cautiously: permissions errors can otherwise wait indefinitely.";
        if (setting == "DefaultOverwritePolicy") return L"Default action when a destination file exists. Stop on conflict leaves the existing file untouched and marks the run incomplete.";
        if (setting == "DefaultSymlinkHandling") return L"Internal-only follows links whose resolved target remains beneath the source. External following can copy unrelated folders or another volume and requires an attended run.";
        if (setting == "DefaultAllowRecoveredOverwriteExisting") return L"Expert override: permit a file containing synthetic salvage bytes to replace an existing destination. Leave disabled so recovery output is published as a clearly named sidecar.";
        if (setting == "DefaultAllowDecryptedDestination") return L"Permit EFS-encrypted source data to be published without encryption. Leave disabled unless plaintext output is an explicit requirement.";
        if (setting == "DefaultSalvageFillPattern") return L"Bytes written where source data could not be recovered. The result is always marked non-exact regardless of pattern.";
        if (setting == "DefaultUseBadRangeMap") return L"Load remembered unreadable ranges for the same source so repeated runs can make safer decisions.";
        if (setting == "DefaultSkipKnownBadRanges") return L"Avoid rereading compatible map ranges already known to be bad. This reduces media stress but cannot reconstruct those bytes.";
        if (setting == "DefaultUpdateBadRangeMapFromRun") return L"Persist unreadable ranges found during scans or copies for later runs. Disable for a read-only, non-updating map policy.";
        if (setting == "DefaultUseExperimentalRawDiskScan") return L"Read local NTFS file extents through the raw volume when running as Administrator. Unsupported layouts safely fall back to file reads.";
        if (setting == "DefaultBadRangeMapMaxAgeDays") return L"Reject old map entries as skip hints after this many days. Zero never expires them and should be used only for stable media.";

        if (setting == "DefaultUseAdaptiveBuffer") return L"Adjust buffer size from observed throughput and failures instead of keeping the manual size fixed.";
        if (setting == "DefaultTransferEnginePolicy") return L"Auto selects a safe engine for the active policies. Managed enables rescue behavior; Native Fast attempts CopyFileEx first.";
        if (setting == "DefaultScanPerformanceProfile") return L"Fast uses parallel file reads with precise fallback around failures. Precise performs detailed unreadable-range detection. Automatic currently chooses Fast.";
        if (setting == "DefaultWorkerProcessPriorityClass") return L"Windows scheduling priority for the isolated worker process. High can reduce responsiveness of other applications.";
        if (setting == "DefaultBufferSizeMb") return L"Base I/O buffer in MiB. Larger values favor healthy throughput; smaller values isolate damaged areas more precisely.";
        if (setting == "DefaultMaxRetries") return L"Additional application-level attempts for retry-capable I/O. Windows and the device may already retry internally; large values can heavily stress failing media.";
        if (setting == "DefaultOperationTimeoutSeconds") return L"Timeout for one I/O operation. This is not a whole-file or whole-job timeout.";
        if (setting == "DefaultPerFileTimeoutSeconds") return L"Optional total processing deadline for one file. Zero disables the per-file deadline.";
        if (setting == "DefaultMaxThroughputMbPerSecond") return L"Throttle aggregate transfer throughput to reduce device load. Zero removes the throughput cap.";
        if (setting == "DefaultParallelSmallFileWorkers") return L"Concurrent workers used for eligible small-file Native Fast copies. Zero uses one worker on seek-based disks and up to eight on non-seek storage.";
        if (setting == "DefaultParallelScanWorkers") return L"Concurrent workers used by Fast assessments. Zero uses one worker on seek-based disks and up to eight on non-seek storage; explicit higher values can stress media.";
        if (setting == "DefaultSmallFileThresholdKb") return L"Largest file eligible for the parallel small-file path. Files above this size use the regular staged pipeline.";
        if (setting == "DefaultFragileMediaMode") return L"Enable conservative behavior for unstable media, including clustered-failure protection and reduced read pressure.";
        if (setting == "DefaultSkipFileOnFirstReadError") return L"In fragile mode, stop reading a file after its first read error instead of running aggressive rescue passes.";
        if (setting == "DefaultPersistFragileSkipsAcrossResume") return L"Remember fragile-mode file skips in the journal so a resumed run does not immediately stress them again.";
        if (setting == "DefaultFragileFailureWindowSeconds") return L"Time window used to detect a cluster of failures that may indicate the whole device is becoming unstable.";
        if (setting == "DefaultFragileFailureThreshold") return L"Number of failures within the configured window that activates the fragile-media cooldown.";
        if (setting == "DefaultFragileCooldownSeconds") return L"Pause after a clustered failure event to let unstable media recover. Zero disables the cooldown delay.";
        if (setting == "DefaultLockContentionProbeIntervalMs") return L"Delay between lock/media availability probes. Smaller values react faster but create more filesystem traffic.";
        if (setting == "DefaultSourceMutationPolicy") return L"Action when a source file changes during copying: fail it, skip it visibly, or wait for the expected file to reappear.";
        if (setting == "DefaultRescueFastScanChunkKb") return L"Override the broad first rescue-pass chunk size. Zero lets XactCopy derive it from the active buffer.";
        if (setting == "DefaultRescueTrimChunkKb") return L"Override the TrimSweep chunk used around failed-range edges. Zero selects the built-in adaptive size.";
        if (setting == "DefaultRescueScrapeChunkKb") return L"Override the fine Scrape pass chunk size for damaged ranges. Zero selects the built-in adaptive size.";
        if (setting == "DefaultRescueRetryChunkKb") return L"Override the chunk size used when retrying remaining bad ranges. Zero selects the built-in adaptive size.";
        if (setting == "DefaultRescueSplitMinimumKb") return L"Smallest range the rescue engine may split while isolating damage. Zero uses the engine default.";
        if (setting == "DefaultRescueFastScanRetries") return L"Additional attempts made by the broad FastScan rescue pass.";
        if (setting == "DefaultRescueTrimRetries") return L"Additional attempts made while trimming the boundaries of failed ranges.";
        if (setting == "DefaultRescueScrapeRetries") return L"Additional fine-grained attempts made by the Scrape rescue pass.";

        if (setting == "WorkerTelemetryProfile") return L"Controls worker diagnostic detail. Debug is useful for investigation but produces substantially more log traffic.";
        if (setting == "WorkerProgressIntervalMs") return L"Minimum interval between worker progress events. Lower values update more smoothly but increase UI and IPC work.";
        if (setting == "WorkerMaxLogsPerSecond") return L"Rate-limit worker log events to keep the UI responsive during noisy failures. Zero disables the cap.";
        if (setting == "UiShowDiagnostics") return L"Show the live diagnostics strip on the main window.";
        if (setting == "UiDiagnosticsRefreshMs") return L"Refresh interval for main-window diagnostic counters. Lower values consume more UI-thread time.";
        if (setting == "UiMaxLogLines") return L"Maximum operations-log lines retained in memory. Older lines are discarded when the limit is reached.";
        if (setting == "CompactJournalsOnCompletion") return L"Remove rotated journal backups after a completed run while retaining the current trusted snapshot.";
        if (setting == "JournalRetentionDays") return L"Delete eligible old journal sets after this many days. Zero keeps them indefinitely.";
        if (setting == "JournalKeepMinimum") return L"Always retain at least this many newest journal sets even when they exceed the age limit.";

        if (setting == "DefaultVerifyAfterCopy") return L"Enable post-copy verification by default. Verification occurs against the staged file before it replaces the destination.";
        if (setting == "DefaultVerificationMode") return L"Full hashes all bytes; Sampled checks selected chunks and is faster but provides weaker detection.";
        if (setting == "DefaultVerificationHashAlgorithm") return L"Hash used for verification. SHA-256 is faster and sufficient for copy integrity; SHA-512 provides a larger digest.";
        if (setting == "DefaultSampleVerificationChunkKb") return L"Size of each region compared in Sampled verification mode.";
        if (setting == "DefaultSampleVerificationChunkCount") return L"Number of distributed regions compared in Sampled mode. More samples improve coverage and cost more reads.";

        if (setting == "EnableRecoveryAutostart") return L"Register XactCopy to start at the next logon after an interrupted run so recovery can be offered.";
        if (setting == "AutoRunQueuedJobsOnStartup") return L"Automatically dequeue and run saved jobs when XactCopy starts.";
        if (setting == "PromptResumeAfterCrash") return L"Offer to resume a journaled run after XactCopy or the worker ended unexpectedly.";
        if (setting == "AutoResumeAfterCrash") return L"Resume an interrupted run without asking. Leave disabled when source or destination media may have changed.";
        if (setting == "KeepResumePromptUntilResolved") return L"Continue showing the recovery prompt on later launches until the interrupted run is resumed or dismissed.";
        if (setting == "RecoveryTouchIntervalSeconds") return L"How often the active-run recovery marker is refreshed. Shorter intervals reduce stale-state uncertainty but write more often.";
        if (setting == "EnableExplorerContextMenu") return L"Add XactCopy copy and scan commands to Windows Explorer for selected files and folders.";

        return L"Configure this XactCopy setting. Changes are applied only after Save is selected.";
    }

    static const wchar_t* section_help(std::string_view section) {
        if (section == "Theme & Accent") return L"Choose XactCopy's light/dark presentation, accent source, and title-bar treatment.";
        if (section == "Layout") return L"Adjust application density and scaling in addition to Windows per-monitor DPI scaling.";
        if (section == "Operations Log") return L"Control the typeface and severity coloring used by the live operations log.";
        if (section == "Grid Appearance") return L"Control Job Manager row spacing, alternating backgrounds, and header emphasis.";
        if (section == "Status & Progress") return L"Choose which live status rows are visible and how progress bars are rendered.";
        if (section == "Default Run Behavior") return L"Defaults applied when a new copy or scan is prepared. They can still be changed on the main window.";
        if (section == "Policies") return L"Default conflict, symbolic-link, and unreadable-byte handling policies for new jobs.";
        if (section == "Bad Range Map") return L"Configure whether source-specific bad-range maps are read, updated, and used to avoid repeated damaged sectors.";
        if (section == "Transfer Tuning") return L"Tune worker counts, buffers, timeouts, throughput, process priority, and scan profile.";
        if (section == "Fragile Media Guard") return L"Reduce repeated reads against deteriorating media after clustered failures are detected.";
        if (section == "Contention & Source Mutation") return L"Control how XactCopy responds to locked files and files that change or disappear during a run.";
        if (section == "Rescue Engine Tuning") return L"Advanced chunk sizes and retry budgets for the staged FastScan, Trim, Scrape, and RetryBad rescue passes.";
        if (section == "Worker Telemetry") return L"Control worker progress frequency and diagnostic-log volume sent to the UI.";
        if (section == "UI Diagnostics") return L"Control main-window diagnostics and the number of log lines retained in memory.";
        if (section == "Journal Storage") return L"Control completed-journal compaction and age-based retention without removing active resumable runs.";
        if (section == "Verification Defaults") return L"Choose the default post-copy verification method and cryptographic hash algorithm.";
        if (section == "Sampled Mode") return L"Tune how many regions Sampled verification checks and the size of each region.";
        if (section == "Startup") return L"Choose whether recovery support or queued jobs may run when XactCopy starts or Windows signs in.";
        if (section == "Interrupted Run Handling") return L"Control prompts, automatic resume, and freshness tracking for interrupted journaled runs.";
        if (section == "Context Menu") return L"Enable or remove XactCopy commands in Windows Explorer.";
        return L"Settings in this group are saved together when you select Save.";
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
    HWND save_button_ = nullptr;
    HWND cancel_button_ = nullptr;
    UINT dpi_ = 96;
    TooltipManager tooltips_;
    int current_page_ = 0;
    bool saved_ = false;
    std::wstring validation_error_;
    HWND validation_control_ = nullptr;

    // Staged values (all pages), keyed by settings key.
    std::map<std::string, bool> staged_bools_;
    std::map<std::string, int> staged_ints_;
    std::map<std::string, std::string> staged_strings_;

    // Live controls for the current page: field-table index -> hwnd.
    std::map<int, HWND> page_controls_;
    std::vector<HWND> page_labels_;
    std::map<int, HWND> page_field_labels_;
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

    void rebuild_fonts(UINT monitor_dpi) {
        dpi_ = ui_layout_dpi(monitor_dpi, theme_.density_percent, theme_.scale_percent);
        if (font_ != nullptr) DeleteObject(font_);
        if (section_font_ != nullptr) DeleteObject(section_font_);
        font_ = CreateFontW(-MulDiv(9, static_cast<int>(dpi_), 72), 0, 0, 0, FW_NORMAL,
                            FALSE, FALSE, FALSE, DEFAULT_CHARSET, OUT_DEFAULT_PRECIS,
                            CLIP_DEFAULT_PRECIS, CLEARTYPE_QUALITY, DEFAULT_PITCH, L"Segoe UI");
        section_font_ = CreateFontW(-MulDiv(10, static_cast<int>(dpi_), 72), 0, 0, 0,
                                    FW_SEMIBOLD, FALSE, FALSE, FALSE, DEFAULT_CHARSET,
                                    OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, CLEARTYPE_QUALITY,
                                    DEFAULT_PITCH, L"Segoe UI");
    }

    void relayout_shell(HWND hwnd) {
        if (nav_list_ == nullptr) return;
        auto scale = [this](int value) { return MulDiv(value, static_cast<int>(dpi_), 96); };
        RECT client{};
        GetClientRect(hwnd, &client);
        MoveWindow(nav_list_, scale(8), scale(12), scale(172),
                   std::max(scale(80), static_cast<int>(client.bottom) - scale(24)), TRUE);
        MoveWindow(save_button_, client.right - scale(200), client.bottom - scale(38), scale(90),
                   scale(28), TRUE);
        MoveWindow(cancel_button_, client.right - scale(100), client.bottom - scale(38), scale(90),
                   scale(28), TRUE);
        tooltips_.update_layout();
    }

    bool run(HWND owner) {
        stage_from_settings();
        host_.set_density_percent(theme_.density_percent);
        host_.set_scale_percent(theme_.scale_percent);
        HWND hwnd = host_.create(
            owner, L"XactCopySettingsDlg", L"Settings", 900, 720,
            [this](HWND h, UINT m, WPARAM w, LPARAM l, bool& handled) {
                return proc(h, m, w, l, handled);
            });
        if (hwnd == nullptr) return false;
        apply_window_icons(hwnd);
        set_dark_title_bar(hwnd, theme_.dark && theme_.themed_chrome);
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
                {
                    std::string value = settings_.get_string(field.key, field.fallback_text);
                    bool recognized = std::any_of(
                        field.values.begin(), field.values.end(), [&](const char* candidate) {
                            return candidate != nullptr &&
                                   models::detail::equals_ignore_case(candidate, value);
                        });
                    staged_strings_[field.key] = recognized ? std::move(value)
                                                            : field.fallback_text;
                    break;
                }
                case FieldKind::ComboInt:
                {
                    int value = settings_.get_int(field.key, INT32_MIN / 2,
                                                  INT32_MAX / 2, field.fallback_int);
                    if (!field.int_values.empty()) {
                        auto nearest = std::min_element(
                            field.int_values.begin(), field.int_values.end(),
                            [value](int left, int right) {
                                return std::llabs(static_cast<long long>(left) - value) <
                                       std::llabs(static_cast<long long>(right) - value);
                            });
                        value = *nearest;
                    }
                    staged_ints_[field.key] = value;
                    break;
                }
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

    bool save_all() {
        validation_error_.clear();
        validation_control_ = nullptr;
        if (!harvest_current_page(true)) {
            themed_message_box(host_.hwnd(), theme_, validation_error_, L"Invalid setting",
                               MessageIcon::Warning);
            if (validation_control_ != nullptr) SetFocus(validation_control_);
            return false;
        }
        if (staged_strings_["AccentColorMode"] == "custom") {
            const std::string& color = staged_strings_["AccentColorHex"];
            const auto is_hex_digit = [](char ch) {
                return (ch >= '0' && ch <= '9') || (ch >= 'a' && ch <= 'f') ||
                       (ch >= 'A' && ch <= 'F');
            };
            const bool valid = color.size() == 7 && color.front() == '#' &&
                               std::all_of(color.begin() + 1, color.end(), is_hex_digit);
            if (!valid) {
                if (current_page_ != 0) {
                    SendMessageW(nav_list_, LB_SETCURSEL, 0, 0);
                    build_page(0);
                }
                const int index = field_index_for_key("AccentColorHex");
                auto control = page_controls_.find(index);
                validation_control_ = control == page_controls_.end() ? nullptr : control->second;
                themed_message_box(host_.hwnd(), theme_,
                                   L"Custom accent color must use the form #RRGGBB.",
                                   L"Invalid setting", MessageIcon::Warning);
                if (validation_control_ != nullptr) SetFocus(validation_control_);
                return false;
            }
        }
        const bool persisted = settings_.update_and_save([&](AppSettings& target) {
            for (const auto& field : fields()) {
                if (field.key == nullptr) continue;
                switch (field.kind) {
                    case FieldKind::Check:
                        target.set_bool(field.key, staged_bools_[field.key]);
                        break;
                    case FieldKind::Combo:
                    case FieldKind::EditText:
                        target.set_string(field.key, staged_strings_[field.key]);
                        break;
                    case FieldKind::ComboInt:
                    case FieldKind::EditInt:
                        target.set_int(field.key, staged_ints_[field.key]);
                        break;
                    default:
                        break;
                }
            }
        });
        if (!persisted) {
            themed_message_box(
                host_.hwnd(), theme_,
                storage::fsutil::utf8_to_wide(settings_.last_save_error()),
                L"Settings could not be saved", MessageIcon::Error);
            return false;
        }
        saved_ = true;
        return true;
    }

    // ---- Page building -----------------------------------------------------

    void destroy_page_controls() {
        for (auto& [index, control] : page_controls_) {
            tooltips_.remove(control);
            DestroyWindow(control);
        }
        for (HWND label : page_labels_) {
            tooltips_.remove(label);
            DestroyWindow(label);
        }
        page_controls_.clear();
        page_labels_.clear();
        page_field_labels_.clear();
        page_checks_.clear();
    }

    int field_index_for_key(std::string_view key) const {
        const auto& table = fields();
        for (int index = 0; index < static_cast<int>(table.size()); ++index) {
            const char* candidate = table[static_cast<std::size_t>(index)].key;
            if (candidate != nullptr && key == candidate) return index;
        }
        return -1;
    }

    bool staged_check(std::string_view key) const {
        const int index = field_index_for_key(key);
        auto live = page_checks_.find(index);
        if (live != page_checks_.end()) return live->second;
        auto staged = staged_bools_.find(std::string(key));
        return staged != staged_bools_.end() && staged->second;
    }

    std::string staged_combo(std::string_view key) const {
        const int index = field_index_for_key(key);
        auto control = page_controls_.find(index);
        if (control != page_controls_.end()) {
            const FieldSpec& field = fields()[static_cast<std::size_t>(index)];
            const int selection = static_cast<int>(
                SendMessageW(control->second, CB_GETCURSEL, 0, 0));
            if (selection >= 0 && selection < static_cast<int>(field.values.size())) {
                return field.values[static_cast<std::size_t>(selection)];
            }
        }
        auto staged = staged_strings_.find(std::string(key));
        return staged == staged_strings_.end() ? std::string() : staged->second;
    }

    void set_field_enabled(std::string_view key, bool enabled) {
        const int index = field_index_for_key(key);
        auto control = page_controls_.find(index);
        if (control != page_controls_.end()) EnableWindow(control->second, enabled ? TRUE : FALSE);
        auto label = page_field_labels_.find(index);
        if (label != page_field_labels_.end()) EnableWindow(label->second, enabled ? TRUE : FALSE);
    }

    void sync_page_dependencies() {
        if (current_page_ == 0) {
            set_field_enabled("AccentColorHex", staged_combo("AccentColorMode") == "custom");
        } else if (current_page_ == 1) {
            const bool salvage = staged_check("DefaultSalvageUnreadableBlocks");
            set_field_enabled("DefaultAllowRecoveredOverwriteExisting", salvage);
            set_field_enabled("DefaultSalvageFillPattern", salvage);
            const int access_denied_index =
                field_index_for_key("DefaultTreatAccessDeniedAsContention");
            if (salvage && staged_check("DefaultTreatAccessDeniedAsContention")) {
                staged_bools_["DefaultTreatAccessDeniedAsContention"] = false;
                auto live = page_checks_.find(access_denied_index);
                if (live != page_checks_.end()) {
                    live->second = false;
                    InvalidateRect(page_controls_[access_denied_index], nullptr, FALSE);
                }
            }
            set_field_enabled("DefaultTreatAccessDeniedAsContention", !salvage);
            const bool map = staged_check("DefaultUseBadRangeMap");
            set_field_enabled("DefaultSkipKnownBadRanges", map);
            set_field_enabled("DefaultBadRangeMapMaxAgeDays", map);
        } else if (current_page_ == 2) {
            const bool fragile = staged_check("DefaultFragileMediaMode");
            for (std::string_view key : {"DefaultSkipFileOnFirstReadError",
                                         "DefaultPersistFragileSkipsAcrossResume",
                                         "DefaultFragileFailureWindowSeconds",
                                         "DefaultFragileFailureThreshold",
                                         "DefaultFragileCooldownSeconds"}) {
                set_field_enabled(key, fragile);
            }
        } else if (current_page_ == 4) {
            const bool verify = staged_check("DefaultVerifyAfterCopy");
            set_field_enabled("DefaultVerificationMode", verify);
            set_field_enabled("DefaultVerificationHashAlgorithm", verify);
            const bool sampled = verify && staged_combo("DefaultVerificationMode") == "sampled";
            set_field_enabled("DefaultSampleVerificationChunkKb", sampled);
            set_field_enabled("DefaultSampleVerificationChunkCount", sampled);
        }
    }

    bool harvest_current_page(bool validate = false) {
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
                    wchar_t* end = nullptr;
                    errno = 0;
                    long parsed = std::wcstol(buffer, &end, 10);
                    const bool valid = end != buffer && end != nullptr && *end == L'\0' &&
                                       errno != ERANGE && parsed >= field.min && parsed <= field.max;
                    if (!valid) {
                        if (validate) {
                            validation_error_ = storage::fsutil::utf8_to_wide(field.label) +
                                                L" must be a whole number from " +
                                                std::to_wstring(field.min) + L" to " +
                                                std::to_wstring(field.max) + L".";
                            validation_control_ = control;
                            return false;
                        }
                        break;
                    }
                    staged_ints_[field.key] = static_cast<int>(parsed);
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
        return true;
    }

    void build_page(int page) {
        destroy_page_controls();
        current_page_ = page;

        HWND hwnd = host_.hwnd();
        UINT dpi = dpi_;
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
                tooltips_.add(label, section_help(field.label));
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
                tooltips_.add(check, field_help(field.key));
                continue;
            }

            // Labeled field: label + control on a shared baseline.
            HWND label = CreateWindowExW(
                0, L"STATIC", storage::fsutil::utf8_to_wide(field.label).c_str(),
                WS_CHILD | WS_VISIBLE | SS_ENDELLIPSIS | SS_NOPREFIX | SS_CENTERIMAGE, x, y,
                label_width, row_height, hwnd, nullptr, GetModuleHandleW(nullptr), nullptr);
            SendMessageW(label, WM_SETFONT, reinterpret_cast<WPARAM>(font_), TRUE);
            page_labels_.push_back(label);
            page_field_labels_[index] = label;
            tooltips_.add(label, field_help(field.key));

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
                const int item_height = scale(22);
                SendMessageW(combo, CB_SETITEMHEIGHT, static_cast<WPARAM>(-1), item_height);
                SendMessageW(combo, CB_SETITEMHEIGHT, 0, item_height);
                apply_combo_theme(combo, theme_.dark);
                page_controls_[index] = combo;
                tooltips_.add(combo, field_help(field.key));
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
            tooltips_.add(edit, field_help(field.key));
        }
        sync_page_dependencies();
        tooltips_.update_layout();
        InvalidateRect(hwnd, nullptr, TRUE);
    }

    void on_create() {
        HWND hwnd = host_.hwnd();
        window_brush_ = CreateSolidBrush(theme_.window);
        edit_brush_ = CreateSolidBrush(theme_.edit);
        panel_brush_ = CreateSolidBrush(theme_.panel);
        rebuild_fonts(GetDpiForWindow(hwnd));
        UINT dpi = dpi_;
        auto scale = [dpi](int value) { return MulDiv(value, static_cast<int>(dpi), 96); };

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
            return button;
        };
        save_button_ = make_button(IdSaveButton, L"Save", client.right - scale(200));
        cancel_button_ = make_button(IdCancelButton, L"Cancel", client.right - scale(100));

        tooltips_.create(hwnd, dpi_, theme_.dark);
        tooltips_.add(nav_list_, L"Choose a settings category. Values on the current page are retained when you switch categories.");
        tooltips_.add(save_button_, L"Validate and save every settings page, then apply supported appearance changes to the main window.");
        tooltips_.add(cancel_button_, L"Discard all staged changes from every settings page and close the dialog.");

        relayout_shell(hwnd);
        build_page(0);
    }

    void draw_nav_item(const DRAWITEMSTRUCT& draw) {
        const bool selected = (draw.itemState & ODS_SELECTED) != 0;
        const UINT dpi = dpi_ == 0 ? 96 : dpi_;
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
            case WM_SIZE:
                relayout_shell(hwnd);
                handled = true;
                return 0;
            case WM_DPICHANGED: {
                harvest_current_page();
                rebuild_fonts(HIWORD(wparam) != 0 ? HIWORD(wparam) : LOWORD(wparam));
                const RECT* suggested = reinterpret_cast<const RECT*>(lparam);
                SetWindowPos(hwnd, nullptr, suggested->left, suggested->top,
                             suggested->right - suggested->left,
                             suggested->bottom - suggested->top,
                             SWP_NOZORDER | SWP_NOACTIVATE);
                SendMessageW(nav_list_, WM_SETFONT, reinterpret_cast<WPARAM>(font_), TRUE);
                SendMessageW(save_button_, WM_SETFONT, reinterpret_cast<WPARAM>(font_), TRUE);
                SendMessageW(cancel_button_, WM_SETFONT, reinterpret_cast<WPARAM>(font_), TRUE);
                SendMessageW(nav_list_, LB_SETITEMHEIGHT, 0,
                             MulDiv(32, static_cast<int>(dpi_), 96));
                tooltips_.update_dpi(dpi_);
                relayout_shell(hwnd);
                build_page(current_page_);
                apply_window_icons(hwnd);
                RedrawWindow(hwnd, nullptr, nullptr,
                             RDW_INVALIDATE | RDW_ERASE | RDW_ALLCHILDREN);
                handled = true;
                return 0;
            }
            case WM_ERASEBKGND: {
                HDC dc = reinterpret_cast<HDC>(wparam);
                RECT client;
                GetClientRect(hwnd, &client);
                themedraw::fill_rect(dc, client, theme_.window);
                // Left navigation rail: a panel-colored sidebar with a 1px edge.
                UINT dpi = dpi_;
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
                        sync_page_dependencies();
                        handled = true;
                        return 0;
                    }
                }
                if (id >= IdFieldBase && HIWORD(wparam) == CBN_SELCHANGE) {
                    sync_page_dependencies();
                    handled = true;
                    return 0;
                }
                if (id == IdSaveButton) {
                    if (save_all()) host_.destroy();
                    handled = true;
                    return 0;
                }
                if (id == IdCancelButton) {
                    host_.destroy();
                    handled = true;
                    return 0;
                }
                return 0;
            }
            case WM_CLOSE:
                host_.destroy();
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
    UINT dpi_ = 96;
    TooltipManager tooltips_;

    AboutDialog(const ThemePalette& theme, AppSettings& settings, UINT check_updates_command)
        : theme_(theme), settings_(settings), check_updates_command_(check_updates_command) {
        auto_update_ = settings_.get_bool("CheckUpdatesOnLaunch", false);
    }

    ~AboutDialog() {
        if (font_ != nullptr) DeleteObject(font_);
        if (title_font_ != nullptr) DeleteObject(title_font_);
        if (window_brush_ != nullptr) DeleteObject(window_brush_);
    }

    void rebuild_fonts(UINT monitor_dpi) {
        dpi_ = ui_layout_dpi(monitor_dpi, theme_.density_percent, theme_.scale_percent);
        if (font_ != nullptr) DeleteObject(font_);
        if (title_font_ != nullptr) DeleteObject(title_font_);
        font_ = CreateFontW(-MulDiv(9, static_cast<int>(dpi_), 72), 0, 0, 0, FW_NORMAL,
                            FALSE, FALSE, FALSE, DEFAULT_CHARSET, OUT_DEFAULT_PRECIS,
                            CLIP_DEFAULT_PRECIS, CLEARTYPE_QUALITY, DEFAULT_PITCH, L"Segoe UI");
        title_font_ = CreateFontW(-MulDiv(22, static_cast<int>(dpi_), 72), 0, 0, 0,
                                  FW_SEMIBOLD, FALSE, FALSE, FALSE, DEFAULT_CHARSET,
                                  OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, CLEARTYPE_QUALITY,
                                  DEFAULT_PITCH, L"Segoe UI");
    }

    void update_large_icon() {
        if (icon_ == nullptr) return;
        const int size = MulDiv(48, static_cast<int>(dpi_), 96);
        SendMessageW(icon_, STM_SETICON,
                     reinterpret_cast<WPARAM>(load_app_icon(size, size)), 0);
    }

    void run(HWND owner) {
        owner_ = owner;
        host_.set_density_percent(theme_.density_percent);
        host_.set_scale_percent(theme_.scale_percent);
        HWND hwnd = host_.create(
            owner, L"XactCopyAboutDlg", L"About XactCopy", 480, 400,
            [this](HWND h, UINT m, WPARAM w, LPARAM l, bool& handled) {
                return proc(h, m, w, l, handled);
            });
        if (hwnd == nullptr) return;
        apply_window_icons(hwnd);
        set_dark_title_bar(hwnd, theme_.dark && theme_.themed_chrome);
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
                rebuild_fonts(GetDpiForWindow(hwnd));
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
                update_large_icon();
                title_ = make_static(L"XactCopy", SS_NOPREFIX);
                description_ = make_static(
                    L"Verified file copier and recovery tool for unstable media.",
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
                tooltips_.create(hwnd, dpi_, theme_.dark);
                tooltips_.add(icon_, L"XactCopy application icon.");
                tooltips_.add(title_, L"XactCopy version and build information is listed below.");
                tooltips_.add(description_, L"XactCopy is designed for resilient copying, recovery-oriented reads, and read-only media scans.");
                tooltips_.add(metadata_[0], L"The semantic version of this XactCopy executable.");
                tooltips_.add(metadata_[1], L"The date and time at which this executable was compiled.");
                tooltips_.add(metadata_[2], L"The primary author of XactCopy.");
                tooltips_.add(metadata_[3], L"The licence under which XactCopy is distributed.");
                tooltips_.add(components_title_, L"Native Windows services and libraries used by XactCopy are listed below.");
                tooltips_.add(components_, L"Native Windows components used for hashing, updates, themed controls, and supervised worker communication.");
                tooltips_.add(auto_update_check_, L"Check GitHub for a newer XactCopy release when the application starts.");
                tooltips_.add(check_btn_, L"Check for a newer release now. If one is available, XactCopy shows its notes and verified installer package.");
                tooltips_.add(ok_btn_, L"Close the About dialog.");
                layout(hwnd);
                SetFocus(ok_btn_);
                handled = true;
                return 0;
            }
            case WM_SIZE:
                if (ok_btn_ != nullptr) layout(hwnd);
                handled = true;
                return 0;
            case WM_DPICHANGED: {
                rebuild_fonts(HIWORD(wparam) != 0 ? HIWORD(wparam) : LOWORD(wparam));
                const RECT* suggested = reinterpret_cast<const RECT*>(lparam);
                SetWindowPos(hwnd, nullptr, suggested->left, suggested->top,
                             suggested->right - suggested->left,
                             suggested->bottom - suggested->top,
                             SWP_NOZORDER | SWP_NOACTIVATE);
                for (HWND control : {title_, description_, metadata_[0], metadata_[1], metadata_[2],
                                     metadata_[3], components_title_, components_,
                                     auto_update_check_, check_btn_, ok_btn_}) {
                    SendMessageW(control, WM_SETFONT,
                                 reinterpret_cast<WPARAM>(control == title_ ? title_font_ : font_),
                                 TRUE);
                }
                update_large_icon();
                tooltips_.update_dpi(dpi_);
                apply_window_icons(hwnd);
                layout(hwnd);
                RedrawWindow(hwnd, nullptr, nullptr,
                             RDW_INVALIDATE | RDW_ERASE | RDW_ALLCHILDREN);
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
                RECT client;
                GetClientRect(hwnd, &client);
                UINT dpi = dpi_;
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
                    const bool previous = auto_update_;
                    auto_update_ = !previous;
                    if (!settings_.update_and_save([&](AppSettings& target) {
                            target.set_bool("CheckUpdatesOnLaunch", auto_update_);
                        })) {
                        auto_update_ = previous;
                        themed_message_box(
                            hwnd, theme_,
                            storage::fsutil::utf8_to_wide(settings_.last_save_error()),
                            L"Setting could not be saved", MessageIcon::Error);
                    }
                    InvalidateRect(auto_update_check_, nullptr, FALSE);
                    handled = true;
                    return 0;
                }
                if (id == IdOkButton) {
                    host_.destroy();
                } else if (id == IdCheckButton && check_updates_command_ != 0) {
                    HWND owner = owner_;
                    UINT cmd = check_updates_command_;
                    host_.destroy();
                    PostMessageW(owner, WM_COMMAND, MAKEWPARAM(cmd, 0), 0);
                }
                handled = true;
                return 0;
            }
            case WM_CLOSE:
                host_.destroy();
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
        UINT dpi = dpi_;
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
        tooltips_.update_layout();
        InvalidateRect(hwnd, nullptr, TRUE);
    }
};

} // namespace xact::ui
