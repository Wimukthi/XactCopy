// -----------------------------------------------------------------------------
// File: src\ui\theme.h
// Purpose: XactCopy-specific palettes and owner-drawn controls. Generic Win32
//          dark-mode integration is provided by Wimukthi.Win32Theme.
// -----------------------------------------------------------------------------

#pragma once

#include <algorithm>
#include <string>

#ifndef NOMINMAX
#define NOMINMAX
#endif
#ifndef WIN32_LEAN_AND_MEAN
#define WIN32_LEAN_AND_MEAN
#endif
#include <windows.h>
#include <dwmapi.h>
#include <gdiplus.h>

#include "icons.h"
#include <wimukthi/win32_theme.hpp>

#if defined(_MSC_VER)
#pragma comment(lib, "dwmapi.lib")
#pragma comment(lib, "gdiplus.lib")
#pragma comment(lib, "msimg32.lib")
#endif

namespace xact::ui {

// GDI+ is used only for the antialiased glyphs (checkmarks, radio dots); it is
// started once on first use and shut down at exit.
namespace gdip {

class Session {
public:
    Session() {
        Gdiplus::GdiplusStartupInput input;
        ready_ = Gdiplus::GdiplusStartup(&token_, &input, nullptr) == Gdiplus::Ok;
    }
    ~Session() {
        if (ready_) Gdiplus::GdiplusShutdown(token_);
    }
    Session(const Session&) = delete;
    Session& operator=(const Session&) = delete;
    bool ready() const { return ready_; }

private:
    ULONG_PTR token_ = 0;
    bool ready_ = false;
};

inline bool ready() {
    static Session session;
    return session.ready();
}

inline Gdiplus::Color color_of(COLORREF color) {
    return Gdiplus::Color(255, GetRValue(color), GetGValue(color), GetBValue(color));
}

// Gamma-corrected coverage keeps saturated accents from looking washed out
// while retaining smooth round edges.
inline void configure(Gdiplus::Graphics& graphics) {
    graphics.SetCompositingMode(Gdiplus::CompositingModeSourceOver);
    graphics.SetCompositingQuality(Gdiplus::CompositingQualityGammaCorrected);
    graphics.SetSmoothingMode(Gdiplus::SmoothingModeAntiAlias);
    graphics.SetPixelOffsetMode(Gdiplus::PixelOffsetModeHighQuality);
}

} // namespace gdip

struct ThemePalette {
    bool dark = true;
    // Settings that affect top-level windows/layout are carried with the
    // palette so every native dialog opened by the main window observes the
    // same appearance choice.
    bool themed_chrome = true;
    int density_percent = 100;
    int scale_percent = 100;
    COLORREF accent = RGB(0x5A, 0x78, 0xC8);
    COLORREF window = RGB(32, 32, 32);
    COLORREF panel = RGB(45, 45, 48);
    COLORREF edit = RGB(37, 37, 38);
    COLORREF text = RGB(241, 241, 241);
    COLORREF muted_text = RGB(180, 180, 180);
    COLORREF border = RGB(64, 64, 64);
    COLORREF button = RGB(45, 45, 48);
    COLORREF button_hot = 0;
    COLORREF button_pressed = 0;
    COLORREF selection = 0;
    COLORREF selection_text = RGB(255, 255, 255);
    COLORREF focus = 0;
    // Log severity colors.
    COLORREF log_error = RGB(240, 110, 110);
    COLORREF log_warning = RGB(228, 190, 100);
    COLORREF log_success = RGB(120, 205, 130);
    COLORREF log_supervisor = RGB(130, 170, 235);
};

inline COLORREF blend_color(COLORREF base, COLORREF overlay, int overlay_percent) {
    overlay_percent = std::clamp(overlay_percent, 0, 100);
    const int base_percent = 100 - overlay_percent;
    return RGB((GetRValue(base) * base_percent + GetRValue(overlay) * overlay_percent) / 100,
               (GetGValue(base) * base_percent + GetGValue(overlay) * overlay_percent) / 100,
               (GetBValue(base) * base_percent + GetBValue(overlay) * overlay_percent) / 100);
}

inline COLORREF readable_text_color(COLORREF background) {
    const int luminance = GetRValue(background) * 299 + GetGValue(background) * 587 +
                          GetBValue(background) * 114;
    return luminance > 150000 ? RGB(0, 0, 0) : RGB(255, 255, 255);
}

inline bool high_contrast_enabled() {
    return wimukthi::win32_theme::is_high_contrast();
}

inline bool system_prefers_dark_mode() {
    if (high_contrast_enabled()) return false;
    return wimukthi::win32_theme::is_dark();
}

// Parses "#RRGGBB" (or "RRGGBB"); falls back on malformed input.
inline COLORREF parse_color_hex(const std::string& text, COLORREF fallback) {
    std::string hex = text;
    if (!hex.empty() && hex[0] == '#') hex = hex.substr(1);
    if (hex.size() != 6) return fallback;
    unsigned value = 0;
    for (char c : hex) {
        value <<= 4;
        if (c >= '0' && c <= '9') value |= static_cast<unsigned>(c - '0');
        else if (c >= 'a' && c <= 'f') value |= static_cast<unsigned>(c - 'a' + 10);
        else if (c >= 'A' && c <= 'F') value |= static_cast<unsigned>(c - 'A' + 10);
        else return fallback;
    }
    return RGB((value >> 16) & 0xFF, (value >> 8) & 0xFF, value & 0xFF);
}

inline int ui_density_percent(const std::string& setting) {
    std::string normalized;
    normalized.reserve(setting.size());
    for (char c : setting) {
        if (c >= 'A' && c <= 'Z') c = static_cast<char>(c - 'A' + 'a');
        normalized.push_back(c);
    }
    if (normalized == "compact") return 92;
    if (normalized == "comfortable") return 108;
    return 100;
}

inline UINT ui_layout_dpi(UINT monitor_dpi, int density_percent, int scale_percent = 100) {
    const int density = std::clamp(density_percent, 80, 120);
    const int scale = std::clamp(scale_percent, 50, 250);
    const int combined = MulDiv(scale, density, 100);
    return static_cast<UINT>(MulDiv(static_cast<int>(monitor_dpi), combined, 100));
}

inline COLORREF system_accent_color(COLORREF fallback) {
    DWORD colorization = 0;
    BOOL opaque = FALSE;
    if (SUCCEEDED(DwmGetColorizationColor(&colorization, &opaque))) {
        return RGB((colorization >> 16) & 0xFF, (colorization >> 8) & 0xFF, colorization & 0xFF);
    }
    return fallback;
}

// Builds the palette from the XactCopy settings values: Theme = dark|light|
// classic|system, AccentColorMode = auto|system|custom, AccentColorHex.
inline ThemePalette make_theme(const std::string& theme_setting,
                               const std::string& accent_mode,
                               const std::string& accent_hex) {
    ThemePalette theme;
    if (theme_setting == "dark") theme.dark = true;
    else if (theme_setting == "light" || theme_setting == "classic") theme.dark = false;
    else theme.dark = system_prefers_dark_mode();

    if (high_contrast_enabled()) {
        theme.dark = false;
        theme.window = GetSysColor(COLOR_WINDOW);
        theme.panel = GetSysColor(COLOR_BTNFACE);
        theme.edit = GetSysColor(COLOR_WINDOW);
        theme.text = GetSysColor(COLOR_WINDOWTEXT);
        theme.muted_text = GetSysColor(COLOR_GRAYTEXT);
        theme.border = GetSysColor(COLOR_WINDOWFRAME);
        theme.button = GetSysColor(COLOR_BTNFACE);
        theme.accent = GetSysColor(COLOR_HIGHLIGHT);
        theme.selection = GetSysColor(COLOR_HIGHLIGHT);
        theme.selection_text = GetSysColor(COLOR_HIGHLIGHTTEXT);
        theme.button_hot = theme.button;
        theme.button_pressed = theme.button;
        theme.focus = GetSysColor(COLOR_HOTLIGHT);
        return theme;
    }

    COLORREF default_accent = parse_color_hex(accent_hex, RGB(0x5A, 0x78, 0xC8));
    if (accent_mode == "system") theme.accent = system_accent_color(default_accent);
    else if (accent_mode == "custom") theme.accent = default_accent;
    else theme.accent = default_accent;

    if (theme.dark) {
        theme.window = RGB(32, 32, 32);
        theme.panel = RGB(45, 45, 48);
        theme.edit = RGB(37, 37, 38);
        theme.text = RGB(241, 241, 241);
        theme.muted_text = RGB(180, 180, 180);
        theme.border = RGB(64, 64, 64);
        theme.button = RGB(45, 45, 48);
        theme.button_hot = blend_color(theme.button, theme.accent, 16);
        theme.button_pressed = blend_color(theme.button, theme.accent, 28);
        theme.selection = blend_color(theme.window, theme.accent, 44);
    } else {
        theme.window = RGB(250, 250, 250);
        theme.panel = RGB(238, 238, 240);
        theme.edit = RGB(255, 255, 255);
        theme.text = RGB(20, 20, 20);
        theme.muted_text = RGB(96, 96, 100);
        theme.border = RGB(190, 190, 195);
        theme.button = RGB(238, 238, 240);
        theme.button_hot = blend_color(theme.button, theme.accent, 12);
        theme.button_pressed = blend_color(theme.button, theme.accent, 22);
        theme.selection = theme.accent;
        theme.log_error = RGB(190, 40, 40);
        theme.log_warning = RGB(150, 110, 10);
        theme.log_success = RGB(30, 130, 55);
        theme.log_supervisor = RGB(40, 80, 175);
    }
    theme.selection_text = readable_text_color(theme.selection);
    theme.focus = theme.accent;
    return theme;
}

inline wimukthi::win32_theme::Mode shared_theme_mode(const std::string& setting) {
    using Mode = wimukthi::win32_theme::Mode;
    if (setting == "dark") return Mode::dark;
    if (setting == "light") return Mode::light;
    if (setting == "classic") return Mode::classic;
    return Mode::system;
}

inline void configure_theme_engine(const std::string& setting, const ThemePalette& theme) {
    using namespace wimukthi::win32_theme;
    Configuration configuration;
    configuration.mode = shared_theme_mode(setting);
    configuration.use_custom_palette = !high_contrast_enabled() && setting != "classic";
    configuration.palette = {
        theme.window,
        theme.edit,
        theme.button_hot,
        theme.window,
        blend_color(theme.window, theme.log_error, 18),
        theme.text,
        theme.muted_text,
        theme.muted_text,
        theme.accent,
        theme.border,
        theme.focus,
        theme.border,
        theme.selection,
        theme.edit,
        theme.text,
        theme.border,
        theme.panel,
        theme.button_hot,
        theme.text,
        theme.border,
    };
    configure(configuration);
}

inline bool handle_theme_setting_change(LPARAM lparam) {
    return wimukthi::win32_theme::handle_setting_change(lparam);
}

inline void attach_theme_window(HWND hwnd) {
    wimukthi::win32_theme::AttachOptions options;
    options.child_controls = false;
    options.control_colours = false;
    options.custom_draw_notifications = false;
    wimukthi::win32_theme::attach(hwnd, options);
}

// Compatibility shim for existing dialog code. The shared framework owns the
// per-window UxTheme integration.
inline void allow_dark_mode_for_window(HWND hwnd, bool allow = true) {
    (void)allow;
    wimukthi::win32_theme::apply_control(hwnd);
}

inline void set_dark_title_bar(HWND hwnd, bool dark) {
    if (hwnd == nullptr) return;

    // The shared theme framework intentionally applies the configured system
    // title-bar mode. Standard chrome is the one setting it cannot infer from
    // the palette, so explicitly clear the Windows dark-title-bar attribute
    // when the user asks for an unthemed title bar.
    if (!dark) {
        constexpr DWORD UseImmersiveDarkMode = 20;
        constexpr DWORD UseImmersiveDarkModeLegacy = 19;
        BOOL enabled = FALSE;
        HRESULT result = DwmSetWindowAttribute(hwnd, UseImmersiveDarkMode, &enabled,
                                               sizeof(enabled));
        if (FAILED(result)) {
            DwmSetWindowAttribute(hwnd, UseImmersiveDarkModeLegacy, &enabled, sizeof(enabled));
        }
        return;
    }
    wimukthi::win32_theme::apply_title_bar(hwnd);
}

// DarkMode_Explorer gives native dark scrollbars on list/edit controls; the
// per-window opt-in must precede it so it renders consistently on a Light OS.
inline void apply_control_theme(HWND control, bool dark) {
    (void)dark;
    wimukthi::win32_theme::apply_control(control);
}

// Combos need the CFD theme class for a dark frame + dropdown arrow.
inline void apply_combo_theme(HWND combo, bool dark) {
    (void)dark;
    wimukthi::win32_theme::apply_combo_box(combo);
}

// ---------------------------------------------------------------------------
// Owner-drawn control painting (BS_OWNERDRAW / CBS_OWNERDRAWFIXED)
// ---------------------------------------------------------------------------

namespace themedraw {

inline void fill_rect(HDC dc, const RECT& rect, COLORREF color) {
    HBRUSH brush = CreateSolidBrush(color);
    FillRect(dc, &rect, brush);
    DeleteObject(brush);
}

inline void frame_rect(HDC dc, const RECT& rect, COLORREF color) {
    HBRUSH brush = CreateSolidBrush(color);
    FrameRect(dc, &rect, brush);
    DeleteObject(brush);
}

inline void draw_button(const DRAWITEMSTRUCT& draw, const ThemePalette& theme) {
    const bool disabled = (draw.itemState & ODS_DISABLED) != 0;
    const bool pressed = (draw.itemState & ODS_SELECTED) != 0;
    const bool hot = (draw.itemState & ODS_HOTLIGHT) != 0;
    const bool focused = (draw.itemState & ODS_FOCUS) != 0;

    RECT rect = draw.rcItem;
    const COLORREF fill = pressed ? theme.button_pressed
                          : (hot || focused) ? theme.button_hot
                                             : theme.button;
    fill_rect(draw.hDC, rect, fill);
    frame_rect(draw.hDC, rect, (focused || hot || pressed) ? theme.focus : theme.border);

    wchar_t text[128]{};
    GetWindowTextW(draw.hwndItem, text, 128);
    HFONT font = reinterpret_cast<HFONT>(SendMessageW(draw.hwndItem, WM_GETFONT, 0, 0));
    HGDIOBJ old_font = font != nullptr ? SelectObject(draw.hDC, font) : nullptr;
    SetBkMode(draw.hDC, TRANSPARENT);
    const COLORREF content = disabled ? theme.muted_text : theme.text;
    SetTextColor(draw.hDC, content);
    if (pressed) OffsetRect(&rect, 1, 1);

    // Pair the caption with a Fluent glyph when there's room for both; otherwise
    // fall back to centred text so narrow buttons never truncate.
    const ButtonIcon icon = icon_for_button(draw.hwndItem);
    bool drew_with_icon = false;
    if (icon != ButtonIcon::None) {
        const UINT dpi = GetDpiForWindow(draw.hwndItem);
        auto scale = [dpi](int v) { return MulDiv(v, static_cast<int>(dpi), 96); };
        SIZE text_size{};
        GetTextExtentPoint32W(draw.hDC, text, static_cast<int>(wcslen(text)), &text_size);
        const int icon_size = scale(16);
        const int gap = scale(6);
        const int content_width = icon_size + gap + text_size.cx;
        const int available = (rect.right - rect.left) - scale(10);
        if (content_width <= available) {
            const int left = rect.left + ((rect.right - rect.left) - content_width) / 2;
            RECT icon_rect{left, rect.top, left + icon_size, rect.bottom};
            draw_button_icon(draw.hDC, icon, icon_rect, content, icon_size);
            RECT text_rect = rect;
            text_rect.left = icon_rect.right + gap;
            text_rect.right = text_rect.left + text_size.cx;
            DrawTextW(draw.hDC, text, -1, &text_rect,
                      DT_LEFT | DT_VCENTER | DT_SINGLELINE | DT_NOPREFIX);
            drew_with_icon = true;
        }
    }
    if (!drew_with_icon) {
        DrawTextW(draw.hDC, text, -1, &rect,
                  DT_CENTER | DT_VCENTER | DT_SINGLELINE | DT_NOPREFIX | DT_END_ELLIPSIS);
    }
    if (old_font != nullptr) SelectObject(draw.hDC, old_font);
}

// Antialiased checkmark via GDI+ (round caps/joins), falling back to a plain GDI
// polyline if GDI+ is unavailable.
inline void draw_checkmark(HDC dc, const RECT& box, COLORREF color, UINT dpi) {
    auto scale = [dpi](int value) { return MulDiv(value, static_cast<int>(dpi), 96); };
    const int width = box.right - box.left;
    const int height = box.bottom - box.top;
    const POINT points[3] = {
        {box.left + width * 22 / 100, box.top + height * 52 / 100},
        {box.left + width * 42 / 100, box.top + height * 72 / 100},
        {box.left + width * 78 / 100, box.top + height * 30 / 100},
    };
    if (!gdip::ready()) {
        HPEN pen = CreatePen(PS_SOLID, std::max(1, scale(2)), color);
        HGDIOBJ old_pen = SelectObject(dc, pen);
        MoveToEx(dc, points[0].x, points[0].y, nullptr);
        LineTo(dc, points[1].x, points[1].y);
        LineTo(dc, points[2].x, points[2].y);
        SelectObject(dc, old_pen);
        DeleteObject(pen);
        return;
    }
    Gdiplus::Graphics graphics(dc);
    gdip::configure(graphics);
    const Gdiplus::REAL stroke = static_cast<Gdiplus::REAL>(scale(2));
    Gdiplus::Pen pen(gdip::color_of(color), stroke < 1.75f ? 1.75f : stroke);
    pen.SetLineCap(Gdiplus::LineCapRound, Gdiplus::LineCapRound, Gdiplus::DashCapRound);
    pen.SetLineJoin(Gdiplus::LineJoinRound);
    const Gdiplus::PointF path[3] = {
        {static_cast<Gdiplus::REAL>(points[0].x), static_cast<Gdiplus::REAL>(points[0].y)},
        {static_cast<Gdiplus::REAL>(points[1].x), static_cast<Gdiplus::REAL>(points[1].y)},
        {static_cast<Gdiplus::REAL>(points[2].x), static_cast<Gdiplus::REAL>(points[2].y)},
    };
    graphics.DrawLines(&pen, path, 3);
}

inline void draw_checkbox(const DRAWITEMSTRUCT& draw, const ThemePalette& theme, bool checked) {
    const bool disabled = (draw.itemState & ODS_DISABLED) != 0;
    const bool focused = (draw.itemState & ODS_FOCUS) != 0;
    fill_rect(draw.hDC, draw.rcItem, theme.window);

    const UINT dpi = GetDpiForWindow(draw.hwndItem);
    auto scale = [dpi](int value) { return MulDiv(value, static_cast<int>(dpi), 96); };
    const int box_size = scale(15);
    RECT box{
        draw.rcItem.left + scale(2),
        draw.rcItem.top + (draw.rcItem.bottom - draw.rcItem.top - box_size) / 2,
        draw.rcItem.left + scale(2) + box_size,
        draw.rcItem.top + (draw.rcItem.bottom - draw.rcItem.top + box_size) / 2,
    };
    fill_rect(draw.hDC, box, theme.edit);
    frame_rect(draw.hDC, box, focused ? theme.focus : theme.border);
    if (checked) {
        draw_checkmark(draw.hDC, box, disabled ? theme.muted_text : theme.accent, dpi);
    }

    wchar_t text[128]{};
    GetWindowTextW(draw.hwndItem, text, 128);
    HFONT font = reinterpret_cast<HFONT>(SendMessageW(draw.hwndItem, WM_GETFONT, 0, 0));
    HGDIOBJ old_font = font != nullptr ? SelectObject(draw.hDC, font) : nullptr;
    SetBkMode(draw.hDC, TRANSPARENT);
    SetTextColor(draw.hDC, disabled ? theme.muted_text : theme.text);
    RECT text_rect = draw.rcItem;
    text_rect.left = box.right + scale(6);
    DrawTextW(draw.hDC, text, -1, &text_rect,
              DT_LEFT | DT_VCENTER | DT_SINGLELINE | DT_NOPREFIX | DT_END_ELLIPSIS);
    if (old_font != nullptr) SelectObject(draw.hDC, old_font);
}

inline void draw_combo_item(const DRAWITEMSTRUCT& draw, const ThemePalette& theme) {
    const bool disabled = (draw.itemState & ODS_DISABLED) != 0;
    const bool selected = (draw.itemState & ODS_SELECTED) != 0;
    const COLORREF background = selected ? theme.selection : theme.edit;
    fill_rect(draw.hDC, draw.rcItem, background);

    if (draw.itemID != static_cast<UINT>(-1)) {
        const LRESULT length = SendMessageW(draw.hwndItem, CB_GETLBTEXTLEN, draw.itemID, 0);
        if (length >= 0) {
            std::wstring text(static_cast<std::size_t>(length) + 1, L'\0');
            SendMessageW(draw.hwndItem, CB_GETLBTEXT, draw.itemID,
                         reinterpret_cast<LPARAM>(text.data()));
            text.resize(static_cast<std::size_t>(length));

            const UINT dpi = GetDpiForWindow(draw.hwndItem);
            HFONT font = reinterpret_cast<HFONT>(SendMessageW(draw.hwndItem, WM_GETFONT, 0, 0));
            HGDIOBJ old_font = font != nullptr ? SelectObject(draw.hDC, font) : nullptr;
            SetBkMode(draw.hDC, TRANSPARENT);
            SetTextColor(draw.hDC, disabled ? theme.muted_text
                                            : selected ? theme.selection_text : theme.text);
            RECT text_rect = draw.rcItem;
            text_rect.left += MulDiv(7, static_cast<int>(dpi), 96);
            text_rect.right -= MulDiv(4, static_cast<int>(dpi), 96);
            DrawTextW(draw.hDC, text.c_str(), -1, &text_rect,
                      DT_LEFT | DT_VCENTER | DT_SINGLELINE | DT_NOPREFIX);
            if (old_font != nullptr) SelectObject(draw.hDC, old_font);
        }
    }

    if ((draw.itemState & ODS_FOCUS) != 0 && (draw.itemState & ODS_COMBOBOXEDIT) == 0) {
        RECT focus = draw.rcItem;
        InflateRect(&focus, -2, -2);
        frame_rect(draw.hDC, focus, theme.focus);
    }
}

} // namespace themedraw

// ---------------------------------------------------------------------------
// Log severity classification (mirrors the .NET UI's colorize-by-severity).
// ---------------------------------------------------------------------------

enum class LogSeverity { Normal, Error, Warning, Success, Supervisor };

inline LogSeverity classify_log_line(const std::string& line) {
    auto contains_ci = [&line](const char* fragment) {
        std::string lower;
        lower.reserve(line.size());
        for (char c : line) {
            lower.push_back(c >= 'A' && c <= 'Z' ? static_cast<char>(c - 'A' + 'a') : c);
        }
        return lower.find(fragment) != std::string::npos;
    };
    if (contains_ci("fatal") || contains_ci("failed") || contains_ci("error") ||
        contains_ci("mismatch")) {
        return LogSeverity::Error;
    }
    if (contains_ci("recovered") || contains_ci("salvag") || contains_ci("succeeded") ||
        contains_ci("completed:")) {
        return LogSeverity::Success;
    }
    if (contains_ci("retry") || contains_ci("warning") || contains_ci("stale") ||
        contains_ci("skipped") || contains_ci("waiting") || contains_ci("fallback")) {
        return LogSeverity::Warning;
    }
    if (contains_ci("[supervisor]") || contains_ci("[devfault]") ||
        contains_ci("[workertelemetry]")) {
        return LogSeverity::Supervisor;
    }
    return LogSeverity::Normal;
}

inline COLORREF severity_color(const ThemePalette& theme, LogSeverity severity) {
    switch (severity) {
        case LogSeverity::Error: return theme.log_error;
        case LogSeverity::Warning: return theme.log_warning;
        case LogSeverity::Success: return theme.log_success;
        case LogSeverity::Supervisor: return theme.log_supervisor;
        default: return theme.text;
    }
}

} // namespace xact::ui
