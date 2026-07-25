// -----------------------------------------------------------------------------
// File: cpp\src\ui\icons.h
// Purpose: Fluent UI System Icons for owner-drawn buttons. Icons ship as 24x24
//          8-bit coverage masks (generated from the pinned Fluent SVG subset);
//          at draw time a mask is bilinearly resampled to the requested size,
//          tinted with the theme color into a premultiplied 32bpp DIB, cached by
//          (icon, color, size), and AlphaBlended so edges stay smooth at any DPI.
// -----------------------------------------------------------------------------

#pragma once

#ifndef NOMINMAX
#define NOMINMAX
#endif
#include <windows.h>

#include <algorithm>
#include <array>
#include <cmath>
#include <cstdint>
#include <string>
#include <vector>

namespace xact::ui {

enum class ButtonIcon {
    None,
    Start,     // play / resume
    Pause,
    Cancel,
    Open,      // browse / folder
    Settings,
    Info,      // about
    Refresh,
    Delete,
    Save,      // archive glyph
    Up,
    Down,
    View,
    Verify,    // check / test
};

namespace icon_detail {

constexpr int kSourceSize = 24;

#include "icon_masks.inc"

inline const std::uint8_t* mask_for_icon(ButtonIcon icon) {
    switch (icon) {
        case ButtonIcon::Start: return kResumeMask.data();
        case ButtonIcon::Pause: return kPauseMask.data();
        case ButtonIcon::Cancel: return kCancelMask.data();
        case ButtonIcon::Open: return kOpenMask.data();
        case ButtonIcon::Settings: return kSettingsMask.data();
        case ButtonIcon::Info: return kInfoMask.data();
        case ButtonIcon::Refresh: return kRefreshMask.data();
        case ButtonIcon::Delete: return kDeleteMask.data();
        case ButtonIcon::Save: return kArchiveMask.data();
        case ButtonIcon::Up: return kUpMask.data();
        case ButtonIcon::Down: return kExtractMask.data();
        case ButtonIcon::View: return kViewMask.data();
        case ButtonIcon::Verify: return kTestMask.data();
        case ButtonIcon::None: break;
    }
    return nullptr;
}

// Bilinear sample of the 24x24 coverage mask.
inline std::uint8_t sample_mask(const std::uint8_t* mask, double source_x, double source_y) {
    const int x0 = std::clamp(static_cast<int>(std::floor(source_x)), 0, kSourceSize - 1);
    const int y0 = std::clamp(static_cast<int>(std::floor(source_y)), 0, kSourceSize - 1);
    const int x1 = std::min(x0 + 1, kSourceSize - 1);
    const int y1 = std::min(y0 + 1, kSourceSize - 1);
    const double fx = std::clamp(source_x - std::floor(source_x), 0.0, 1.0);
    const double fy = std::clamp(source_y - std::floor(source_y), 0.0, 1.0);
    auto at = [mask](int x, int y) { return static_cast<double>(mask[y * kSourceSize + x]); };
    const double top = at(x0, y0) * (1.0 - fx) + at(x1, y0) * fx;
    const double bottom = at(x0, y1) * (1.0 - fx) + at(x1, y1) * fx;
    return static_cast<std::uint8_t>(std::lround(top * (1.0 - fy) + bottom * fy));
}

struct CachedIcon {
    ButtonIcon icon = ButtonIcon::None;
    COLORREF color = 0;
    int size = 0;
    HBITMAP bitmap = nullptr;
};

class IconCache {
public:
    HBITMAP get(ButtonIcon icon, COLORREF color, int size) {
        for (const auto& entry : entries_) {
            if (entry.icon == icon && entry.color == color && entry.size == size) {
                return entry.bitmap;
            }
        }
        HBITMAP bitmap = create(icon, color, size);
        if (bitmap != nullptr) entries_.push_back({icon, color, size, bitmap});
        return bitmap;
    }

private:
    static HBITMAP create(ButtonIcon icon, COLORREF color, int size) {
        const std::uint8_t* mask = mask_for_icon(icon);
        if (mask == nullptr || size <= 0) return nullptr;

        BITMAPINFO info{};
        info.bmiHeader.biSize = sizeof(info.bmiHeader);
        info.bmiHeader.biWidth = size;
        info.bmiHeader.biHeight = -size; // top-down
        info.bmiHeader.biPlanes = 1;
        info.bmiHeader.biBitCount = 32;
        info.bmiHeader.biCompression = BI_RGB;

        void* bits = nullptr;
        HBITMAP bitmap = CreateDIBSection(nullptr, &info, DIB_RGB_COLORS, &bits, nullptr, 0);
        if (bitmap == nullptr || bits == nullptr) return nullptr;

        auto* pixels = static_cast<std::uint32_t*>(bits);
        const double scale = static_cast<double>(kSourceSize) / static_cast<double>(size);
        const int red = GetRValue(color);
        const int green = GetGValue(color);
        const int blue = GetBValue(color);
        for (int y = 0; y < size; ++y) {
            const double source_y = (y + 0.5) * scale - 0.5;
            for (int x = 0; x < size; ++x) {
                const double source_x = (x + 0.5) * scale - 0.5;
                const std::uint8_t alpha = sample_mask(mask, source_x, source_y);
                // Premultiplied BGRA for AC_SRC_ALPHA blending.
                pixels[y * size + x] =
                    (static_cast<std::uint32_t>(alpha) << 24) |
                    (static_cast<std::uint32_t>(red * alpha / 255) << 16) |
                    (static_cast<std::uint32_t>(green * alpha / 255) << 8) |
                    static_cast<std::uint32_t>(blue * alpha / 255);
            }
        }
        return bitmap;
    }

    std::vector<CachedIcon> entries_;
};

inline IconCache& icon_cache() {
    static IconCache cache;
    return cache;
}

inline std::wstring lower_text(const wchar_t* text) {
    std::wstring out(text != nullptr ? text : L"");
    for (wchar_t& c : out) c = static_cast<wchar_t>(towlower(c));
    return out;
}

inline bool contains(const std::wstring& haystack, const wchar_t* needle) {
    return haystack.find(needle) != std::wstring::npos;
}

} // namespace icon_detail

// Draws a theme-tinted Fluent icon centred in `bounds` (which should be square).
inline void draw_button_icon(HDC dc, ButtonIcon icon, const RECT& bounds, COLORREF color,
                             int size) {
    if (dc == nullptr || icon == ButtonIcon::None || size <= 0) return;
    HBITMAP bitmap = icon_detail::icon_cache().get(icon, color, size);
    if (bitmap == nullptr) return;
    HDC memory_dc = CreateCompatibleDC(dc);
    if (memory_dc == nullptr) return;
    HGDIOBJ old_bitmap = SelectObject(memory_dc, bitmap);
    const int x = bounds.left + (bounds.right - bounds.left - size) / 2;
    const int y = bounds.top + (bounds.bottom - bounds.top - size) / 2;
    BLENDFUNCTION blend{AC_SRC_OVER, 0, 255, AC_SRC_ALPHA};
    AlphaBlend(dc, x, y, size, size, memory_dc, 0, 0, size, size, blend);
    SelectObject(memory_dc, old_bitmap);
    DeleteDC(memory_dc);
}

// Maps a button's caption (and dialog id) to an icon, so owner-drawn buttons
// pick one up without every call site being rewritten.
inline ButtonIcon icon_for_button(HWND button) {
    wchar_t raw[256]{};
    GetWindowTextW(button, raw, static_cast<int>(std::size(raw)));
    const std::wstring text = icon_detail::lower_text(raw);
    using icon_detail::contains;
    const int id = GetDlgCtrlID(button);

    if (text.empty()) return ButtonIcon::None;
    if (id == IDCANCEL || contains(text, L"cancel") || contains(text, L"close") ||
        text == L"no") {
        return ButtonIcon::Cancel;
    }
    if (contains(text, L"pause")) return ButtonIcon::Pause;
    if (contains(text, L"start") || contains(text, L"resume") || contains(text, L"run") ||
        contains(text, L"install") || text == L"yes" || id == IDYES) {
        return ButtonIcon::Start;
    }
    if (contains(text, L"browse") || contains(text, L"open") || contains(text, L"folder")) {
        return ButtonIcon::Open;
    }
    if (contains(text, L"settings") || contains(text, L"options")) return ButtonIcon::Settings;
    if (contains(text, L"about")) return ButtonIcon::Info;
    if (contains(text, L"refresh") || contains(text, L"update") || contains(text, L"check")) {
        return ButtonIcon::Refresh;
    }
    if (contains(text, L"delete") || contains(text, L"remove") || contains(text, L"clear")) {
        return ButtonIcon::Delete;
    }
    if (contains(text, L"save")) return ButtonIcon::Save;
    if (contains(text, L"queue")) return ButtonIcon::Down;
    if (contains(text, L"view")) return ButtonIcon::View;
    if (text == L"ok") return ButtonIcon::Verify;
    return ButtonIcon::None;
}

} // namespace xact::ui
