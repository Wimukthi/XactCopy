// Application-icon helpers: load the embedded XactCopy icon at a requested size
// and stamp it onto window classes / window frames (title-bar + Alt-Tab icons).
#pragma once

#ifndef NOMINMAX
#define NOMINMAX
#endif
#include <windows.h>

#include "resource.h"

namespace xact::ui {

// Loads the embedded IDI_XACTCOPY icon at the requested pixel size, falling back
// to the generic application icon if the resource is missing.
inline HICON load_app_icon(int width, int height) {
    HINSTANCE instance = GetModuleHandleW(nullptr);
    HICON icon = static_cast<HICON>(LoadImageW(instance, MAKEINTRESOURCEW(IDI_XACTCOPY), IMAGE_ICON,
                                               width, height, LR_DEFAULTCOLOR | LR_SHARED));
    return icon != nullptr ? icon : LoadIconW(nullptr, IDI_APPLICATION);
}

// Fills a window class's large/small icons before RegisterClass.
inline void assign_window_class_icons(WNDCLASSW& window_class) {
    window_class.hIcon =
        load_app_icon(GetSystemMetrics(SM_CXICON), GetSystemMetrics(SM_CYICON));
}

// Stamps DPI-appropriate title-bar (small) and Alt-Tab (big) icons on a window.
inline void apply_window_icons(HWND window) {
    const UINT dpi = GetDpiForWindow(window);
    SendMessageW(window, WM_SETICON, ICON_BIG,
                 reinterpret_cast<LPARAM>(load_app_icon(GetSystemMetricsForDpi(SM_CXICON, dpi),
                                                        GetSystemMetricsForDpi(SM_CYICON, dpi))));
    SendMessageW(window, WM_SETICON, ICON_SMALL,
                 reinterpret_cast<LPARAM>(load_app_icon(GetSystemMetricsForDpi(SM_CXSMICON, dpi),
                                                        GetSystemMetricsForDpi(SM_CYSMICON, dpi))));
}

} // namespace xact::ui
