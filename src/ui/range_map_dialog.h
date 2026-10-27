#pragma once

#include <atomic>
#include <condition_variable>
#include <memory>
#include <mutex>
#include <thread>
#include <windowsx.h>

#include "dialogs.h"
#include "range_map.h"

namespace xact::ui {

struct RangeMapFile {
    std::wstring path;
    std::int64_t size = 0;
    std::vector<storage::RescueRange> ranges;
};

class RangeMapDialog {
public:
    static void show(HWND owner, const ThemePalette& theme,
                     std::vector<RangeMapFile> files) {
        RangeMapDialog dialog(theme, std::move(files));
        dialog.dpi_ = ui_layout_dpi(owner != nullptr ? GetDpiForWindow(owner) : GetDpiForSystem(),
                                    theme.density_percent, theme.scale_percent);
        dialog.host_.set_density_percent(theme.density_percent);
        dialog.host_.set_scale_percent(theme.scale_percent);
        HWND hwnd = dialog.host_.create(owner, L"XactCopyRangeMapDlg", L"Rescue range map",
            760, 500, [&dialog](HWND h, UINT m, WPARAM w, LPARAM l, bool& handled) {
                return dialog.proc(h, m, w, l, handled);
            }, true);
        if (hwnd == nullptr) return;
        apply_window_icons(hwnd);
        set_dark_title_bar(hwnd, theme.dark && theme.themed_chrome);
        dialog.host_.run_modal();
    }

private:
    static constexpr int IdFiles = 3201;
    static constexpr int IdMap = 3202;
    static constexpr int IdClose = 3203;
    static constexpr UINT GeometryReady = WM_APP + 0x5a1;
    struct GeometryResult {
        std::uint64_t generation = 0;
        std::vector<RangeMapSpan> spans;
    };

    RangeMapDialog(const ThemePalette& theme, std::vector<RangeMapFile> files)
        : theme_(theme), files_(std::move(files)) {}

    ~RangeMapDialog() {
        stop_worker();
        if (font_ != nullptr) DeleteObject(font_);
        if (window_brush_ != nullptr) DeleteObject(window_brush_);
        if (edit_brush_ != nullptr) DeleteObject(edit_brush_);
    }

    int scale(int value) const { return MulDiv(value, static_cast<int>(dpi_), 96); }

    void rebuild_font(UINT dpi) {
        dpi_ = ui_layout_dpi(dpi, theme_.density_percent, theme_.scale_percent);
        if (font_ != nullptr) DeleteObject(font_);
        font_ = CreateFontW(-MulDiv(9, static_cast<int>(dpi_), 72), 0, 0, 0, FW_NORMAL,
                            FALSE, FALSE, FALSE, DEFAULT_CHARSET, OUT_DEFAULT_PRECIS,
                            CLIP_DEFAULT_PRECIS, CLEARTYPE_QUALITY, DEFAULT_PITCH, L"Segoe UI");
    }

    void layout(HWND hwnd) {
        if (list_ == nullptr) return;
        RECT client{}; GetClientRect(hwnd, &client);
        const int margin = scale(16);
        const int width = std::max(1, static_cast<int>(client.right) - margin * 2);
        const int list_height = std::max(scale(80), static_cast<int>(client.bottom) - scale(205));
        MoveWindow(label_, margin, margin, width, scale(22), TRUE);
        MoveWindow(list_, margin, margin + scale(25), width, list_height, TRUE);
        const int map_y = margin + scale(33) + list_height;
        MoveWindow(map_, margin, map_y, width, scale(38), TRUE);
        MoveWindow(note_, margin, map_y + scale(43), width, scale(38), TRUE);
        MoveWindow(close_, client.right - margin - scale(90), client.bottom - scale(38),
                   scale(90), scale(28), TRUE);
        request_geometry();
        InvalidateRect(hwnd, nullptr, TRUE);
    }

    COLORREF color(storage::RescueRangeState state) const {
        switch (state) {
            case storage::RescueRangeState::Good: return theme_.log_success;
            case storage::RescueRangeState::Recovered: return theme_.accent;
            case storage::RescueRangeState::Bad: return theme_.log_error;
            case storage::RescueRangeState::KnownBad: return theme_.log_warning;
            default: return theme_.muted_text;
        }
    }

    static const wchar_t* state_name(storage::RescueRangeState state) {
        switch (state) {
            case storage::RescueRangeState::Good: return L"Good";
            case storage::RescueRangeState::Recovered: return L"Recovered";
            case storage::RescueRangeState::Bad: return L"Bad";
            case storage::RescueRangeState::KnownBad: return L"KnownBad";
            default: return L"Pending";
        }
    }

    void request_geometry() {
        if (map_ == nullptr || list_ == nullptr) return;
        const int index = static_cast<int>(SendMessageW(list_, LB_GETCURSEL, 0, 0));
        RECT client{}; GetClientRect(map_, &client);
        if (index < 0 || static_cast<std::size_t>(index) >= files_.size() || client.right <= 0) return;
        {
            std::lock_guard<std::mutex> guard(worker_lock_);
            requested_index_ = index;
            requested_width_ = client.right;
            ++generation_;
            pending_ = true;
        }
        spans_.clear();
        pixel_span_.clear();
        InvalidateRect(map_, nullptr, TRUE);
        worker_ready_.notify_one();
    }

    void geometry_worker(HWND hwnd) {
        for (;;) {
            int index, width;
            std::uint64_t generation;
            {
                std::unique_lock<std::mutex> guard(worker_lock_);
                worker_ready_.wait(guard, [this] { return stopping_ || pending_; });
                if (stopping_) return;
                index = requested_index_;
                width = requested_width_;
                generation = generation_;
                pending_ = false;
            }
            const auto& file = files_[static_cast<std::size_t>(index)];
            auto result = std::make_unique<GeometryResult>();
            result->generation = generation;
            result->spans = range_map_spans(file.size, file.ranges, width);
            if (!closing_.load() && PostMessageW(hwnd, GeometryReady, 0,
                    reinterpret_cast<LPARAM>(result.get()))) result.release();
        }
    }

    void stop_worker() {
        closing_ = true;
        {
            std::lock_guard<std::mutex> guard(worker_lock_);
            stopping_ = true;
        }
        worker_ready_.notify_one();
        if (worker_.joinable()) worker_.join();
    }

    void update_tooltip(int x, int y) {
        if (x < 0 || static_cast<std::size_t>(x) >= pixel_span_.size()) {
            SendMessageW(tooltip_, TTM_TRACKACTIVATE, FALSE, reinterpret_cast<LPARAM>(&tool_));
            return;
        }
        const auto& span = spans_[pixel_span_[static_cast<std::size_t>(x)]];
        tip_text_ = std::wstring(state_name(span.state)) + L" - offset " +
                    std::to_wstring(span.offset) + L", length " + std::to_wstring(span.length);
        tool_.lpszText = tip_text_.data();
        SendMessageW(tooltip_, TTM_UPDATETIPTEXTW, 0, reinterpret_cast<LPARAM>(&tool_));
        POINT point{x + scale(12), y + scale(20)}; ClientToScreen(map_, &point);
        SendMessageW(tooltip_, TTM_TRACKPOSITION, 0, MAKELPARAM(point.x, point.y));
        SendMessageW(tooltip_, TTM_TRACKACTIVATE, TRUE, reinterpret_cast<LPARAM>(&tool_));
        TRACKMOUSEEVENT track{sizeof(track), TME_LEAVE, map_, 0}; TrackMouseEvent(&track);
    }

    static LRESULT CALLBACK map_subclass(HWND hwnd, UINT message, WPARAM wparam, LPARAM lparam,
                                        UINT_PTR, DWORD_PTR data) {
        auto* self = reinterpret_cast<RangeMapDialog*>(data);
        if (message == WM_MOUSEMOVE) self->update_tooltip(GET_X_LPARAM(lparam), GET_Y_LPARAM(lparam));
        if (message == WM_MOUSELEAVE)
            SendMessageW(self->tooltip_, TTM_TRACKACTIVATE, FALSE,
                         reinterpret_cast<LPARAM>(&self->tool_));
        return DefSubclassProc(hwnd, message, wparam, lparam);
    }

    LRESULT proc(HWND hwnd, UINT message, WPARAM wparam, LPARAM lparam, bool& handled) {
        switch (message) {
            case WM_CREATE: {
                rebuild_font(GetDpiForWindow(hwnd));
                window_brush_ = CreateSolidBrush(theme_.window);
                edit_brush_ = CreateSolidBrush(theme_.edit);
                auto create = [&](const wchar_t* cls, const wchar_t* text, DWORD style, int id) {
                    HWND child = CreateWindowExW(0, cls, text, WS_CHILD | WS_VISIBLE | style,
                        0, 0, 10, 10, hwnd, reinterpret_cast<HMENU>(static_cast<INT_PTR>(id)),
                        GetModuleHandleW(nullptr), nullptr);
                    SendMessageW(child, WM_SETFONT, reinterpret_cast<WPARAM>(font_), TRUE);
                    return child;
                };
                label_ = create(L"STATIC", L"Files with rescue ranges:", SS_NOPREFIX, 0);
                list_ = create(L"LISTBOX", L"", WS_TABSTOP | WS_VSCROLL | WS_BORDER |
                    LBS_NODATA | LBS_OWNERDRAWFIXED | LBS_NOINTEGRALHEIGHT | LBS_NOTIFY, IdFiles);
                SendMessageW(list_, LB_SETITEMHEIGHT, 0, scale(22));
                SendMessageW(list_, LB_SETCOUNT, files_.size(), 0);
                map_ = create(L"STATIC", L"", SS_OWNERDRAW | SS_NOTIFY, IdMap);
                note_ = create(L"STATIC", L"Recovered can include synthetic fill; hover for byte offsets and lengths.",
                               SS_NOPREFIX, 0);
                close_ = create(L"BUTTON", L"Close", WS_TABSTOP | BS_OWNERDRAW, IdClose);
                apply_control_theme(list_, theme_.dark);
                tooltip_ = CreateWindowExW(WS_EX_TOPMOST, TOOLTIPS_CLASSW, nullptr,
                    WS_POPUP | TTS_NOPREFIX | TTS_ALWAYSTIP, CW_USEDEFAULT, CW_USEDEFAULT,
                    CW_USEDEFAULT, CW_USEDEFAULT, hwnd, nullptr, GetModuleHandleW(nullptr), nullptr);
                tool_.cbSize = sizeof(tool_); tool_.uFlags = TTF_TRACK | TTF_ABSOLUTE;
                tool_.hwnd = map_; tool_.uId = 1; tool_.lpszText = const_cast<wchar_t*>(L"");
                SendMessageW(tooltip_, TTM_ADDTOOLW, 0, reinterpret_cast<LPARAM>(&tool_));
                SetWindowSubclass(map_, &map_subclass, 1, reinterpret_cast<DWORD_PTR>(this));
                worker_ = std::thread([this, hwnd] { geometry_worker(hwnd); });
                if (!files_.empty()) SendMessageW(list_, LB_SETCURSEL, 0, 0);
                layout(hwnd);
                handled = true; return 0;
            }
            case WM_SIZE: layout(hwnd); handled = true; return 0;
            case WM_GETMINMAXINFO: {
                auto* limits = reinterpret_cast<MINMAXINFO*>(lparam);
                limits->ptMinTrackSize.x = scale(500);
                limits->ptMinTrackSize.y = scale(300);
                handled = true; return 0;
            }
            case WM_DPICHANGED: {
                rebuild_font(HIWORD(wparam));
                for (HWND child : {label_, list_, map_, note_, close_})
                    SendMessageW(child, WM_SETFONT, reinterpret_cast<WPARAM>(font_), TRUE);
                SendMessageW(list_, LB_SETITEMHEIGHT, 0, scale(22));
                const RECT* rect = reinterpret_cast<const RECT*>(lparam);
                SetWindowPos(hwnd, nullptr, rect->left, rect->top, rect->right - rect->left,
                             rect->bottom - rect->top, SWP_NOZORDER | SWP_NOACTIVATE);
                layout(hwnd); handled = true; return 0;
            }
            case WM_ERASEBKGND: {
                RECT r{}; GetClientRect(hwnd, &r);
                themedraw::fill_rect(reinterpret_cast<HDC>(wparam), r, theme_.window);
                handled = true; return 1;
            }
            case WM_CTLCOLORSTATIC:
                SetTextColor(reinterpret_cast<HDC>(wparam), theme_.text);
                SetBkColor(reinterpret_cast<HDC>(wparam), theme_.window);
                handled = true; return reinterpret_cast<LRESULT>(window_brush_);
            case WM_CTLCOLORLISTBOX:
                SetTextColor(reinterpret_cast<HDC>(wparam), theme_.text);
                SetBkColor(reinterpret_cast<HDC>(wparam), theme_.edit);
                handled = true; return reinterpret_cast<LRESULT>(edit_brush_);
            case WM_DRAWITEM: {
                const auto& draw = *reinterpret_cast<DRAWITEMSTRUCT*>(lparam);
                if (draw.CtlID == IdFiles) {
                    const bool selected = (draw.itemState & ODS_SELECTED) != 0;
                    themedraw::fill_rect(draw.hDC, draw.rcItem, selected ? theme_.selection : theme_.edit);
                    if (draw.itemID < files_.size()) {
                        HGDIOBJ old = SelectObject(draw.hDC, font_);
                        SetBkMode(draw.hDC, TRANSPARENT);
                        SetTextColor(draw.hDC, selected ? theme_.selection_text : theme_.text);
                        RECT text = draw.rcItem; text.left += scale(5);
                        DrawTextW(draw.hDC, files_[draw.itemID].path.c_str(), -1, &text,
                                  DT_LEFT | DT_VCENTER | DT_SINGLELINE | DT_NOPREFIX | DT_END_ELLIPSIS);
                        SelectObject(draw.hDC, old);
                    }
                } else if (draw.CtlID == IdMap) {
                    themedraw::fill_rect(draw.hDC, draw.rcItem, theme_.edit);
                    for (const auto& span : spans_) {
                        RECT r{draw.rcItem.left + span.left, draw.rcItem.top,
                               draw.rcItem.left + span.right, draw.rcItem.bottom};
                        themedraw::fill_rect(draw.hDC, r, color(span.state));
                    }
                    themedraw::frame_rect(draw.hDC, draw.rcItem, theme_.border);
                } else themedraw::draw_button(draw, theme_);
                handled = true; return TRUE;
            }
            case WM_PAINT: {
                PAINTSTRUCT paint{}; HDC dc = BeginPaint(hwnd, &paint);
                RECT client{}; GetClientRect(hwnd, &client);
                int x = scale(16), y = client.bottom - scale(63);
                HGDIOBJ old = SelectObject(dc, font_); SetBkMode(dc, TRANSPARENT);
                SetTextColor(dc, theme_.text);
                for (auto state : {storage::RescueRangeState::Good, storage::RescueRangeState::Recovered,
                                   storage::RescueRangeState::Bad, storage::RescueRangeState::KnownBad,
                                   storage::RescueRangeState::Pending}) {
                    RECT swatch{x, y, x + scale(12), y + scale(12)};
                    themedraw::fill_rect(dc, swatch, color(state));
                    x += scale(17); const wchar_t* name = state_name(state);
                    TextOutW(dc, x, y - scale(2), name, static_cast<int>(wcslen(name)));
                    SIZE size{}; GetTextExtentPoint32W(dc, name, static_cast<int>(wcslen(name)), &size);
                    x += size.cx + scale(14);
                }
                SelectObject(dc, old); EndPaint(hwnd, &paint); handled = true; return 0;
            }
            case WM_COMMAND:
                if (LOWORD(wparam) == IdFiles && HIWORD(wparam) == LBN_SELCHANGE) request_geometry();
                else if (LOWORD(wparam) == IdClose || LOWORD(wparam) == IDCANCEL) host_.destroy();
                handled = true; return 0;
            case GeometryReady: {
                std::unique_ptr<GeometryResult> result(reinterpret_cast<GeometryResult*>(lparam));
                std::lock_guard<std::mutex> guard(worker_lock_);
                if (result->generation == generation_) {
                    spans_ = std::move(result->spans);
                    pixel_span_.clear();
                    if (!spans_.empty()) {
                        pixel_span_.resize(static_cast<std::size_t>(spans_.back().right));
                        for (std::size_t i = 0; i < spans_.size(); ++i) {
                            for (int x = spans_[i].left; x < spans_[i].right; ++x)
                                pixel_span_[static_cast<std::size_t>(x)] = i;
                        }
                    }
                    InvalidateRect(map_, nullptr, TRUE);
                }
                handled = true; return 0;
            }
            case WM_CLOSE: host_.destroy(); handled = true; return 0;
            case WM_DESTROY: {
                stop_worker();
                MSG pending{};
                while (PeekMessageW(&pending, hwnd, GeometryReady, GeometryReady, PM_REMOVE))
                    delete reinterpret_cast<GeometryResult*>(pending.lParam);
                handled = true; return 0;
            }
        }
        return 0;
    }

    ThemePalette theme_;
    std::vector<RangeMapFile> files_;
    dialog_detail::ModalHost host_;
    UINT dpi_ = 96;
    HFONT font_ = nullptr;
    HBRUSH window_brush_ = nullptr, edit_brush_ = nullptr;
    HWND label_ = nullptr, list_ = nullptr, map_ = nullptr, note_ = nullptr, close_ = nullptr;
    HWND tooltip_ = nullptr;
    TOOLINFOW tool_{};
    std::wstring tip_text_;
    std::vector<RangeMapSpan> spans_;
    std::vector<std::size_t> pixel_span_;
    std::thread worker_;
    std::mutex worker_lock_;
    std::condition_variable worker_ready_;
    bool pending_ = false, stopping_ = false;
    std::atomic<bool> closing_{false};
    int requested_index_ = 0, requested_width_ = 0;
    std::uint64_t generation_ = 0;
};

} // namespace xact::ui
