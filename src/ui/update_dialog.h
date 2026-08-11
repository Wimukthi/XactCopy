// -----------------------------------------------------------------------------
// File: src\ui\update_dialog.h
// Purpose: Themed "Update Available" dialog (port of XactCopy.UI UpdateForm).
//          Shows the release summary + notes, downloads the best release asset
//          over WinHTTP with progress, verifies its SHA-256, and applies the
//          update in place: an .exe asset is launched directly; a .zip asset is
//          handed to an apply-update.ps1 that waits for this process to exit,
//          extracts + robocopies the payload over the install folder (with a
//          backup/restore fallback), and relaunches. Built on the shared
//          ModalHost; owner-drawn buttons / progress bar / notes list.
// -----------------------------------------------------------------------------

#pragma once

#include <atomic>
#include <memory>
#include <mutex>
#include <string>
#include <thread>
#include <vector>

#include "dialogs.h"      // dialog_detail::ModalHost, themedraw
#include "settings.h"
#include "theme.h"
#include "update_service.h"

namespace xact::ui {

class UpdateDialog {
public:
    // Returns true if the caller should exit the application so the update can
    // be applied (the updater/installer has been launched).
    static bool show(HWND owner, const UpdateReleaseInfo& release, const AppVersion& current,
                     AppSettings& settings, const ThemePalette& theme) {
        UpdateDialog dialog(release, current, settings, theme);
        dialog.run(owner);
        return dialog.apply_launched_;
    }

private:
    enum : int {
        IdNotesList = 3301,
        IdProgressBar = 3302,
        IdDownloadBtn = 3303,
        IdViewReleaseBtn = 3304,
        IdCloseBtn = 3305,
    };
    static constexpr UINT WM_APP_DL_PROGRESS = WM_APP + 1;
    static constexpr UINT WM_APP_DL_DONE = WM_APP + 2; // wParam: Outcome; lparam: std::string* msg

    enum class Outcome { Ok, Cancelled, Failed };
    enum class Phase { Idle, Downloading, Verifying, Applying, Done };

    UpdateReleaseInfo release_;
    AppVersion current_;
    AppSettings& settings_;
    ThemePalette theme_;
    const UpdateAssetInfo* asset_ = nullptr;

    dialog_detail::ModalHost host_;
    HWND hwnd_ = nullptr;
    HFONT font_ = nullptr;
    HFONT bold_font_ = nullptr;
    HBRUSH window_brush_ = nullptr;
    HBRUSH edit_brush_ = nullptr;

    HWND title_label_ = nullptr;
    HWND current_label_ = nullptr;
    HWND latest_label_ = nullptr;
    HWND package_label_ = nullptr;
    HWND notes_list_ = nullptr;
    HWND progress_bar_ = nullptr;
    HWND status_label_ = nullptr;
    HWND detail_label_ = nullptr;
    HWND download_btn_ = nullptr;
    HWND view_btn_ = nullptr;
    HWND close_btn_ = nullptr;
    UINT monitor_dpi_ = 0;
    UINT layout_dpi_ = 0;
    TooltipManager tooltips_;

    // Notes, pre-wrapped to the list width. Parallel heading flags.
    std::vector<std::string> notes_lines_;
    std::vector<bool> notes_heading_;

    Phase phase_ = Phase::Idle;
    int progress_permille_ = 0;
    std::wstring temp_root_;
    std::wstring download_path_;

    std::thread worker_;
    std::atomic<bool> cancel_flag_{false};
    std::atomic<bool> progress_posted_{false};
    std::mutex progress_mutex_;
    DownloadProgress latest_progress_;

    UpdateDialog(const UpdateReleaseInfo& release, const AppVersion& current, AppSettings& settings,
                 const ThemePalette& theme)
        : release_(release), current_(current), settings_(settings), theme_(theme) {
        asset_ = UpdateService::select_best_asset(release_);
    }

    ~UpdateDialog() {
        cancel_flag_ = true;
        if (worker_.joinable()) worker_.join();
        if (!apply_launched_) remove_private_temp_tree(temp_root_);
        if (font_ != nullptr) DeleteObject(font_);
        if (bold_font_ != nullptr) DeleteObject(bold_font_);
        if (window_brush_ != nullptr) DeleteObject(window_brush_);
        if (edit_brush_ != nullptr) DeleteObject(edit_brush_);
    }

    bool apply_launched_ = false;

    void run(HWND owner) {
        host_.set_density_percent(theme_.density_percent);
        host_.set_scale_percent(theme_.scale_percent);
        HWND hwnd = host_.create(
            owner, L"XactCopyUpdateDlg", L"XactCopy Update", 640, 560,
            [this](HWND h, UINT m, WPARAM w, LPARAM l, bool& handled) {
                return proc(h, m, w, l, handled);
            },
            /*resizable=*/true);
        if (hwnd == nullptr) return;
        apply_window_icons(hwnd);
        set_dark_title_bar(hwnd, theme_.dark && theme_.themed_chrome);
        host_.run_modal();
    }

    // --- font helpers -------------------------------------------------------

    void build_fonts(UINT dpi) {
        layout_dpi_ = dpi == 0 ? 96 : dpi;
        if (font_ != nullptr) DeleteObject(font_);
        if (bold_font_ != nullptr) DeleteObject(bold_font_);
        auto make = [this](int pt, int weight) {
            return CreateFontW(-MulDiv(pt, static_cast<int>(layout_dpi_), 72), 0, 0, 0, weight, FALSE, FALSE,
                               FALSE, DEFAULT_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS,
                               CLEARTYPE_QUALITY, DEFAULT_PITCH, L"Segoe UI");
        };
        font_ = make(9, FW_NORMAL);
        bold_font_ = make(9, FW_BOLD);
    }

    void set_font(HWND control, HFONT font) {
        if (control != nullptr) SendMessageW(control, WM_SETFONT, reinterpret_cast<WPARAM>(font), TRUE);
    }

    // --- creation -----------------------------------------------------------

    LRESULT on_create(HWND hwnd) {
        hwnd_ = hwnd;
        window_brush_ = CreateSolidBrush(theme_.window);
        edit_brush_ = CreateSolidBrush(theme_.edit);
        monitor_dpi_ = GetDpiForWindow(hwnd);
        UINT dpi = ui_layout_dpi(monitor_dpi_, theme_.density_percent,
                                 theme_.scale_percent);
        build_fonts(dpi);
        HINSTANCE inst = GetModuleHandleW(nullptr);

        auto make_static = [&](int id, DWORD extra) {
            HWND h = CreateWindowExW(0, L"STATIC", L"", WS_CHILD | WS_VISIBLE | extra, 0, 0, 0, 0,
                                     hwnd, reinterpret_cast<HMENU>(static_cast<INT_PTR>(id)), inst,
                                     nullptr);
            return h;
        };
        auto make_button = [&](int id, const wchar_t* text) {
            HWND h = CreateWindowExW(0, L"BUTTON", text,
                                     WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_OWNERDRAW, 0, 0, 0, 0,
                                     hwnd, reinterpret_cast<HMENU>(static_cast<INT_PTR>(id)), inst,
                                     nullptr);
            return h;
        };

        title_label_ = make_static(0, SS_LEFT);
        current_label_ = make_static(0, SS_LEFT);
        latest_label_ = make_static(0, SS_LEFT);
        package_label_ = make_static(0, SS_LEFT | SS_ENDELLIPSIS);
        SetWindowTextW(title_label_, L"Update Available");
        SetWindowTextW(current_label_,
                       (L"Current version:  " + format_version(current_, std::string())).c_str());
        SetWindowTextW(latest_label_,
                       (L"Latest version:  " + format_version(release_.version, release_.tag_name))
                           .c_str());
        SetWindowTextW(package_label_,
                       (L"Package:  " +
                        storage::fsutil::utf8_to_wide(
                            asset_ != nullptr ? asset_->name
                                              : std::string("No compatible package found.")))
                           .c_str());

        notes_list_ = CreateWindowExW(
            0, L"LISTBOX", L"",
            WS_CHILD | WS_VISIBLE | WS_VSCROLL | LBS_NOINTEGRALHEIGHT | LBS_OWNERDRAWFIXED |
                LBS_NOSEL,
            0, 0, 0, 0, hwnd, reinterpret_cast<HMENU>(static_cast<INT_PTR>(IdNotesList)), inst,
            nullptr);
        allow_dark_mode_for_window(notes_list_, theme_.dark);
        apply_control_theme(notes_list_, theme_.dark);

        progress_bar_ = CreateWindowExW(
            0, L"STATIC", L"", WS_CHILD | WS_VISIBLE | SS_OWNERDRAW, 0, 0, 0, 0, hwnd,
            reinterpret_cast<HMENU>(static_cast<INT_PTR>(IdProgressBar)), inst, nullptr);

        status_label_ = make_static(0, SS_LEFT | SS_ENDELLIPSIS);
        detail_label_ = make_static(0, SS_LEFT | SS_ENDELLIPSIS);
        SetWindowTextW(status_label_, asset_ != nullptr ? L"Ready to download."
                                                        : L"No compatible package is available.");
        SetWindowTextW(detail_label_, L"");

        // Owner-drawn buttons draw with DT_NOPREFIX, so use a literal single '&'.
        download_btn_ = make_button(IdDownloadBtn, L"Download & Install");
        view_btn_ = make_button(IdViewReleaseBtn, L"View Release");
        close_btn_ = make_button(IdCloseBtn, L"Close");
        if (asset_ == nullptr) EnableWindow(download_btn_, FALSE);
        if (release_.html_url.empty()) EnableWindow(view_btn_, FALSE);

        for (HWND h : {title_label_, current_label_, latest_label_, package_label_, notes_list_,
                       progress_bar_, status_label_, detail_label_, download_btn_, view_btn_,
                       close_btn_}) {
            set_font(h, h == title_label_ ? bold_font_ : font_);
        }

        int line_height = MulDiv(18, static_cast<int>(dpi), 96);
        SendMessageW(notes_list_, LB_SETITEMHEIGHT, 0, MAKELPARAM(line_height, 0));

        tooltips_.create(hwnd, layout_dpi_, theme_.dark);
        tooltips_.add(title_label_, L"A newer XactCopy release is available. Review the version, package, and release notes before installing.");
        tooltips_.add(current_label_, L"The XactCopy version currently installed on this computer.");
        tooltips_.add(latest_label_, L"The newest published release detected by the update service.");
        tooltips_.add(package_label_, L"The installer package selected for this Windows architecture. Its published SHA-256 checksum is verified before launch.");
        tooltips_.add(notes_list_, L"Scrollable release notes supplied with the selected release.");
        tooltips_.add(progress_bar_, L"Download and checksum-verification progress for the update package.");
        tooltips_.add(status_label_, L"Current update phase, such as ready, downloading, verifying, launching, cancelled, or failed.");
        tooltips_.add(detail_label_, L"Downloaded bytes, total package size, transfer speed, or a detailed update result.");
        tooltips_.add(download_btn_, L"Download the selected installer, verify its published SHA-256 checksum, and launch it after confirmation.");
        tooltips_.add(view_btn_, L"Open this release's GitHub page in your default browser without downloading it.");
        tooltips_.add(close_btn_, L"Close this window. During a download, this button requests cancellation instead.");

        // Lay out first so the notes list has its real width, then wrap.
        relayout(hwnd);
        rebuild_notes();
        return 0;
    }

    // --- notes wrapping -----------------------------------------------------

    // Classify + strip markdown, then greedy word-wrap to the list width.
    void rebuild_notes() {
        notes_lines_.clear();
        notes_heading_.clear();
        std::string notes = release_.notes.empty() ? std::string("No release notes provided.")
                                                    : release_.notes;

        int width = notes_client_width();
        HDC dc = GetDC(notes_list_);
        HGDIOBJ old_font = SelectObject(dc, font_);

        std::string normalized;
        normalized.reserve(notes.size());
        for (char c : notes) {
            if (c != '\r') normalized.push_back(c);
        }
        std::size_t start = 0;
        while (start <= normalized.size()) {
            std::size_t nl = normalized.find('\n', start);
            std::string raw = normalized.substr(
                start, nl == std::string::npos ? std::string::npos : nl - start);
            start = (nl == std::string::npos) ? normalized.size() + 1 : nl + 1;

            // right-trim
            while (!raw.empty() && (raw.back() == ' ' || raw.back() == '\t')) raw.pop_back();

            bool heading = false;
            std::string line = raw;
            if (line.rfind("## ", 0) == 0) {
                heading = true;
                line = line.substr(3);
            } else if (line.rfind("### ", 0) == 0) {
                heading = true;
                line = line.substr(4);
            } else if (line == "---") {
                notes_lines_.emplace_back("");
                notes_heading_.push_back(false);
                continue;
            } else if (line.rfind("- ", 0) == 0) {
                line = "\xE2\x80\xA2 " + line.substr(2); // bullet
            }
            line = strip_bold_markers(line);

            HFONT wrap_font = heading ? bold_font_ : font_;
            SelectObject(dc, wrap_font);
            for (const std::string& seg : wrap_line(dc, line, width)) {
                notes_lines_.push_back(seg);
                notes_heading_.push_back(heading);
            }
        }
        SelectObject(dc, old_font);
        ReleaseDC(notes_list_, dc);

        SendMessageW(notes_list_, WM_SETREDRAW, FALSE, 0);
        SendMessageW(notes_list_, LB_RESETCONTENT, 0, 0);
        for (std::size_t i = 0; i < notes_lines_.size(); ++i) {
            SendMessageW(notes_list_, LB_ADDSTRING, 0, reinterpret_cast<LPARAM>(L""));
        }
        SendMessageW(notes_list_, WM_SETREDRAW, TRUE, 0);
        InvalidateRect(notes_list_, nullptr, TRUE);
    }

    static std::string strip_bold_markers(const std::string& s) {
        std::string out;
        out.reserve(s.size());
        for (std::size_t i = 0; i < s.size(); ++i) {
            if (i + 1 < s.size() && s[i] == '*' && s[i + 1] == '*') {
                ++i;
                continue;
            }
            out.push_back(s[i]);
        }
        return out;
    }

    std::vector<std::string> wrap_line(HDC dc, const std::string& line, int max_width) {
        std::vector<std::string> out;
        if (line.empty() || max_width <= 0) {
            out.push_back(line);
            return out;
        }
        std::string current;
        std::size_t i = 0;
        while (i < line.size()) {
            std::size_t sp = line.find(' ', i);
            std::string word = line.substr(i, sp == std::string::npos ? std::string::npos : sp - i);
            std::string candidate = current.empty() ? word : current + " " + word;
            if (text_width(dc, candidate) <= max_width || current.empty()) {
                current = candidate;
            } else {
                out.push_back(current);
                current = word;
            }
            if (sp == std::string::npos) break;
            i = sp + 1;
        }
        if (!current.empty() || out.empty()) out.push_back(current);
        return out;
    }

    static int text_width(HDC dc, const std::string& utf8) {
        std::wstring w = storage::fsutil::utf8_to_wide(utf8);
        SIZE size{};
        GetTextExtentPoint32W(dc, w.c_str(), static_cast<int>(w.size()), &size);
        return size.cx;
    }

    int notes_client_width() {
        RECT rc{};
        GetClientRect(notes_list_, &rc);
        int w = rc.right - rc.left - MulDiv(24, dpi(), 96); // padding + scrollbar
        return w > MulDiv(80, dpi(), 96) ? w : MulDiv(80, dpi(), 96);
    }

    UINT dpi() const { return layout_dpi_ == 0 ? 96 : layout_dpi_; }

    // --- layout -------------------------------------------------------------

    void relayout(HWND hwnd) {
        RECT client{};
        GetClientRect(hwnd, &client);
        const int w = client.right;
        const int h = client.bottom;
        const UINT d = dpi();
        auto s = [d](int v) { return MulDiv(v, static_cast<int>(d), 96); };
        const int pad = s(14);
        const int line = s(20);
        const int x = pad;
        const int cw = w - pad * 2;

        int y = pad;
        MoveWindow(title_label_, x, y, cw, s(22), TRUE);
        y += s(28);
        MoveWindow(current_label_, x, y, cw, line, TRUE);
        y += line;
        MoveWindow(latest_label_, x, y, cw, line, TRUE);
        y += line;
        MoveWindow(package_label_, x, y, cw, line, TRUE);
        y += line + s(8);

        // Footer buttons (bottom), right-aligned.
        const int btn_h = s(30);
        const int btn_w = s(150);
        const int btn_gap = s(8);
        const int footer_y = h - pad - btn_h;
        int bx = w - pad - btn_w;
        MoveWindow(download_btn_, bx, footer_y, btn_w, btn_h, TRUE);
        bx -= btn_w + btn_gap;
        MoveWindow(view_btn_, bx, footer_y, s(120), btn_h, TRUE);
        MoveWindow(close_btn_, x, footer_y, s(100), btn_h, TRUE);

        // Progress block sits above the footer.
        const int detail_h = s(18);
        const int bar_h = s(18);
        int py = footer_y - s(10) - detail_h - s(4) - detail_h - s(4) - bar_h;
        MoveWindow(progress_bar_, x, py, cw, bar_h, TRUE);
        py += bar_h + s(4);
        MoveWindow(status_label_, x, py, cw, detail_h, TRUE);
        py += detail_h + s(4);
        MoveWindow(detail_label_, x, py, cw, detail_h, TRUE);

        // Notes list fills the gap between the header and the progress block.
        int notes_top = y;
        int notes_bottom = progress_bar_ != nullptr ? (footer_y - s(10) - detail_h - s(4) -
                                                       detail_h - s(4) - bar_h - s(8))
                                                     : notes_top;
        if (notes_bottom < notes_top + line) notes_bottom = notes_top + line;
        MoveWindow(notes_list_, x, notes_top, cw, notes_bottom - notes_top, TRUE);
        tooltips_.update_layout();
        InvalidateRect(hwnd, nullptr, TRUE);
    }

    // --- drawing ------------------------------------------------------------

    void draw_progress_bar(const DRAWITEMSTRUCT& draw) {
        RECT rect = draw.rcItem;
        int clamped = std::clamp(progress_permille_, 0, 1000);
        themedraw::fill_rect(draw.hDC, rect, theme_.panel);
        int span = rect.right - rect.left;
        int filled = MulDiv(span, clamped, 1000);
        if (filled > 0) {
            RECT fill = rect;
            fill.right = rect.left + filled;
            themedraw::fill_rect(draw.hDC, fill, theme_.accent);
        }
        themedraw::frame_rect(draw.hDC, rect, theme_.border);
        wchar_t text[8];
        swprintf(text, 8, L"%d%%", (clamped + 5) / 10);
        HGDIOBJ old_font = SelectObject(draw.hDC, font_);
        SetBkMode(draw.hDC, TRANSPARENT);
        SetTextColor(draw.hDC, readable_text_color(clamped >= 500 ? theme_.accent : theme_.panel));
        DrawTextW(draw.hDC, text, -1, &rect, DT_CENTER | DT_VCENTER | DT_SINGLELINE);
        SelectObject(draw.hDC, old_font);
    }

    void draw_notes_item(const DRAWITEMSTRUCT& draw) {
        themedraw::fill_rect(draw.hDC, draw.rcItem, theme_.edit);
        UINT idx = draw.itemID;
        if (idx == static_cast<UINT>(-1) || idx >= notes_lines_.size()) return;
        bool heading = notes_heading_[idx];
        HGDIOBJ old_font = SelectObject(draw.hDC, heading ? bold_font_ : font_);
        SetBkMode(draw.hDC, TRANSPARENT);
        SetTextColor(draw.hDC, heading ? theme_.text : theme_.muted_text);
        RECT rect = draw.rcItem;
        rect.left += MulDiv(8, static_cast<int>(dpi()), 96);
        std::wstring line = storage::fsutil::utf8_to_wide(notes_lines_[idx]);
        DrawTextW(draw.hDC, line.c_str(), -1, &rect,
                  DT_LEFT | DT_VCENTER | DT_SINGLELINE | DT_NOPREFIX);
        SelectObject(draw.hDC, old_font);
    }

    // --- version formatting -------------------------------------------------

    static std::wstring format_version(const AppVersion& v, const std::string& tag) {
        int build = std::max(0, v.build);
        int revision = std::max(0, v.revision);
        if (v.major == 0 && v.minor == 0 && build == 0 && revision == 0 && !tag.empty()) {
            return storage::fsutil::utf8_to_wide(tag);
        }
        return std::to_wstring(v.major) + L"." + std::to_wstring(v.minor) + L"." +
               std::to_wstring(build) + L"." + std::to_wstring(revision);
    }

    // --- download + apply ---------------------------------------------------

    void set_status(const std::wstring& text) { SetWindowTextW(status_label_, text.c_str()); }

    void start_download() {
        if (asset_ == nullptr || phase_ != Phase::Idle) return;
        phase_ = Phase::Downloading;
        cancel_flag_ = false;
        progress_permille_ = 0;
        EnableWindow(download_btn_, FALSE);
        EnableWindow(view_btn_, FALSE);
        SetWindowTextW(close_btn_, L"Cancel");
        set_status(L"Downloading update package...");
        SetWindowTextW(detail_label_, L"");
        InvalidateRect(progress_bar_, nullptr, FALSE);

        remove_private_temp_tree(temp_root_);
        temp_root_.clear();
        download_path_.clear();
        temp_root_ = create_temp_root();
        if (temp_root_.empty()) {
            phase_ = Phase::Idle;
            SetWindowTextW(close_btn_, L"Close");
            EnableWindow(download_btn_, TRUE);
            EnableWindow(view_btn_, !release_.html_url.empty());
            set_status(L"Update failed: a private temporary folder could not be created.");
            return;
        }
        download_path_ = temp_root_ + L"\\" + storage::fsutil::utf8_to_wide(asset_->name);

        UpdateReleaseInfo release = release_;
        UpdateAssetInfo asset = *asset_;
        std::wstring dest = download_path_;
        HWND hwnd = hwnd_;
        std::int64_t total = asset.size;

        if (worker_.joinable()) worker_.join();
        worker_ = std::thread([this, hwnd, release, asset, dest, total]() {
            try {
                auto on_progress = [this, hwnd](const DownloadProgress& p) {
                    {
                        std::lock_guard<std::mutex> lock(progress_mutex_);
                        latest_progress_ = p;
                    }
                    if (!progress_posted_.exchange(true)) {
                        PostMessageW(hwnd, WM_APP_DL_PROGRESS, 0, 0);
                    }
                };
                auto cancelled = [this]() { return cancel_flag_.load(); };

                std::string actual = UpdateService::download_asset(
                    asset.download_url, dest, total, on_progress, cancelled);
                if (cancel_flag_.load()) {
                    post_done(hwnd, Outcome::Cancelled, "Download canceled.");
                    return;
                }
                if (actual.empty()) {
                    post_done(hwnd, Outcome::Failed, "Download failed.");
                    return;
                }
                // Verify against a release-published digest before any package
                // is launched or extracted.
                std::string expected = UpdateService::resolve_asset_sha256(release, asset);
                if (expected.empty()) {
                    post_done(hwnd, Outcome::Failed,
                              "Update package checksum is unavailable for this release.");
                    return;
                }
                if (!update_detail::equals_hex_ci(expected, actual)) {
                    post_done(hwnd, Outcome::Failed,
                              "Update package integrity check failed.");
                    return;
                }
                post_done(hwnd, Outcome::Ok, std::string());
            } catch (const std::exception& ex) {
                post_done(hwnd, Outcome::Failed,
                          std::string("Update processing failed: ") + ex.what());
            } catch (...) {
                post_done(hwnd, Outcome::Failed, "Update processing failed unexpectedly.");
            }
        });
    }

    static void post_done(HWND hwnd, Outcome outcome, const std::string& message) {
        auto* msg = new std::string(message);
        if (!PostMessageW(hwnd, WM_APP_DL_DONE, static_cast<WPARAM>(outcome),
                          reinterpret_cast<LPARAM>(msg))) {
            delete msg;
        }
    }

    void on_progress_message() {
        progress_posted_ = false;
        DownloadProgress p;
        {
            std::lock_guard<std::mutex> lock(progress_mutex_);
            p = latest_progress_;
        }
        int permille = 0;
        if (p.total > 0) permille = static_cast<int>((p.received * 1000) / p.total);
        progress_permille_ = std::clamp(permille, 0, 1000);
        InvalidateRect(progress_bar_, nullptr, FALSE);
        std::wstring received = storage::fsutil::utf8_to_wide(UpdateService::format_bytes(p.received));
        std::wstring total = p.total > 0
                                 ? storage::fsutil::utf8_to_wide(UpdateService::format_bytes(p.total))
                                 : std::wstring(L"Unknown");
        std::wstring speed = p.bytes_per_second > 0
                                 ? storage::fsutil::utf8_to_wide(UpdateService::format_bytes(
                                       static_cast<std::int64_t>(p.bytes_per_second))) +
                                       L"/s"
                                 : std::wstring(L"-");
        SetWindowTextW(detail_label_, (received + L" / " + total + L"     Speed: " + speed).c_str());
    }

    void on_done_message(Outcome outcome, std::string message) {
        if (worker_.joinable()) worker_.join();
        if (outcome == Outcome::Cancelled) {
            phase_ = Phase::Idle;
            set_status(L"Download canceled.");
            SetWindowTextW(close_btn_, L"Close");
            EnableWindow(download_btn_, asset_ != nullptr);
            EnableWindow(view_btn_, !release_.html_url.empty());
            return;
        }
        if (outcome == Outcome::Failed) {
            phase_ = Phase::Idle;
            set_status(L"Update failed: " + storage::fsutil::utf8_to_wide(message));
            SetWindowTextW(close_btn_, L"Close");
            EnableWindow(download_btn_, asset_ != nullptr);
            EnableWindow(view_btn_, !release_.html_url.empty());
            return;
        }
        // Ok -> apply.
        phase_ = Phase::Applying;
        EnableWindow(close_btn_, FALSE);
        set_status(L"Preparing update...");
        std::wstring error;
        if (apply_update(error)) {
            apply_launched_ = true;
            phase_ = Phase::Done;
            host_.destroy(); // returns to caller, which exits the app
        } else {
            phase_ = Phase::Idle;
            set_status(L"Update failed: " + error);
            SetWindowTextW(close_btn_, L"Close");
            EnableWindow(close_btn_, TRUE);
            EnableWindow(download_btn_, asset_ != nullptr);
            EnableWindow(view_btn_, !release_.html_url.empty());
        }
    }

    // Launch the installer (.exe) or the apply-update.ps1 (.zip). Returns true
    // on launch; the caller then exits the app so the update can proceed.
    bool apply_update(std::wstring& error) {
        if (download_path_.empty() || GetFileAttributesW(download_path_.c_str()) == INVALID_FILE_ATTRIBUTES) {
            error = L"Downloaded update package was not found.";
            return false;
        }
        wchar_t exe_path[MAX_PATH]{};
        GetModuleFileNameW(nullptr, exe_path, MAX_PATH);
        std::wstring target_dir = directory_of(exe_path);
        std::wstring exe_name = file_name_of(exe_path);

        std::wstring ext = to_lower_w(extension_of(download_path_));
        if (ext == L".exe") {
            if (!file_version_matches(download_path_, release_.version)) {
                error = L"The installer version does not match the selected release.";
                return false;
            }
            HINSTANCE rc = ShellExecuteW(nullptr, L"open", download_path_.c_str(), nullptr, nullptr,
                                         SW_SHOWNORMAL);
            if (reinterpret_cast<INT_PTR>(rc) <= 32) {
                error = L"Could not launch the installer.";
                return false;
            }
            return true;
        }
        if (ext != L".zip") {
            error = L"Unsupported update package format.";
            return false;
        }
        if (!folder_writable(target_dir)) {
            error = L"The portable application folder is not writable. Use the installer package "
                    L"or move the portable copy to a writable folder.";
            return false;
        }
        return launch_update_script(target_dir, exe_name, error);
    }

    bool launch_update_script(const std::wstring& target_dir, const std::wstring& exe_name,
                              std::wstring& error) {
        std::wstring normalized = target_dir;
        while (normalized.size() > 3 &&
               (normalized.back() == L'\\' || normalized.back() == L'/')) {
            normalized.pop_back();
        }
        std::wstring script_path = temp_root_ + L"\\apply-update.ps1";
        std::wstring backup = temp_root_ + L"\\backup";
        std::wstring payload = temp_root_ + L"\\payload";
        std::wstring log = temp_root_ + L"\\update.log";
        std::wstring executable = normalized + L"\\" + exe_name;
        std::wstring expected_version =
            std::to_wstring(release_.version.major) + L"." +
            std::to_wstring(release_.version.minor) + L"." +
            std::to_wstring(std::max(0, release_.version.build)) + L"." +
            std::to_wstring(std::max(0, release_.version.revision));
        DWORD pid = GetCurrentProcessId();

        auto q = [](std::wstring s) {
            std::wstring out;
            for (wchar_t c : s) {
                if (c == L'\'') out.push_back(L'\'');
                out.push_back(c);
            }
            return out;
        };

        std::wstring script;
        script += L"$ErrorActionPreference = 'Stop'\r\n";
        script += L"$TargetPid  = " + std::to_wstring(pid) + L"\r\n";
        script += L"$Zip        = '" + q(download_path_) + L"'\r\n";
        script += L"$Target     = '" + q(normalized) + L"'\r\n";
        script += L"$Backup     = '" + q(backup) + L"'\r\n";
        script += L"$Payload    = '" + q(payload) + L"'\r\n";
        script += L"$Log        = '" + q(log) + L"'\r\n";
        script += L"$TempRoot   = '" + q(temp_root_) + L"'\r\n";
        script += L"$Executable = '" + q(executable) + L"'\r\n";
        script += L"$ExeName    = '" + q(exe_name) + L"'\r\n";
        script += L"$ExpectedVersion = [Version]'" + q(expected_version) + L"'\r\n";
        script += L"$PackageRelative = @()\r\n";
        script += L"$UpdateCopyStarted = $false\r\n\r\n";
        script += L"function Assert-SafeTargetPath([string]$Relative) {\r\n";
        script += L"    $parts = @($Relative -split '\\\\')\r\n";
        script += L"    $cursor = $Target\r\n";
        script += L"    for ($index = 0; $index -lt $parts.Count; $index++) {\r\n";
        script += L"        $cursor = Join-Path $cursor $parts[$index]\r\n";
        script += L"        if (Test-Path -LiteralPath $cursor) {\r\n";
        script += L"            $existing = Get-Item -LiteralPath $cursor -Force\r\n";
        script += L"            if (($existing.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) { throw ('Refusing target reparse point: ' + $Relative) }\r\n";
        script += L"            if ($index -lt ($parts.Count - 1) -and -not $existing.PSIsContainer) { throw ('File/directory conflict: ' + $Relative) }\r\n";
        script += L"        }\r\n";
        script += L"    }\r\n";
        script += L"}\r\n\r\n";
        script += L"'XactCopy update log' | Out-File -LiteralPath $Log -Encoding UTF8\r\n";
        script += L"while ($null -ne (Get-Process -Id $TargetPid -ErrorAction SilentlyContinue)) { "
                  L"Start-Sleep -Milliseconds 500 }\r\n";
        script += L"Start-Sleep -Seconds 1\r\n";
        script += L"'Process exited. Starting update.' | Out-File -LiteralPath $Log -Append "
                  L"-Encoding UTF8\r\n\r\n";
        script += L"if (Test-Path -LiteralPath $Payload) { Remove-Item -LiteralPath $Payload "
                  L"-Recurse -Force }\r\n";
        script += L"$null = New-Item -ItemType Directory -Path $Payload -Force\r\n";
        script += L"Add-Type -AssemblyName System.IO.Compression.FileSystem\r\n";
        script += L"try {\r\n";
        script += L"    $PayloadRoot = [IO.Path]::GetFullPath($Payload + [IO.Path]::DirectorySeparatorChar)\r\n";
        script += L"    $archive = [System.IO.Compression.ZipFile]::OpenRead($Zip)\r\n";
        script += L"    try {\r\n";
        script += L"        $entryPaths = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)\r\n";
        script += L"        foreach ($entry in $archive.Entries) {\r\n";
        script += L"            $entryName = $entry.FullName.Replace('/', '\\')\r\n";
        script += L"            if ([string]::IsNullOrWhiteSpace($entryName) -or [IO.Path]::IsPathRooted($entryName) -or $entryName -match '(^|\\\\)\\.\\.(\\\\|$)') { throw 'Unsafe path in update archive.' }\r\n";
        script += L"            $trimmedEntry = $entryName.TrimEnd('\\')\r\n";
        script += L"            $segments = @($trimmedEntry -split '\\\\')\r\n";
        script += L"            if ($segments.Count -eq 0) { throw 'Empty path in update archive.' }\r\n";
        script += L"            foreach ($segment in $segments) {\r\n";
        script += L"                if ([string]::IsNullOrWhiteSpace($segment) -or $segment.EndsWith('.') -or $segment.EndsWith(' ') -or $segment.IndexOfAny([IO.Path]::GetInvalidFileNameChars()) -ge 0 -or $segment -match '^(?i:con|prn|aux|nul|com[1-9]|lpt[1-9])(\\..*)?$') { throw 'Unsupported Windows path in update archive.' }\r\n";
        script += L"            }\r\n";
        script += L"            $candidate = [IO.Path]::GetFullPath((Join-Path $Payload $entryName))\r\n";
        script += L"            if (-not $candidate.StartsWith($PayloadRoot, [StringComparison]::OrdinalIgnoreCase)) { throw 'Unsafe path in update archive.' }\r\n";
        script += L"            if (-not $entryPaths.Add($candidate.TrimEnd('\\'))) { throw 'Duplicate path in update archive.' }\r\n";
        script += L"        }\r\n";
        script += L"    } finally { $archive.Dispose() }\r\n";
        script += L"    [System.IO.Compression.ZipFile]::ExtractToDirectory($Zip, $Payload)\r\n";
        script += L"    $matches = @(Get-ChildItem -LiteralPath $Payload -Filter $ExeName -Recurse -File)\r\n";
        script += L"    if ($matches.Count -ne 1) { throw 'Portable package must contain exactly one application executable.' }\r\n";
        script += L"    $found = $matches[0]\r\n";
        script += L"    $actualVersion = [Version]([regex]::Match($found.VersionInfo.FileVersion, '\\d+(\\.\\d+){1,3}').Value)\r\n";
        script += L"    if ($actualVersion -ne $ExpectedVersion) { throw ('Package version mismatch: ' + $actualVersion) }\r\n";
        script += L"    $Source = $found.DirectoryName\r\n";
        script += L"    if (-not (Test-Path -LiteralPath (Join-Path $Source 'XactCopyExecutive.exe') -PathType Leaf)) { throw 'Portable package is missing XactCopyExecutive.exe.' }\r\n";
        script += L"    $packageItems = @(Get-ChildItem -LiteralPath $Source -Recurse -Force)\r\n";
        script += L"    $packageFiles = @($packageItems | Where-Object { -not $_.PSIsContainer })\r\n";
        script += L"    if ($packageFiles.Count -eq 0) { throw 'Portable package contains no files.' }\r\n";
        script += L"    if (@($packageItems | Where-Object { ($_.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0 }).Count -ne 0) { throw 'Portable package contains a reparse point.' }\r\n";
        script += L"    foreach ($directory in @($packageItems | Where-Object { $_.PSIsContainer })) {\r\n";
        script += L"        $relativeDirectory = $directory.FullName.Substring($Source.Length).TrimStart('\\')\r\n";
        script += L"        if (-not [string]::IsNullOrWhiteSpace($relativeDirectory)) { Assert-SafeTargetPath $relativeDirectory }\r\n";
        script += L"    }\r\n";
        script += L"    if (Test-Path -LiteralPath $Backup) { Remove-Item -LiteralPath $Backup -Recurse -Force }\r\n";
        script += L"    $null = New-Item -ItemType Directory -Path $Backup -Force\r\n";
        script += L"    foreach ($file in $packageFiles) {\r\n";
        script += L"        $relative = $file.FullName.Substring($Source.Length).TrimStart('\\')\r\n";
        script += L"        if ([string]::IsNullOrWhiteSpace($relative)) { throw 'Invalid package file path.' }\r\n";
        script += L"        Assert-SafeTargetPath $relative\r\n";
        script += L"        $PackageRelative += $relative\r\n";
        script += L"        $targetFile = Join-Path $Target $relative\r\n";
        script += L"        if (Test-Path -LiteralPath $targetFile) {\r\n";
        script += L"            if ((Get-Item -LiteralPath $targetFile -Force).PSIsContainer) { throw ('File/directory conflict: ' + $relative) }\r\n";
        script += L"            $backupFile = Join-Path $Backup $relative\r\n";
        script += L"            $null = New-Item -ItemType Directory -Path ([IO.Path]::GetDirectoryName($backupFile)) -Force\r\n";
        script += L"            Copy-Item -LiteralPath $targetFile -Destination $backupFile -Force -ErrorAction Stop\r\n";
        script += L"        }\r\n";
        script += L"    }\r\n";
        script += L"    $UpdateCopyStarted = $true\r\n";
        script += L"    robocopy $Source $Target /E /COPY:DAT /R:2 /W:1 /NFL /NDL /NJH /NJS /NP >> $Log\r\n";
        script += L"    if ($LASTEXITCODE -ge 8) { throw ('Robocopy failed with exit code ' + $LASTEXITCODE) }\r\n";
        script += L"    foreach ($relative in $PackageRelative) {\r\n";
        script += L"        $sourceFile = Join-Path $Source $relative\r\n";
        script += L"        $installedFile = Join-Path $Target $relative\r\n";
        script += L"        if (-not (Test-Path -LiteralPath $installedFile -PathType Leaf)) { throw ('Installed package file is missing: ' + $relative) }\r\n";
        script += L"        $sourceHash = (Get-FileHash -LiteralPath $sourceFile -Algorithm SHA256).Hash\r\n";
        script += L"        $installedHash = (Get-FileHash -LiteralPath $installedFile -Algorithm SHA256).Hash\r\n";
        script += L"        if ($sourceHash -ne $installedHash) { throw ('Installed package hash mismatch: ' + $relative) }\r\n";
        script += L"    }\r\n";
        script += L"    $installedVersion = [Version]([regex]::Match((Get-Item -LiteralPath $Executable).VersionInfo.FileVersion, '\\d+(\\.\\d+){1,3}').Value)\r\n";
        script += L"    if ($installedVersion -ne $ExpectedVersion) { throw ('Installed version validation failed: ' + $installedVersion) }\r\n";
        script += L"} catch {\r\n";
        script += L"    $_ | Out-File -LiteralPath $Log -Append -Encoding UTF8\r\n";
        script += L"    if ($UpdateCopyStarted) {\r\n";
        script += L"        foreach ($relative in $PackageRelative) { Remove-Item -LiteralPath (Join-Path $Target $relative) -Force -ErrorAction SilentlyContinue }\r\n";
        script += L"        robocopy $Backup $Target /E /COPY:DAT /R:2 /W:1 /NFL /NDL /NJH /NJS /NP >> $Log\r\n";
        script += L"    }\r\n";
        script += L"    if (Test-Path -LiteralPath $Executable -PathType Leaf) { Start-Process -FilePath $Executable }\r\n";
        script += L"    return\r\n";
        script += L"}\r\n\r\n";
        script += L"Start-Process -FilePath $Executable\r\n";
        script += L"Start-Sleep -Seconds 10\r\n";
        script += L"Remove-Item -LiteralPath $TempRoot -Recurse -Force -ErrorAction "
                  L"SilentlyContinue\r\n";

        if (!write_utf8_no_bom(script_path, script)) {
            error = L"Could not write the update script.";
            return false;
        }

        std::wstring args = L"-NonInteractive -WindowStyle Hidden -ExecutionPolicy Bypass -File \"" +
                            script_path + L"\"";
        // UseShellExecute equivalent: ShellExecuteW starts PowerShell outside this
        // process's job object, so it survives our exit.
        wchar_t system_directory[MAX_PATH]{};
        UINT system_length = GetSystemDirectoryW(system_directory, MAX_PATH);
        if (system_length == 0 || system_length >= MAX_PATH) {
            error = L"Could not locate Windows PowerShell.";
            return false;
        }
        std::wstring powershell = std::wstring(system_directory, system_length) +
                                  L"\\WindowsPowerShell\\v1.0\\powershell.exe";
        HINSTANCE rc = ShellExecuteW(nullptr, L"open", powershell.c_str(), args.c_str(), nullptr,
                                     SW_HIDE);
        if (reinterpret_cast<INT_PTR>(rc) <= 32) {
            error = L"Could not start the update process.";
            return false;
        }
        return true;
    }

    // --- small filesystem helpers ------------------------------------------

    static std::wstring create_temp_root() {
        wchar_t temp[MAX_PATH]{};
        DWORD length = GetTempPathW(MAX_PATH, temp);
        if (length == 0 || length >= MAX_PATH) return std::wstring();
        std::wstring base = std::wstring(temp) + L"XactCopyUpdate";
        storage::fsutil::create_directories(base);
        DWORD attributes = GetFileAttributesW(base.c_str());
        if (attributes == INVALID_FILE_ATTRIBUTES ||
            (attributes & FILE_ATTRIBUTE_DIRECTORY) == 0 ||
            (attributes & FILE_ATTRIBUTE_REPARSE_POINT) != 0) {
            return std::wstring();
        }
        for (int attempt = 0; attempt < 8; ++attempt) {
            std::wstring candidate =
                base + L"\\" + storage::fsutil::random_temp_suffix();
            if (CreateDirectoryW(candidate.c_str(), nullptr)) return candidate;
            if (GetLastError() != ERROR_ALREADY_EXISTS) return std::wstring();
        }
        return std::wstring();
    }

    // Never follow a junction/symbolic link while cleaning the private update
    // directory. This routine is used only for the cryptographically random
    // directory created above, not for a caller-supplied path.
    static void remove_private_temp_tree(const std::wstring& path) {
        if (path.empty()) return;
        DWORD root_attributes = GetFileAttributesW(path.c_str());
        if (root_attributes == INVALID_FILE_ATTRIBUTES) return;
        if ((root_attributes & FILE_ATTRIBUTE_DIRECTORY) == 0) return;
        if ((root_attributes & FILE_ATTRIBUTE_REPARSE_POINT) != 0) {
            RemoveDirectoryW(path.c_str());
            return;
        }

        WIN32_FIND_DATAW data{};
        HANDLE find = FindFirstFileExW((path + L"\\*").c_str(), FindExInfoBasic, &data,
                                       FindExSearchNameMatch, nullptr, 0);
        if (find != INVALID_HANDLE_VALUE) {
            do {
                const std::wstring name = data.cFileName;
                if (name.empty() || name == L"." || name == L"..") continue;
                const std::wstring child = path + L"\\" + name;
                if ((data.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) != 0) {
                    if ((data.dwFileAttributes & FILE_ATTRIBUTE_REPARSE_POINT) != 0) {
                        RemoveDirectoryW(child.c_str());
                    } else {
                        remove_private_temp_tree(child);
                    }
                } else {
                    SetFileAttributesW(child.c_str(), FILE_ATTRIBUTE_NORMAL);
                    DeleteFileW(child.c_str());
                }
            } while (FindNextFileW(find, &data));
            FindClose(find);
        }
        SetFileAttributesW(path.c_str(), FILE_ATTRIBUTE_NORMAL);
        RemoveDirectoryW(path.c_str());
    }

    static bool folder_writable(const std::wstring& folder) {
        std::wstring probe = folder + L"\\write-probe.tmp";
        HANDLE h = CreateFileW(probe.c_str(), GENERIC_WRITE, 0, nullptr, CREATE_ALWAYS,
                               FILE_ATTRIBUTE_TEMPORARY | FILE_FLAG_DELETE_ON_CLOSE, nullptr);
        if (h == INVALID_HANDLE_VALUE) return false;
        CloseHandle(h);
        return true;
    }

    static bool file_version_matches(const std::wstring& path,
                                     const AppVersion& expected) {
        DWORD ignored = 0;
        DWORD size = GetFileVersionInfoSizeW(path.c_str(), &ignored);
        if (size == 0) return false;
        std::vector<unsigned char> data(size);
        if (!GetFileVersionInfoW(path.c_str(), 0, size, data.data())) return false;
        VS_FIXEDFILEINFO* info = nullptr;
        UINT info_size = 0;
        if (!VerQueryValueW(data.data(), L"\\", reinterpret_cast<void**>(&info),
                            &info_size) || info == nullptr ||
            info_size < sizeof(VS_FIXEDFILEINFO) ||
            info->dwSignature != 0xFEEF04BD) {
            return false;
        }
        return static_cast<int>(HIWORD(info->dwFileVersionMS)) == expected.major &&
               static_cast<int>(LOWORD(info->dwFileVersionMS)) == expected.minor &&
               static_cast<int>(HIWORD(info->dwFileVersionLS)) ==
                   std::max(0, expected.build) &&
               static_cast<int>(LOWORD(info->dwFileVersionLS)) ==
                   std::max(0, expected.revision);
    }

    static bool write_utf8_no_bom(const std::wstring& path, const std::wstring& text) {
        std::string utf8 = storage::fsutil::wide_to_utf8(text);
        return storage::fsutil::write_atomic_bytes(path, utf8);
    }

    static std::wstring directory_of(const std::wstring& path) {
        std::size_t slash = path.find_last_of(L"\\/");
        return slash == std::wstring::npos ? path : path.substr(0, slash);
    }
    static std::wstring file_name_of(const std::wstring& path) {
        std::size_t slash = path.find_last_of(L"\\/");
        return slash == std::wstring::npos ? path : path.substr(slash + 1);
    }
    static std::wstring extension_of(const std::wstring& path) {
        std::size_t dot = path.find_last_of(L'.');
        std::size_t slash = path.find_last_of(L"\\/");
        if (dot == std::wstring::npos || (slash != std::wstring::npos && dot < slash)) return L"";
        return path.substr(dot);
    }
    static std::wstring to_lower_w(std::wstring s) {
        for (wchar_t& c : s) c = static_cast<wchar_t>(towlower(c));
        return s;
    }

    // --- close / cancel -----------------------------------------------------

    void on_close_button() {
        if (phase_ == Phase::Downloading) {
            set_status(L"Canceling...");
            cancel_flag_ = true;
            return;
        }
        if (phase_ == Phase::Applying) {
            set_status(L"Applying update. XactCopy will close and restart automatically.");
            return;
        }
        host_.destroy();
    }

    // --- window proc --------------------------------------------------------

    LRESULT proc(HWND hwnd, UINT message, WPARAM wparam, LPARAM lparam, bool& handled) {
        switch (message) {
            case WM_CREATE:
                handled = true;
                return on_create(hwnd);
            case WM_SIZE:
                if (hwnd_ != nullptr) {
                    relayout(hwnd);
                    rebuild_notes();
                }
                handled = true;
                return 0;
            case WM_GETMINMAXINFO: {
                auto* mmi = reinterpret_cast<MINMAXINFO*>(lparam);
                const UINT d = layout_dpi_ == 0
                                   ? ui_layout_dpi(GetDpiForWindow(hwnd),
                                                   theme_.density_percent,
                                                   theme_.scale_percent)
                                   : layout_dpi_;
                mmi->ptMinTrackSize.x = MulDiv(560, static_cast<int>(d), 96);
                mmi->ptMinTrackSize.y = MulDiv(460, static_cast<int>(d), 96);
                handled = true;
                return 0;
            }
            case WM_ERASEBKGND: {
                RECT rc;
                GetClientRect(hwnd, &rc);
                themedraw::fill_rect(reinterpret_cast<HDC>(wparam), rc, theme_.window);
                handled = true;
                return 1;
            }
            case WM_CTLCOLORSTATIC: {
                HDC dc = reinterpret_cast<HDC>(wparam);
                HWND ctl = reinterpret_cast<HWND>(lparam);
                SetBkColor(dc, theme_.window);
                SetTextColor(dc, ctl == title_label_ ? theme_.accent : theme_.text);
                handled = true;
                return reinterpret_cast<LRESULT>(window_brush_);
            }
            case WM_CTLCOLORLISTBOX: {
                HDC dc = reinterpret_cast<HDC>(wparam);
                SetBkColor(dc, theme_.edit);
                SetTextColor(dc, theme_.text);
                handled = true;
                return reinterpret_cast<LRESULT>(edit_brush_);
            }
            case WM_DRAWITEM: {
                auto* draw = reinterpret_cast<DRAWITEMSTRUCT*>(lparam);
                if (draw->CtlID == IdProgressBar) {
                    draw_progress_bar(*draw);
                } else if (draw->CtlID == IdNotesList) {
                    draw_notes_item(*draw);
                } else {
                    themedraw::draw_button(*draw, theme_);
                }
                handled = true;
                return TRUE;
            }
            case WM_APP_DL_PROGRESS:
                on_progress_message();
                handled = true;
                return 0;
            case WM_APP_DL_DONE: {
                std::unique_ptr<std::string> msg(reinterpret_cast<std::string*>(lparam));
                on_done_message(static_cast<Outcome>(wparam), *msg);
                handled = true;
                return 0;
            }
            case WM_COMMAND: {
                int id = LOWORD(wparam);
                if (id == IdDownloadBtn) start_download();
                else if (id == IdViewReleaseBtn) view_release();
                else if (id == IdCloseBtn) on_close_button();
                handled = true;
                return 0;
            }
            case WM_DPICHANGED: {
                monitor_dpi_ = HIWORD(wparam) != 0 ? HIWORD(wparam) : LOWORD(wparam);
                UINT d = ui_layout_dpi(monitor_dpi_, theme_.density_percent,
                                       theme_.scale_percent);
                build_fonts(d);
                for (HWND h : {title_label_, current_label_, latest_label_, package_label_,
                               notes_list_, progress_bar_, status_label_, detail_label_,
                               download_btn_, view_btn_, close_btn_}) {
                    set_font(h, h == title_label_ ? bold_font_ : font_);
                }
                SendMessageW(notes_list_, LB_SETITEMHEIGHT, 0,
                             MAKELPARAM(MulDiv(18, static_cast<int>(d), 96), 0));
                RECT* r = reinterpret_cast<RECT*>(lparam);
                SetWindowPos(hwnd, nullptr, r->left, r->top, r->right - r->left, r->bottom - r->top,
                             SWP_NOZORDER | SWP_NOACTIVATE);
                relayout(hwnd);
                rebuild_notes();
                tooltips_.update_dpi(layout_dpi_);
                apply_window_icons(hwnd);
                RedrawWindow(hwnd, nullptr, nullptr,
                             RDW_INVALIDATE | RDW_ERASE | RDW_FRAME | RDW_ALLCHILDREN);
                handled = true;
                return 0;
            }
            case WM_CLOSE:
                on_close_button();
                handled = true;
                return 0;
            case WM_DESTROY:
                cancel_flag_ = true;
                if (worker_.joinable()) worker_.join();
                handled = true;
                return 0;
            default:
                return 0;
        }
    }

    void view_release() {
        if (release_.html_url.empty()) return;
        ShellExecuteW(hwnd_, L"open", storage::fsutil::utf8_to_wide(release_.html_url).c_str(),
                      nullptr, nullptr, SW_SHOWNORMAL);
    }
};

} // namespace xact::ui
