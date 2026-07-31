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
        if (font_ != nullptr) DeleteObject(font_);
        if (bold_font_ != nullptr) DeleteObject(bold_font_);
        if (window_brush_ != nullptr) DeleteObject(window_brush_);
        if (edit_brush_ != nullptr) DeleteObject(edit_brush_);
    }

    bool apply_launched_ = false;

    void run(HWND owner) {
        HWND hwnd = host_.create(
            owner, L"XactCopyUpdateDlg", L"XactCopy Update", 640, 560,
            [this](HWND h, UINT m, WPARAM w, LPARAM l, bool& handled) {
                return proc(h, m, w, l, handled);
            },
            /*resizable=*/true);
        if (hwnd == nullptr) return;
        apply_window_icons(hwnd);
        set_dark_title_bar(hwnd, theme_.dark);
        host_.run_modal();
    }

    // --- font helpers -------------------------------------------------------

    void build_fonts(UINT dpi) {
        if (font_ != nullptr) DeleteObject(font_);
        if (bold_font_ != nullptr) DeleteObject(bold_font_);
        auto make = [dpi](int pt, int weight) {
            return CreateFontW(-MulDiv(pt, static_cast<int>(dpi), 72), 0, 0, 0, weight, FALSE, FALSE,
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
        UINT dpi = GetDpiForWindow(hwnd);
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

    UINT dpi() { return hwnd_ != nullptr ? GetDpiForWindow(hwnd_) : 96; }

    // --- layout -------------------------------------------------------------

    void relayout(HWND hwnd) {
        RECT client{};
        GetClientRect(hwnd, &client);
        const int w = client.right;
        const int h = client.bottom;
        UINT d = GetDpiForWindow(hwnd);
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
        rect.left += MulDiv(8, static_cast<int>(GetDpiForWindow(hwnd_)), 96);
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

        temp_root_ = create_temp_root();
        create_directory_tree(temp_root_);
        download_path_ = temp_root_ + L"\\" + storage::fsutil::utf8_to_wide(asset_->name);

        UpdateReleaseInfo release = release_;
        UpdateAssetInfo asset = *asset_;
        std::wstring dest = download_path_;
        HWND hwnd = hwnd_;
        std::int64_t total = asset.size;

        if (worker_.joinable()) worker_.join();
        worker_ = std::thread([this, hwnd, release, asset, dest, total]() {
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

            std::string actual = UpdateService::download_asset(asset.download_url, dest, total,
                                                              on_progress, cancelled);
            if (cancel_flag_.load()) {
                post_done(hwnd, Outcome::Cancelled, "Download canceled.");
                return;
            }
            if (actual.empty()) {
                post_done(hwnd, Outcome::Failed, "Download failed.");
                return;
            }
            // Verify.
            std::string expected = UpdateService::resolve_asset_sha256(release, asset);
            if (expected.empty()) {
                post_done(hwnd, Outcome::Failed,
                          "Update package checksum is unavailable for this release.");
                return;
            }
            if (!update_detail::equals_hex_ci(expected, actual)) {
                post_done(hwnd, Outcome::Failed, "Update package integrity check failed.");
                return;
            }
            post_done(hwnd, Outcome::Ok, std::string());
        });
    }

    static void post_done(HWND hwnd, Outcome outcome, const std::string& message) {
        auto* msg = new std::string(message);
        PostMessageW(hwnd, WM_APP_DL_DONE, static_cast<WPARAM>(outcome),
                     reinterpret_cast<LPARAM>(msg));
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
            DestroyWindow(hwnd_); // returns to caller, which exits the app
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

        if (!folder_writable(target_dir)) {
            error = L"The application folder is not writable. Run XactCopy with write access to "
                    L"install updates.";
            return false;
        }

        std::wstring ext = to_lower_w(extension_of(download_path_));
        if (ext == L".exe") {
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
        script += L"$ErrorActionPreference = 'Continue'\r\n";
        script += L"$TargetPid  = " + std::to_wstring(pid) + L"\r\n";
        script += L"$Zip        = '" + q(download_path_) + L"'\r\n";
        script += L"$Target     = '" + q(normalized) + L"'\r\n";
        script += L"$Backup     = '" + q(backup) + L"'\r\n";
        script += L"$Payload    = '" + q(payload) + L"'\r\n";
        script += L"$Log        = '" + q(log) + L"'\r\n";
        script += L"$TempRoot   = '" + q(temp_root_) + L"'\r\n";
        script += L"$Executable = '" + q(executable) + L"'\r\n";
        script += L"$ExeName    = '" + q(exe_name) + L"'\r\n\r\n";
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
        script += L"try { [System.IO.Compression.ZipFile]::ExtractToDirectory($Zip, $Payload) } "
                  L"catch { $_ | Out-File -LiteralPath $Log -Append -Encoding UTF8; return }\r\n\r\n";
        script += L"$Source = $Payload\r\n";
        script += L"$found = Get-ChildItem -LiteralPath $Payload -Filter $ExeName -Recurse -File | "
                  L"Select-Object -First 1\r\n";
        script += L"if ($found) { $Source = $found.DirectoryName }\r\n\r\n";
        script += L"if (Test-Path -LiteralPath $Backup) { Remove-Item -LiteralPath $Backup "
                  L"-Recurse -Force }\r\n";
        script += L"$null = New-Item -ItemType Directory -Path $Backup -Force\r\n";
        script += L"robocopy $Target $Backup /E /COPY:DAT /R:2 /W:1 /NFL /NDL /NJH /NJS /NP >> "
                  L"$Log\r\n\r\n";
        script += L"robocopy $Source $Target /E /COPY:DAT /R:2 /W:1 /NFL /NDL /NJH /NJS /NP >> "
                  L"$Log\r\n";
        script += L"$copyRc = $LASTEXITCODE\r\n";
        script += L"if ($copyRc -ge 8) {\r\n";
        script += L"    'Update copy failed. Restoring backup.' | Out-File -LiteralPath $Log "
                  L"-Append -Encoding UTF8\r\n";
        script += L"    robocopy $Backup $Target /E /COPY:DAT /R:2 /W:1 /NFL /NDL /NJH /NJS /NP >> "
                  L"$Log\r\n";
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
        HINSTANCE rc = ShellExecuteW(nullptr, L"open", L"powershell.exe", args.c_str(), nullptr,
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
        GetTempPathW(MAX_PATH, temp);
        std::wstring base = std::wstring(temp) + L"XactCopyUpdate";
        wchar_t suffix[32];
        swprintf(suffix, 32, L"%08X%08X", static_cast<unsigned>(GetCurrentProcessId()),
                 static_cast<unsigned>(GetTickCount()));
        return base + L"\\" + suffix;
    }

    static void create_directory_tree(const std::wstring& path) {
        std::wstring accum;
        for (std::size_t i = 0; i < path.size(); ++i) {
            accum.push_back(path[i]);
            if (path[i] == L'\\' || i + 1 == path.size()) {
                if (accum.size() > 3) CreateDirectoryW(accum.c_str(), nullptr);
            }
        }
    }

    static bool folder_writable(const std::wstring& folder) {
        std::wstring probe = folder + L"\\write-probe.tmp";
        HANDLE h = CreateFileW(probe.c_str(), GENERIC_WRITE, 0, nullptr, CREATE_ALWAYS,
                               FILE_ATTRIBUTE_TEMPORARY | FILE_FLAG_DELETE_ON_CLOSE, nullptr);
        if (h == INVALID_HANDLE_VALUE) return false;
        CloseHandle(h);
        return true;
    }

    static bool write_utf8_no_bom(const std::wstring& path, const std::wstring& text) {
        std::string utf8 = storage::fsutil::wide_to_utf8(text);
        HANDLE h = CreateFileW(path.c_str(), GENERIC_WRITE, 0, nullptr, CREATE_ALWAYS,
                               FILE_ATTRIBUTE_NORMAL, nullptr);
        if (h == INVALID_HANDLE_VALUE) return false;
        DWORD written = 0;
        bool ok = WriteFile(h, utf8.data(), static_cast<DWORD>(utf8.size()), &written, nullptr) &&
                  written == utf8.size();
        CloseHandle(h);
        return ok;
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
        DestroyWindow(hwnd_);
    }

    // --- window proc --------------------------------------------------------

    LRESULT proc(HWND hwnd, UINT message, WPARAM wparam, LPARAM lparam, bool& handled) {
        switch (message) {
            case WM_CREATE:
                handled = true;
                return on_create(hwnd);
            case WM_SIZE:
                if (hwnd_ != nullptr) {
                    rebuild_notes();
                    relayout(hwnd);
                }
                handled = true;
                return 0;
            case WM_GETMINMAXINFO: {
                auto* mmi = reinterpret_cast<MINMAXINFO*>(lparam);
                UINT d = GetDpiForWindow(hwnd);
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
                UINT d = HIWORD(wparam);
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
                rebuild_notes();
                relayout(hwnd);
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
