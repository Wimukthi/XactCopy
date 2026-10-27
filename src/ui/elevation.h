// -----------------------------------------------------------------------------
// File: src\ui\elevation.h
// Purpose: Administrator elevation for direct-NTFS (raw volume) assessments.
//          Opening \\.\X: needs an elevated token, so a raw assessment started
//          from a normal window relaunches XactCopy through UAC and hands the
//          exact start request to the elevated instance. The handoff file's
//          SHA-256 travels on the elevated command line, which UAC fixes at
//          consent time, so the file cannot be swapped after approval.
//          Also: UIPI message filters an elevated window needs, and launching
//          URLs/folders as the signed-in user rather than as Administrator.
// -----------------------------------------------------------------------------

#pragma once

#include <chrono>
#include <cstdint>
#include <string>
#include <vector>

#ifndef NOMINMAX
#define NOMINMAX
#endif
#ifndef WIN32_LEAN_AND_MEAN
#define WIN32_LEAN_AND_MEAN
#endif
#include <windows.h>
#include <ole2.h> // exdisp.h/shldisp.h need OLE types that WIN32_LEAN_AND_MEAN drops
#include <exdisp.h>
#include <sddl.h>
#include <shellapi.h>
#include <shldisp.h>
#include <shlobj.h>

#include "../core/crypto.h"
#include "../core/json.h"
#include "../core/models.h"
#include "../storage/stores.h"

namespace xact::ui::elevation {

// Command-line switches understood by wWinMain.
inline constexpr const wchar_t* HandoffSwitch = L"--elevated-handoff";
inline constexpr const wchar_t* OriginSidSwitch = L"--origin-sid";

// Exit codes an elevated instance uses to tell the waiting original why it
// refused the handoff.
inline constexpr DWORD ExitHandoffWrongUser = 3;
inline constexpr DWORD ExitHandoffRejected = 4;
inline constexpr DWORD ExitHandoffMutexTimeout = 5;

// A handoff older than this is refused; UAC consent plus startup takes seconds.
inline constexpr std::int64_t HandoffMaximumAgeMs = 10 * 60 * 1000;

inline bool is_process_elevated() {
    HANDLE token = nullptr;
    if (!OpenProcessToken(GetCurrentProcess(), TOKEN_QUERY, &token)) return false;
    TOKEN_ELEVATION elevation{};
    DWORD size = 0;
    const bool elevated =
        GetTokenInformation(token, TokenElevation, &elevation, sizeof(elevation), &size) &&
        elevation.TokenIsElevated != 0;
    CloseHandle(token);
    return elevated;
}

inline std::string current_user_sid() {
    HANDLE token = nullptr;
    if (!OpenProcessToken(GetCurrentProcess(), TOKEN_QUERY, &token)) return std::string();
    DWORD size = 0;
    GetTokenInformation(token, TokenUser, nullptr, 0, &size);
    std::vector<unsigned char> buffer(size);
    std::string result;
    if (size != 0 && GetTokenInformation(token, TokenUser, buffer.data(), size, &size)) {
        const auto* user = reinterpret_cast<const TOKEN_USER*>(buffer.data());
        LPWSTR text = nullptr;
        if (ConvertSidToStringSidW(user->User.Sid, &text) && text != nullptr) {
            result = storage::fsutil::wide_to_utf8(text);
            LocalFree(text);
        }
    }
    CloseHandle(token);
    return result;
}

inline std::int64_t unix_now_ms() {
    return std::chrono::duration_cast<std::chrono::milliseconds>(
               std::chrono::system_clock::now().time_since_epoch())
        .count();
}

// True when this run would use direct NTFS reads but cannot, only because the
// process lacks elevation. Sources that can never use the raw backend (UNC,
// non-NTFS, unresolvable) return false with `unsupported_reason` set, so they
// do not cost the user a pointless UAC prompt; the worker falls back to
// standard reads for them as it always has.
inline bool raw_scan_requires_elevation(const models::CopyJobOptions& options,
                                        bool process_elevated,
                                        std::string& unsupported_reason) {
    unsupported_reason.clear();
    if (options.OperationMode != models::JobOperationMode::ScanOnly ||
        !options.UseExperimentalRawDiskScan) {
        return false;
    }
    std::wstring source =
        storage::fsutil::get_full_path(storage::fsutil::utf8_to_wide(options.SourceRoot));
    if (source.size() < 2 || source[1] != L':') {
        unsupported_reason = "direct NTFS reads need a local drive-letter source";
        return false;
    }
    wchar_t volume_root[MAX_PATH]{};
    if (!GetVolumePathNameW(storage::fsutil::long_path(source).c_str(), volume_root, MAX_PATH)) {
        wcsncpy_s(volume_root, (source.substr(0, 2) + L"\\").c_str(), _TRUNCATE);
    }
    wchar_t file_system[64]{};
    if (!GetVolumeInformationW(volume_root, nullptr, 0, nullptr, nullptr, nullptr, file_system,
                               static_cast<DWORD>(std::size(file_system)))) {
        unsupported_reason = "the source volume could not be inspected";
        return false;
    }
    if (_wcsicmp(file_system, L"NTFS") != 0) {
        unsupported_reason = "the source volume is " + storage::fsutil::wide_to_utf8(file_system) +
                             ", not NTFS";
        return false;
    }
    return !process_elevated;
}

// Everything the elevated instance needs to start the same run the user asked
// for in the unelevated window.
struct HandoffRequest {
    models::CopyJobOptions Options;
    bool FromRecovery = false;
    std::string ManagedRunId;
    std::string DisplayName;
    std::string RecoveryRunId; // crash-recovery run this start resumes, if any
    std::string UserSid;
    std::int64_t CreatedUnixMs = 0;

    std::string to_json() const {
        json::Writer w;
        w.begin_object();
        w.key("SchemaVersion"); w.value(static_cast<std::int32_t>(1));
        w.key("Options"); Options.to_json(w);
        w.key("FromRecovery"); w.value(FromRecovery);
        w.key("ManagedRunId"); w.value(ManagedRunId);
        w.key("DisplayName"); w.value(DisplayName);
        w.key("RecoveryRunId"); w.value(RecoveryRunId);
        w.key("UserSid"); w.value(UserSid);
        w.key("CreatedUnixMs"); w.value(CreatedUnixMs);
        w.end_object();
        return w.take();
    }

    static bool from_json(std::string_view text, HandoffRequest& out) {
        try {
            const json::Value parsed = json::parse(text);
            const json::Object* object = parsed.as_object();
            if (object == nullptr) return false;
            const json::Object& root = *object;
            const json::Value* schema = root.find("SchemaVersion");
            if (schema == nullptr || schema->as_int64() != 1) return false;
            const json::Value* options = root.find("Options");
            if (options == nullptr || !options->is_object()) return false;
            out.Options = models::CopyJobOptions::from_json(*options);
            auto text_of = [&root](const char* key) {
                const json::Value* value = root.find(key);
                return value != nullptr && value->is_string() ? value->as_string() : std::string();
            };
            const json::Value* from_recovery = root.find("FromRecovery");
            out.FromRecovery = from_recovery != nullptr && from_recovery->as_bool();
            out.ManagedRunId = text_of("ManagedRunId");
            out.DisplayName = text_of("DisplayName");
            out.RecoveryRunId = text_of("RecoveryRunId");
            out.UserSid = text_of("UserSid");
            const json::Value* created = root.find("CreatedUnixMs");
            out.CreatedUnixMs = created != nullptr ? created->as_int64() : 0;
            return true;
        } catch (const std::exception&) {
            return false;
        }
    }
};

inline std::wstring handoff_directory() {
    return storage::fsutil::local_app_data() + L"\\XactCopy\\handoff";
}

inline bool is_handoff_token(const std::string& token) {
    if (token.size() != 64) return false;
    for (char c : token) {
        if (!((c >= '0' && c <= '9') || (c >= 'a' && c <= 'f'))) return false;
    }
    return true;
}

inline std::wstring handoff_path(const std::string& token) {
    return handoff_directory() + L"\\" + storage::fsutil::utf8_to_wide(token) + L".json";
}

// Writes the request and returns its token (lowercase SHA-256 hex of the file
// bytes), or an empty string on failure.
inline std::string write_handoff(const HandoffRequest& request) {
    const std::string payload = request.to_json();
    const std::string token = crypto::to_hex_lower(crypto::sha256(payload));
    storage::fsutil::create_directories(handoff_directory());
    const std::wstring path = handoff_path(token);
    if (!storage::fsutil::write_file_raw(path,
                                         reinterpret_cast<const unsigned char*>(payload.data()),
                                         payload.size(), true, false)) {
        return std::string();
    }
    return token;
}

inline void discard_handoff(const std::string& token) {
    if (!is_handoff_token(token)) return;
    DeleteFileW(storage::fsutil::long_path(handoff_path(token)).c_str());
}

inline bool handoff_pending(const std::string& token) {
    return is_handoff_token(token) &&
           GetFileAttributesW(storage::fsutil::long_path(handoff_path(token)).c_str()) !=
               INVALID_FILE_ATTRIBUTES;
}

enum class HandoffStatus { Ok, Invalid, Missing, Tampered, Expired, WrongUser, NotRawScan };

inline const char* describe(HandoffStatus status) {
    switch (status) {
        case HandoffStatus::Ok: return "accepted";
        case HandoffStatus::Invalid: return "the request token is malformed";
        case HandoffStatus::Missing: return "the request file was not found";
        case HandoffStatus::Tampered: return "the request file does not match the approved request";
        case HandoffStatus::Expired: return "the request is too old";
        case HandoffStatus::WrongUser: return "the request belongs to a different Windows account";
        case HandoffStatus::NotRawScan:
            return "only direct-NTFS readability assessments are started through elevation";
    }
    return "unknown";
}

// Reads, validates, and deletes the handoff. Deleting the file is also the
// acceptance signal the waiting unelevated instance polls for, so it happens
// for every outcome once the file has been read.
inline HandoffStatus consume_handoff(const std::string& token, HandoffRequest& out,
                                     std::int64_t now_ms = unix_now_ms(),
                                     std::string expected_sid = std::string()) {
    if (!is_handoff_token(token)) return HandoffStatus::Invalid;
    const std::wstring path = handoff_path(token);
    auto bytes = storage::fsutil::read_all_bytes(path, 4ULL * 1024ULL * 1024ULL);
    if (!bytes.has_value()) return HandoffStatus::Missing;
    DeleteFileW(storage::fsutil::long_path(path).c_str());

    if (crypto::to_hex_lower(crypto::sha256(*bytes)) != token) return HandoffStatus::Tampered;
    const std::string text(bytes->begin(), bytes->end());
    if (!HandoffRequest::from_json(text, out)) return HandoffStatus::Invalid;
    if (expected_sid.empty()) expected_sid = current_user_sid();
    if (out.UserSid.empty() || out.UserSid != expected_sid) return HandoffStatus::WrongUser;
    const std::int64_t age = now_ms - out.CreatedUnixMs;
    if (out.CreatedUnixMs <= 0 || age < -60'000 || age > HandoffMaximumAgeMs) {
        return HandoffStatus::Expired;
    }
    if (out.Options.OperationMode != models::JobOperationMode::ScanOnly ||
        !out.Options.UseExperimentalRawDiskScan) {
        return HandoffStatus::NotRawScan;
    }
    return HandoffStatus::Ok;
}

// Removes handoff files a crashed or refused relaunch left behind.
inline void cleanup_stale_handoffs() {
    const auto files = storage::fsutil::enumerate_files(handoff_directory(), L"*.json");
    FILETIME now_ft{};
    GetSystemTimeAsFileTime(&now_ft);
    const ULONGLONG now = (static_cast<ULONGLONG>(now_ft.dwHighDateTime) << 32) | now_ft.dwLowDateTime;
    for (const auto& file : files) {
        WIN32_FILE_ATTRIBUTE_DATA data{};
        if (!GetFileAttributesExW(storage::fsutil::long_path(file).c_str(), GetFileExInfoStandard,
                                  &data)) {
            continue;
        }
        const ULONGLONG written = (static_cast<ULONGLONG>(data.ftLastWriteTime.dwHighDateTime) << 32) |
                                  data.ftLastWriteTime.dwLowDateTime;
        constexpr ULONGLONG OneDay100ns = 24ULL * 60ULL * 60ULL * 10'000'000ULL;
        if (now > written && now - written > OneDay100ns) {
            DeleteFileW(storage::fsutil::long_path(file).c_str());
        }
    }
}

enum class LaunchOutcome { Started, Declined, Failed };

// Starts an elevated copy of this executable with the handoff token. Blocks
// while the UAC consent prompt is shown. On Started, `process` receives the
// new process handle (SYNCHRONIZE/query rights are enough to wait on it).
inline LaunchOutcome relaunch_elevated(HWND owner, const std::string& token,
                                       HANDLE& process, DWORD& error) {
    process = nullptr;
    error = 0;
    wchar_t module[MAX_PATH * 4]{};
    const DWORD length = GetModuleFileNameW(nullptr, module, static_cast<DWORD>(std::size(module)));
    if (length == 0 || length >= std::size(module)) {
        error = GetLastError();
        return LaunchOutcome::Failed;
    }
    const std::wstring parameters = std::wstring(HandoffSwitch) + L" " +
                                    storage::fsutil::utf8_to_wide(token) + L" " +
                                    OriginSidSwitch + L" " +
                                    storage::fsutil::utf8_to_wide(current_user_sid());
    SHELLEXECUTEINFOW info{};
    info.cbSize = sizeof(info);
    info.fMask = SEE_MASK_NOCLOSEPROCESS | SEE_MASK_NOASYNC;
    info.hwnd = owner;
    info.lpVerb = L"runas";
    info.lpFile = module;
    info.lpParameters = parameters.c_str();
    info.nShow = SW_SHOWNORMAL;
    if (!ShellExecuteExW(&info)) {
        error = GetLastError();
        return error == ERROR_CANCELLED ? LaunchOutcome::Declined : LaunchOutcome::Failed;
    }
    if (info.hProcess == nullptr) {
        error = ERROR_INVALID_HANDLE;
        return LaunchOutcome::Failed;
    }
    process = info.hProcess;
    return LaunchOutcome::Started;
}

// An elevated window runs at High integrity, and UIPI drops WM_COPYDATA and
// drag-and-drop messages from normal-integrity senders such as Explorer and a
// second XactCopy launched from an Explorer verb. Admit exactly those.
inline void allow_unelevated_messages(HWND window) {
    constexpr UINT WM_COPYGLOBALDATA = 0x0049;
    for (UINT message : {static_cast<UINT>(WM_COPYDATA), static_cast<UINT>(WM_DROPFILES),
                         WM_COPYGLOBALDATA}) {
        ChangeWindowMessageFilterEx(window, message, MSGFLT_ALLOW, nullptr);
    }
}

namespace detail {

// Asks the desktop shell (which runs unelevated) to perform the launch, so
// a browser or Explorer window opened from an elevated XactCopy does not run
// as Administrator. Returns false when the shell is not available.
inline bool shell_execute_via_desktop(const std::wstring& target, const std::wstring& parameters) {
    const HRESULT init = CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED | COINIT_DISABLE_OLE1DDE);
    const bool uninitialize = SUCCEEDED(init);
    bool launched = false;
    IShellWindows* windows = nullptr;
    if (SUCCEEDED(CoCreateInstance(CLSID_ShellWindows, nullptr, CLSCTX_LOCAL_SERVER,
                                   IID_PPV_ARGS(&windows)))) {
        VARIANT empty{};
        VariantInit(&empty);
        long hwnd = 0;
        IDispatch* desktop_dispatch = nullptr;
        if (windows->FindWindowSW(&empty, &empty, SWC_DESKTOP, &hwnd, SWFO_NEEDDISPATCH,
                                  &desktop_dispatch) == S_OK &&
            desktop_dispatch != nullptr) {
            IServiceProvider* provider = nullptr;
            IShellBrowser* browser = nullptr;
            IShellView* view = nullptr;
            IDispatch* background = nullptr;
            IShellFolderViewDual* folder_view = nullptr;
            IDispatch* application = nullptr;
            IShellDispatch2* shell = nullptr;
            if (SUCCEEDED(desktop_dispatch->QueryInterface(IID_PPV_ARGS(&provider))) &&
                SUCCEEDED(provider->QueryService(SID_STopLevelBrowser, IID_PPV_ARGS(&browser))) &&
                SUCCEEDED(browser->QueryActiveShellView(&view)) &&
                SUCCEEDED(view->GetItemObject(SVGIO_BACKGROUND, IID_PPV_ARGS(&background))) &&
                SUCCEEDED(background->QueryInterface(IID_PPV_ARGS(&folder_view))) &&
                SUCCEEDED(folder_view->get_Application(&application)) &&
                SUCCEEDED(application->QueryInterface(IID_PPV_ARGS(&shell)))) {
                BSTR file = SysAllocString(target.c_str());
                VARIANT args{};
                VariantInit(&args);
                if (!parameters.empty()) {
                    args.vt = VT_BSTR;
                    args.bstrVal = SysAllocString(parameters.c_str());
                }
                VARIANT none{};
                VariantInit(&none);
                VARIANT show{};
                VariantInit(&show);
                show.vt = VT_I4;
                show.lVal = SW_SHOWNORMAL;
                launched = file != nullptr && SUCCEEDED(shell->ShellExecute(file, args, none, none, show));
                VariantClear(&args);
                if (file != nullptr) SysFreeString(file);
            }
            if (shell != nullptr) shell->Release();
            if (application != nullptr) application->Release();
            if (folder_view != nullptr) folder_view->Release();
            if (background != nullptr) background->Release();
            if (view != nullptr) view->Release();
            if (browser != nullptr) browser->Release();
            if (provider != nullptr) provider->Release();
            desktop_dispatch->Release();
        }
        windows->Release();
    }
    if (uninitialize) CoUninitialize();
    return launched;
}

} // namespace detail

// Opens a URL, document, or folder for the user. From an elevated instance it
// goes through the unelevated desktop shell; if that is unavailable the launch
// is still performed, elevated, rather than silently doing nothing.
inline bool open_as_user(HWND owner, const std::wstring& target,
                         const std::wstring& parameters = std::wstring()) {
    if (is_process_elevated() && detail::shell_execute_via_desktop(target, parameters)) {
        return true;
    }
    const HINSTANCE result = ShellExecuteW(owner, L"open", target.c_str(),
                                           parameters.empty() ? nullptr : parameters.c_str(),
                                           nullptr, SW_SHOWNORMAL);
    return reinterpret_cast<INT_PTR>(result) > 32;
}

} // namespace xact::ui::elevation
