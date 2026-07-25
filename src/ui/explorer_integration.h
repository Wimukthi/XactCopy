// ExplorerIntegrationService: comprehensive per-user Windows shell integration.
// All registration lives under HKCU (no elevation), so it can be toggled from
// the running app via the EnableExplorerContextMenu setting. It registers:
//   * "Copy with XactCopy" verbs on files, folders, drives, and folder
//     backgrounds (the original context-menu integration).
//   * "Scan for Bad Blocks with XactCopy" verbs on folders and drives.
//   * An Applications\XactCopy.exe entry (FriendlyAppName + DefaultIcon + open
//     command + SupportedTypes) so XactCopy appears in "Open with" and accepts
//     files dragged onto it.
//   * An App Paths entry so "XactCopy" launches from the Run dialog / search.
//   * A "Send to > XactCopy" shortcut.
// unregister() removes every one of these.
#pragma once

#ifndef NOMINMAX
#define NOMINMAX
#endif
#include <windows.h>
#include <shlobj.h>
#include <objbase.h>

#include <string>
#include <vector>

namespace xact::ui {

class ExplorerIntegrationService {
public:
    struct Target {
        const wchar_t* shell_path;
        const wchar_t* command_arguments;
        bool supports_multi_select;
    };

    // The original "Copy with XactCopy" verbs (kept for is_registered()).
    static const std::vector<Target>& targets() {
        static const std::vector<Target> table = {
            {L"Directory", L"--from-explorer \"%1\" %*", true},
            {L"Drive", L"--from-explorer \"%1\" %*", true},
            {L"*", L"--from-explorer \"%1\" %*", true},
            {L"Directory\\Background", L"--from-explorer-folder \"%V\"", false},
        };
        return table;
    }

    // "Scan for Bad Blocks with XactCopy" verbs (folders + drives).
    static const std::vector<Target>& scan_targets() {
        static const std::vector<Target> table = {
            {L"Directory", L"--scan-from-explorer \"%1\" %*", true},
            {L"Drive", L"--scan-from-explorer \"%1\" %*", true},
        };
        return table;
    }

    // Registers when enabled, removes otherwise. Returns false (without throwing)
    // if registration was requested with an invalid executable path.
    static bool sync(bool enabled, const std::wstring& executable_path) {
        std::wstring resolved = trim(executable_path);
        if (enabled) {
            if (resolved.empty() || !file_exists(resolved)) return false;
            return register_all(resolved);
        }
        unregister_all();
        return true;
    }

    static bool is_registered() {
        for (const auto& target : targets()) {
            if (!key_exists(build_menu_path(target.shell_path, MenuKeyName))) return false;
        }
        // Also require the comprehensive keys, so an upgrade from a copy-only
        // install re-syncs the scan verb + Applications entry.
        if (!key_exists(build_menu_path(L"Drive", ScanKeyName))) return false;
        if (!key_exists(std::wstring(ShellClassesRoot) + L"\\Applications\\" + AppExeName)) {
            return false;
        }
        return true;
    }

private:
    static constexpr const wchar_t* ShellClassesRoot = L"Software\\Classes";
    static constexpr const wchar_t* MenuKeyName = L"XactCopy";
    static constexpr const wchar_t* ScanKeyName = L"XactCopyScan";
    static constexpr const wchar_t* AppExeName = L"XactCopy.exe";

    static std::wstring build_menu_path(const std::wstring& shell_path, const wchar_t* key_name) {
        return std::wstring(ShellClassesRoot) + L"\\" + shell_path + L"\\shell\\" + key_name;
    }

    static bool key_exists(const std::wstring& path) {
        HKEY key = nullptr;
        if (RegOpenKeyExW(HKEY_CURRENT_USER, path.c_str(), 0, KEY_READ, &key) != ERROR_SUCCESS) {
            return false;
        }
        RegCloseKey(key);
        return true;
    }

    static bool register_all(const std::wstring& executable_path) {
        bool ok = true;
        ok &= register_verbs(targets(), MenuKeyName, L"Copy with XactCopy", executable_path);
        ok &= register_verbs(scan_targets(), ScanKeyName, L"Scan for Bad Blocks with XactCopy",
                             executable_path);
        register_application(executable_path);
        register_app_paths(executable_path);
        register_send_to(executable_path);
        return ok;
    }

    static void unregister_all() {
        for (const auto& target : targets()) {
            RegDeleteTreeW(HKEY_CURRENT_USER, build_menu_path(target.shell_path, MenuKeyName).c_str());
        }
        for (const auto& target : scan_targets()) {
            RegDeleteTreeW(HKEY_CURRENT_USER, build_menu_path(target.shell_path, ScanKeyName).c_str());
        }
        RegDeleteTreeW(HKEY_CURRENT_USER,
                       (std::wstring(ShellClassesRoot) + L"\\Applications\\" + AppExeName).c_str());
        RegDeleteTreeW(HKEY_CURRENT_USER,
                       (std::wstring(L"Software\\Microsoft\\Windows\\CurrentVersion\\App Paths\\") +
                        AppExeName)
                           .c_str());
        remove_send_to();
    }

    // Writes a context-menu verb (caption + Icon + MultiSelectModel + command)
    // for each target under the given key name.
    static bool register_verbs(const std::vector<Target>& table, const wchar_t* key_name,
                               const wchar_t* caption, const std::wstring& executable_path) {
        bool all_ok = true;
        for (const auto& target : table) {
            std::wstring menu_path = build_menu_path(target.shell_path, key_name);
            HKEY menu_key = nullptr;
            if (RegCreateKeyExW(HKEY_CURRENT_USER, menu_path.c_str(), 0, nullptr,
                                REG_OPTION_NON_VOLATILE, KEY_WRITE, nullptr, &menu_key,
                                nullptr) != ERROR_SUCCESS) {
                all_ok = false;
                continue;
            }
            set_string(menu_key, nullptr, caption);
            set_string(menu_key, L"Icon", executable_path.c_str());
            if (target.supports_multi_select) {
                set_string(menu_key, L"MultiSelectModel", L"Player");
            } else {
                RegDeleteValueW(menu_key, L"MultiSelectModel");
            }
            HKEY command_key = nullptr;
            if (RegCreateKeyExW(menu_key, L"command", 0, nullptr, REG_OPTION_NON_VOLATILE, KEY_WRITE,
                                nullptr, &command_key, nullptr) == ERROR_SUCCESS) {
                std::wstring command = L"\"" + executable_path + L"\" " + target.command_arguments;
                set_string(command_key, nullptr, command.c_str());
                RegCloseKey(command_key);
            } else {
                all_ok = false;
            }
            RegCloseKey(menu_key);
        }
        return all_ok;
    }

    // Applications\XactCopy.exe: enables "Open with" + drag-and-drop onto the app.
    static void register_application(const std::wstring& executable_path) {
        std::wstring root = std::wstring(ShellClassesRoot) + L"\\Applications\\" + AppExeName;
        HKEY key = nullptr;
        if (RegCreateKeyExW(HKEY_CURRENT_USER, root.c_str(), 0, nullptr, REG_OPTION_NON_VOLATILE,
                            KEY_WRITE, nullptr, &key, nullptr) == ERROR_SUCCESS) {
            set_string(key, L"FriendlyAppName", L"XactCopy");
            RegCloseKey(key);
        }
        HKEY icon_key = nullptr;
        if (RegCreateKeyExW(HKEY_CURRENT_USER, (root + L"\\DefaultIcon").c_str(), 0, nullptr,
                            REG_OPTION_NON_VOLATILE, KEY_WRITE, nullptr, &icon_key, nullptr) ==
            ERROR_SUCCESS) {
            set_string(icon_key, nullptr, (L"\"" + executable_path + L"\",0").c_str());
            RegCloseKey(icon_key);
        }
        HKEY command_key = nullptr;
        if (RegCreateKeyExW(HKEY_CURRENT_USER, (root + L"\\shell\\open\\command").c_str(), 0, nullptr,
                            REG_OPTION_NON_VOLATILE, KEY_WRITE, nullptr, &command_key, nullptr) ==
            ERROR_SUCCESS) {
            set_string(command_key, nullptr,
                       (L"\"" + executable_path + L"\" --from-explorer \"%1\" %*").c_str());
            RegCloseKey(command_key);
        }
        // SupportedTypes: an empty ".*" value lets XactCopy show for any file.
        HKEY types_key = nullptr;
        if (RegCreateKeyExW(HKEY_CURRENT_USER, (root + L"\\SupportedTypes").c_str(), 0, nullptr,
                            REG_OPTION_NON_VOLATILE, KEY_WRITE, nullptr, &types_key, nullptr) ==
            ERROR_SUCCESS) {
            set_string(types_key, L".*", L"");
            RegCloseKey(types_key);
        }
    }

    // App Paths: lets "XactCopy" resolve from the Run dialog and Start search.
    static void register_app_paths(const std::wstring& executable_path) {
        std::wstring path = std::wstring(L"Software\\Microsoft\\Windows\\CurrentVersion\\App Paths\\") +
                            AppExeName;
        HKEY key = nullptr;
        if (RegCreateKeyExW(HKEY_CURRENT_USER, path.c_str(), 0, nullptr, REG_OPTION_NON_VOLATILE,
                            KEY_WRITE, nullptr, &key, nullptr) == ERROR_SUCCESS) {
            set_string(key, nullptr, executable_path.c_str());
            set_string(key, L"Path", directory_of(executable_path).c_str());
            RegCloseKey(key);
        }
    }

    // "Send to > XactCopy" shortcut in the user's SendTo folder.
    static void register_send_to(const std::wstring& executable_path) {
        std::wstring link = send_to_link_path();
        if (link.empty()) return;
        bool com_initialized = SUCCEEDED(CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED));
        IShellLinkW* shell_link = nullptr;
        if (SUCCEEDED(CoCreateInstance(CLSID_ShellLink, nullptr, CLSCTX_INPROC_SERVER,
                                       IID_IShellLinkW, reinterpret_cast<void**>(&shell_link)))) {
            shell_link->SetPath(executable_path.c_str());
            shell_link->SetDescription(L"Copy with XactCopy");
            shell_link->SetIconLocation(executable_path.c_str(), 0);
            IPersistFile* persist = nullptr;
            if (SUCCEEDED(shell_link->QueryInterface(IID_IPersistFile,
                                                     reinterpret_cast<void**>(&persist)))) {
                persist->Save(link.c_str(), TRUE);
                persist->Release();
            }
            shell_link->Release();
        }
        if (com_initialized) CoUninitialize();
    }

    static void remove_send_to() {
        std::wstring link = send_to_link_path();
        if (!link.empty()) DeleteFileW(link.c_str());
    }

    static std::wstring send_to_link_path() {
        wchar_t folder[MAX_PATH]{};
        if (SHGetFolderPathW(nullptr, CSIDL_SENDTO, nullptr, SHGFP_TYPE_CURRENT, folder) != S_OK) {
            return std::wstring();
        }
        return std::wstring(folder) + L"\\XactCopy.lnk";
    }

    static void set_string(HKEY key, const wchar_t* name, const wchar_t* value) {
        DWORD bytes = static_cast<DWORD>((wcslen(value) + 1) * sizeof(wchar_t));
        RegSetValueExW(key, name, 0, REG_SZ, reinterpret_cast<const BYTE*>(value), bytes);
    }

    static bool file_exists(const std::wstring& path) {
        DWORD attributes = GetFileAttributesW(path.c_str());
        return attributes != INVALID_FILE_ATTRIBUTES &&
               (attributes & FILE_ATTRIBUTE_DIRECTORY) == 0;
    }

    static std::wstring directory_of(const std::wstring& path) {
        std::size_t slash = path.find_last_of(L"\\/");
        return slash == std::wstring::npos ? path : path.substr(0, slash);
    }

    static std::wstring trim(const std::wstring& text) {
        std::size_t begin = 0;
        std::size_t end = text.size();
        auto is_space = [](wchar_t c) {
            return c == L' ' || c == L'\t' || c == L'\r' || c == L'\n';
        };
        while (begin < end && is_space(text[begin])) ++begin;
        while (end > begin && is_space(text[end - 1])) --end;
        return text.substr(begin, end - begin);
    }
};

} // namespace xact::ui
