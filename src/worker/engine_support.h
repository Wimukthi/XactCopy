// -----------------------------------------------------------------------------
// File: cpp\src\worker\engine_support.h
// Purpose: Support types for the native copy engine: error/cancellation model,
//          pause control, file metadata, directory scanner (FindFirstFileEx),
//          transfer session with timed positional I/O, and the env-driven
//          fault injector mirroring the .NET worker's DevFaultInjector.
// -----------------------------------------------------------------------------

#pragma once

#include <atomic>
#include <charconv>
#include <cstdint>
#include <functional>
#include <mutex>
#include <optional>
#include <random>
#include <stdexcept>
#include <string>
#include <string_view>
#include <unordered_set>
#include <vector>

#ifndef NOMINMAX
#define NOMINMAX
#endif
#ifndef WIN32_LEAN_AND_MEAN
#define WIN32_LEAN_AND_MEAN
#endif
#include <windows.h>

#include "../core/dotnet_time.h"
#include "../storage/stores.h"

namespace xact::engine {

using storage::fsutil::utf8_to_wide;
using storage::fsutil::wide_to_utf8;

inline constexpr int MinimumRescueBlockSize = 4096;
inline constexpr int MaximumIoBufferSize = 256 * 1024 * 1024;

// Win32 error codes the retry policy reasons about (mirrors the VB constants).
inline constexpr DWORD ErrFileNotFound = 2;
inline constexpr DWORD ErrPathNotFound = 3;
inline constexpr DWORD ErrAccessDenied = 5;
inline constexpr DWORD ErrNotReady = 21;
inline constexpr DWORD ErrSharingViolation = 32;
inline constexpr DWORD ErrLockViolation = 33;
inline constexpr DWORD ErrHandleDiskFull = 39;
inline constexpr DWORD ErrBadNetPath = 53;
inline constexpr DWORD ErrNetNameDeleted = 64;
inline constexpr DWORD ErrBadNetName = 67;
inline constexpr DWORD ErrDiskFull = 112;
inline constexpr DWORD ErrDeviceNotConnected = 1167;

// ---------------------------------------------------------------------------
// Exceptions
// ---------------------------------------------------------------------------

// Category tags matching the .NET exception types the classifiers test for.
enum class IoErrorKind : std::uint8_t {
    General,          // IOException
    Timeout,          // TimeoutException
    DriveNotFound,    // DriveNotFoundException
    DirectoryNotFound,
    FileNotFound,
    UnauthorizedAccess,
    NotSupported,
    InvalidArgument
};

class IoError : public std::runtime_error {
public:
    IoError(std::string message, DWORD win32_code = 0, IoErrorKind error_kind = IoErrorKind::General)
        : std::runtime_error(std::move(message)), win32(win32_code), kind(error_kind) {}

    DWORD win32;
    IoErrorKind kind;

    static IoError from_win32(std::string message, DWORD code) {
        IoErrorKind kind = IoErrorKind::General;
        switch (code) {
            case ErrFileNotFound: kind = IoErrorKind::FileNotFound; break;
            case ErrPathNotFound: kind = IoErrorKind::DirectoryNotFound; break;
            case ErrAccessDenied: kind = IoErrorKind::UnauthorizedAccess; break;
            default: break;
        }
        return IoError(message + " (Win32 " + std::to_string(code) + ")", code, kind);
    }
};

struct OperationCanceled {
    bool user_requested = true; // false = per-file/operation timeout deadline
};

class SourceMutationSkipped : public std::runtime_error {
public:
    explicit SourceMutationSkipped(std::string message) : std::runtime_error(std::move(message)) {}
};

class FragileReadSkip : public std::runtime_error {
public:
    explicit FragileReadSkip(std::string message) : std::runtime_error(std::move(message)) {}
};

// ---------------------------------------------------------------------------
// Cancellation + pause control
// ---------------------------------------------------------------------------

// Cancellation flag + optional absolute deadline (per-file timeout). Matches
// the linked CancellationTokenSource pattern of the .NET worker: a deadline
// hit raises OperationCanceled{user=false}, a flag hit {user=true}.
struct CancelContext {
    const std::atomic<bool>* cancel_flag = nullptr;
    ULONGLONG deadline_tick = 0; // GetTickCount64-based; 0 = none

    bool is_cancelled() const noexcept {
        return cancel_flag != nullptr && cancel_flag->load(std::memory_order_relaxed);
    }

    bool deadline_passed() const noexcept {
        return deadline_tick != 0 && GetTickCount64() >= deadline_tick;
    }

    void throw_if_cancelled() const {
        if (is_cancelled()) throw OperationCanceled{true};
        if (deadline_passed()) throw OperationCanceled{false};
    }

    CancelContext with_deadline(std::int64_t timeout_ms) const {
        CancelContext child = *this;
        if (timeout_ms > 0) {
            ULONGLONG candidate = GetTickCount64() + static_cast<ULONGLONG>(timeout_ms);
            if (child.deadline_tick == 0 || candidate < child.deadline_tick) {
                child.deadline_tick = candidate;
            }
        }
        return child;
    }

    // Sleeps up to delay_ms, waking early on cancellation (throws) like
    // Task.Delay(token).
    void delay(std::int64_t delay_ms) const {
        ULONGLONG until = GetTickCount64() + static_cast<ULONGLONG>(delay_ms > 0 ? delay_ms : 0);
        while (true) {
            throw_if_cancelled();
            ULONGLONG now = GetTickCount64();
            if (now >= until) return;
            DWORD slice = static_cast<DWORD>(std::min<ULONGLONG>(100, until - now));
            Sleep(slice);
        }
    }
};

class ExecutionControl {
public:
    bool is_paused() const noexcept { return paused_.load(std::memory_order_acquire); }
    void pause() noexcept { paused_.store(true, std::memory_order_release); }
    void resume() noexcept { paused_.store(false, std::memory_order_release); }

    void wait_if_paused(const CancelContext& cancel) const {
        while (is_paused()) {
            cancel.throw_if_cancelled();
            Sleep(100);
        }
    }

private:
    std::atomic<bool> paused_{false};
};

// ---------------------------------------------------------------------------
// File metadata helpers
// ---------------------------------------------------------------------------

struct ExistingFileMetadata {
    std::int64_t length = 0;
    std::int64_t last_write_utc_ticks = 0; // .NET DateTime UTC ticks
};

inline std::int64_t filetime_to_ticks(DWORD low, DWORD high) {
    std::uint64_t filetime = (static_cast<std::uint64_t>(high) << 32) | low;
    if (filetime == 0) return 0; // DateTime.MinValue
    return static_cast<std::int64_t>(filetime) + time::FileTimeEpochTicks;
}

inline FILETIME ticks_to_filetime(std::int64_t ticks) {
    std::int64_t filetime = ticks - time::FileTimeEpochTicks;
    if (filetime < 0) filetime = 0;
    FILETIME result;
    result.dwLowDateTime = static_cast<DWORD>(filetime & 0xFFFFFFFF);
    result.dwHighDateTime = static_cast<DWORD>(static_cast<std::uint64_t>(filetime) >> 32);
    return result;
}

inline bool try_get_existing_file_metadata(const std::wstring& path, ExistingFileMetadata& metadata) {
    WIN32_FILE_ATTRIBUTE_DATA data{};
    if (!GetFileAttributesExW(path.c_str(), GetFileExInfoStandard, &data)) return false;
    if ((data.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) != 0) return false;
    metadata.length = static_cast<std::int64_t>(
        (static_cast<std::uint64_t>(data.nFileSizeHigh) << 32) | data.nFileSizeLow);
    metadata.last_write_utc_ticks =
        filetime_to_ticks(data.ftLastWriteTime.dwLowDateTime, data.ftLastWriteTime.dwHighDateTime);
    return true;
}

inline std::int64_t get_existing_file_length(const std::wstring& path) {
    ExistingFileMetadata metadata;
    if (!try_get_existing_file_metadata(path, metadata)) return 0;
    return metadata.length;
}

inline bool directory_exists(const std::wstring& path) {
    DWORD attributes = GetFileAttributesW(path.c_str());
    return attributes != INVALID_FILE_ATTRIBUTES && (attributes & FILE_ATTRIBUTE_DIRECTORY) != 0;
}

inline bool set_last_write_time_utc(const std::wstring& path, std::int64_t ticks) {
    HANDLE handle = CreateFileW(path.c_str(), FILE_WRITE_ATTRIBUTES,
                                FILE_SHARE_READ | FILE_SHARE_WRITE, nullptr,
                                OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, nullptr);
    if (handle == INVALID_HANDLE_VALUE) return false;
    FILETIME write_time = ticks_to_filetime(ticks);
    BOOL ok = SetFileTime(handle, nullptr, nullptr, &write_time);
    CloseHandle(handle);
    return ok != FALSE;
}

// Matches the VB FormatBytes helper ("{value} B", then "0.##" KB/MB/GB/TB/PB).
inline std::string format_bytes(std::int64_t value) {
    if (value < 1024) return std::to_string(value) + " B";
    static const char* units[] = {"KB", "MB", "GB", "TB", "PB"};
    double size = static_cast<double>(value);
    int unit_index = -1;
    do {
        size /= 1024.0;
        ++unit_index;
    } while (size >= 1024.0 && unit_index < 4);

    // "0.##": up to two decimals, trailing zeros trimmed (1.50 -> "1.5").
    long long scaled = static_cast<long long>(size * 100.0 + 0.5);
    long long whole = scaled / 100;
    int fraction = static_cast<int>(scaled % 100);
    std::string text = std::to_string(whole);
    if (fraction != 0) {
        text.push_back('.');
        text.push_back(static_cast<char>('0' + fraction / 10));
        if (fraction % 10 != 0) {
            text.push_back(static_cast<char>('0' + fraction % 10));
        }
    }
    return text + " " + units[unit_index];
}

// ---------------------------------------------------------------------------
// Directory scanner (port of DirectoryScanner.vb)
// ---------------------------------------------------------------------------

struct SourceFileDescriptor {
    std::string relative_path;      // UTF-8, backslash separators
    std::wstring full_path;
    std::int64_t length = 0;
    std::int64_t last_write_utc_ticks = 0;
};

struct SourceScanResult {
    std::vector<SourceFileDescriptor> files;
    std::vector<std::string> directories; // relative, sorted
};

namespace detail {

inline std::wstring trim_trailing_separators(std::wstring value) {
    while (!value.empty() && (value.back() == L'\\' || value.back() == L'/')) value.pop_back();
    return value;
}

inline std::string normalize_relative_path(std::string text) {
    // Trim ASCII whitespace, convert '/' to '\\', trim separators.
    auto is_space = [](char c) { return c == ' ' || c == '\t' || c == '\r' || c == '\n'; };
    std::size_t begin = 0, end = text.size();
    while (begin < end && is_space(text[begin])) ++begin;
    while (end > begin && is_space(text[end - 1])) --end;
    text = text.substr(begin, end - begin);
    if (text.empty() || text == ".") return std::string();
    for (char& c : text) {
        if (c == '/') c = '\\';
    }
    begin = 0; end = text.size();
    while (begin < end && text[begin] == '\\') ++begin;
    while (end > begin && text[end - 1] == '\\') --end;
    return text.substr(begin, end - begin);
}

inline bool ordinal_ignore_case_less(const std::wstring& a, const std::wstring& b) {
    int result = CompareStringOrdinal(a.c_str(), static_cast<int>(a.size()),
                                      b.c_str(), static_cast<int>(b.size()), TRUE);
    return result == CSTR_LESS_THAN;
}

inline bool starts_with_ignore_case(std::string_view text, std::string_view prefix) {
    if (text.size() < prefix.size()) return false;
    return models::detail::equals_ignore_case(text.substr(0, prefix.size()), prefix);
}

inline bool is_reparse_point(const std::wstring& path) {
    DWORD attributes = GetFileAttributesW(path.c_str());
    return attributes != INVALID_FILE_ATTRIBUTES && (attributes & FILE_ATTRIBUTE_REPARSE_POINT) != 0;
}

inline std::wstring resolve_final_path(const std::wstring& path) {
    HANDLE handle = CreateFileW(path.c_str(), 0,
                                FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE, nullptr,
                                OPEN_EXISTING, FILE_FLAG_BACKUP_SEMANTICS, nullptr);
    if (handle == INVALID_HANDLE_VALUE) return std::wstring();
    wchar_t buffer[1024];
    DWORD length = GetFinalPathNameByHandleW(handle, buffer, 1024, FILE_NAME_NORMALIZED);
    CloseHandle(handle);
    if (length == 0 || length >= 1024) return std::wstring();
    std::wstring result(buffer, length);
    if (result.rfind(L"\\\\?\\", 0) == 0) result = result.substr(4);
    return result;
}

} // namespace detail

// Selection filter for Explorer-driven subsets (port of SelectionFilter).
class SelectionFilter {
public:
    static SelectionFilter create(const std::wstring& root_full,
                                  const std::vector<std::string>& requested,
                                  const std::function<void(const std::string&)>& log) {
        SelectionFilter filter;
        if (requested.empty()) return filter;

        std::string root_utf8 = wide_to_utf8(detail::trim_trailing_separators(root_full));

        for (const auto& raw : requested) {
            std::string candidate = detail::normalize_relative_path(raw);
            if (candidate.empty()) return SelectionFilter(); // select-all sentinel

            std::string relative = candidate;
            bool rooted = candidate.size() >= 2 &&
                          ((candidate[1] == ':') || (candidate[0] == '\\' && candidate[1] == '\\'));
            if (rooted) {
                std::string prefix = root_utf8 + "\\";
                if (models::detail::equals_ignore_case(candidate, root_utf8)) return SelectionFilter();
                if (!detail::starts_with_ignore_case(candidate, prefix)) {
                    if (log) log("Ignored selected path outside source root: " + candidate);
                    continue;
                }
                relative = detail::normalize_relative_path(candidate.substr(prefix.size()));
                if (relative.empty()) return SelectionFilter();
            }

            // Reject traversal escapes.
            if (relative.find("..") != std::string::npos) {
                if (log) log("Ignored selected relative path outside source root: " + relative);
                continue;
            }
            bool exists = false;
            for (const auto& entry : filter.entries_) {
                if (models::detail::equals_ignore_case(entry, relative)) { exists = true; break; }
            }
            if (!exists) filter.entries_.push_back(relative);
        }
        return filter;
    }

    bool should_include_directory(const std::string& relative_directory) const {
        return should_traverse_directory(relative_directory);
    }

    bool should_traverse_directory(const std::string& relative_directory) const {
        if (entries_.empty()) return true;
        std::string normalized = detail::normalize_relative_path(relative_directory);
        if (normalized.empty()) return true;

        std::string directory_prefix = normalized + "\\";
        for (const auto& entry : entries_) {
            if (models::detail::equals_ignore_case(entry, normalized)) return true;
            if (detail::starts_with_ignore_case(entry, directory_prefix)) return true;
            std::string entry_prefix = entry + "\\";
            if (detail::starts_with_ignore_case(normalized, entry_prefix)) return true;
        }
        return false;
    }

    bool should_include_file(const std::string& relative_file) const {
        if (entries_.empty()) return true;
        std::string normalized = detail::normalize_relative_path(relative_file);
        if (normalized.empty()) return false;

        for (const auto& entry : entries_) {
            if (models::detail::equals_ignore_case(entry, normalized)) return true;
            std::string entry_prefix = entry + "\\";
            if (detail::starts_with_ignore_case(normalized, entry_prefix)) return true;
        }
        return false;
    }

private:
    std::vector<std::string> entries_;
};

// Recursive Win32 scan matching DirectoryScanner.ScanSource: iterative stack,
// symlink policy, visited-identity cycle guard, selection filter, sorted output.
inline SourceScanResult scan_source(
    const std::wstring& root_path,
    const std::vector<std::string>& selected_relative_paths,
    models::SymlinkHandlingMode symlink_handling,
    bool copy_empty_directories,
    const std::function<void(const std::string&)>& log,
    const std::function<bool()>& should_cancel = nullptr) {

    std::wstring normalized_root = storage::fsutil::get_full_path(root_path);
    SourceScanResult result;
    std::vector<std::string> discovered_dirs;
    // Ordinal-ignore-case dedup sets (the .NET DirectoryScanner uses
    // HashSet(OrdinalIgnoreCase)); a linear scan here is O(n^2) and turns a
    // whole-drive enumeration into an apparent hang.
    std::unordered_set<std::wstring> visited_identities;
    std::unordered_set<std::string> discovered_file_keys;
    std::size_t enumerated_files = 0;
    ULONGLONG last_heartbeat_tick = GetTickCount64();

    // Enumeration can dominate a large drive; poll cancellation and emit a
    // periodic heartbeat so Cancel is responsive and the UI shows activity.
    auto pump = [&](bool force_heartbeat) {
        if (should_cancel && should_cancel()) throw OperationCanceled{true};
        ULONGLONG now = GetTickCount64();
        if (log && (force_heartbeat || now - last_heartbeat_tick >= 1000)) {
            last_heartbeat_tick = now;
            log("Enumerating source: " + std::to_string(enumerated_files) + " file(s) so far...");
        }
    };

    auto directory_identity = [&](const std::wstring& path) -> std::wstring {
        std::wstring normalized =
            storage::fsutil::to_upper_invariant(detail::trim_trailing_separators(storage::fsutil::get_full_path(path)));
        if (symlink_handling != models::SymlinkHandlingMode::Follow || !detail::is_reparse_point(path)) {
            return normalized;
        }
        std::wstring target = detail::resolve_final_path(path);
        if (!target.empty()) {
            return storage::fsutil::to_upper_invariant(detail::trim_trailing_separators(target));
        }
        if (log) log("Unable to resolve directory symbolic-link target for " + wide_to_utf8(path));
        return normalized;
    };

    SelectionFilter filter = SelectionFilter::create(normalized_root, selected_relative_paths, log);
    visited_identities.insert(directory_identity(normalized_root));

    std::wstring root_prefix = detail::trim_trailing_separators(normalized_root) + L"\\";
    auto relative_of = [&](const std::wstring& full) -> std::string {
        if (full.size() >= root_prefix.size()) {
            return detail::normalize_relative_path(wide_to_utf8(full.substr(root_prefix.size())));
        }
        return std::string();
    };

    std::vector<std::wstring> pending;
    pending.push_back(normalized_root);

    while (!pending.empty()) {
        pump(false);
        std::wstring current = pending.back();
        pending.pop_back();

        std::string current_relative = relative_of(current);
        if (copy_empty_directories && !current_relative.empty() &&
            filter.should_include_directory(current_relative)) {
            discovered_dirs.push_back(current_relative);
        }

        std::wstring search = detail::trim_trailing_separators(current) + L"\\*";
        WIN32_FIND_DATAW find_data{};
        HANDLE find = FindFirstFileExW(search.c_str(), FindExInfoBasic, &find_data,
                                       FindExSearchNameMatch, nullptr, FIND_FIRST_EX_LARGE_FETCH);
        if (find == INVALID_HANDLE_VALUE) {
            DWORD error = GetLastError();
            if (error == ERROR_FILE_NOT_FOUND || error == ERROR_PATH_NOT_FOUND || error == ERROR_NO_MORE_FILES) {
                continue;
            }
            if (log) {
                std::string dir_text = wide_to_utf8(current);
                log("Directory skipped: " + dir_text + " (Native enumeration failed (" + std::to_string(error) + ").)");
                log("Files skipped in: " + dir_text + " (Native enumeration failed (" + std::to_string(error) + ").)");
            }
            continue;
        }

        do {
            std::wstring name = find_data.cFileName;
            if (name.empty() || name == L"." || name == L"..") continue;

            std::wstring full_path = detail::trim_trailing_separators(current) + L"\\" + name;
            bool is_directory = (find_data.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) != 0;
            bool is_reparse = (find_data.dwFileAttributes & FILE_ATTRIBUTE_REPARSE_POINT) != 0;

            if (is_directory) {
                std::string relative_dir = relative_of(full_path);
                if (!filter.should_traverse_directory(relative_dir)) continue;
                if (is_reparse && symlink_handling == models::SymlinkHandlingMode::Skip) {
                    if (log) log("Skipping symbolic-link directory: " + wide_to_utf8(full_path));
                    continue;
                }
                std::wstring identity = directory_identity(full_path);
                if (!visited_identities.insert(identity).second) {
                    if (log) log("Skipping already visited directory target: " + wide_to_utf8(full_path));
                    continue;
                }
                pending.push_back(full_path);
                continue;
            }

            std::string relative_file = relative_of(full_path);
            if (!filter.should_include_file(relative_file)) continue;
            if (is_reparse && symlink_handling == models::SymlinkHandlingMode::Skip) {
                if (log) log("Skipping symbolic-link file: " + wide_to_utf8(full_path));
                continue;
            }

            std::string dedup_key = storage::fsutil::wide_to_utf8(
                storage::fsutil::to_upper_invariant(utf8_to_wide(relative_file)));
            if (!discovered_file_keys.insert(dedup_key).second) {
                if (log) log("Skipping duplicate file entry: " + relative_file);
                continue;
            }

            ++enumerated_files;
            pump(false);

            SourceFileDescriptor descriptor;
            descriptor.relative_path = relative_file;
            descriptor.full_path = full_path;
            std::uint64_t raw_length = (static_cast<std::uint64_t>(find_data.nFileSizeHigh) << 32) |
                                       find_data.nFileSizeLow;
            descriptor.length = raw_length > static_cast<std::uint64_t>(INT64_MAX)
                                    ? INT64_MAX
                                    : static_cast<std::int64_t>(raw_length);
            descriptor.last_write_utc_ticks = filetime_to_ticks(
                find_data.ftLastWriteTime.dwLowDateTime, find_data.ftLastWriteTime.dwHighDateTime);
            result.files.push_back(std::move(descriptor));
        } while (FindNextFileW(find, &find_data));
        FindClose(find);
    }

    // Sort by relative path, ordinal-ignore-case (same order the journal is
    // populated and files are processed in on the .NET side).
    std::stable_sort(result.files.begin(), result.files.end(),
                     [](const SourceFileDescriptor& a, const SourceFileDescriptor& b) {
                         return detail::ordinal_ignore_case_less(
                             utf8_to_wide(a.relative_path), utf8_to_wide(b.relative_path));
                     });

    std::stable_sort(discovered_dirs.begin(), discovered_dirs.end(),
                     [](const std::string& a, const std::string& b) {
                         return detail::ordinal_ignore_case_less(utf8_to_wide(a), utf8_to_wide(b));
                     });
    discovered_dirs.erase(std::unique(discovered_dirs.begin(), discovered_dirs.end(),
                                      [](const std::string& a, const std::string& b) {
                                          return models::detail::equals_ignore_case(a, b);
                                      }),
                          discovered_dirs.end());
    result.directories = std::move(discovered_dirs);
    return result;
}

// ---------------------------------------------------------------------------
// Fault injector (port of DevFaultInjector; format:
//   XACTCOPY_DEV_FAULT_RULES = "op,offset,length,trigger,kind[,param];..."
//   op: read|write; trigger: once|always|percent|every; kind: io|timeout|offline)
// ---------------------------------------------------------------------------

class FaultInjector {
public:
    static std::optional<FaultInjector> try_create_from_environment() {
        wchar_t raw[4096];
        DWORD length = GetEnvironmentVariableW(L"XACTCOPY_DEV_FAULT_RULES", raw, 4096);
        if (length == 0 || length >= 4096) return std::nullopt;

        FaultInjector injector;
        std::string rules_text = wide_to_utf8(std::wstring(raw, length));
        std::size_t start = 0;
        while (start <= rules_text.size()) {
            std::size_t end = rules_text.find(';', start);
            std::string token = rules_text.substr(
                start, end == std::string::npos ? std::string::npos : end - start);
            if (Rule rule; Rule::try_parse(token, rule)) {
                injector.rules_.push_back(rule);
            }
            if (end == std::string::npos) break;
            start = end + 1;
        }
        if (injector.rules_.empty()) return std::nullopt;

        int seed = 1337;
        wchar_t raw_seed[64];
        DWORD seed_length = GetEnvironmentVariableW(L"XACTCOPY_DEV_FAULT_SEED", raw_seed, 64);
        if (seed_length > 0 && seed_length < 64) {
            std::string seed_text = wide_to_utf8(std::wstring(raw_seed, seed_length));
            std::from_chars(seed_text.data(), seed_text.data() + seed_text.size(), seed);
        }
        injector.random_.seed(static_cast<unsigned int>(seed));
        return injector;
    }

    std::string description() const {
        return std::to_string(rules_.size()) + " rule(s) via XACTCOPY_DEV_FAULT_RULES";
    }

    // Returns nullopt (no fault) or the IoError to throw.
    std::optional<IoError> create_read_fault(const std::string& relative_path,
                                             std::int64_t offset, std::int32_t length) {
        return create_fault(false, relative_path, offset, length);
    }

    std::optional<IoError> create_write_fault(const std::string& relative_path,
                                              std::int64_t offset, std::int32_t length) {
        return create_fault(true, relative_path, offset, length);
    }

private:
    struct Rule {
        bool is_write = false;
        std::int64_t offset = 0;
        std::int64_t length = 0; // 0 = whole file
        int trigger = 0;         // 0=once 1=always 2=percent 3=every
        int error_kind = 0;      // 0=io 1=timeout 2=offline
        int parameter = 0;
        int match_count = 0;
        int fire_count = 0;

        static bool try_parse(const std::string& text, Rule& rule) {
            std::vector<std::string> parts;
            std::size_t start = 0;
            while (start <= text.size()) {
                std::size_t end = text.find(',', start);
                std::string part = text.substr(
                    start, end == std::string::npos ? std::string::npos : end - start);
                // trim
                std::size_t b = 0, e = part.size();
                while (b < e && (part[b] == ' ' || part[b] == '\t')) ++b;
                while (e > b && (part[e - 1] == ' ' || part[e - 1] == '\t')) --e;
                parts.push_back(part.substr(b, e - b));
                if (end == std::string::npos) break;
                start = end + 1;
            }
            if (parts.size() < 5) return false;

            auto lower = [](std::string value) {
                for (char& c : value) {
                    if (c >= 'A' && c <= 'Z') c = static_cast<char>(c - 'A' + 'a');
                }
                return value;
            };

            std::string op = lower(parts[0]);
            if (op == "read") rule.is_write = false;
            else if (op == "write") rule.is_write = true;
            else return false;

            auto parse_i64 = [](const std::string& value, std::int64_t& out) {
                auto result = std::from_chars(value.data(), value.data() + value.size(), out);
                return result.ec == std::errc() && result.ptr == value.data() + value.size();
            };
            if (!parse_i64(parts[1], rule.offset)) return false;
            if (!parse_i64(parts[2], rule.length)) return false;
            rule.offset = rule.offset < 0 ? 0 : rule.offset;
            rule.length = rule.length < 0 ? 0 : rule.length;

            std::string trigger = lower(parts[3]);
            if (trigger == "once") rule.trigger = 0;
            else if (trigger == "always") rule.trigger = 1;
            else if (trigger == "percent") rule.trigger = 2;
            else if (trigger == "every") rule.trigger = 3;
            else return false;

            std::string kind = lower(parts[4]);
            if (kind == "io") rule.error_kind = 0;
            else if (kind == "timeout") rule.error_kind = 1;
            else if (kind == "offline") rule.error_kind = 2;
            else return false;

            if (parts.size() >= 6) {
                std::int64_t parameter = 0;
                parse_i64(parts[5], parameter);
                rule.parameter = static_cast<int>(parameter);
            }
            if (rule.trigger == 2) rule.parameter = std::clamp(rule.parameter, 1, 100);
            else if (rule.trigger == 3) rule.parameter = std::max(1, rule.parameter);
            else rule.parameter = std::max(0, rule.parameter);
            return true;
        }

        bool overlaps(std::int64_t request_offset, std::int32_t request_length) const {
            std::int64_t request_start = request_offset < 0 ? 0 : request_offset;
            std::int64_t request_end = request_start + (request_length < 0 ? 0 : request_length);
            if (request_end <= request_start) return false;
            std::int64_t rule_end = length <= 0 ? INT64_MAX : offset + length;
            return request_start < rule_end && request_end > offset;
        }

        bool should_trigger(std::mt19937& random) {
            switch (trigger) {
                case 0: return fire_count == 0;
                case 1: return true;
                case 2: return static_cast<int>(random() % 100) < parameter;
                case 3: return (match_count % std::max(1, parameter)) == 0;
                default: return false;
            }
        }
    };

    std::vector<Rule> rules_;
    std::mt19937 random_{1337};
    // Heap-held so FaultInjector stays movable inside std::optional.
    std::shared_ptr<std::mutex> lock_ = std::make_shared<std::mutex>();

    std::optional<IoError> create_fault(bool is_write, const std::string& relative_path,
                                        std::int64_t offset, std::int32_t length) {
        if (length <= 0) return std::nullopt;
        std::lock_guard<std::mutex> guard(*lock_);
        for (auto& rule : rules_) {
            if (rule.is_write != is_write) continue;
            if (!rule.overlaps(offset, length)) continue;
            rule.match_count++;
            if (!rule.should_trigger(random_)) continue;
            rule.fire_count++;

            std::string label = relative_path.empty() ? "?" : relative_path;
            std::string message = "[DevFault] Injected " + std::string(is_write ? "write" : "read") +
                                  " fault on " + label + " at " + format_bytes(offset) +
                                  " (" + std::to_string(length) + " B).";
            switch (rule.error_kind) {
                case 1: return IoError(message, 0, IoErrorKind::Timeout);
                case 2: return IoError(message, 0, IoErrorKind::DriveNotFound);
                default: return IoError(message, 0, IoErrorKind::General);
            }
        }
        return std::nullopt;
    }
};

// ---------------------------------------------------------------------------
// Timed positional I/O + transfer session
// ---------------------------------------------------------------------------

// One overlapped positional operation with an operation timeout and
// cancellation polling. Mirrors RandomAccess.Read/WriteAsync + CancelAfter +
// CancelIoEx from the .NET worker.
namespace timed_io {

enum class Outcome { Success, Timeout, EndOfFile };

inline Outcome wait_for_completion(HANDLE file, OVERLAPPED& overlapped, DWORD& transferred,
                                   std::int64_t timeout_ms, const CancelContext& cancel) {
    ULONGLONG deadline = GetTickCount64() + static_cast<ULONGLONG>(timeout_ms > 0 ? timeout_ms : 0);
    while (true) {
        if (cancel.is_cancelled() || cancel.deadline_passed()) {
            CancelIoEx(file, &overlapped);
            GetOverlappedResult(file, &overlapped, &transferred, TRUE);
            cancel.throw_if_cancelled();
        }
        ULONGLONG now = GetTickCount64();
        if (timeout_ms > 0 && now >= deadline) {
            CancelIoEx(file, &overlapped);
            GetOverlappedResult(file, &overlapped, &transferred, TRUE);
            return Outcome::Timeout;
        }
        DWORD slice = 100;
        if (timeout_ms > 0) {
            slice = static_cast<DWORD>(std::min<ULONGLONG>(slice, deadline - now));
        }
        DWORD wait_result = WaitForSingleObject(overlapped.hEvent, slice);
        if (wait_result == WAIT_OBJECT_0) {
            if (!GetOverlappedResult(file, &overlapped, &transferred, FALSE)) {
                DWORD error = GetLastError();
                if (error == ERROR_HANDLE_EOF || error == ERROR_BROKEN_PIPE) {
                    transferred = 0;
                    return Outcome::EndOfFile;
                }
                throw IoError::from_win32("Positional I/O failed.", error);
            }
            return Outcome::Success;
        }
    }
}

// Reads exactly `count` bytes at `offset` unless EOF; Timeout outcome throws
// IoError{kind=Timeout}. Returns bytes actually read (may be < count at EOF).
inline std::int32_t read_at(HANDLE file, std::int64_t offset, unsigned char* buffer,
                            std::int32_t count, std::int64_t timeout_ms,
                            const CancelContext& cancel) {
    std::int32_t total = 0;
    while (total < count) {
        OVERLAPPED overlapped{};
        std::int64_t position = offset + total;
        overlapped.Offset = static_cast<DWORD>(position & 0xFFFFFFFF);
        overlapped.OffsetHigh = static_cast<DWORD>(static_cast<std::uint64_t>(position) >> 32);
        overlapped.hEvent = CreateEventW(nullptr, TRUE, FALSE, nullptr);
        if (overlapped.hEvent == nullptr) {
            throw IoError::from_win32("CreateEvent failed for read.", GetLastError());
        }

        DWORD transferred = 0;
        BOOL ok = ReadFile(file, buffer + total, static_cast<DWORD>(count - total), nullptr, &overlapped);
        DWORD error = ok ? ERROR_SUCCESS : GetLastError();
        Outcome outcome = Outcome::Success;
        try {
            if (!ok && error == ERROR_IO_PENDING) {
                outcome = wait_for_completion(file, overlapped, transferred, timeout_ms, cancel);
            } else if (!ok && error == ERROR_HANDLE_EOF) {
                transferred = 0;
                outcome = Outcome::EndOfFile;
            } else if (!ok) {
                throw IoError::from_win32("ReadFile failed.", error);
            } else {
                GetOverlappedResult(file, &overlapped, &transferred, TRUE);
            }
        } catch (...) {
            CloseHandle(overlapped.hEvent);
            throw;
        }
        CloseHandle(overlapped.hEvent);

        if (outcome == Outcome::Timeout) {
            throw IoError("Read timed out.", 0, IoErrorKind::Timeout);
        }
        if (transferred == 0) break; // EOF
        total += static_cast<std::int32_t>(transferred);
    }
    return total;
}

inline void write_at(HANDLE file, std::int64_t offset, const unsigned char* buffer,
                     std::int32_t count, std::int64_t timeout_ms,
                     const CancelContext& cancel) {
    std::int32_t total = 0;
    while (total < count) {
        OVERLAPPED overlapped{};
        std::int64_t position = offset + total;
        overlapped.Offset = static_cast<DWORD>(position & 0xFFFFFFFF);
        overlapped.OffsetHigh = static_cast<DWORD>(static_cast<std::uint64_t>(position) >> 32);
        overlapped.hEvent = CreateEventW(nullptr, TRUE, FALSE, nullptr);
        if (overlapped.hEvent == nullptr) {
            throw IoError::from_win32("CreateEvent failed for write.", GetLastError());
        }

        DWORD transferred = 0;
        BOOL ok = WriteFile(file, buffer + total, static_cast<DWORD>(count - total), nullptr, &overlapped);
        DWORD error = ok ? ERROR_SUCCESS : GetLastError();
        Outcome outcome = Outcome::Success;
        try {
            if (!ok && error == ERROR_IO_PENDING) {
                outcome = wait_for_completion(file, overlapped, transferred, timeout_ms, cancel);
            } else if (!ok) {
                throw IoError::from_win32("WriteFile failed.", error);
            } else {
                GetOverlappedResult(file, &overlapped, &transferred, TRUE);
            }
        } catch (...) {
            CloseHandle(overlapped.hEvent);
            throw;
        }
        CloseHandle(overlapped.hEvent);

        if (outcome == Outcome::Timeout) {
            throw IoError("Write timed out.", 0, IoErrorKind::Timeout);
        }
        if (transferred == 0) {
            throw IoError("Write made no progress.", ERROR_WRITE_FAULT, IoErrorKind::General);
        }
        total += static_cast<std::int32_t>(transferred);
    }
}

} // namespace timed_io

// Lazily opened source/destination handles for one file transfer, invalidated
// on errors so the next attempt reopens (port of FileTransferSession).
class FileTransferSession {
public:
    FileTransferSession(std::wstring source_path, std::wstring destination_path)
        : source_path_(std::move(source_path)), destination_path_(std::move(destination_path)) {}

    ~FileTransferSession() {
        invalidate_source();
        invalidate_destination();
    }

    FileTransferSession(const FileTransferSession&) = delete;
    FileTransferSession& operator=(const FileTransferSession&) = delete;

    HANDLE source_handle() {
        if (source_ == INVALID_HANDLE_VALUE) {
            source_ = CreateFileW(source_path_.c_str(), GENERIC_READ,
                                  FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE, nullptr,
                                  OPEN_EXISTING,
                                  FILE_ATTRIBUTE_NORMAL | FILE_FLAG_OVERLAPPED | FILE_FLAG_SEQUENTIAL_SCAN,
                                  nullptr);
            if (source_ == INVALID_HANDLE_VALUE) {
                throw IoError::from_win32("Unable to open source file.", GetLastError());
            }
        }
        return source_;
    }

    HANDLE destination_handle() {
        if (destination_ == INVALID_HANDLE_VALUE) {
            destination_ = CreateFileW(destination_path_.c_str(), GENERIC_WRITE,
                                       FILE_SHARE_READ, nullptr, OPEN_ALWAYS,
                                       FILE_ATTRIBUTE_NORMAL | FILE_FLAG_OVERLAPPED | FILE_FLAG_SEQUENTIAL_SCAN,
                                       nullptr);
            if (destination_ == INVALID_HANDLE_VALUE) {
                throw IoError::from_win32("Unable to open destination file.", GetLastError());
            }
        }
        return destination_;
    }

    HANDLE try_get_open_source() const noexcept { return source_; }
    HANDLE try_get_open_destination() const noexcept { return destination_; }

    void invalidate_source() noexcept {
        if (source_ != INVALID_HANDLE_VALUE) {
            CancelIoEx(source_, nullptr);
            CloseHandle(source_);
            source_ = INVALID_HANDLE_VALUE;
        }
    }

    void invalidate_destination() noexcept {
        if (destination_ != INVALID_HANDLE_VALUE) {
            CancelIoEx(destination_, nullptr);
            CloseHandle(destination_);
            destination_ = INVALID_HANDLE_VALUE;
        }
    }

private:
    std::wstring source_path_;
    std::wstring destination_path_;
    HANDLE source_ = INVALID_HANDLE_VALUE;
    HANDLE destination_ = INVALID_HANDLE_VALUE;
};

} // namespace xact::engine
