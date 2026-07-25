// -----------------------------------------------------------------------------
// File: cpp\src\storage\stores.h
// Purpose: C++ port of XactCopy.Storage: JobJournalStore and BadRangeMapStore
//          with the full hardening pipeline — atomic write-through snapshots,
//          rotating backups, mirror snapshots, HMAC-signed envelopes, and the
//          hash-chained journal ledger — file-compatible with the .NET stores.
// -----------------------------------------------------------------------------

#pragma once

#include <cstdint>
#include <mutex>
#include <optional>
#include <string>
#include <string_view>
#include <vector>

#ifndef NOMINMAX
#define NOMINMAX
#endif
#ifndef WIN32_LEAN_AND_MEAN
#define WIN32_LEAN_AND_MEAN
#endif
#include <windows.h>
#include <shlobj.h>

#include "../core/crypto.h"
#include "../core/json.h"
#include "storage_models.h"

namespace xact::storage {

// ---------------------------------------------------------------------------
// Path and file utilities
// ---------------------------------------------------------------------------

namespace fsutil {

inline std::wstring utf8_to_wide(std::string_view text) {
    if (text.empty()) return std::wstring();
    int needed = MultiByteToWideChar(CP_UTF8, 0, text.data(), static_cast<int>(text.size()), nullptr, 0);
    std::wstring out(static_cast<std::size_t>(needed), L'\0');
    MultiByteToWideChar(CP_UTF8, 0, text.data(), static_cast<int>(text.size()), out.data(), needed);
    return out;
}

inline std::string wide_to_utf8(std::wstring_view text) {
    if (text.empty()) return std::string();
    int needed = WideCharToMultiByte(CP_UTF8, 0, text.data(), static_cast<int>(text.size()),
                                     nullptr, 0, nullptr, nullptr);
    std::string out(static_cast<std::size_t>(needed), '\0');
    WideCharToMultiByte(CP_UTF8, 0, text.data(), static_cast<int>(text.size()),
                        out.data(), needed, nullptr, nullptr);
    return out;
}

inline bool file_exists(const std::wstring& path) {
    DWORD attributes = GetFileAttributesW(path.c_str());
    return attributes != INVALID_FILE_ATTRIBUTES && (attributes & FILE_ATTRIBUTE_DIRECTORY) == 0;
}

inline void create_directories(const std::wstring& path) {
    if (path.empty()) return;
    // Walk the path creating each component; ignore already-exists.
    std::wstring current;
    current.reserve(path.size());
    for (std::size_t i = 0; i <= path.size(); ++i) {
        if (i == path.size() || path[i] == L'\\' || path[i] == L'/') {
            if (!current.empty() && current.back() != L':') {
                CreateDirectoryW(current.c_str(), nullptr);
            }
            if (i < path.size()) current.push_back(L'\\');
            continue;
        }
        current.push_back(path[i]);
    }
}

inline std::wstring get_directory_name(const std::wstring& path) {
    std::size_t pos = path.find_last_of(L"\\/");
    if (pos == std::wstring::npos) return std::wstring();
    return path.substr(0, pos);
}

inline std::wstring get_file_name(const std::wstring& path) {
    std::size_t pos = path.find_last_of(L"\\/");
    if (pos == std::wstring::npos) return path;
    return path.substr(pos + 1);
}

inline void split_extension(const std::wstring& file_name, std::wstring& base, std::wstring& extension) {
    std::size_t pos = file_name.find_last_of(L'.');
    if (pos == std::wstring::npos || pos == 0) {
        base = file_name;
        extension.clear();
        return;
    }
    base = file_name.substr(0, pos);
    extension = file_name.substr(pos);
}

inline std::wstring get_full_path(const std::wstring& path) {
    if (path.empty()) return path;
    DWORD needed = GetFullPathNameW(path.c_str(), 0, nullptr, nullptr);
    if (needed == 0) return path;
    std::wstring out(static_cast<std::size_t>(needed), L'\0');
    DWORD written = GetFullPathNameW(path.c_str(), needed, out.data(), nullptr);
    if (written == 0 || written >= needed) return path;
    out.resize(written);
    return out;
}

// Full Unicode invariant uppercase, matching String.ToUpperInvariant.
inline std::wstring to_upper_invariant(const std::wstring& text) {
    if (text.empty()) return text;
    int needed = LCMapStringEx(LOCALE_NAME_INVARIANT, LCMAP_UPPERCASE,
                               text.c_str(), static_cast<int>(text.size()),
                               nullptr, 0, nullptr, nullptr, 0);
    if (needed <= 0) return text;
    std::wstring out(static_cast<std::size_t>(needed), L'\0');
    LCMapStringEx(LOCALE_NAME_INVARIANT, LCMAP_UPPERCASE,
                  text.c_str(), static_cast<int>(text.size()),
                  out.data(), needed, nullptr, nullptr, 0);
    return out;
}

inline std::wstring trim(const std::wstring& text) {
    std::size_t begin = 0;
    std::size_t end = text.size();
    auto is_space = [](wchar_t c) { return c == L' ' || c == L'\t' || c == L'\r' || c == L'\n'; };
    while (begin < end && is_space(text[begin])) ++begin;
    while (end > begin && is_space(text[end - 1])) --end;
    return text.substr(begin, end - begin);
}

// GetFullPath(path).Trim().ToUpperInvariant() — the store's canonical form for
// fingerprints, mirror names, and cache keys.
inline std::wstring normalize_path(const std::wstring& path) {
    std::wstring trimmed = trim(path);
    if (trimmed.empty()) return std::wstring();
    return to_upper_invariant(trim(get_full_path(trimmed)));
}

inline std::wstring local_app_data() {
    PWSTR folder = nullptr;
    std::wstring result;
    if (SHGetKnownFolderPath(FOLDERID_LocalAppData, 0, nullptr, &folder) == S_OK && folder != nullptr) {
        result = folder;
    }
    if (folder != nullptr) CoTaskMemFree(folder);
    return result;
}

inline std::optional<std::vector<unsigned char>> read_all_bytes(const std::wstring& path) {
    HANDLE handle = CreateFileW(path.c_str(), GENERIC_READ, FILE_SHARE_READ | FILE_SHARE_WRITE,
                                nullptr, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, nullptr);
    if (handle == INVALID_HANDLE_VALUE) return std::nullopt;

    LARGE_INTEGER size{};
    if (!GetFileSizeEx(handle, &size) || size.QuadPart < 0 || size.QuadPart > (1LL << 31)) {
        CloseHandle(handle);
        return std::nullopt;
    }

    std::vector<unsigned char> data(static_cast<std::size_t>(size.QuadPart));
    std::size_t offset = 0;
    while (offset < data.size()) {
        DWORD chunk = 0;
        DWORD request = static_cast<DWORD>(std::min<std::size_t>(data.size() - offset, 1 << 20));
        if (!ReadFile(handle, data.data() + offset, request, &chunk, nullptr) || chunk == 0) {
            CloseHandle(handle);
            return std::nullopt;
        }
        offset += chunk;
    }
    CloseHandle(handle);
    return data;
}

inline bool write_file_raw(const std::wstring& path, const unsigned char* data, std::size_t size,
                           bool write_through, bool fail_if_exists) {
    HANDLE handle = CreateFileW(path.c_str(), GENERIC_WRITE, 0, nullptr,
                                fail_if_exists ? CREATE_NEW : CREATE_ALWAYS,
                                FILE_ATTRIBUTE_NORMAL | (write_through ? FILE_FLAG_WRITE_THROUGH : 0),
                                nullptr);
    if (handle == INVALID_HANDLE_VALUE) return false;

    std::size_t offset = 0;
    while (offset < size) {
        DWORD written = 0;
        DWORD request = static_cast<DWORD>(std::min<std::size_t>(size - offset, 1 << 20));
        if (!WriteFile(handle, data + offset, request, &written, nullptr) || written == 0) {
            CloseHandle(handle);
            DeleteFileW(path.c_str());
            return false;
        }
        offset += written;
    }
    if (write_through) FlushFileBuffers(handle);
    CloseHandle(handle);
    return true;
}

inline std::wstring random_temp_suffix() {
    std::vector<unsigned char> raw(16);
    crypto::fill_random(raw);
    static const wchar_t hex[] = L"0123456789abcdef";
    std::wstring out;
    out.reserve(32);
    for (unsigned char b : raw) {
        out.push_back(hex[b >> 4]);
        out.push_back(hex[b & 0xF]);
    }
    return out;
}

// Write-through temp file + rename commit, matching WriteAtomicBytesAsync.
inline bool write_atomic_bytes(const std::wstring& path, const unsigned char* data, std::size_t size) {
    std::wstring temp_path = path + L".tmp." + random_temp_suffix();
    if (!write_file_raw(temp_path, data, size, /*write_through*/ true, /*fail_if_exists*/ true)) {
        return false;
    }
    if (!MoveFileExW(temp_path.c_str(), path.c_str(), MOVEFILE_REPLACE_EXISTING)) {
        DeleteFileW(temp_path.c_str());
        return false;
    }
    return true;
}

inline bool write_atomic_bytes(const std::wstring& path, const std::string& data) {
    return write_atomic_bytes(path, reinterpret_cast<const unsigned char*>(data.data()), data.size());
}

inline void set_hidden(const std::wstring& path) {
    DWORD attributes = GetFileAttributesW(path.c_str());
    if (attributes != INVALID_FILE_ATTRIBUTES) {
        SetFileAttributesW(path.c_str(), attributes | FILE_ATTRIBUTE_HIDDEN);
    }
}

} // namespace fsutil

// ---------------------------------------------------------------------------
// Shared store plumbing
// ---------------------------------------------------------------------------

namespace detail {

struct SecurityContext {
    std::string path_fingerprint; // sha256 hex of UTF-8(normalized path)
    std::vector<unsigned char> hmac_key;
};

inline std::string fingerprint_for_path(const std::wstring& path) {
    return crypto::sha256_hex(fsutil::wide_to_utf8(fsutil::normalize_path(path)));
}

inline std::wstring backup_path(const std::wstring& snapshot_path, int generation) {
    return snapshot_path + L".bak" + std::to_wstring(generation);
}

inline void add_unique_path(std::vector<std::wstring>& paths, const std::wstring& candidate) {
    if (candidate.empty()) return;
    std::wstring upper_candidate = fsutil::to_upper_invariant(candidate);
    for (const auto& existing : paths) {
        if (fsutil::to_upper_invariant(existing) == upper_candidate) return;
    }
    paths.push_back(candidate);
}

// Mirror path under %LOCALAPPDATA%\XactCopy\<mirror_dir>\{base}-{pathhash16}{ext}.
inline std::wstring build_mirror_path(const std::wstring& primary_path,
                                      const wchar_t* mirror_dir,
                                      const wchar_t* fallback_base) {
    std::wstring root = fsutil::local_app_data() + L"\\XactCopy\\" + mirror_dir;
    fsutil::create_directories(root);

    std::string path_hash = fingerprint_for_path(primary_path);
    std::wstring file_name = fsutil::get_file_name(primary_path);
    if (file_name.empty()) file_name = std::wstring(fallback_base) + L".json";

    std::wstring base, extension;
    fsutil::split_extension(file_name, base, extension);
    if (extension.empty()) extension = L".json";
    if (base.empty()) base = fallback_base;

    std::wstring hash16 = fsutil::utf8_to_wide(path_hash.substr(0, 16));
    return root + L"\\" + base + L"-" + hash16 + extension;
}

inline void rotate_backups(const std::wstring& snapshot_path, int generation_count) {
    for (int generation = generation_count; generation >= 1; --generation) {
        std::wstring destination = backup_path(snapshot_path, generation);
        std::wstring source = generation == 1 ? snapshot_path : backup_path(snapshot_path, generation - 1);

        if (generation > 1) {
            if (fsutil::file_exists(source)) {
                MoveFileExW(source.c_str(), destination.c_str(), MOVEFILE_REPLACE_EXISTING);
            } else if (fsutil::file_exists(destination)) {
                DeleteFileW(destination.c_str());
            }
        } else {
            if (fsutil::file_exists(destination)) {
                DeleteFileW(destination.c_str());
            }
            if (fsutil::file_exists(source)) {
                CopyFileW(source.c_str(), destination.c_str(), FALSE);
            }
        }
    }
}

inline void save_snapshot_set(const std::wstring& snapshot_path, const std::string& payload,
                              int generation_count) {
    std::wstring directory = fsutil::get_directory_name(snapshot_path);
    if (!directory.empty()) fsutil::create_directories(directory);
    rotate_backups(snapshot_path, generation_count);
    fsutil::write_atomic_bytes(snapshot_path, payload);
}

// FNV-1a 32-bit, matching ComputeChecksum32.
inline std::uint32_t checksum32(const unsigned char* data, std::size_t size) {
    std::uint32_t hash = 0x811C9DC5u;
    for (std::size_t i = 0; i < size; ++i) {
        hash = (hash ^ data[i]) * 0x01000193u;
    }
    return hash;
}

inline std::string new_guid_n() {
    std::vector<unsigned char> raw(16);
    crypto::fill_random(raw);
    raw[6] = static_cast<unsigned char>((raw[6] & 0x0F) | 0x40); // version 4
    raw[8] = static_cast<unsigned char>((raw[8] & 0x3F) | 0x80); // RFC variant
    return crypto::to_hex_lower(raw);
}

// Shared HMAC key loader. The journal key is stored raw; the bad-range-map key
// is DPAPI-protected with legacy raw-key migration. Both live under
// %LOCALAPPDATA%\XactCopy\security and are cached per process.
class HmacKeyCache {
public:
    HmacKeyCache(const wchar_t* file_name, bool dpapi_protected)
        : file_name_(file_name), dpapi_protected_(dpapi_protected) {}

    std::vector<unsigned char> get_or_create() {
        std::lock_guard<std::mutex> guard(lock_);
        if (cached_.size() == KeySizeBytes) return cached_;

        std::wstring security_root = fsutil::local_app_data() + L"\\XactCopy\\security";
        fsutil::create_directories(security_root);
        std::wstring key_path = security_root + L"\\" + file_name_;

        std::vector<unsigned char> loaded;
        if (auto file_bytes = fsutil::read_all_bytes(key_path)) {
            if (dpapi_protected_) {
                if (file_bytes->size() == KeySizeBytes) {
                    // Legacy raw key: accept and re-protect in place.
                    loaded = *file_bytes;
                    try_save_dpapi(key_path, loaded);
                } else {
                    try {
                        auto raw = crypto::dpapi_unprotect(*file_bytes);
                        if (raw.size() == KeySizeBytes) loaded = raw;
                    } catch (const std::exception&) {
                        loaded.clear();
                    }
                }
            } else if (file_bytes->size() == KeySizeBytes) {
                loaded = *file_bytes;
            }
        }

        if (loaded.size() != KeySizeBytes) {
            loaded.assign(KeySizeBytes, 0);
            crypto::fill_random(loaded);
            if (dpapi_protected_) {
                try_save_dpapi(key_path, loaded);
            } else {
                save_raw(key_path, loaded);
            }
        }

        cached_ = loaded;
        return cached_;
    }

private:
    static constexpr std::size_t KeySizeBytes = 32;

    std::wstring file_name_;
    bool dpapi_protected_;
    std::mutex lock_;
    std::vector<unsigned char> cached_;

    static void save_raw(const std::wstring& key_path, std::vector<unsigned char>& key) {
        std::wstring temp_path = key_path + L".tmp." + fsutil::random_temp_suffix();
        if (!fsutil::write_file_raw(temp_path, key.data(), key.size(), false, true)) return;
        DeleteFileW(key_path.c_str());
        if (!MoveFileW(temp_path.c_str(), key_path.c_str())) {
            DeleteFileW(temp_path.c_str());
            // Lost the race with another process — adopt its key instead.
            if (auto existing = fsutil::read_all_bytes(key_path);
                existing.has_value() && existing->size() == KeySizeBytes) {
                key = *existing;
            }
        }
        fsutil::set_hidden(key_path);
    }

    static void try_save_dpapi(const std::wstring& key_path, const std::vector<unsigned char>& key) {
        try {
            auto protected_bytes = crypto::dpapi_protect(key);
            std::wstring temp_path = key_path + L".tmp." + fsutil::random_temp_suffix();
            if (!fsutil::write_file_raw(temp_path, protected_bytes.data(), protected_bytes.size(),
                                        false, true)) {
                return;
            }
            DeleteFileW(key_path.c_str());
            if (!MoveFileW(temp_path.c_str(), key_path.c_str())) {
                DeleteFileW(temp_path.c_str());
            }
            fsutil::set_hidden(key_path);
        } catch (const std::exception&) {
            // DPAPI unavailable — key stays unprotected this session (matches .NET).
        }
    }
};

inline HmacKeyCache& journal_key_cache() {
    static HmacKeyCache cache(L"journal-hmac.key", /*dpapi*/ false);
    return cache;
}

inline HmacKeyCache& badmap_key_cache() {
    static HmacKeyCache cache(L"badmap-hmac.key", /*dpapi*/ true);
    return cache;
}

} // namespace detail

// ---------------------------------------------------------------------------
// BadRangeMapStore
// ---------------------------------------------------------------------------

class BadRangeMapStore {
public:
    static constexpr int BackupGenerationCount = 2;
    static constexpr std::int32_t CurrentEnvelopeSchemaVersion = 1;

    static std::wstring get_default_map_path(const std::wstring& source_root) {
        std::string hash = detail::fingerprint_for_path(source_root);
        std::wstring root = fsutil::local_app_data() + L"\\XactCopy\\badmaps";
        fsutil::create_directories(root);
        return root + L"\\map-" + fsutil::utf8_to_wide(hash.substr(0, 20)) + L".json";
    }

    std::optional<BadRangeMap> load(const std::wstring& map_path) {
        if (map_path.empty()) return std::nullopt;

        std::vector<std::wstring> candidates = build_candidates(map_path);
        detail::SecurityContext context = create_security_context(map_path);

        for (const auto& candidate : candidates) {
            if (auto map = try_load_envelope(candidate, context)) return map;
        }
        // Legacy pre-envelope snapshots.
        for (const auto& candidate : candidates) {
            if (auto map = try_load_legacy(candidate)) return map;
        }
        return std::nullopt;
    }

    // save_utc lets tests pin the UpdatedUtc/SavedUtcTicks values; production
    // callers omit it (matches .NET DateTimeOffset.UtcNow behavior).
    void save(const std::wstring& map_path, BadRangeMap map,
              std::optional<DateTimeOffset> save_utc = std::nullopt) {
        if (map_path.empty()) {
            throw std::invalid_argument("Map path is required.");
        }

        DateTimeOffset now = save_utc.value_or(DateTimeOffset::now_utc());
        normalize_map(map);
        map.UpdatedUtc = now;

        detail::SecurityContext context = create_security_context(map_path);
        std::string payload = serialize_map(map);
        std::string payload_sha = crypto::sha256_hex(payload);
        std::int64_t saved_ticks = now.utc_ticks();
        std::string signature = compute_envelope_signature(payload_sha, saved_ticks, context);

        json::Writer w(/*indented*/ true);
        w.begin_object();
        w.key("SchemaVersion"); w.value(CurrentEnvelopeSchemaVersion);
        w.key("SavedUtcTicks"); w.value(saved_ticks);
        w.key("PayloadSha256"); w.value(payload_sha);
        w.key("Signature"); w.value(signature);
        w.key("Payload");
        map.to_json(w);
        w.end_object();
        std::string envelope = w.take();

        std::wstring mirror_path = detail::build_mirror_path(map_path, L"badmaps-mirror", L"badmap");
        detail::save_snapshot_set(map_path, envelope, BackupGenerationCount);
        detail::save_snapshot_set(mirror_path, envelope, BackupGenerationCount);
    }

    // Serializes exactly like the .NET store's payload serializer (indented).
    static std::string serialize_map(const BadRangeMap& map) {
        json::Writer w(/*indented*/ true);
        map.to_json(w);
        return w.take();
    }

    static void normalize_map(BadRangeMap& map) {
        if (map.SchemaVersion <= 0) map.SchemaVersion = 1;
        map.SourceRoot = trim_utf8(map.SourceRoot);
        map.SourceIdentity = trim_utf8(map.SourceIdentity);
        if (map.UpdatedUtc == DateTimeOffset::min_value()) map.UpdatedUtc = DateTimeOffset::now_utc();

        OrderedFileMap<BadRangeMapFileEntry> normalized;
        for (auto& [key, entry] : map.Files.entries) {
            if (is_blank(key)) continue;

            BadRangeMapFileEntry item = entry;
            item.RelativePath = trim_utf8(item.RelativePath.empty() ? key : item.RelativePath);
            if (item.RelativePath.empty()) item.RelativePath = key;
            if (item.LastScanUtc == DateTimeOffset::min_value()) item.LastScanUtc = map.UpdatedUtc;

            std::vector<ByteRange> ranges;
            for (const auto& range : item.BadRanges) {
                if (range.Length <= 0) continue;
                ByteRange clamped;
                clamped.Offset = range.Offset < 0 ? 0 : range.Offset;
                clamped.Length = range.Length < 0 ? 0 : range.Length;
                if (clamped.Length > 0) ranges.push_back(clamped);
            }
            std::stable_sort(ranges.begin(), ranges.end(),
                             [](const ByteRange& a, const ByteRange& b) { return a.Offset < b.Offset; });
            item.BadRanges = std::move(ranges);

            // Copy the key before moving the entry (argument evaluation order
            // would otherwise read a moved-from RelativePath).
            std::string entry_key = item.RelativePath;
            normalized.set(std::move(entry_key), std::move(item));
        }
        map.Files = std::move(normalized);
    }

private:
    static bool is_blank(std::string_view text) {
        for (char c : text) {
            if (c != ' ' && c != '\t' && c != '\r' && c != '\n') return false;
        }
        return true;
    }

    static std::string trim_utf8(const std::string& text) {
        std::size_t begin = 0, end = text.size();
        auto is_space = [](char c) { return c == ' ' || c == '\t' || c == '\r' || c == '\n'; };
        while (begin < end && is_space(text[begin])) ++begin;
        while (end > begin && is_space(text[end - 1])) --end;
        return text.substr(begin, end - begin);
    }

    static detail::SecurityContext create_security_context(const std::wstring& map_path) {
        detail::SecurityContext context;
        context.path_fingerprint = detail::fingerprint_for_path(map_path);
        context.hmac_key = detail::badmap_key_cache().get_or_create();
        return context;
    }

    static std::string compute_envelope_signature(const std::string& payload_sha,
                                                  std::int64_t saved_ticks,
                                                  const detail::SecurityContext& context) {
        std::string payload = payload_sha + "|" + std::to_string(saved_ticks) + "|" + context.path_fingerprint;
        return crypto::to_base64(crypto::hmac_sha256(context.hmac_key, payload));
    }

    static std::vector<std::wstring> build_candidates(const std::wstring& map_path) {
        std::vector<std::wstring> candidates;
        detail::add_unique_path(candidates, map_path);
        for (int generation = 1; generation <= BackupGenerationCount; ++generation) {
            detail::add_unique_path(candidates, detail::backup_path(map_path, generation));
        }
        std::wstring mirror = detail::build_mirror_path(map_path, L"badmaps-mirror", L"badmap");
        detail::add_unique_path(candidates, mirror);
        for (int generation = 1; generation <= BackupGenerationCount; ++generation) {
            detail::add_unique_path(candidates, detail::backup_path(mirror, generation));
        }
        return candidates;
    }

    std::optional<BadRangeMap> try_load_envelope(const std::wstring& snapshot_path,
                                                 const detail::SecurityContext& context) {
        auto bytes = fsutil::read_all_bytes(snapshot_path);
        if (!bytes.has_value() || bytes->empty()) return std::nullopt;

        try {
            std::string_view text(reinterpret_cast<const char*>(bytes->data()), bytes->size());
            json::Value root = json::parse(text);
            const auto* obj = root.as_object();
            if (obj == nullptr) return std::nullopt;

            std::int32_t schema_version = 0;
            std::int64_t saved_ticks = 0;
            std::string payload_sha, signature;
            models::detail::read(obj, "SchemaVersion", schema_version);
            models::detail::read(obj, "SavedUtcTicks", saved_ticks);
            models::detail::read(obj, "PayloadSha256", payload_sha);
            models::detail::read(obj, "Signature", signature);

            const json::Value* payload_value = obj->find("Payload");
            if (payload_value == nullptr || !payload_value->is_object()) return std::nullopt;
            if (schema_version <= 0) return std::nullopt;

            // Verify by re-serializing the payload exactly as .NET does (the
            // envelope-embedded payload is indented at a deeper level than the
            // standalone bytes that were hashed on save).
            BadRangeMap map = BadRangeMap::from_json(*payload_value);
            std::string reserialized = serialize_map(map);
            std::string actual_sha = crypto::sha256_hex(reserialized);
            if (!crypto::secure_equals(actual_sha, payload_sha)) return std::nullopt;

            std::string expected_signature = compute_envelope_signature(actual_sha, saved_ticks, context);
            if (!crypto::secure_equals(expected_signature, signature)) return std::nullopt;

            normalize_map(map);
            return map;
        } catch (const std::exception&) {
            return std::nullopt;
        }
    }

    std::optional<BadRangeMap> try_load_legacy(const std::wstring& snapshot_path) {
        auto bytes = fsutil::read_all_bytes(snapshot_path);
        if (!bytes.has_value() || bytes->empty()) return std::nullopt;

        try {
            std::string_view text(reinterpret_cast<const char*>(bytes->data()), bytes->size());
            json::Value root = json::parse(text);
            if (!root.is_object()) return std::nullopt;
            BadRangeMap map = BadRangeMap::from_json(root);
            normalize_map(map);
            return map;
        } catch (const std::exception&) {
            return std::nullopt;
        }
    }
};

// ---------------------------------------------------------------------------
// JobJournalStore
// ---------------------------------------------------------------------------

class JobJournalStore {
public:
    static constexpr int BackupGenerationCount = 3;
    static constexpr std::uint32_t LedgerMagic = 0x58434A4Cu; // "LJCX" on disk (LE)
    static constexpr unsigned char LedgerVersion = 1;
    static constexpr int LedgerFrameHeaderSize = 9;
    static constexpr int LedgerFrameTrailerSize = 4;
    static constexpr std::int64_t LedgerMaxBytes = 8LL * 1024 * 1024;
    static constexpr int LedgerCompactRecordCount = 2048;

    struct LedgerRecord {
        std::int64_t Sequence = 0;
        std::int64_t UpdatedUtcTicks = 0;
        std::string SnapshotHash;
        std::int64_t SnapshotLength = 0;
        std::string PreviousRecordHash;
        std::string RecordHash;
        std::string Signature;

        void to_json(json::Writer& w) const {
            w.begin_object();
            w.key("Sequence"); w.value(Sequence);
            w.key("UpdatedUtcTicks"); w.value(UpdatedUtcTicks);
            w.key("SnapshotHash"); w.value(SnapshotHash);
            w.key("SnapshotLength"); w.value(SnapshotLength);
            w.key("PreviousRecordHash"); w.value(PreviousRecordHash);
            w.key("RecordHash"); w.value(RecordHash);
            w.key("Signature"); w.value(Signature);
            w.end_object();
        }

        static LedgerRecord from_json(const json::Value& value) {
            LedgerRecord r;
            const auto* obj = value.as_object();
            if (obj == nullptr) return r;
            models::detail::read(obj, "Sequence", r.Sequence);
            models::detail::read(obj, "UpdatedUtcTicks", r.UpdatedUtcTicks);
            models::detail::read(obj, "SnapshotHash", r.SnapshotHash);
            models::detail::read(obj, "SnapshotLength", r.SnapshotLength);
            models::detail::read(obj, "PreviousRecordHash", r.PreviousRecordHash);
            models::detail::read(obj, "RecordHash", r.RecordHash);
            models::detail::read(obj, "Signature", r.Signature);
            return r;
        }
    };

    static std::string build_job_id(const std::wstring& source_root, const std::wstring& destination_root) {
        std::wstring payload_wide = fsutil::to_upper_invariant(fsutil::trim(source_root)) + L"|" +
                                    fsutil::to_upper_invariant(fsutil::trim(destination_root));
        std::string hash = crypto::sha256_hex(fsutil::wide_to_utf8(payload_wide));
        return hash.substr(0, 20);
    }

    static std::wstring get_default_journal_path(const std::string& job_id) {
        std::wstring root = fsutil::local_app_data() + L"\\XactCopy\\journals";
        fsutil::create_directories(root);
        return root + L"\\job-" + fsutil::utf8_to_wide(job_id) + L".json";
    }

    std::optional<JobJournal> load(const std::wstring& journal_path) {
        if (journal_path.empty()) return std::nullopt;

        std::vector<std::wstring> candidates = build_candidates(journal_path);
        bool any_exists = false;
        for (const auto& candidate : candidates) {
            if (fsutil::file_exists(candidate)) { any_exists = true; break; }
        }
        if (!any_exists) return std::nullopt;

        detail::SecurityContext context = create_security_context(journal_path);

        // Trusted hashes: hash -> highest ledger sequence, insertion-ordered.
        std::vector<std::pair<std::string, std::int64_t>> trusted;
        auto merge_trusted = [&trusted](const LedgerRecord& record) {
            if (record.SnapshotHash.empty()) return;
            for (auto& item : trusted) {
                if (models::detail::equals_ignore_case(item.first, record.SnapshotHash)) {
                    if (record.Sequence > item.second) item.second = record.Sequence;
                    return;
                }
            }
            trusted.emplace_back(record.SnapshotHash, record.Sequence);
        };

        std::wstring mirror_path = detail::build_mirror_path(journal_path, L"journals-mirror", L"journal");
        for (const auto& record : read_ledger_records(ledger_path(journal_path), context)) merge_trusted(record);
        for (const auto& record : read_ledger_records(ledger_path(mirror_path), context)) merge_trusted(record);

        if (!trusted.empty()) {
            // First readable snapshot per trusted hash, in candidate order.
            std::vector<std::pair<std::string, JobJournal>> trusted_snapshots;
            for (const auto& candidate : candidates) {
                auto snapshot = try_read_snapshot(candidate);
                if (!snapshot.has_value()) continue;
                bool hash_trusted = false;
                for (const auto& item : trusted) {
                    if (models::detail::equals_ignore_case(item.first, snapshot->first)) {
                        hash_trusted = true;
                        break;
                    }
                }
                if (!hash_trusted) continue;
                bool already = false;
                for (const auto& existing : trusted_snapshots) {
                    if (models::detail::equals_ignore_case(existing.first, snapshot->first)) {
                        already = true;
                        break;
                    }
                }
                if (!already) trusted_snapshots.emplace_back(snapshot->first, std::move(snapshot->second));
            }

            std::stable_sort(trusted.begin(), trusted.end(),
                             [](const auto& a, const auto& b) { return a.second > b.second; });
            for (const auto& [hash, sequence] : trusted) {
                for (auto& [snapshot_hash, journal] : trusted_snapshots) {
                    if (models::detail::equals_ignore_case(snapshot_hash, hash)) {
                        return std::move(journal);
                    }
                }
            }
        }

        // Legacy/plain snapshot fallback.
        for (const auto& candidate : candidates) {
            if (auto snapshot = try_read_snapshot(candidate)) {
                return std::move(snapshot->second);
            }
        }
        return std::nullopt;
    }

    void save(const std::wstring& journal_path, JobJournal journal,
              std::optional<DateTimeOffset> save_utc = std::nullopt) {
        if (journal_path.empty()) {
            throw std::invalid_argument("Journal path is required.");
        }

        std::wstring directory = fsutil::get_directory_name(journal_path);
        if (!directory.empty()) fsutil::create_directories(directory);

        DateTimeOffset now = save_utc.value_or(DateTimeOffset::now_utc());
        normalize_journal(journal);
        journal.UpdatedUtc = now;

        json::Writer w(/*indented*/ true);
        journal.to_json(w);
        std::string payload = w.take();
        std::string snapshot_hash = crypto::sha256_hex(payload);

        std::wstring mirror_path = detail::build_mirror_path(journal_path, L"journals-mirror", L"journal");
        detail::save_snapshot_set(journal_path, payload, BackupGenerationCount);
        detail::save_snapshot_set(mirror_path, payload, BackupGenerationCount);

        detail::SecurityContext context = create_security_context(journal_path);
        std::wstring primary_anchor = anchor_path(journal_path);
        std::wstring mirror_anchor = anchor_path(mirror_path);
        std::wstring primary_ledger = ledger_path(journal_path);
        std::wstring mirror_ledger = ledger_path(mirror_path);
        std::wstring seed_cache_key = fsutil::normalize_path(journal_path);

        SeedState seed;
        if (!try_get_cached_seed(seed_cache_key, seed)) {
            seed = resolve_seed_state(primary_anchor, mirror_anchor, primary_ledger, mirror_ledger, context);
            set_cached_seed(seed_cache_key, seed);
        }

        LedgerRecord record;
        record.Sequence = seed.sequence + 1;
        record.UpdatedUtcTicks = now.utc_ticks();
        record.SnapshotHash = snapshot_hash;
        record.SnapshotLength = static_cast<std::int64_t>(payload.size());
        record.PreviousRecordHash = seed.last_record_hash;
        record.RecordHash = compute_record_hash(record);
        record.Signature = compute_record_signature(record, context);

        append_ledger_record(primary_ledger, record);
        append_ledger_record(mirror_ledger, record);

        std::string anchor_json = build_anchor_json(record.Sequence, record.RecordHash,
                                                    record.UpdatedUtcTicks, context);
        write_anchor(primary_anchor, anchor_json);
        write_anchor(mirror_anchor, anchor_json);

        compact_ledger_if_needed(primary_ledger, context);
        compact_ledger_if_needed(mirror_ledger, context);

        set_cached_seed(seed_cache_key, SeedState{record.Sequence, record.RecordHash});
    }

    static void normalize_journal(JobJournal& journal) {
        if (is_blank(journal.JobId)) journal.JobId = detail::new_guid_n();
        if (journal.CreatedUtc == DateTimeOffset::min_value()) journal.CreatedUtc = DateTimeOffset::now_utc();
        if (journal.UpdatedUtc == DateTimeOffset::min_value()) journal.UpdatedUtc = journal.CreatedUtc;

        OrderedFileMap<JournalFileEntry> normalized;
        for (auto& [key, entry] : journal.Files.entries) {
            if (is_blank(key)) continue;
            JournalFileEntry item = entry;
            if (is_blank(item.RelativePath)) item.RelativePath = key;
            normalized.set(key, std::move(item));
        }
        journal.Files = std::move(normalized);
    }

    // Exposed for tests/tools.
    std::vector<LedgerRecord> read_ledger(const std::wstring& journal_path) {
        detail::SecurityContext context = create_security_context(journal_path);
        return read_ledger_records(ledger_path(journal_path), context);
    }

    static std::wstring ledger_path(const std::wstring& snapshot_path) { return snapshot_path + L".ledger"; }
    static std::wstring anchor_path(const std::wstring& snapshot_path) { return snapshot_path + L".anchor"; }

private:
    struct SeedState {
        std::int64_t sequence = 0;
        std::string last_record_hash;
    };

    std::mutex seed_lock_;
    std::vector<std::pair<std::wstring, SeedState>> seed_cache_;

    static bool is_blank(std::string_view text) {
        for (char c : text) {
            if (c != ' ' && c != '\t' && c != '\r' && c != '\n') return false;
        }
        return true;
    }

    static detail::SecurityContext create_security_context(const std::wstring& journal_path) {
        detail::SecurityContext context;
        context.path_fingerprint = detail::fingerprint_for_path(journal_path);
        context.hmac_key = detail::journal_key_cache().get_or_create();
        return context;
    }

    static std::vector<std::wstring> build_candidates(const std::wstring& journal_path) {
        std::vector<std::wstring> candidates;
        detail::add_unique_path(candidates, journal_path);
        for (int generation = 1; generation <= BackupGenerationCount; ++generation) {
            detail::add_unique_path(candidates, detail::backup_path(journal_path, generation));
        }
        std::wstring mirror = detail::build_mirror_path(journal_path, L"journals-mirror", L"journal");
        detail::add_unique_path(candidates, mirror);
        for (int generation = 1; generation <= BackupGenerationCount; ++generation) {
            detail::add_unique_path(candidates, detail::backup_path(mirror, generation));
        }
        return candidates;
    }

    // Returns (snapshot hash, journal) or nullopt.
    std::optional<std::pair<std::string, JobJournal>> try_read_snapshot(const std::wstring& snapshot_path) {
        auto bytes = fsutil::read_all_bytes(snapshot_path);
        if (!bytes.has_value() || bytes->empty()) return std::nullopt;

        try {
            std::string snapshot_hash = crypto::sha256_hex(*bytes);
            std::string_view text(reinterpret_cast<const char*>(bytes->data()), bytes->size());
            json::Value root = json::parse(text);
            if (!root.is_object()) return std::nullopt;
            JobJournal journal = JobJournal::from_json(root);
            normalize_journal(journal);
            return std::make_pair(std::move(snapshot_hash), std::move(journal));
        } catch (const std::exception&) {
            return std::nullopt;
        }
    }

    static std::string compute_record_hash(const LedgerRecord& record) {
        std::string payload = std::to_string(record.Sequence) + "|" +
                              std::to_string(record.UpdatedUtcTicks) + "|" +
                              record.SnapshotHash + "|" +
                              std::to_string(record.SnapshotLength) + "|" +
                              record.PreviousRecordHash;
        return crypto::sha256_hex(payload);
    }

    static std::string compute_record_signature(const LedgerRecord& record,
                                                const detail::SecurityContext& context) {
        std::string payload = record.RecordHash + "|" + context.path_fingerprint;
        return crypto::to_base64(crypto::hmac_sha256(context.hmac_key, payload));
    }

    static std::string compute_anchor_signature(std::int64_t sequence,
                                                const std::string& last_record_hash,
                                                std::int64_t updated_ticks,
                                                const detail::SecurityContext& context) {
        std::string payload = std::to_string(sequence) + "|" + last_record_hash + "|" +
                              std::to_string(updated_ticks) + "|" + context.path_fingerprint;
        return crypto::to_base64(crypto::hmac_sha256(context.hmac_key, payload));
    }

    static bool is_record_valid(const LedgerRecord& record, const LedgerRecord* previous,
                                const detail::SecurityContext& context) {
        if (record.Sequence <= 0) return false;
        if (record.SnapshotHash.size() != 64) return false;
        if (record.RecordHash.size() != 64) return false;
        if (!crypto::secure_equals(record.RecordHash, compute_record_hash(record))) return false;
        if (!crypto::secure_equals(record.Signature, compute_record_signature(record, context))) return false;
        if (previous != nullptr) {
            if (record.Sequence <= previous->Sequence) return false;
            if (!crypto::secure_equals(record.PreviousRecordHash, previous->RecordHash)) return false;
        }
        return true;
    }

    static std::vector<unsigned char> build_frame(const LedgerRecord& record) {
        json::Writer w(/*indented*/ true);
        record.to_json(w);
        std::string payload = w.take();

        std::uint32_t checksum = detail::checksum32(
            reinterpret_cast<const unsigned char*>(payload.data()), payload.size());

        std::vector<unsigned char> frame;
        frame.reserve(LedgerFrameHeaderSize + payload.size() + LedgerFrameTrailerSize);
        auto push_u32 = [&frame](std::uint32_t value) {
            frame.push_back(static_cast<unsigned char>(value & 0xFF));
            frame.push_back(static_cast<unsigned char>((value >> 8) & 0xFF));
            frame.push_back(static_cast<unsigned char>((value >> 16) & 0xFF));
            frame.push_back(static_cast<unsigned char>((value >> 24) & 0xFF));
        };
        push_u32(LedgerMagic);
        frame.push_back(LedgerVersion);
        push_u32(static_cast<std::uint32_t>(payload.size()));
        frame.insert(frame.end(), payload.begin(), payload.end());
        push_u32(checksum);
        return frame;
    }

    static void append_ledger_record(const std::wstring& path, const LedgerRecord& record) {
        std::wstring directory = fsutil::get_directory_name(path);
        if (!directory.empty()) fsutil::create_directories(directory);

        std::vector<unsigned char> frame = build_frame(record);
        HANDLE handle = CreateFileW(path.c_str(), FILE_APPEND_DATA, FILE_SHARE_READ, nullptr,
                                    OPEN_ALWAYS, FILE_ATTRIBUTE_NORMAL | FILE_FLAG_WRITE_THROUGH,
                                    nullptr);
        if (handle == INVALID_HANDLE_VALUE) return;
        DWORD written = 0;
        WriteFile(handle, frame.data(), static_cast<DWORD>(frame.size()), &written, nullptr);
        FlushFileBuffers(handle);
        CloseHandle(handle);
    }

    static std::vector<LedgerRecord> read_ledger_records(const std::wstring& path,
                                                         const detail::SecurityContext& context) {
        std::vector<LedgerRecord> records;
        auto bytes = fsutil::read_all_bytes(path);
        if (!bytes.has_value()) return records;

        const unsigned char* data = bytes->data();
        std::size_t size = bytes->size();
        std::size_t position = 0;
        const LedgerRecord* last_accepted = nullptr;

        auto read_u32 = [](const unsigned char* p) {
            return static_cast<std::uint32_t>(p[0]) |
                   (static_cast<std::uint32_t>(p[1]) << 8) |
                   (static_cast<std::uint32_t>(p[2]) << 16) |
                   (static_cast<std::uint32_t>(p[3]) << 24);
        };

        try {
            while (position + LedgerFrameHeaderSize <= size) {
                std::uint32_t magic = read_u32(data + position);
                if (magic != LedgerMagic) break;
                unsigned char version = data[position + 4];
                if (version != LedgerVersion) break;
                std::int32_t payload_length = static_cast<std::int32_t>(read_u32(data + position + 5));
                if (payload_length <= 0 || payload_length > 1024 * 1024) break;

                std::size_t payload_start = position + LedgerFrameHeaderSize;
                std::size_t trailer_start = payload_start + static_cast<std::size_t>(payload_length);
                if (trailer_start + LedgerFrameTrailerSize > size) break;

                std::uint32_t expected_checksum = read_u32(data + trailer_start);
                std::uint32_t actual_checksum = detail::checksum32(
                    data + payload_start, static_cast<std::size_t>(payload_length));
                if (expected_checksum != actual_checksum) break;

                std::string_view payload_text(
                    reinterpret_cast<const char*>(data + payload_start),
                    static_cast<std::size_t>(payload_length));
                LedgerRecord record = LedgerRecord::from_json(json::parse(payload_text));
                if (!is_record_valid(record, last_accepted, context)) break;

                records.push_back(std::move(record));
                last_accepted = &records.back();
                position = trailer_start + LedgerFrameTrailerSize;
            }
        } catch (const std::exception&) {
            return {};
        }
        return records;
    }

    static void compact_ledger_if_needed(const std::wstring& path,
                                         const detail::SecurityContext& context) {
        HANDLE handle = CreateFileW(path.c_str(), GENERIC_READ,
                                    FILE_SHARE_READ | FILE_SHARE_WRITE, nullptr,
                                    OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, nullptr);
        if (handle == INVALID_HANDLE_VALUE) return;
        LARGE_INTEGER size{};
        BOOL got_size = GetFileSizeEx(handle, &size);
        CloseHandle(handle);
        if (!got_size || size.QuadPart <= LedgerMaxBytes) return;

        std::vector<LedgerRecord> records = read_ledger_records(path, context);
        if (records.size() <= static_cast<std::size_t>(LedgerCompactRecordCount)) return;

        std::vector<LedgerRecord> retained(
            records.end() - LedgerCompactRecordCount, records.end());

        std::wstring temp_path = path + L".tmp." + fsutil::random_temp_suffix();
        std::vector<unsigned char> all;
        for (const auto& record : retained) {
            std::vector<unsigned char> frame = build_frame(record);
            all.insert(all.end(), frame.begin(), frame.end());
        }
        if (fsutil::write_file_raw(temp_path, all.data(), all.size(), true, true)) {
            if (!MoveFileExW(temp_path.c_str(), path.c_str(), MOVEFILE_REPLACE_EXISTING)) {
                DeleteFileW(temp_path.c_str());
            }
        }
    }

    std::string build_anchor_json(std::int64_t sequence, const std::string& last_record_hash,
                                  std::int64_t updated_ticks,
                                  const detail::SecurityContext& context) {
        json::Writer w(/*indented*/ true);
        w.begin_object();
        w.key("Sequence"); w.value(sequence);
        w.key("LastRecordHash"); w.value(last_record_hash);
        w.key("UpdatedUtcTicks"); w.value(updated_ticks);
        w.key("Signature"); w.value(compute_anchor_signature(sequence, last_record_hash, updated_ticks, context));
        w.end_object();
        return w.take();
    }

    static void write_anchor(const std::wstring& path, const std::string& payload) {
        std::wstring directory = fsutil::get_directory_name(path);
        if (!directory.empty()) fsutil::create_directories(directory);
        fsutil::write_atomic_bytes(path, payload);
    }

    std::optional<SeedState> read_anchor(const std::wstring& path,
                                         const detail::SecurityContext& context) {
        auto bytes = fsutil::read_all_bytes(path);
        if (!bytes.has_value() || bytes->empty()) return std::nullopt;

        try {
            std::string_view text(reinterpret_cast<const char*>(bytes->data()), bytes->size());
            json::Value root = json::parse(text);
            const auto* obj = root.as_object();
            if (obj == nullptr) return std::nullopt;

            std::int64_t sequence = 0, updated_ticks = 0;
            std::string last_record_hash, signature;
            models::detail::read(obj, "Sequence", sequence);
            models::detail::read(obj, "LastRecordHash", last_record_hash);
            models::detail::read(obj, "UpdatedUtcTicks", updated_ticks);
            models::detail::read(obj, "Signature", signature);

            if (sequence < 0) return std::nullopt;
            if (sequence > 0 && last_record_hash.size() != 64) return std::nullopt;
            std::string expected = compute_anchor_signature(sequence, last_record_hash, updated_ticks, context);
            if (!crypto::secure_equals(signature, expected)) return std::nullopt;

            return SeedState{sequence, last_record_hash};
        } catch (const std::exception&) {
            return std::nullopt;
        }
    }

    SeedState resolve_seed_state(const std::wstring& primary_anchor, const std::wstring& mirror_anchor,
                                 const std::wstring& primary_ledger, const std::wstring& mirror_ledger,
                                 const detail::SecurityContext& context) {
        std::vector<SeedState> states;
        if (auto anchor = read_anchor(primary_anchor, context)) states.push_back(*anchor);
        if (auto anchor = read_anchor(mirror_anchor, context)) states.push_back(*anchor);

        auto primary_records = read_ledger_records(primary_ledger, context);
        if (!primary_records.empty()) {
            const auto& last = primary_records.back();
            states.push_back(SeedState{last.Sequence, last.RecordHash});
        }
        auto mirror_records = read_ledger_records(mirror_ledger, context);
        if (!mirror_records.empty()) {
            const auto& last = mirror_records.back();
            states.push_back(SeedState{last.Sequence, last.RecordHash});
        }

        if (states.empty()) return SeedState{};
        SeedState best = states.front();
        for (std::size_t i = 1; i < states.size(); ++i) {
            if (states[i].sequence > best.sequence) best = states[i];
        }
        return best;
    }

    bool try_get_cached_seed(const std::wstring& key, SeedState& state) {
        if (key.empty()) return false;
        std::lock_guard<std::mutex> guard(seed_lock_);
        for (const auto& item : seed_cache_) {
            if (item.first == key) {
                state = item.second;
                return true;
            }
        }
        return false;
    }

    void set_cached_seed(const std::wstring& key, const SeedState& state) {
        if (key.empty()) return;
        std::lock_guard<std::mutex> guard(seed_lock_);
        for (auto& item : seed_cache_) {
            if (item.first == key) {
                item.second = state;
                return;
            }
        }
        seed_cache_.emplace_back(key, state);
    }
};

} // namespace xact::storage
