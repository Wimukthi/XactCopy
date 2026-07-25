// UpdateService (port of the check portion of XactCopy.UI UpdateService.vb):
// fetches the latest GitHub release over WinHTTP and compares versions. The
// download/apply/elevation flow is intentionally out of scope here — the UI
// offers "open the release page" instead.
#pragma once

#ifndef NOMINMAX
#define NOMINMAX
#endif
#include <windows.h>
#include <winhttp.h>

#include <algorithm>
#include <array>
#include <cctype>
#include <cstdint>
#include <cstdio>
#include <functional>
#include <optional>
#include <string>
#include <vector>

#include "../core/crypto.h"
#include "../core/json.h"
#include "../storage/stores.h" // fsutil::utf8_to_wide / wide_to_utf8

namespace xact::ui {

struct UpdateAssetInfo {
    std::string name;
    std::string download_url;
    std::int64_t size = 0;
    std::string expected_sha256_hex; // from the release "digest", if present
};

struct DownloadProgress {
    std::int64_t received = 0;
    std::int64_t total = 0;
    double bytes_per_second = 0.0;
};

struct AppVersion {
    int major = 0, minor = 0, build = 0, revision = 0;

    // Normalized (build/revision floored at 0) lexicographic comparison.
    int compare(const AppVersion& o) const {
        std::array<int, 4> a{major, minor, std::max(0, build), std::max(0, revision)};
        std::array<int, 4> b{o.major, o.minor, std::max(0, o.build), std::max(0, o.revision)};
        for (int i = 0; i < 4; ++i) {
            if (a[i] != b[i]) return a[i] < b[i] ? -1 : 1;
        }
        return 0;
    }
};

// Native build version — the "current version" for update comparisons.
inline constexpr AppVersion kNativeVersion{2, 0, 0, 2};

struct UpdateReleaseInfo {
    std::string tag_name;
    std::string title;
    std::string notes;
    std::string html_url;
    AppVersion version;
    std::vector<UpdateAssetInfo> assets;
    bool ok = false; // fetch + parse succeeded
    std::string error;
};

namespace update_detail {

// Mirrors ParseVersionSafe/NormalizeVersionText: strip a leading v/version,
// keep the leading dotted-number run.
inline AppVersion parse_version(const std::string& text) {
    std::string s;
    std::size_t i = 0;
    while (i < text.size() && !(text[i] >= '0' && text[i] <= '9')) ++i; // skip "v"/"version "
    for (; i < text.size(); ++i) {
        char c = text[i];
        if ((c >= '0' && c <= '9') || c == '.') s.push_back(c);
        else break;
    }
    AppVersion v;
    int part = 0;
    int value = 0;
    bool has_digit = false;
    auto commit = [&] {
        switch (part) {
            case 0: v.major = value; break;
            case 1: v.minor = value; break;
            case 2: v.build = value; break;
            case 3: v.revision = value; break;
            default: break;
        }
        value = 0;
        has_digit = false;
        ++part;
    };
    for (char c : s) {
        if (c == '.') {
            commit();
            if (part > 3) break;
        } else {
            value = value * 10 + (c - '0');
            has_digit = true;
        }
    }
    if (has_digit && part <= 3) commit();
    return v;
}

inline std::string json_string(const json::Object* obj, std::string_view key) {
    if (obj == nullptr) return std::string();
    const json::Value* v = obj->find(key);
    return (v != nullptr && v->is_string()) ? v->as_string() : std::string();
}

// GitHub asset "digest" is "sha256:<64 hex>"; extract the hex if present.
inline std::string parse_sha256_candidate(const std::string& digest) {
    std::string s = digest;
    const std::string prefix = "sha256:";
    if (s.size() > prefix.size() &&
        models::detail::equals_ignore_case(s.substr(0, prefix.size()), prefix)) {
        s = s.substr(prefix.size());
    }
    if (s.size() != 64) return std::string();
    for (char c : s) {
        if (!((c >= '0' && c <= '9') || (c >= 'a' && c <= 'f') || (c >= 'A' && c <= 'F'))) {
            return std::string();
        }
    }
    return s;
}

// Case-insensitive hex comparison (SHA-256 digests differ only in a/A..f/F).
inline bool equals_hex_ci(const std::string& a, const std::string& b) {
    if (a.size() != b.size()) return false;
    for (std::size_t i = 0; i < a.size(); ++i) {
        if (std::tolower(static_cast<unsigned char>(a[i])) !=
            std::tolower(static_cast<unsigned char>(b[i]))) {
            return false;
        }
    }
    return true;
}

inline std::string to_lower_ascii(std::string s) {
    for (char& c : s) c = static_cast<char>(std::tolower(static_cast<unsigned char>(c)));
    return s;
}

inline bool ends_with(const std::string& s, const std::string& suffix) {
    return s.size() >= suffix.size() && s.compare(s.size() - suffix.size(), suffix.size(), suffix) == 0;
}

// File name after the last '/' or '\\'.
inline std::string base_name(const std::string& path) {
    std::size_t slash = path.find_last_of("/\\");
    return slash == std::string::npos ? path : path.substr(slash + 1);
}

// Mirrors IsChecksumAssetName.
inline bool is_checksum_asset_name(const std::string& name) {
    std::string lower = to_lower_ascii(name);
    // trim
    std::size_t b = lower.find_first_not_of(" \t\r\n");
    if (b == std::string::npos) return false;
    lower = lower.substr(b, lower.find_last_not_of(" \t\r\n") - b + 1);
    if (lower.empty()) return false;
    return ends_with(lower, ".sha256") || ends_with(lower, ".sha256sum") ||
           ends_with(lower, ".sha256.txt") || ends_with(lower, "checksums.txt") ||
           ends_with(lower, "checksums.sha256") || lower.find("sha256") != std::string::npos;
}

// Mirrors ScoreChecksumAssetName.
inline int score_checksum_asset_name(const std::string& checksum_name, const std::string& target_name) {
    std::string cn = to_lower_ascii(checksum_name);
    std::string tn = to_lower_ascii(target_name);
    int score = 0;
    if (cn.empty()) return score;
    if (!tn.empty()) {
        if (cn == tn + ".sha256" || cn == tn + ".sha256sum" || cn == tn + ".sha256.txt") score += 100;
        if (cn.find(tn) != std::string::npos) score += 25;
    }
    if (cn.find("checksums") != std::string::npos) score += 10;
    if (ends_with(cn, ".sha256") || ends_with(cn, ".sha256sum")) score += 10;
    return score;
}

// Mirrors ParseSha256FromChecksumText: matches "<hash>  filename" and
// "SHA256 (filename) = <hash>" styles, with a single-line fallback.
inline std::string parse_sha256_from_checksum_text(const std::string& text,
                                                   const std::string& target_asset_name) {
    if (text.empty()) return std::string();
    std::string target_file = to_lower_ascii(base_name(target_asset_name));
    std::string fallback;
    std::string normalized;
    normalized.reserve(text.size());
    for (char c : text) {
        if (c != '\r') normalized.push_back(c);
    }
    std::size_t line_count = 0;
    std::size_t start = 0;
    while (start <= normalized.size()) {
        std::size_t nl = normalized.find('\n', start);
        std::string line = normalized.substr(start, nl == std::string::npos ? std::string::npos : nl - start);
        start = (nl == std::string::npos) ? normalized.size() + 1 : nl + 1;
        // trim
        std::size_t lb = line.find_first_not_of(" \t");
        if (lb == std::string::npos) continue;
        line = line.substr(lb, line.find_last_not_of(" \t") - lb + 1);
        if (line.empty()) continue;
        ++line_count;

        std::size_t tok_end = line.find_first_of(" \t");
        std::string first_token = line.substr(0, tok_end);
        std::string parsed = parse_sha256_candidate(first_token);
        if (!parsed.empty()) {
            if (target_file.empty()) {
                if (fallback.empty()) fallback = parsed;
            } else {
                std::string remainder = (tok_end == std::string::npos) ? std::string() : line.substr(tok_end);
                std::size_t rb = remainder.find_first_not_of(" \t*\"");
                if (rb != std::string::npos) {
                    remainder = remainder.substr(rb);
                    std::size_t re = remainder.find_last_not_of(" \t\"");
                    remainder = remainder.substr(0, re + 1);
                }
                if (to_lower_ascii(base_name(remainder)) == target_file) return parsed;
                if (fallback.empty()) fallback = parsed;
            }
        }

        // "SHA256 (filename) = <hash>" style.
        std::string lower_line = to_lower_ascii(line);
        std::size_t marker = lower_line.find("sha256 (");
        std::size_t eq = line.find('=');
        if (marker != std::string::npos && eq != std::string::npos && eq > marker) {
            std::size_t name_start = marker + 8; // strlen("sha256 (")
            std::size_t name_end = line.find(')', name_start);
            if (name_end != std::string::npos && name_end > name_start) {
                std::string cand_name = line.substr(name_start, name_end - name_start);
                std::string hash_part = line.substr(eq + 1);
                std::size_t hb = hash_part.find_first_not_of(" \t");
                if (hb != std::string::npos) hash_part = hash_part.substr(hb);
                std::string cand_hash = parse_sha256_candidate(
                    hash_part.substr(0, hash_part.find_first_of(" \t")));
                if (!cand_hash.empty()) {
                    if (target_file.empty() || to_lower_ascii(base_name(cand_name)) == target_file) {
                        return cand_hash;
                    }
                    if (fallback.empty()) fallback = cand_hash;
                }
            }
        }
    }
    return line_count == 1 ? fallback : std::string();
}

// WinHTTP GET returning the response body (empty optional on any failure).
inline std::optional<std::string> https_get(const std::wstring& url,
                                            const std::wstring& user_agent) {
    URL_COMPONENTS uc{};
    uc.dwStructSize = sizeof(uc);
    wchar_t host[256]{};
    wchar_t path[2048]{};
    uc.lpszHostName = host;
    uc.dwHostNameLength = std::size(host);
    uc.lpszUrlPath = path;
    uc.dwUrlPathLength = std::size(path);
    if (!WinHttpCrackUrl(url.c_str(), 0, 0, &uc)) return std::nullopt;

    HINTERNET session = WinHttpOpen(user_agent.c_str(), WINHTTP_ACCESS_TYPE_AUTOMATIC_PROXY,
                                    WINHTTP_NO_PROXY_NAME, WINHTTP_NO_PROXY_BYPASS, 0);
    if (session == nullptr) return std::nullopt;
    DWORD timeout = 20000;
    WinHttpSetTimeouts(session, timeout, timeout, timeout, timeout);

    std::optional<std::string> result;
    HINTERNET connect = WinHttpConnect(session, host, uc.nPort, 0);
    if (connect != nullptr) {
        DWORD flags = (uc.nScheme == INTERNET_SCHEME_HTTPS) ? WINHTTP_FLAG_SECURE : 0;
        HINTERNET request = WinHttpOpenRequest(connect, L"GET", path, nullptr, WINHTTP_NO_REFERER,
                                               WINHTTP_DEFAULT_ACCEPT_TYPES, flags);
        if (request != nullptr) {
            static const wchar_t* accept = L"Accept: application/vnd.github+json\r\n";
            WinHttpAddRequestHeaders(request, accept, static_cast<DWORD>(-1),
                                     WINHTTP_ADDREQ_FLAG_ADD);
            if (WinHttpSendRequest(request, WINHTTP_NO_ADDITIONAL_HEADERS, 0, WINHTTP_NO_REQUEST_DATA,
                                   0, 0, 0) &&
                WinHttpReceiveResponse(request, nullptr)) {
                DWORD status = 0;
                DWORD size = sizeof(status);
                WinHttpQueryHeaders(request,
                                    WINHTTP_QUERY_STATUS_CODE | WINHTTP_QUERY_FLAG_NUMBER,
                                    WINHTTP_HEADER_NAME_BY_INDEX, &status, &size,
                                    WINHTTP_NO_HEADER_INDEX);
                if (status == 200) {
                    std::string body;
                    DWORD available = 0;
                    while (WinHttpQueryDataAvailable(request, &available) && available > 0) {
                        std::vector<char> chunk(available);
                        DWORD read = 0;
                        if (!WinHttpReadData(request, chunk.data(), available, &read) || read == 0) {
                            break;
                        }
                        body.append(chunk.data(), read);
                    }
                    result = std::move(body);
                }
            }
            WinHttpCloseHandle(request);
        }
        WinHttpCloseHandle(connect);
    }
    WinHttpCloseHandle(session);
    return result;
}

} // namespace update_detail

inline constexpr const char* kDefaultUpdateReleaseUrl =
    "https://api.github.com/repos/Wimukthi/XactCopy/releases/latest";

class UpdateService {
public:
    static bool is_update_available(const AppVersion& current, const AppVersion& latest) {
        return latest.compare(current) > 0;
    }

    // Mirrors SelectBestAsset: prefer Windows + this arch + zip/exe.
    static const UpdateAssetInfo* select_best_asset(const UpdateReleaseInfo& release) {
        const UpdateAssetInfo* best = nullptr;
        int best_score = -1;
        const std::string arch = sizeof(void*) == 8 ? "x64" : "x86";
        for (const auto& asset : release.assets) {
            std::string lower = asset.name;
            for (char& c : lower) c = static_cast<char>(std::tolower(static_cast<unsigned char>(c)));
            int score = 0;
            if (lower.find("win") != std::string::npos) score += 2;
            if (lower.find(arch) != std::string::npos) score += 3;
            if (lower.size() >= 4 && lower.compare(lower.size() - 4, 4, ".zip") == 0) score += 3;
            else if (lower.size() >= 4 && lower.compare(lower.size() - 4, 4, ".exe") == 0) score += 2;
            if (score > best_score) {
                best_score = score;
                best = &asset;
            }
        }
        return best;
    }

    // Resolves the expected SHA-256 for the asset: direct release digest first,
    // then any checksum sidecar asset in the same release (mirrors
    // ResolveAssetSha256Async). Returns lowercase hex, or "" if unavailable.
    static std::string resolve_asset_sha256(const UpdateReleaseInfo& release,
                                            const UpdateAssetInfo& asset) {
        std::string direct = update_detail::parse_sha256_candidate(asset.expected_sha256_hex);
        if (!direct.empty()) return direct;

        // Rank candidate checksum assets, best first.
        std::vector<const UpdateAssetInfo*> candidates;
        for (const auto& c : release.assets) {
            if (!c.download_url.empty() && update_detail::is_checksum_asset_name(c.name)) {
                candidates.push_back(&c);
            }
        }
        std::stable_sort(candidates.begin(), candidates.end(),
                         [&](const UpdateAssetInfo* a, const UpdateAssetInfo* b) {
                             return update_detail::score_checksum_asset_name(a->name, asset.name) >
                                    update_detail::score_checksum_asset_name(b->name, asset.name);
                         });
        for (const UpdateAssetInfo* c : candidates) {
            auto text = update_detail::https_get(storage::fsutil::utf8_to_wide(c->download_url),
                                                 L"XactCopy/2.0.0.2-native");
            if (!text.has_value()) continue;
            std::string parsed = update_detail::parse_sha256_from_checksum_text(*text, asset.name);
            if (!parsed.empty()) return parsed;
        }
        return std::string();
    }

    // Mirrors UpdateService.FormatBytes ("{size:F1} {unit}").
    static std::string format_bytes(std::int64_t value) {
        static const char* units[] = {"B", "KB", "MB", "GB", "TB"};
        double size = static_cast<double>(value);
        int unit = 0;
        while (size >= 1024.0 && unit < 4) {
            size /= 1024.0;
            ++unit;
        }
        char buf[64];
        std::snprintf(buf, sizeof(buf), "%.1f %s", size, units[unit]);
        return std::string(buf);
    }

    // Blocking download to `destination` with progress + SHA-256. Returns the
    // lowercase hex digest of the downloaded bytes (empty on failure).
    static std::string download_asset(const std::string& download_url,
                                      const std::wstring& destination, std::int64_t expected_total,
                                      const std::function<void(const DownloadProgress&)>& on_progress,
                                      const std::function<bool()>& cancelled) {
        std::wstring url = storage::fsutil::utf8_to_wide(download_url);
        URL_COMPONENTS uc{};
        uc.dwStructSize = sizeof(uc);
        wchar_t host[256]{};
        wchar_t path[4096]{};
        uc.lpszHostName = host;
        uc.dwHostNameLength = std::size(host);
        uc.lpszUrlPath = path;
        uc.dwUrlPathLength = std::size(path);
        if (!WinHttpCrackUrl(url.c_str(), 0, 0, &uc)) return std::string();

        HINTERNET session = WinHttpOpen(L"XactCopy/2.0.0.2-native", WINHTTP_ACCESS_TYPE_AUTOMATIC_PROXY,
                                        WINHTTP_NO_PROXY_NAME, WINHTTP_NO_PROXY_BYPASS, 0);
        if (session == nullptr) return std::string();
        DWORD to = 600000;
        WinHttpSetTimeouts(session, 20000, 20000, to, to);

        std::string digest;
        HINTERNET connect = WinHttpConnect(session, host, uc.nPort, 0);
        if (connect != nullptr) {
            DWORD flags = (uc.nScheme == INTERNET_SCHEME_HTTPS) ? WINHTTP_FLAG_SECURE : 0;
            HINTERNET request = WinHttpOpenRequest(connect, L"GET", path, nullptr, WINHTTP_NO_REFERER,
                                                   WINHTTP_DEFAULT_ACCEPT_TYPES, flags);
            if (request != nullptr) {
                if (WinHttpSendRequest(request, WINHTTP_NO_ADDITIONAL_HEADERS, 0,
                                       WINHTTP_NO_REQUEST_DATA, 0, 0, 0) &&
                    WinHttpReceiveResponse(request, nullptr)) {
                    DWORD status = 0, size = sizeof(status);
                    WinHttpQueryHeaders(request,
                                        WINHTTP_QUERY_STATUS_CODE | WINHTTP_QUERY_FLAG_NUMBER,
                                        WINHTTP_HEADER_NAME_BY_INDEX, &status, &size,
                                        WINHTTP_NO_HEADER_INDEX);
                    if (status == 200) {
                        digest = stream_to_file(request, destination, expected_total, on_progress,
                                                cancelled);
                    }
                }
                WinHttpCloseHandle(request);
            }
            WinHttpCloseHandle(connect);
        }
        WinHttpCloseHandle(session);
        return digest;
    }

private:
    static std::string stream_to_file(HINTERNET request, const std::wstring& destination,
                                      std::int64_t expected_total,
                                      const std::function<void(const DownloadProgress&)>& on_progress,
                                      const std::function<bool()>& cancelled) {
        HANDLE file = CreateFileW(destination.c_str(), GENERIC_WRITE, 0, nullptr, CREATE_ALWAYS,
                                  FILE_ATTRIBUTE_NORMAL, nullptr);
        if (file == INVALID_HANDLE_VALUE) return std::string();

        crypto::detail::StreamingHash hash(/*sha512=*/false);
        std::int64_t received = 0;
        ULONGLONG start = GetTickCount64();
        std::vector<char> buffer(64 * 1024);
        bool ok = true;
        DWORD available = 0;
        while (WinHttpQueryDataAvailable(request, &available) && available > 0) {
            if (cancelled && cancelled()) { ok = false; break; }
            DWORD want = std::min<DWORD>(available, static_cast<DWORD>(buffer.size()));
            DWORD read = 0;
            if (!WinHttpReadData(request, buffer.data(), want, &read) || read == 0) break;
            hash.update(reinterpret_cast<const unsigned char*>(buffer.data()), read);
            DWORD written = 0;
            if (!WriteFile(file, buffer.data(), read, &written, nullptr) || written != read) {
                ok = false;
                break;
            }
            received += read;
            if (on_progress) {
                double elapsed = static_cast<double>(GetTickCount64() - start) / 1000.0;
                DownloadProgress p;
                p.received = received;
                p.total = expected_total;
                p.bytes_per_second = elapsed > 0 ? static_cast<double>(received) / elapsed : 0.0;
                on_progress(p);
            }
        }
        CloseHandle(file);
        if (!ok) {
            DeleteFileW(destination.c_str());
            return std::string();
        }
        return crypto::to_hex_lower(hash.finish());
    }

public:

    // Blocking GitHub-release fetch (call from a background thread).
    static UpdateReleaseInfo get_latest_release(const std::string& release_url) {
        UpdateReleaseInfo info;
        if (release_url.empty()) {
            info.error = "Update release URL is not configured.";
            return info;
        }
        std::wstring url = storage::fsutil::utf8_to_wide(release_url);
        std::wstring user_agent = L"XactCopy/2.0.0.2-native";
        auto body = update_detail::https_get(url, user_agent);
        if (!body.has_value()) {
            info.error = "Could not reach the update server.";
            return info;
        }
        try {
            json::Value parsed = json::parse(*body);
            const json::Object* obj = parsed.as_object();
            if (obj == nullptr) {
                info.error = "Unexpected update response.";
                return info;
            }
            info.tag_name = update_detail::json_string(obj, "tag_name");
            info.title = update_detail::json_string(obj, "name");
            info.notes = update_detail::json_string(obj, "body");
            info.html_url = update_detail::json_string(obj, "html_url");
            info.version = update_detail::parse_version(info.tag_name);
            if (const json::Value* assets = obj->find("assets");
                assets != nullptr && assets->as_array() != nullptr) {
                for (const auto& entry : *assets->as_array()) {
                    const json::Object* a = entry.as_object();
                    if (a == nullptr) continue;
                    UpdateAssetInfo asset;
                    asset.name = update_detail::json_string(a, "name");
                    asset.download_url = update_detail::json_string(a, "browser_download_url");
                    if (const json::Value* sz = a->find("size"); sz != nullptr && sz->is_number()) {
                        asset.size = sz->as_int64();
                    }
                    asset.expected_sha256_hex =
                        update_detail::parse_sha256_candidate(update_detail::json_string(a, "digest"));
                    if (!asset.name.empty() && !asset.download_url.empty()) {
                        info.assets.push_back(std::move(asset));
                    }
                }
            }
            info.ok = true;
        } catch (const std::exception&) {
            info.error = "Could not parse the update response.";
        }
        return info;
    }
};

} // namespace xact::ui
