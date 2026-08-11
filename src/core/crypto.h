// -----------------------------------------------------------------------------
// File: src\core\crypto.h
// Purpose: Crypto primitives for XactCopy storage signing: SHA-256 and
//          HMAC-SHA256 via Windows CNG (BCrypt), DPAPI key protection,
//          secure random bytes, hex/base64 encoding, constant-time compare.
// -----------------------------------------------------------------------------

#pragma once

#include <cstdint>
#include <stdexcept>
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
#include <bcrypt.h>
#include <dpapi.h>

#if defined(_MSC_VER)
#pragma comment(lib, "bcrypt.lib")
#pragma comment(lib, "crypt32.lib")
#endif

namespace xact::crypto {

class CryptoError : public std::runtime_error {
public:
    CryptoError(const std::string& message, long status)
        : std::runtime_error(message + " (status 0x" + to_hex32(static_cast<std::uint32_t>(status)) + ")") {}

private:
    static std::string to_hex32(std::uint32_t value) {
        static const char digits[] = "0123456789ABCDEF";
        std::string out(8, '0');
        for (int i = 7; i >= 0; --i) {
            out[static_cast<std::size_t>(i)] = digits[value & 0xF];
            value >>= 4;
        }
        return out;
    }
};

namespace detail {

// Cached CNG algorithm handles (opened once per process; BCrypt handles are
// thread-safe for hash creation). Index: 0=SHA256, 1=SHA256-HMAC, 2=SHA512.
inline BCRYPT_ALG_HANDLE hash_algorithm(int slot_index) {
    static BCRYPT_ALG_HANDLE slots[3] = {nullptr, nullptr, nullptr};
    BCRYPT_ALG_HANDLE& slot = slots[slot_index];
    if (slot == nullptr) {
        BCRYPT_ALG_HANDLE handle = nullptr;
        const wchar_t* algorithm = slot_index == 2 ? BCRYPT_SHA512_ALGORITHM : BCRYPT_SHA256_ALGORITHM;
        NTSTATUS status = BCryptOpenAlgorithmProvider(
            &handle, algorithm, nullptr,
            slot_index == 1 ? BCRYPT_ALG_HANDLE_HMAC_FLAG : 0);
        if (status != 0) {
            throw CryptoError("BCryptOpenAlgorithmProvider failed.", status);
        }
        // Benign race: a second thread may open a duplicate handle; leak it.
        if (InterlockedCompareExchangePointer(
                reinterpret_cast<PVOID*>(&slot), handle, nullptr) != nullptr) {
            BCryptCloseAlgorithmProvider(handle, 0);
        }
    }
    return slot;
}

// Incremental hasher for chunked/stream hashing (SHA-256 or SHA-512).
class StreamingHash {
public:
    explicit StreamingHash(bool sha512) : digest_size_(sha512 ? 64u : 32u) {
        NTSTATUS status = BCryptCreateHash(
            hash_algorithm(sha512 ? 2 : 0), &hash_, nullptr, 0, nullptr, 0, 0);
        if (status != 0) {
            throw CryptoError("BCryptCreateHash failed.", status);
        }
    }

    ~StreamingHash() {
        if (hash_ != nullptr) BCryptDestroyHash(hash_);
    }

    StreamingHash(const StreamingHash&) = delete;
    StreamingHash& operator=(const StreamingHash&) = delete;

    void update(const unsigned char* data, std::size_t size) {
        NTSTATUS status = BCryptHashData(hash_, const_cast<unsigned char*>(data),
                                         static_cast<ULONG>(size), 0);
        if (status != 0) {
            throw CryptoError("BCryptHashData failed.", status);
        }
    }

    std::vector<unsigned char> finish() {
        std::vector<unsigned char> digest(digest_size_);
        NTSTATUS status = BCryptFinishHash(hash_, digest.data(), static_cast<ULONG>(digest.size()), 0);
        if (status != 0) {
            throw CryptoError("BCryptFinishHash failed.", status);
        }
        return digest;
    }

private:
    BCRYPT_HASH_HANDLE hash_ = nullptr;
    unsigned int digest_size_;
};

inline std::vector<unsigned char> hash_sha256(
    const unsigned char* data, std::size_t size,
    const unsigned char* key, std::size_t key_size) {

    const bool hmac = key != nullptr;
    BCRYPT_ALG_HANDLE algorithm = hash_algorithm(hmac ? 1 : 0);

    BCRYPT_HASH_HANDLE hash = nullptr;
    NTSTATUS status = BCryptCreateHash(
        algorithm, &hash, nullptr, 0,
        const_cast<unsigned char*>(key), static_cast<ULONG>(key_size), 0);
    if (status != 0) {
        throw CryptoError("BCryptCreateHash failed.", status);
    }

    status = BCryptHashData(hash, const_cast<unsigned char*>(data), static_cast<ULONG>(size), 0);
    if (status != 0) {
        BCryptDestroyHash(hash);
        throw CryptoError("BCryptHashData failed.", status);
    }

    std::vector<unsigned char> digest(32);
    status = BCryptFinishHash(hash, digest.data(), static_cast<ULONG>(digest.size()), 0);
    BCryptDestroyHash(hash);
    if (status != 0) {
        throw CryptoError("BCryptFinishHash failed.", status);
    }
    return digest;
}

} // namespace detail

// One-shot digest of a memory buffer with SHA-256 (sha512=false) or SHA-512.
inline std::vector<unsigned char> hash_buffer(const unsigned char* data, std::size_t size, bool sha512) {
    detail::StreamingHash hasher(sha512);
    if (size > 0) hasher.update(data, size);
    return hasher.finish();
}

inline std::vector<unsigned char> sha256(std::string_view data) {
    return detail::hash_sha256(
        reinterpret_cast<const unsigned char*>(data.data()), data.size(), nullptr, 0);
}

inline std::vector<unsigned char> sha256(const std::vector<unsigned char>& data) {
    return detail::hash_sha256(data.data(), data.size(), nullptr, 0);
}

inline std::vector<unsigned char> hmac_sha256(
    const std::vector<unsigned char>& key, std::string_view data) {
    return detail::hash_sha256(
        reinterpret_cast<const unsigned char*>(data.data()), data.size(),
        key.data(), key.size());
}

// Lowercase hex, matching Convert.ToHexString(...).ToLowerInvariant().
inline std::string to_hex_lower(const std::vector<unsigned char>& bytes) {
    static const char digits[] = "0123456789abcdef";
    std::string out;
    out.reserve(bytes.size() * 2);
    for (unsigned char b : bytes) {
        out.push_back(digits[b >> 4]);
        out.push_back(digits[b & 0xF]);
    }
    return out;
}

inline std::string sha256_hex(std::string_view data) { return to_hex_lower(sha256(data)); }
inline std::string sha256_hex(const std::vector<unsigned char>& data) { return to_hex_lower(sha256(data)); }

// Standard base64 with padding, matching Convert.ToBase64String.
inline std::string to_base64(const std::vector<unsigned char>& bytes) {
    static const char table[] = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/";
    std::string out;
    out.reserve(((bytes.size() + 2) / 3) * 4);
    std::size_t i = 0;
    while (i + 3 <= bytes.size()) {
        std::uint32_t chunk = (static_cast<std::uint32_t>(bytes[i]) << 16) |
                              (static_cast<std::uint32_t>(bytes[i + 1]) << 8) |
                              bytes[i + 2];
        out.push_back(table[(chunk >> 18) & 0x3F]);
        out.push_back(table[(chunk >> 12) & 0x3F]);
        out.push_back(table[(chunk >> 6) & 0x3F]);
        out.push_back(table[chunk & 0x3F]);
        i += 3;
    }
    if (i + 1 == bytes.size()) {
        std::uint32_t chunk = static_cast<std::uint32_t>(bytes[i]) << 16;
        out.push_back(table[(chunk >> 18) & 0x3F]);
        out.push_back(table[(chunk >> 12) & 0x3F]);
        out.append("==");
    } else if (i + 2 == bytes.size()) {
        std::uint32_t chunk = (static_cast<std::uint32_t>(bytes[i]) << 16) |
                              (static_cast<std::uint32_t>(bytes[i + 1]) << 8);
        out.push_back(table[(chunk >> 18) & 0x3F]);
        out.push_back(table[(chunk >> 12) & 0x3F]);
        out.push_back(table[(chunk >> 6) & 0x3F]);
        out.push_back('=');
    }
    return out;
}

// Constant-time string equality, matching CryptographicOperations.FixedTimeEquals
// over the UTF-8 bytes (length mismatch returns false immediately, as in .NET).
inline bool secure_equals(std::string_view left, std::string_view right) {
    if (left.size() != right.size()) return false;
    unsigned char accumulator = 0;
    for (std::size_t i = 0; i < left.size(); ++i) {
        accumulator = static_cast<unsigned char>(
            accumulator | (static_cast<unsigned char>(left[i]) ^ static_cast<unsigned char>(right[i])));
    }
    return accumulator == 0;
}

inline void fill_random(std::vector<unsigned char>& buffer) {
    NTSTATUS status = BCryptGenRandom(
        nullptr, buffer.data(), static_cast<ULONG>(buffer.size()),
        BCRYPT_USE_SYSTEM_PREFERRED_RNG);
    if (status != 0) {
        throw CryptoError("BCryptGenRandom failed.", status);
    }
}

// ---------------------------------------------------------------------------
// DPAPI (CurrentUser scope, no UI) — used for map/catalog HMAC key files.
// ---------------------------------------------------------------------------

inline std::vector<unsigned char> dpapi_protect(const std::vector<unsigned char>& data) {
    DATA_BLOB in{static_cast<DWORD>(data.size()), const_cast<BYTE*>(data.data())};
    DATA_BLOB out{};
    if (!CryptProtectData(&in, nullptr, nullptr, nullptr, nullptr,
                          CRYPTPROTECT_UI_FORBIDDEN, &out)) {
        throw CryptoError("CryptProtectData failed.", static_cast<long>(GetLastError()));
    }
    std::vector<unsigned char> result(out.pbData, out.pbData + out.cbData);
    LocalFree(out.pbData);
    return result;
}

inline std::vector<unsigned char> dpapi_unprotect(const std::vector<unsigned char>& data) {
    DATA_BLOB in{static_cast<DWORD>(data.size()), const_cast<BYTE*>(data.data())};
    DATA_BLOB out{};
    if (!CryptUnprotectData(&in, nullptr, nullptr, nullptr, nullptr,
                            CRYPTPROTECT_UI_FORBIDDEN, &out)) {
        throw CryptoError("CryptUnprotectData failed.", static_cast<long>(GetLastError()));
    }
    std::vector<unsigned char> result(out.pbData, out.pbData + out.cbData);
    LocalFree(out.pbData);
    return result;
}

} // namespace xact::crypto
