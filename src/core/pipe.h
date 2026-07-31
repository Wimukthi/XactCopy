// -----------------------------------------------------------------------------
// File: src\core\pipe.h
// Purpose: Win32 named-pipe transport for XactCopy IPC. Reproduces the .NET
//          JsonMessagePipe framing (4-byte little-endian length prefix, 16 MB
//          frame ceiling) and the CurrentUserOnly server security descriptor.
// -----------------------------------------------------------------------------

#pragma once

#include <cstdint>
#include <optional>
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
#include <sddl.h>

namespace xact::pipe {

inline constexpr std::int32_t MaximumPayloadBytes = 16 * 1024 * 1024;

class PipeError : public std::runtime_error {
public:
    PipeError(const std::string& message, DWORD error_code)
        : std::runtime_error(message + " (Win32 error " + std::to_string(error_code) + ")"),
          code(error_code) {}

    DWORD code;
};

// Signals pending overlapped operations to abort. One instance is shared
// between the pipe connection and whichever threads may request shutdown.
class CancellationEvent {
public:
    CancellationEvent() {
        handle_ = CreateEventW(nullptr, TRUE, FALSE, nullptr);
        if (handle_ == nullptr) {
            throw PipeError("CreateEvent failed for cancellation event.", GetLastError());
        }
    }

    ~CancellationEvent() {
        if (handle_ != nullptr) CloseHandle(handle_);
    }

    CancellationEvent(const CancellationEvent&) = delete;
    CancellationEvent& operator=(const CancellationEvent&) = delete;

    void cancel() noexcept { SetEvent(handle_); }
    void reset() noexcept { ResetEvent(handle_); }
    bool is_cancelled() const noexcept { return WaitForSingleObject(handle_, 0) == WAIT_OBJECT_0; }
    HANDLE handle() const noexcept { return handle_; }

private:
    HANDLE handle_ = nullptr;
};

namespace detail {

// Owner + protected DACL granting GENERIC_ALL to the current user only —
// the same shape PipeOptions.CurrentUserOnly produces on the .NET side.
inline PSECURITY_DESCRIPTOR build_current_user_only_descriptor() {
    HANDLE token = nullptr;
    if (!OpenProcessToken(GetCurrentProcess(), TOKEN_QUERY, &token)) {
        throw PipeError("OpenProcessToken failed.", GetLastError());
    }

    DWORD required = 0;
    GetTokenInformation(token, TokenUser, nullptr, 0, &required);
    std::vector<unsigned char> buffer(required);
    if (!GetTokenInformation(token, TokenUser, buffer.data(), required, &required)) {
        DWORD error = GetLastError();
        CloseHandle(token);
        throw PipeError("GetTokenInformation failed.", error);
    }
    CloseHandle(token);

    auto* token_user = reinterpret_cast<TOKEN_USER*>(buffer.data());
    LPWSTR sid_text = nullptr;
    if (!ConvertSidToStringSidW(token_user->User.Sid, &sid_text)) {
        throw PipeError("ConvertSidToStringSid failed.", GetLastError());
    }

    std::wstring sddl = L"O:";
    sddl += sid_text;
    sddl += L"D:P(A;;GA;;;";
    sddl += sid_text;
    sddl += L")";
    LocalFree(sid_text);

    PSECURITY_DESCRIPTOR descriptor = nullptr;
    if (!ConvertStringSecurityDescriptorToSecurityDescriptorW(
            sddl.c_str(), SDDL_REVISION_1, &descriptor, nullptr)) {
        throw PipeError("ConvertStringSecurityDescriptorToSecurityDescriptor failed.", GetLastError());
    }
    return descriptor; // Caller frees with LocalFree.
}

} // namespace detail

// A connected message pipe endpoint (server or client side). All reads/writes
// use overlapped I/O so a CancellationEvent can abort them at any point.
class MessagePipe {
public:
    MessagePipe() = default;

    MessagePipe(const MessagePipe&) = delete;
    MessagePipe& operator=(const MessagePipe&) = delete;

    MessagePipe(MessagePipe&& other) noexcept { swap(other); }
    MessagePipe& operator=(MessagePipe&& other) noexcept {
        if (this != &other) { close(); swap(other); }
        return *this;
    }

    ~MessagePipe() { close(); }

    bool is_open() const noexcept { return handle_ != INVALID_HANDLE_VALUE && handle_ != nullptr; }

    void close() noexcept {
        if (is_open()) {
            CancelIoEx(handle_, nullptr);
            CloseHandle(handle_);
        }
        handle_ = INVALID_HANDLE_VALUE;
    }

    // ---- Server-side lifecycle -------------------------------------------

    // Creates the pipe exactly like the .NET worker: single instance, byte
    // mode, duplex, overlapped, current-user-only security descriptor.
    static MessagePipe create_server(const std::wstring& pipe_name) {
        std::wstring full_name = L"\\\\.\\pipe\\" + pipe_name;

        PSECURITY_DESCRIPTOR descriptor = detail::build_current_user_only_descriptor();
        SECURITY_ATTRIBUTES attributes{};
        attributes.nLength = sizeof(attributes);
        attributes.lpSecurityDescriptor = descriptor;
        attributes.bInheritHandle = FALSE;

        HANDLE handle = CreateNamedPipeW(
            full_name.c_str(),
            PIPE_ACCESS_DUPLEX | FILE_FLAG_OVERLAPPED | FILE_FLAG_FIRST_PIPE_INSTANCE,
            PIPE_TYPE_BYTE | PIPE_READMODE_BYTE | PIPE_WAIT | PIPE_REJECT_REMOTE_CLIENTS,
            1,
            64 * 1024,
            64 * 1024,
            0,
            &attributes);
        DWORD error = GetLastError();
        LocalFree(descriptor);

        if (handle == INVALID_HANDLE_VALUE) {
            throw PipeError("CreateNamedPipe failed.", error);
        }

        MessagePipe pipe;
        pipe.handle_ = handle;
        pipe.is_server_ = true;
        return pipe;
    }

    // Waits for a client connection; abortable via the cancellation event.
    void wait_for_connection(CancellationEvent& cancellation) {
        OVERLAPPED overlapped{};
        overlapped.hEvent = CreateEventW(nullptr, TRUE, FALSE, nullptr);
        if (overlapped.hEvent == nullptr) {
            throw PipeError("CreateEvent failed for ConnectNamedPipe.", GetLastError());
        }

        BOOL connected = ConnectNamedPipe(handle_, &overlapped);
        DWORD error = connected ? ERROR_SUCCESS : GetLastError();
        if (!connected && error == ERROR_IO_PENDING) {
            error = wait_overlapped(overlapped, cancellation, nullptr) ? ERROR_SUCCESS : ERROR_OPERATION_ABORTED;
        } else if (!connected && error == ERROR_PIPE_CONNECTED) {
            error = ERROR_SUCCESS;
        }

        CloseHandle(overlapped.hEvent);
        if (error != ERROR_SUCCESS) {
            throw PipeError("ConnectNamedPipe failed or was cancelled.", error);
        }
    }

    // ---- Client-side lifecycle -------------------------------------------

    static MessagePipe connect_client(const std::wstring& pipe_name, DWORD timeout_ms) {
        std::wstring full_name = L"\\\\.\\pipe\\" + pipe_name;
        const ULONGLONG deadline = GetTickCount64() + timeout_ms;

        while (true) {
            HANDLE handle = CreateFileW(
                full_name.c_str(),
                GENERIC_READ | GENERIC_WRITE,
                0,
                nullptr,
                OPEN_EXISTING,
                FILE_FLAG_OVERLAPPED,
                nullptr);
            if (handle != INVALID_HANDLE_VALUE) {
                MessagePipe pipe;
                pipe.handle_ = handle;
                return pipe;
            }

            DWORD error = GetLastError();
            if (error != ERROR_PIPE_BUSY && error != ERROR_FILE_NOT_FOUND) {
                throw PipeError("Named pipe connect failed.", error);
            }
            if (GetTickCount64() >= deadline) {
                throw PipeError("Named pipe connect timed out.", ERROR_TIMEOUT);
            }
            Sleep(50);
        }
    }

    // ---- Framed messages --------------------------------------------------

    // Writes one frame: 4-byte little-endian length prefix + UTF-8 payload.
    void write_message(std::string_view message, CancellationEvent& cancellation) {
        if (message.size() > static_cast<std::size_t>(MaximumPayloadBytes)) {
            throw std::length_error(
                "IPC payload length " + std::to_string(message.size()) +
                " exceeds max allowed size " + std::to_string(MaximumPayloadBytes) + ".");
        }

        const std::uint32_t length = static_cast<std::uint32_t>(message.size());
        unsigned char prefix[4] = {
            static_cast<unsigned char>(length & 0xFF),
            static_cast<unsigned char>((length >> 8) & 0xFF),
            static_cast<unsigned char>((length >> 16) & 0xFF),
            static_cast<unsigned char>((length >> 24) & 0xFF)};

        write_all(prefix, 4, cancellation);
        if (length > 0) {
            write_all(reinterpret_cast<const unsigned char*>(message.data()), length, cancellation);
        }
        FlushFileBuffers(handle_);
    }

    // Reads one frame. Returns std::nullopt on clean end-of-stream (peer
    // closed at a frame boundary); throws on mid-frame EOF or invalid length.
    std::optional<std::string> read_message(CancellationEvent& cancellation) {
        unsigned char prefix[4];
        std::size_t got = read_some(prefix, 4, cancellation);
        if (got == 0) return std::nullopt;
        if (got != 4) throw std::runtime_error("IPC stream ended while reading message length.");

        const std::int32_t payload_length =
            static_cast<std::int32_t>(prefix[0]) |
            (static_cast<std::int32_t>(prefix[1]) << 8) |
            (static_cast<std::int32_t>(prefix[2]) << 16) |
            (static_cast<std::int32_t>(prefix[3]) << 24);
        if (payload_length < 0 || payload_length > MaximumPayloadBytes) {
            throw std::runtime_error(
                "IPC payload length is invalid or exceeds limit: " + std::to_string(payload_length) +
                " (max " + std::to_string(MaximumPayloadBytes) + ").");
        }
        if (payload_length == 0) return std::string();

        std::string payload(static_cast<std::size_t>(payload_length), '\0');
        got = read_some(reinterpret_cast<unsigned char*>(payload.data()), payload.size(), cancellation);
        if (got != payload.size()) {
            throw std::runtime_error("IPC stream ended while reading message payload.");
        }
        return payload;
    }

private:
    HANDLE handle_ = INVALID_HANDLE_VALUE;
    bool is_server_ = false;

    void swap(MessagePipe& other) noexcept {
        HANDLE h = handle_; handle_ = other.handle_; other.handle_ = h;
        bool s = is_server_; is_server_ = other.is_server_; other.is_server_ = s;
    }

    // Waits for an overlapped operation or cancellation. Returns false when
    // cancelled (the operation is aborted with CancelIoEx first).
    bool wait_overlapped(OVERLAPPED& overlapped, CancellationEvent& cancellation, DWORD* transferred) {
        HANDLE waits[2] = {overlapped.hEvent, cancellation.handle()};
        DWORD wait_result = WaitForMultipleObjects(2, waits, FALSE, INFINITE);
        if (wait_result == WAIT_OBJECT_0 + 1) {
            CancelIoEx(handle_, &overlapped);
            DWORD ignored = 0;
            GetOverlappedResult(handle_, &overlapped, &ignored, TRUE);
            return false;
        }

        DWORD bytes = 0;
        if (!GetOverlappedResult(handle_, &overlapped, &bytes, TRUE)) {
            DWORD error = GetLastError();
            if (error == ERROR_BROKEN_PIPE || error == ERROR_PIPE_NOT_CONNECTED || error == ERROR_HANDLE_EOF) {
                if (transferred != nullptr) *transferred = 0;
                return true;
            }
            throw PipeError("Overlapped pipe operation failed.", error);
        }
        if (transferred != nullptr) *transferred = bytes;
        return true;
    }

    // Reads exactly `size` bytes unless the stream ends first; returns the
    // number of bytes actually read. Throws when cancelled.
    std::size_t read_some(unsigned char* buffer, std::size_t size, CancellationEvent& cancellation) {
        std::size_t offset = 0;
        while (offset < size) {
            OVERLAPPED overlapped{};
            overlapped.hEvent = CreateEventW(nullptr, TRUE, FALSE, nullptr);
            if (overlapped.hEvent == nullptr) {
                throw PipeError("CreateEvent failed for pipe read.", GetLastError());
            }

            DWORD requested = static_cast<DWORD>(size - offset);
            DWORD transferred = 0;
            BOOL ok = ReadFile(handle_, buffer + offset, requested, nullptr, &overlapped);
            DWORD error = ok ? ERROR_SUCCESS : GetLastError();

            bool completed = true;
            if (!ok && error == ERROR_IO_PENDING) {
                completed = wait_overlapped(overlapped, cancellation, &transferred);
            } else if (!ok && (error == ERROR_BROKEN_PIPE || error == ERROR_PIPE_NOT_CONNECTED || error == ERROR_HANDLE_EOF)) {
                transferred = 0;
            } else if (!ok) {
                CloseHandle(overlapped.hEvent);
                throw PipeError("ReadFile on pipe failed.", error);
            } else {
                GetOverlappedResult(handle_, &overlapped, &transferred, TRUE);
            }

            CloseHandle(overlapped.hEvent);
            if (!completed) {
                throw PipeError("Pipe read cancelled.", ERROR_OPERATION_ABORTED);
            }
            if (transferred == 0) {
                return offset; // End of stream.
            }
            offset += transferred;
        }
        return offset;
    }

    void write_all(const unsigned char* buffer, std::size_t size, CancellationEvent& cancellation) {
        std::size_t offset = 0;
        while (offset < size) {
            OVERLAPPED overlapped{};
            overlapped.hEvent = CreateEventW(nullptr, TRUE, FALSE, nullptr);
            if (overlapped.hEvent == nullptr) {
                throw PipeError("CreateEvent failed for pipe write.", GetLastError());
            }

            DWORD requested = static_cast<DWORD>(size - offset);
            DWORD transferred = 0;
            BOOL ok = WriteFile(handle_, buffer + offset, requested, nullptr, &overlapped);
            DWORD error = ok ? ERROR_SUCCESS : GetLastError();

            bool completed = true;
            if (!ok && error == ERROR_IO_PENDING) {
                completed = wait_overlapped(overlapped, cancellation, &transferred);
            } else if (!ok) {
                CloseHandle(overlapped.hEvent);
                throw PipeError("WriteFile on pipe failed.", error);
            } else {
                GetOverlappedResult(handle_, &overlapped, &transferred, TRUE);
            }

            CloseHandle(overlapped.hEvent);
            if (!completed) {
                throw PipeError("Pipe write cancelled.", ERROR_OPERATION_ABORTED);
            }
            if (transferred == 0) {
                throw PipeError("Pipe write made no progress.", ERROR_WRITE_FAULT);
            }
            offset += transferred;
        }
    }
};

} // namespace xact::pipe
