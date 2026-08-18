// -----------------------------------------------------------------------------
// File: src\worker\engine.h
// Purpose: Native port of ResilientCopyService — the XactCopy copy engine:
//          journal-merged resumable copies, rescue-pass pipeline (FastScan /
//          TrimSweep / TrimSweepReverse / Scrape / RetryBad), salvage fill,
//          retry/backoff with contention and availability policies, media
//          identity guard, native CopyFileEx fast path, and verification.
//          ScanOnly, bad-range-map hints, parallel small-file acceleration,
//          and adaptive buffering are included; raw-disk requests explicitly
//          use the read-only allocated-file raw-volume backend with standard
//          file-I/O fallback when the source layout is unsupported.
// -----------------------------------------------------------------------------

#pragma once

#include <algorithm>
#include <cmath>
#include <cstdint>
#include <deque>
#include <functional>
#include <mutex>
#include <optional>
#include <string>
#include <thread>
#include <unordered_set>
#include <vector>

#include "../core/crypto.h"
#include "../core/models.h"
#include "../storage/stores.h"
#include "engine_support.h"
#include "raw_disk_scan.h"

namespace xact::engine {

using models::CopyJobOptions;
using models::CopyJobResult;
using models::CopyProgressSnapshot;
using storage::ByteRange;
using storage::FileCopyState;
using storage::JobJournal;
using storage::JournalFileEntry;
using storage::RescueRange;
using storage::RescueRangeState;

inline constexpr std::int32_t AdaptiveAutoMinimumBufferSize = 64 * 1024;
inline constexpr std::int32_t AdaptiveAutoInitialSmallBufferSize = 256 * 1024;
inline constexpr std::int32_t AdaptiveAutoInitialCopyBufferSize = 4 * 1024 * 1024;
inline constexpr std::int32_t AdaptiveAutoInitialFastScanBufferSize = 2 * 1024 * 1024;
inline constexpr std::int32_t AdaptiveAutoMaxCopyBufferSize = 64 * 1024 * 1024;
inline constexpr std::int32_t AdaptiveAutoMaxScanBufferSize = 32 * 1024 * 1024;
inline constexpr std::int32_t AdaptiveAutoMaxFragileBufferSize = 2 * 1024 * 1024;
inline constexpr std::int32_t AdaptiveAutoParallelScanBufferBudget = 128 * 1024 * 1024;

enum class BufferPurpose { Copy, PreciseScan, FastHealthScan };

// A copy is built beside the final path and published only after the bytes
// have been verified and flushed. Keeping the temporary in the destination
// directory makes MoveFileExW a same-volume replacement, so a failed read,
// write, verification, or flush leaves the previous destination untouched.
class DestinationStage {
public:
    DestinationStage(std::wstring final_path, const CancelContext& cancel)
        : final_path_(std::move(final_path)), cancel_(&cancel) {
        working_path_ = final_path_ + L".xactcopy-stage." + storage::fsutil::random_temp_suffix();
        lock_path_ = working_path_ + L".lock";
        lock_handle_ = CreateFileW(lock_path_.c_str(), GENERIC_READ | GENERIC_WRITE,
                                   FILE_SHARE_READ, nullptr, CREATE_NEW,
                                   FILE_ATTRIBUTE_HIDDEN | FILE_FLAG_DELETE_ON_CLOSE, nullptr);
        if (lock_handle_ == INVALID_HANDLE_VALUE) {
            throw IoError::from_win32("Unable to reserve destination staging path.", GetLastError());
        }
    }

    ~DestinationStage() {
        if (!committed_) {
            SetFileAttributesW(working_path_.c_str(), FILE_ATTRIBUTE_NORMAL);
            DeleteFileW(working_path_.c_str());
        }
        if (lock_handle_ != INVALID_HANDLE_VALUE) {
            CloseHandle(lock_handle_);
            lock_handle_ = INVALID_HANDLE_VALUE;
        }
    }

    DestinationStage(const DestinationStage&) = delete;
    DestinationStage& operator=(const DestinationStage&) = delete;

    const std::wstring& working_path() const noexcept { return working_path_; }
    const std::wstring& final_path() const noexcept { return final_path_; }

    bool final_exists_now() const noexcept {
        DWORD attributes = GetFileAttributesW(final_path_.c_str());
        return attributes != INVALID_FILE_ATTRIBUTES &&
               (attributes & FILE_ATTRIBUTE_DIRECTORY) == 0;
    }

    DWORD post_publish_attribute_error() const noexcept {
        return post_publish_attribute_error_;
    }

    const std::wstring& published_path() const noexcept { return published_path_; }

    void make_writable() {
        DWORD attributes = GetFileAttributesW(working_path_.c_str());
        if (attributes == INVALID_FILE_ATTRIBUTES) {
            throw IoError::from_win32("Unable to inspect staged destination attributes.", GetLastError());
        }
        if ((attributes & FILE_ATTRIBUTE_READONLY) == 0) return;
        attributes &= ~FILE_ATTRIBUTE_READONLY;
        if (attributes == 0) attributes = FILE_ATTRIBUTE_NORMAL;
        if (!SetFileAttributesW(working_path_.c_str(), attributes)) {
            throw IoError::from_win32("Unable to make staged destination writable.", GetLastError());
        }
    }

    void clone_existing_if_present() {
        DWORD attributes = GetFileAttributesW(final_path_.c_str());
        if (attributes == INVALID_FILE_ATTRIBUTES) {
            DWORD error = GetLastError();
            if (error == ERROR_FILE_NOT_FOUND || error == ERROR_PATH_NOT_FOUND) return;
            throw IoError::from_win32("Unable to inspect existing destination before staging.", error);
        }
        if ((attributes & FILE_ATTRIBUTE_DIRECTORY) != 0) {
            throw IoError("Destination path is a directory, not a file.");
        }

        StageCopyState state{cancel_};
        BOOL cancel_flag = FALSE;
        BOOL ok = CopyFileExW(final_path_.c_str(), working_path_.c_str(),
                              stage_copy_progress, &state, &cancel_flag,
                              COPY_FILE_FAIL_IF_EXISTS);
        if (!ok) {
            if (cancel_->is_cancelled()) throw OperationCanceled{true};
            if (cancel_->deadline_passed()) throw OperationCanceled{false};
            throw IoError::from_win32("Unable to stage the existing destination.", GetLastError());
        }
        make_writable();
    }

    void commit() { commit_to(final_path_, true); }

    std::wstring commit_recovered_sidecar() {
        const std::wstring marker = L".xactcopy-recovered." +
                                    storage::fsutil::random_temp_suffix();
        const std::size_t separator = final_path_.find_last_of(L"\\/");
        const std::size_t extension = final_path_.find_last_of(L'.');
        const bool has_extension = extension != std::wstring::npos &&
                                   (separator == std::wstring::npos || extension > separator + 1);
        std::wstring sidecar = has_extension
                                   ? final_path_.substr(0, extension) + marker +
                                         final_path_.substr(extension)
                                   : final_path_ + marker;
        commit_to(sidecar, false);
        return sidecar;
    }

private:
    void commit_to(const std::wstring& publish_path, bool replace_existing) {
        DWORD staged_attributes = GetFileAttributesW(working_path_.c_str());
        if (staged_attributes == INVALID_FILE_ATTRIBUTES) {
            throw IoError::from_win32("Unable to inspect staged destination before publish.",
                                      GetLastError());
        }
        const bool staged_readonly = (staged_attributes & FILE_ATTRIBUTE_READONLY) != 0;

        DWORD final_attributes = GetFileAttributesW(publish_path.c_str());
        bool final_exists = final_attributes != INVALID_FILE_ATTRIBUTES &&
                            (final_attributes & FILE_ATTRIBUTE_DIRECTORY) == 0;
        if (!replace_existing && final_exists) {
            throw IoError("Recovered-output sidecar already exists.");
        }
        if (!final_exists && final_attributes == INVALID_FILE_ATTRIBUTES) {
            DWORD error = GetLastError();
            if (error != ERROR_FILE_NOT_FOUND && error != ERROR_PATH_NOT_FOUND) {
                throw IoError::from_win32("Unable to inspect existing destination before publish.",
                                          error);
            }
            final_attributes = FILE_ATTRIBUTE_NORMAL;
        }
        const bool final_readonly = final_exists &&
                                    (final_attributes & FILE_ATTRIBUTE_READONLY) != 0;
        bool published = false;

        // READONLY is a final-file policy, not a staging policy. Windows can
        // reject the durable flush or replacement while the bit is set, so
        // clear it only across the flush/publish boundary and put the exact
        // source attributes back after the atomic replacement.
        try {
            if (staged_readonly) {
                DWORD writable_attributes = staged_attributes & ~FILE_ATTRIBUTE_READONLY;
                if (writable_attributes == 0) writable_attributes = FILE_ATTRIBUTE_NORMAL;
                if (!SetFileAttributesW(working_path_.c_str(), writable_attributes)) {
                    throw IoError::from_win32("Unable to prepare read-only stage for publish.",
                                              GetLastError());
                }
            }
            flush_file_path(working_path_);

            if (final_readonly) {
                DWORD writable_attributes = final_attributes & ~FILE_ATTRIBUTE_READONLY;
                if (writable_attributes == 0) writable_attributes = FILE_ATTRIBUTE_NORMAL;
                if (!SetFileAttributesW(publish_path.c_str(), writable_attributes)) {
                    throw IoError::from_win32("Unable to prepare read-only destination for publish.",
                                              GetLastError());
                }
            }
            DWORD move_flags = MOVEFILE_WRITE_THROUGH |
                               (replace_existing ? MOVEFILE_REPLACE_EXISTING : 0);
            if (!MoveFileExW(working_path_.c_str(), publish_path.c_str(), move_flags)) {
                throw IoError::from_win32("Unable to publish staged destination.", GetLastError());
            }
            published = true;
            committed_ = true;
            published_path_ = publish_path;
            wchar_t injected_attribute_failure[2]{};
            const bool fail_attribute_restore_for_test =
                GetEnvironmentVariableW(L"XACTCOPY_DEV_FAIL_POST_PUBLISH_ATTRIBUTES",
                                        injected_attribute_failure,
                                        static_cast<DWORD>(std::size(injected_attribute_failure))) > 0;
            if (staged_readonly &&
                (fail_attribute_restore_for_test ||
                 !SetFileAttributesW(publish_path.c_str(), staged_attributes))) {
                // Publication is already durable and cannot be rolled back by
                // pretending this was a pre-commit failure. Surface the
                // metadata problem separately while keeping the byte result.
                post_publish_attribute_error_ = fail_attribute_restore_for_test
                                                    ? ERROR_ACCESS_DENIED
                                                    : GetLastError();
            }
        } catch (...) {
            if (!published && final_readonly &&
                GetFileAttributesW(publish_path.c_str()) != INVALID_FILE_ATTRIBUTES) {
                SetFileAttributesW(publish_path.c_str(), final_attributes);
            }
            if (!published && staged_readonly &&
                GetFileAttributesW(working_path_.c_str()) != INVALID_FILE_ATTRIBUTES) {
                SetFileAttributesW(working_path_.c_str(), staged_attributes);
            }
            throw;
        }
    }
    struct StageCopyState {
        const CancelContext* cancel;
    };

    static DWORD CALLBACK stage_copy_progress(
        LARGE_INTEGER, LARGE_INTEGER, LARGE_INTEGER, LARGE_INTEGER,
        DWORD, DWORD, HANDLE, HANDLE, LPVOID context) {
        auto* state = static_cast<StageCopyState*>(context);
        if (state != nullptr && state->cancel != nullptr &&
            (state->cancel->is_cancelled() || state->cancel->deadline_passed())) {
            return PROGRESS_CANCEL;
        }
        return PROGRESS_QUIET;
    }

    std::wstring final_path_;
    std::wstring working_path_;
    std::wstring lock_path_;
    std::wstring published_path_;
    DWORD post_publish_attribute_error_ = ERROR_SUCCESS;
    const CancelContext* cancel_ = nullptr;
    HANDLE lock_handle_ = INVALID_HANDLE_VALUE;
    bool committed_ = false;
};

inline bool is_generated_stage_file_name(std::wstring_view name) {
    constexpr std::wstring_view marker = L".xactcopy-stage.";
    constexpr std::size_t suffix_length = 32;
    const std::size_t marker_position = name.rfind(marker);
    if (marker_position == std::wstring_view::npos ||
        name.size() != marker_position + marker.size() + suffix_length) {
        return false;
    }
    for (wchar_t ch : name.substr(marker_position + marker.size())) {
        if (!((ch >= L'0' && ch <= L'9') || (ch >= L'a' && ch <= L'f'))) return false;
    }
    return true;
}

// A process crash can leave a staged file beside a destination. Stages have a
// per-file lock marker while an active copy owns them; cleanup only considers
// the generated name, ignores recent files, and probes the file exclusively
// before removing it. This keeps a concurrent job's live stage out of the
// cleanup sweep while reclaiming abandoned work on the next run.
inline void cleanup_abandoned_staging_files(
    const std::wstring& destination_root,
    const std::function<void(const std::string&)>& log = nullptr,
    const CancelContext* cancel = nullptr) {
    std::wstring normalized = storage::fsutil::get_full_path(destination_root);
    if (normalized.empty()) return;

    constexpr std::uint64_t stale_after_100ns = 60ULL * 60ULL * 10000000ULL;
    FILETIME now_filetime{};
    GetSystemTimeAsFileTime(&now_filetime);
    const std::uint64_t now = (static_cast<std::uint64_t>(now_filetime.dwHighDateTime) << 32) |
                              now_filetime.dwLowDateTime;

    std::vector<std::wstring> pending{normalized};
    while (!pending.empty()) {
        std::wstring current = std::move(pending.back());
        pending.pop_back();

        std::wstring search = detail::trim_trailing_separators(current) + L"\\*";
        WIN32_FIND_DATAW data{};
        HANDLE find = FindFirstFileExW(search.c_str(), FindExInfoBasic, &data,
                                       FindExSearchNameMatch, nullptr, FIND_FIRST_EX_LARGE_FETCH);
        if (find == INVALID_HANDLE_VALUE) continue;

        do {
            if (cancel != nullptr) {
                try {
                    cancel->throw_if_cancelled();
                } catch (...) {
                    FindClose(find);
                    throw;
                }
            }
            std::wstring name = data.cFileName;
            if (name.empty() || name == L"." || name == L"..") continue;
            std::wstring full = detail::trim_trailing_separators(current) + L"\\" + name;
            bool is_directory = (data.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) != 0;
            bool is_reparse = (data.dwFileAttributes & FILE_ATTRIBUTE_REPARSE_POINT) != 0;
            if (is_directory) {
                if (!is_reparse) pending.push_back(std::move(full));
                continue;
            }

            // Only the exact 128-bit lowercase-hex suffix generated by
            // DestinationTransaction is eligible. Merely containing the
            // marker must never make an ordinary user file disposable.
            if (!is_generated_stage_file_name(name)) continue;

            const std::uint64_t written =
                (static_cast<std::uint64_t>(data.ftLastWriteTime.dwHighDateTime) << 32) |
                data.ftLastWriteTime.dwLowDateTime;
            if (written == 0 || now <= written || now - written < stale_after_100ns) continue;
            std::wstring lock_path = full + L".lock";
            if (GetFileAttributesW(lock_path.c_str()) != INVALID_FILE_ATTRIBUTES) {
                // FILE_FLAG_DELETE_ON_CLOSE removes this marker on a normal
                // process exit. If power loss left only the directory entry,
                // an exclusive probe proves no live staging owner remains.
                HANDLE lock_probe = CreateFileW(lock_path.c_str(), GENERIC_READ | GENERIC_WRITE,
                                                0, nullptr, OPEN_EXISTING,
                                                FILE_ATTRIBUTE_NORMAL, nullptr);
                if (lock_probe == INVALID_HANDLE_VALUE) continue;
                CloseHandle(lock_probe);
                SetFileAttributesW(lock_path.c_str(), FILE_ATTRIBUTE_NORMAL);
                DeleteFileW(lock_path.c_str());
            }

            HANDLE probe = CreateFileW(full.c_str(), GENERIC_READ | GENERIC_WRITE, 0, nullptr,
                                       OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, nullptr);
            if (probe == INVALID_HANDLE_VALUE) continue;
            CloseHandle(probe);

            SetFileAttributesW(full.c_str(), FILE_ATTRIBUTE_NORMAL);
            if (DeleteFileW(full.c_str()) && log) {
                log("Removed abandoned destination stage: " + wide_to_utf8(full));
            }
        } while (FindNextFileW(find, &data));
        FindClose(find);
    }
}

struct BufferDecision {
    bool changed = false;
    std::int32_t previous_size = 0;
    std::int32_t current_size = 0;
    const char* reason = "";
};

// Port of AdaptiveBufferController: EWMA-throughput-driven chunk sizing with
// latency guards; static (min==max==configured) when adaptivity is off.
class AdaptiveBufferController {
public:
    static AdaptiveBufferController create(std::int64_t file_length, std::int32_t configured,
                                           bool adaptive_enabled, bool fragile_media,
                                           BufferPurpose purpose, std::int32_t worker_count) {
        std::int32_t manual = normalize(configured);
        if (!adaptive_enabled) {
            return AdaptiveBufferController(manual, manual, manual, false);
        }
        std::int32_t maximum = resolve_maximum(file_length, fragile_media, purpose, worker_count);
        std::int32_t minimum = std::min(AdaptiveAutoMinimumBufferSize, maximum);
        std::int32_t initial = resolve_initial(file_length, fragile_media, purpose, maximum);
        return AdaptiveBufferController(initial, minimum, maximum, true);
    }

    bool is_dynamic() const noexcept { return dynamic_; }
    std::int32_t current_size() const noexcept { return current_; }
    std::int32_t maximum_size() const noexcept { return maximum_; }

    std::int32_t next_chunk_length(std::int64_t remaining) const {
        if (remaining <= 0) return 0;
        return static_cast<std::int32_t>(
            std::min<std::int64_t>(current_, std::min<std::int64_t>(remaining, INT32_MAX)));
    }

    BufferDecision report_success(std::int32_t bytes_transferred, double elapsed_ms) {
        if (!dynamic_ || bytes_transferred <= 0) return {};

        double elapsed_seconds = std::max(0.000001, elapsed_ms / 1000.0);
        double observed = static_cast<double>(bytes_transferred) / elapsed_seconds;
        smoothed_ = smoothed_ <= 0 ? observed : (smoothed_ * 0.75) + (observed * 0.25);
        if (best_ <= 0 || smoothed_ > best_) best_ = smoothed_;

        samples_since_change_ += 1;
        bytes_since_change_ += bytes_transferred;
        bool enough = samples_since_change_ >= 3 ||
                      bytes_since_change_ >= static_cast<std::int64_t>(current_) * 4;
        if (!enough) return {};

        if (elapsed_ms > 1500.0 && current_ > minimum_) {
            return resize(std::max(minimum_, current_ / 2), "I/O latency increased");
        }
        if (current_ < maximum_ && elapsed_ms < 1000.0 && smoothed_ >= best_ * 0.92) {
            return resize(std::min(maximum_, current_ * 2), "throughput stayed stable");
        }
        samples_since_change_ = 0;
        bytes_since_change_ = 0;
        return {};
    }

    BufferDecision report_failure() {
        if (!dynamic_ || current_ <= minimum_) return {};
        return resize(std::max(minimum_, current_ / 2), "read reliability degraded");
    }

private:
    bool dynamic_;
    std::int32_t minimum_;
    std::int32_t maximum_;
    std::int32_t current_;
    double smoothed_ = 0.0;
    double best_ = 0.0;
    std::int32_t samples_since_change_ = 0;
    std::int64_t bytes_since_change_ = 0;

    AdaptiveBufferController(std::int32_t current, std::int32_t minimum, std::int32_t maximum,
                             bool dynamic) {
        minimum_ = normalize(std::min(minimum, maximum));
        maximum_ = normalize(std::max(minimum_, maximum));
        current_ = normalize(std::max(minimum_, std::min(current, maximum_)));
        dynamic_ = dynamic && maximum_ > minimum_;
    }

    static std::int32_t normalize(std::int32_t value) {
        return std::min(MaximumIoBufferSize, std::max(4096, value));
    }

    BufferDecision resize(std::int32_t new_size, const char* reason) {
        std::int32_t normalized = normalize(std::max(minimum_, std::min(new_size, maximum_)));
        samples_since_change_ = 0;
        bytes_since_change_ = 0;
        if (normalized == current_) return {};
        BufferDecision decision{true, current_, normalized, reason};
        current_ = normalized;
        return decision;
    }

    static std::int64_t align_up(std::int64_t value) {
        std::int64_t clamped = std::max<std::int64_t>(
            MinimumRescueBlockSize, std::min<std::int64_t>(MaximumIoBufferSize, value));
        return ((clamped + MinimumRescueBlockSize - 1) / MinimumRescueBlockSize) *
               MinimumRescueBlockSize;
    }

    static std::int32_t resolve_maximum(std::int64_t file_length, bool fragile,
                                        BufferPurpose purpose, std::int32_t worker_count) {
        std::int32_t maximum;
        switch (purpose) {
            case BufferPurpose::Copy:
                maximum = AdaptiveAutoMaxCopyBufferSize;
                break;
            case BufferPurpose::FastHealthScan: {
                std::int32_t safe_workers = std::max(1, worker_count);
                maximum = std::min(AdaptiveAutoMaxScanBufferSize,
                                   std::max(AdaptiveAutoMinimumBufferSize,
                                            AdaptiveAutoParallelScanBufferBudget / safe_workers));
                break;
            }
            default:
                maximum = AdaptiveAutoMaxScanBufferSize;
                break;
        }
        if (fragile) maximum = std::min(maximum, AdaptiveAutoMaxFragileBufferSize);
        if (file_length > 0) {
            maximum = static_cast<std::int32_t>(std::min<std::int64_t>(
                maximum,
                std::max<std::int64_t>(AdaptiveAutoMinimumBufferSize, align_up(file_length))));
        }
        return normalize(maximum);
    }

    static std::int32_t resolve_initial(std::int64_t file_length, bool fragile,
                                        BufferPurpose purpose, std::int32_t maximum) {
        std::int64_t target;
        if (fragile) {
            target = AdaptiveAutoInitialSmallBufferSize;
        } else if (file_length > 0 && file_length <= AdaptiveAutoInitialSmallBufferSize) {
            target = AdaptiveAutoMinimumBufferSize;
        } else if (file_length > 0 && file_length <= 8LL * 1024 * 1024) {
            target = std::min<std::int64_t>(file_length, 1024 * 1024);
        } else if (purpose == BufferPurpose::FastHealthScan) {
            target = AdaptiveAutoInitialFastScanBufferSize;
        } else {
            target = AdaptiveAutoInitialCopyBufferSize;
        }
        return normalize(static_cast<std::int32_t>(std::min<std::int64_t>(maximum, align_up(target))));
    }
};

class ResilientCopyEngine {
public:
    using ProgressCallback = std::function<void(const CopyProgressSnapshot&)>;
    using LogCallback = std::function<void(const std::string&)>;

    ResilientCopyEngine(CopyJobOptions options, ExecutionControl* control,
                        ProgressCallback progress, LogCallback log)
        : options_(std::move(options)),
          control_(control),
          progress_callback_(std::move(progress)),
          log_callback_(std::move(log)),
          fault_injector_(FaultInjector::try_create_from_environment()) {}

    CopyJobResult run(const std::atomic<bool>& cancel_flag) {
        CancelContext cancel{&cancel_flag, 0};
        validate_options();
        run_started_tick_ = GetTickCount64();
        native_fast_path_files_ = 0;
        parallel_native_fast_path_files_ = 0;
        managed_copy_files_ = 0;
        native_fallback_files_ = 0;
        created_directories_.clear();
        fragile_failure_timestamps_.clear();
        integrity_warning_.clear();
        metadata_notice_.clear();
        source_enumeration_incomplete_ = false;
        source_enumeration_error_.clear();
        raw_disk_scan_context_.reset();
        raw_scan_fallback_logged_paths_.clear();
        throttle_window_start_tick_ = 0;
        throttle_window_bytes_ = 0;
        bytes_read_.store(0);
        bytes_written_.store(0);
        bytes_verified_.store(0);

        if (options_.OperationMode == models::JobOperationMode::Copy) {
            const models::VerificationMode verification = resolve_verification_mode();
            if (verification == models::VerificationMode::None) {
                add_integrity_warning(
                    "Destination content verification was disabled for this attended run.");
            } else if (verification == models::VerificationMode::Sampled) {
                add_integrity_warning(
                    "Only sampled destination verification was requested; corruption outside sampled ranges may be missed.");
            }
        }

        if (fault_injector_.has_value()) {
            emit_log("[DevFault] Enabled: " + fault_injector_->description());
        }

        const bool scan_only = options_.OperationMode == models::JobOperationMode::ScanOnly;
        std::wstring source_root = storage::fsutil::get_full_path(utf8_to_wide(options_.SourceRoot));
        std::wstring destination_root = storage::fsutil::get_full_path(utf8_to_wide(options_.DestinationRoot));

        initialize_media_identity_expectations(source_root, destination_root);
        wait_for_media_availability(source_root, destination_root, cancel);
        ensure_media_identity_integrity(source_root, destination_root, !scan_only, true);
        validate_resolved_root_relationship(source_root, destination_root, scan_only);
        emit_media_identity_snapshot();
        initialize_raw_disk_scan_context(scan_only, source_root);
        if (!scan_only) {
            storage::fsutil::create_directories(destination_root);
            cleanup_abandoned_staging_files(
                destination_root,
                [this](const std::string& message) { emit_log(message); }, &cancel);
        }

        initialize_bad_range_map(source_root);

        emit_log("Scanning source: " + wide_to_utf8(source_root));
        auto log_fn = [this](const std::string& text) { emit_log(text); };
        SourceScanResult scan = scan_source(source_root, options_.SelectedRelativePaths,
                                            options_.SymlinkHandling, options_.CopyEmptyDirectories, log_fn,
                                            [&cancel] { return cancel.is_cancelled(); }, nullptr,
                                            scan_only ? std::wstring() : destination_root);
        if (!scan.complete) {
            source_enumeration_incomplete_ = true;
            source_enumeration_error_ = "Source enumeration was incomplete";
            if (!scan.errors.empty()) {
                source_enumeration_error_ += ": " + scan.errors.front();
            }
            emit_log("ERROR: " + source_enumeration_error_ + ".");
            if (!options_.ContinueOnFileError) {
                throw IoError(source_enumeration_error_ + ".");
            }
            emit_log("Continue on error is enabled; the partial source listing will be copied, "
                     "but the run cannot be reported as an exact success.");
        }
        std::int64_t total_bytes = 0;
        for (const auto& file : scan.files) total_bytes += file.length;
        const std::int32_t parallel_small_file_workers =
            resolve_parallel_small_file_workers(static_cast<std::int32_t>(scan.files.size()));
        emit_log("Scan complete. " + std::to_string(scan.files.size()) + " file(s), " +
                 format_bytes(total_bytes) + " total.");
        if (scan_only) {
            emit_log("Operation mode: Read-only allocated-file readability assessment.");
            emit_log("Scan performance profile: " +
                     std::string(models::to_string(resolve_scan_performance_profile())) +
                     "; workers=" +
                     std::to_string(resolve_parallel_scan_workers(
                         static_cast<std::int32_t>(scan.files.size()))) + ".");
        } else {
            emit_destination_capacity_warning(destination_root, total_bytes);
            emit_log("Transfer engine policy: " + std::string(models::to_string(options_.TransferEnginePolicyValue)) + ".");
            if (options_.TransferEnginePolicyValue == models::TransferEnginePolicy::NativeFast &&
                options_.MaxRetries > 0) {
                emit_log("NativeFast performs one CopyFileEx attempt per file; managed retry settings "
                         "apply only after native fallback.");
            }
            emit_log("Small-file acceleration: workers=" + std::to_string(parallel_small_file_workers) +
                     ", threshold=" + format_bytes(options_.SmallFileThresholdBytes) + ".");
        }

        if (!scan_only && options_.CopyEmptyDirectories && !scan.directories.empty()) {
            emit_log("Preparing " + std::to_string(scan.directories.size()) + " destination directory path(s).");
            for (const auto& relative_dir : scan.directories) {
                cancel.throw_if_cancelled();
                if (control_ != nullptr) control_->wait_if_paused(cancel);
                wait_for_media_availability(source_root, destination_root, cancel);
                storage::fsutil::create_directories(
                    detail::trim_trailing_separators(destination_root) + L"\\" + utf8_to_wide(relative_dir));
            }
        }

        std::string job_id = storage::JobJournalStore::build_job_id(source_root, destination_root);
        std::wstring default_journal_path = storage::JobJournalStore::get_default_journal_path(job_id);
        std::wstring journal_path = resolve_writable_journal_path(job_id, default_journal_path);

        std::optional<JobJournal> existing;
        if (options_.ResumeFromJournal) {
            emit_log("Loading resume journal candidates.");
            for (const auto& candidate : journal_load_candidates(default_journal_path, journal_path)) {
                existing = journal_store_.load(candidate);
                if (existing.has_value()) {
                    emit_log("Loaded journal: " + wide_to_utf8(candidate));
                    break;
                }
            }
            if (!existing.has_value()) emit_log("No journal found. Creating a fresh state.");
        } else {
            emit_log("Resume disabled. Starting with a fresh state.");
        }

        emit_log("Preparing journal state for " + std::to_string(scan.files.size()) + " source file(s).");
        JobJournal journal = merge_journal(existing, scan.files,
                                           wide_to_utf8(source_root), wide_to_utf8(destination_root), job_id,
                                           cancel);
        emit_log("Journal state prepared.");

        // A fresh fast scan can contain tens of thousands of files. Do not
        // block worker startup on an all-pending snapshot. run_fast_scan_mode
        // owns a synchronized live checkpoint writer, so its first durable
        // snapshot already contains useful completed/partial coverage.
        const bool defer_initial_scan_checkpoint =
            scan_only && resolve_scan_performance_profile() == models::ScanPerformanceProfile::Fast &&
            scan.files.size() >= 2048;
        if (defer_initial_scan_checkpoint) {
            emit_log("Large fast scan: initial checkpoint deferred to the live progress writer.");
        } else {
            emit_log("Saving initial journal checkpoint.");
            save_journal_now(journal_path, journal);
            emit_log("Initial journal checkpoint saved.");
        }

        ProgressAccumulator progress;
        progress.total_files = static_cast<std::int32_t>(scan.files.size());
        progress.total_bytes = total_bytes;

        // Fast scan profile: parallel healthy-file reads with precise fallback.
        if (scan_only &&
            resolve_scan_performance_profile() == models::ScanPerformanceProfile::Fast) {
            std::string fast_error = run_fast_scan_mode(
                scan.files, source_root, progress, journal, journal_path, cancel);
            flush_bad_range_map();
            save_journal_now(journal_path, journal);
            CopyJobResult fast_result = create_result(
                progress, journal_path,
                fast_error.empty() && progress.failed_files == 0, false, fast_error);
            emit_run_summary(fast_result);
            return fast_result;
        }

        // Parallel small-file bulk phase (copy mode only; mirrors the .NET
        // eligibility gates — sequential loop later treats these as completed).
        std::vector<std::string> parallel_completed;
        if (!scan_only && options_.TransferEnginePolicyValue == models::TransferEnginePolicy::NativeFast &&
            is_native_acceleration_allowed() && parallel_small_file_workers > 1 &&
            !options_.FragileMediaMode && !options_.WaitForFileLockRelease &&
            options_.MaxThroughputBytesPerSecond <= 0 &&
            // The parallel path has no managed rescue loop. Require a full
            // content check before publication so NativeFast parallelism
            // cannot turn an unchecked CopyFileExW result into a clean copy.
            resolve_verification_mode() == models::VerificationMode::Full &&
            !(options_.UseBadRangeMap && options_.SkipKnownBadRanges)) {

            std::vector<SourceFileDescriptor> eligible;
            std::int64_t threshold = std::max(MinimumRescueBlockSize, options_.SmallFileThresholdBytes);
            for (const auto& candidate : scan.files) {
                if (candidate.length <= 0 || candidate.length > threshold) continue;
                JournalFileEntry* candidate_entry = journal.Files.find(candidate.relative_path);
                if (candidate_entry == nullptr) continue;
                if (is_already_completed(*candidate_entry, candidate, destination_root, cancel)) continue;
                if (should_skip_failed_entry_for_fragile_resume(*candidate_entry)) continue;
                std::string parallel_skip_reason;
                if (decide_overwrite_policy(candidate, destination_root, parallel_skip_reason) !=
                    ExistingDestinationDecision::Copy) {
                    continue;
                }
                eligible.push_back(candidate);
            }
            if (eligible.size() > 1) {
                parallel_completed = copy_small_files_parallel(eligible, destination_root, journal,
                                                               journal_path, cancel,
                                                               parallel_small_file_workers);
            }
        }
        auto take_parallel_completed = [&parallel_completed](const std::string& relative) {
            for (std::size_t i = 0; i < parallel_completed.size(); ++i) {
                if (models::detail::equals_ignore_case(parallel_completed[i], relative)) {
                    parallel_completed.erase(parallel_completed.begin() +
                                             static_cast<std::ptrdiff_t>(i));
                    return true;
                }
            }
            return false;
        };

        for (const auto& descriptor : scan.files) {
            cancel.throw_if_cancelled();
            if (control_ != nullptr) control_->wait_if_paused(cancel);
            wait_for_media_availability(source_root, destination_root, cancel);
            ensure_media_identity_integrity(source_root, destination_root, !scan_only, false);

            JournalFileEntry* entry = journal.Files.find(descriptor.relative_path);
            if (entry == nullptr) continue; // merge_journal guarantees presence

            // Overwrite/existing-destination semantics are copy-only; scan
            // mode always attempts source reads to discover bad ranges.
            if (!scan_only) {
                std::string overwrite_reason;
                ExistingDestinationDecision overwrite_decision =
                    decide_overwrite_policy(descriptor, destination_root, overwrite_reason);
                if (overwrite_decision == ExistingDestinationDecision::Reject) {
                    progress.failed_files += 1;
                    entry->State = FileCopyState::Failed;
                    entry->LastError = overwrite_reason;
                    entry->DoNotRetry = false;
                    emit_log("Failed: " + descriptor.relative_path + " (" + overwrite_reason + ")");
                    flush_journal(journal_path, journal, true);
                    if (!options_.ContinueOnFileError) {
                        save_journal_now(journal_path, journal);
                        return create_result(progress, journal_path, false, false, overwrite_reason);
                    }
                    continue;
                }
                if (overwrite_decision == ExistingDestinationDecision::Skip) {
                    progress.skipped_files += 1;
                    progress.total_bytes_copied += descriptor.length;
                    progress.bytes_skipped += descriptor.length;
                    emit_log("Skipped: " + descriptor.relative_path + " (" + overwrite_reason + ")");
                    emit_progress(progress, descriptor.relative_path, descriptor.length,
                                  descriptor.length, 0,
                                  resolve_buffer_size_for_file(descriptor.length));
                    continue;
                }

                if (take_parallel_completed(descriptor.relative_path)) {
                    progress.completed_files += 1;
                    progress.total_bytes_copied += descriptor.length;
                    emit_progress(progress, descriptor.relative_path, descriptor.length,
                                  descriptor.length, 0,
                                  resolve_buffer_size_for_file(descriptor.length),
                                  "ParallelNativeFast", 0, 0);
                    continue;
                }

                if (is_already_completed(*entry, descriptor, destination_root, cancel)) {
                    // A journal reuse is an exact, full-validated completion,
                    // not an omitted source file. Keep it in CompletedFiles
                    // so a clean resume remains a clean success; policy
                    // skips (SkipExisting/IfSourceNewer) are handled above.
                    progress.completed_files += 1;
                    progress.total_bytes_copied += descriptor.length;
                    progress.bytes_reused += descriptor.length;
                    emit_log("Reused: " + descriptor.relative_path + " (verified journal completion).");
                    emit_progress(progress, descriptor.relative_path, descriptor.length,
                                  descriptor.length, 0,
                                  resolve_buffer_size_for_file(descriptor.length),
                                  entry->LastRescuePass, unreadable_region_count(entry->RescueRanges),
                                  rescue_bytes(entry->RescueRanges,
                                               {RescueRangeState::Pending, RescueRangeState::Bad,
                                                RescueRangeState::KnownBad}));
                    continue;
                }
            }

            if (scan_only) {
                if (is_completed_scan_journal_entry(*entry, descriptor)) {
                    progress.completed_files += 1;
                    if (entry->State == FileCopyState::CompletedWithRecovery) {
                        progress.recovered_files += 1;
                    }
                    progress.total_bytes_copied += descriptor.length;
                    progress.bytes_reused += descriptor.length;
                    emit_log("Reused assessment: " + descriptor.relative_path +
                             " (bound journal completion).");
                    emit_progress(
                        progress, descriptor.relative_path, descriptor.length,
                        descriptor.length, 0,
                        resolve_buffer_size_for_file(descriptor.length),
                        "ScanResume", unreadable_region_count(entry->RescueRanges),
                        rescue_bytes(entry->RescueRanges,
                                     {RescueRangeState::Pending, RescueRangeState::Bad,
                                      RescueRangeState::KnownBad}));
                    continue;
                }

                if (should_skip_failed_entry_for_fragile_resume(*entry)) {
                    progress.skipped_files += 1;
                    std::int64_t scan_already = entry->BytesCopied > 0 ? entry->BytesCopied : 0;
                    std::int64_t scan_remaining = descriptor.length - scan_already;
                    const std::int64_t skipped = scan_remaining > 0 ? scan_remaining : 0;
                    progress.total_bytes_copied += skipped;
                    progress.bytes_skipped += skipped;
                    emit_log("Skipped scan: " + descriptor.relative_path + " (persisted fragile skip).");
                    emit_progress(progress, descriptor.relative_path, descriptor.length,
                                  descriptor.length, 0,
                                  resolve_buffer_size_for_file(descriptor.length));
                    continue;
                }

                if (entry->State == FileCopyState::Failed) {
                    entry->State = FileCopyState::Pending;
                    entry->LastError.clear();
                    entry->DoNotRetry = false;
                }

                const bool preserve_scan_coverage =
                    options_.ResumeFromJournal && !entry->RescueRanges.empty() &&
                    entry->SourceChangeUtcTicks != 0 && entry->SourceFileIndex != 0 &&
                    entry->SourceVolumeSerial != 0;
                std::string scan_error = process_precise_scan_file(
                    descriptor, *entry, progress, journal, journal_path,
                    preserve_scan_coverage, /*seed_progress*/ true, cancel);
                if (!scan_error.empty() && !options_.ContinueOnFileError) {
                    save_journal_now(journal_path, journal);
                    return create_result(progress, journal_path, false, false, scan_error);
                }
                continue;
            }

            if (should_skip_failed_entry_for_fragile_resume(*entry)) {
                progress.skipped_files += 1;
                std::int64_t remaining = descriptor.length - (entry->BytesCopied > 0 ? entry->BytesCopied : 0);
                const std::int64_t skipped = remaining > 0 ? remaining : 0;
                progress.total_bytes_copied += skipped;
                progress.bytes_skipped += skipped;
                emit_log("Skipped: " + descriptor.relative_path + " (persisted fragile skip).");
                emit_progress(progress, descriptor.relative_path, descriptor.length, descriptor.length, 0,
                              resolve_buffer_size_for_file(descriptor.length));
                continue;
            }

            if (entry->State == FileCopyState::Failed) {
                entry->State = FileCopyState::Pending;
                entry->LastError.clear();
                entry->DoNotRetry = false;
            }

            std::optional<std::string> copy_error;
            bool recovered = false;

            CancelContext file_cancel = cancel;
            if (options_.PerFileTimeout.ticks > 0) {
                file_cancel = cancel.with_deadline(
                    options_.PerFileTimeout.ticks / time::TicksPerMillisecond);
            }

            try {
                entry->State = FileCopyState::InProgress;
                entry->LastError.clear();
                entry->DoNotRetry = false;
                flush_journal(journal_path, journal, true);

                recovered = copy_single_file(descriptor, *entry, destination_root, progress,
                                             journal, journal_path, file_cancel);
            } catch (const SourceMutationSkipped& ex) {
                progress.skipped_files += 1;
                progress.total_bytes_copied += descriptor.length;
                progress.bytes_skipped += descriptor.length;
                entry->State = FileCopyState::Failed;
                entry->LastError = ex.what();
                entry->DoNotRetry = false;
                emit_log("Skipped: " + descriptor.relative_path + " (" + ex.what() + ")");
                flush_journal(journal_path, journal, true);
                continue;
            } catch (const FragileReadSkip& ex) {
                handle_fragile_read_skip(descriptor, *entry, progress, journal, journal_path,
                                         ex, cancel);
                continue;
            } catch (const OperationCanceled& oc) {
                if (oc.user_requested) throw;
                std::int64_t timeout_seconds =
                    (options_.PerFileTimeout.ticks / time::TicksPerSecond);
                copy_error = "Per-file timeout (" + std::to_string(timeout_seconds) +
                             " sec) reached while copying " + descriptor.relative_path + ".";
            } catch (const std::exception& ex) {
                copy_error = ex.what();
            }

            if (!copy_error.has_value()) {
                progress.completed_files += 1;
                if (recovered) progress.recovered_files += 1;
                entry->State = recovered ? FileCopyState::CompletedWithRecovery : FileCopyState::Completed;
                entry->BytesCopied = descriptor.length;
                entry->LastError.clear();
                entry->DoNotRetry = false;
                emit_progress(progress, descriptor.relative_path, descriptor.length, descriptor.length, 0,
                              resolve_buffer_size_for_file(descriptor.length), entry->LastRescuePass,
                              unreadable_region_count(entry->RescueRanges),
                              rescue_bytes(entry->RescueRanges, {RescueRangeState::Pending, RescueRangeState::Bad, RescueRangeState::KnownBad}));
                try_persist_bad_range_map_entry(descriptor, *entry, true);
                flush_journal(journal_path, journal, true);
                continue;
            }

            progress.failed_files += 1;
            entry->State = FileCopyState::Failed;
            entry->LastError = *copy_error;
            entry->DoNotRetry = false;
            emit_log("Failed: " + descriptor.relative_path + " (" + *copy_error + ")");
            register_fragile_failure_and_maybe_cooldown(descriptor.relative_path, *copy_error, cancel);
            try_persist_bad_range_map_entry(descriptor, *entry, true);
            flush_journal(journal_path, journal, true);

            if (!options_.ContinueOnFileError) {
                save_journal_now(journal_path, journal);
                return create_result(progress, journal_path, false, false, *copy_error);
            }
        }

        std::string directory_metadata_error;
        if (!scan_only && options_.CopyEmptyDirectories && !scan.directories.empty()) {
            directory_metadata_error = preserve_directory_metadata(
                source_root, destination_root, scan.directories, cancel);
            if (!directory_metadata_error.empty()) {
                emit_log("ERROR: " + directory_metadata_error);
            }
        }
        flush_bad_range_map();
        save_journal_now(journal_path, journal);
        CopyJobResult final_result = create_result(progress, journal_path,
                                                   progress.failed_files == 0 &&
                                                       directory_metadata_error.empty(),
                                                   false, directory_metadata_error);
        emit_run_summary(final_result);
        return final_result;
    }

private:
    struct ProgressAccumulator {
        std::int32_t total_files = 0;
        std::int32_t completed_files = 0;
        std::int32_t failed_files = 0;
        std::int32_t recovered_files = 0;
        std::int32_t skipped_files = 0;
        std::int64_t total_bytes = 0;
        std::int64_t total_bytes_copied = 0;
        std::int64_t bytes_skipped = 0;
        std::int64_t bytes_reused = 0;
    };

    // Captures the source hash during an ordinary sequential managed read. If
    // resume coverage or a rescue pass makes reads non-sequential, validation
    // falls back to hashing the source again rather than trusting a partial
    // digest. Healthy managed copies therefore avoid a second source pass.
    class CopyVerificationTracker {
    public:
        explicit CopyVerificationTracker(bool sha512) : hasher_(sha512) {}

        void observe(std::int64_t offset, const unsigned char* data, std::int32_t count) {
            if (!valid_ || count <= 0) return;
            if (offset != next_offset_) {
                valid_ = false;
                return;
            }
            hasher_.update(data, static_cast<std::size_t>(count));
            next_offset_ += count;
        }

        void invalidate() noexcept { valid_ = false; }

        std::optional<std::vector<unsigned char>> finish(std::int64_t expected_length) {
            if (!valid_ || next_offset_ != expected_length) return std::nullopt;
            return hasher_.finish();
        }

    private:
        crypto::detail::StreamingHash hasher_;
        std::int64_t next_offset_ = 0;
        bool valid_ = true;
    };

    struct RescuePassDefinition {
        std::string name;
        RescueRangeState target_state = RescueRangeState::Pending;
        std::int32_t chunk_size_bytes = 0;
        std::int32_t max_read_retries = 0;
        bool split_on_failure = false;
        std::int32_t minimum_split_bytes = 0;
        bool process_descending = false;
    };

    struct RescuePassOutcome {
        std::int64_t attempted_segments = 0;
        std::int64_t recovered_bytes = 0;
        std::int64_t failed_segments = 0;
    };

    CopyJobOptions options_;
    ExecutionControl* control_;
    ProgressCallback progress_callback_;
    LogCallback log_callback_;
    std::optional<FaultInjector> fault_injector_;
    std::unique_ptr<RawDiskScanContext> raw_disk_scan_context_;
    storage::JobJournalStore journal_store_;
    storage::BadRangeMapStore bad_range_map_store_;

    ULONGLONG run_started_tick_ = 0;
    ULONGLONG last_journal_flush_tick_ = 0;
    ULONGLONG last_journal_save_cost_ms_ = 0;
    ULONGLONG throttle_window_start_tick_ = 0;
    std::int64_t throttle_window_bytes_ = 0;
    std::string expected_source_identity_;
    std::string expected_destination_identity_;
    std::mutex media_identity_lock_;
    std::atomic<ULONGLONG> last_source_identity_probe_tick_{0};
    std::atomic<ULONGLONG> last_destination_identity_probe_tick_{0};
    std::atomic<ULONGLONG> last_source_mismatch_log_tick_{0};
    std::atomic<ULONGLONG> last_destination_mismatch_log_tick_{0};
    std::atomic<bool> source_media_identity_accepted_{true};
    std::atomic<bool> destination_media_identity_accepted_{true};
    std::vector<std::wstring> created_directories_;
    std::deque<ULONGLONG> fragile_failure_timestamps_;
    std::atomic<std::int32_t> native_fast_path_files_{0};
    std::atomic<std::int32_t> parallel_native_fast_path_files_{0};
    std::atomic<std::int32_t> managed_copy_files_{0};
    std::atomic<std::int32_t> native_fallback_files_{0};
    std::atomic<std::int64_t> bytes_read_{0};
    std::atomic<std::int64_t> bytes_written_{0};
    std::atomic<std::int64_t> bytes_verified_{0};
    std::mt19937 retry_jitter_{0xC0FFEE};
    std::mutex retry_jitter_lock_;
    bool source_enumeration_incomplete_ = false;
    std::string source_enumeration_error_;
    std::mutex metadata_warning_lock_;
    std::string integrity_warning_;
    std::string metadata_notice_;

    // Bad-range map state (InitializeBadRangeMapAsync et al).
    std::optional<storage::BadRangeMap> bad_range_map_;
    std::wstring bad_range_map_path_;
    bool bad_range_map_loaded_ = false;
    bool bad_range_map_read_hints_enabled_ = false;
    ULONGLONG last_bad_range_map_flush_tick_ = 0;

    // Shared-state locks for fast-scan workers / parallel small-file phase.
    std::mutex journal_lock_;
    std::mutex progress_lock_;
    std::mutex map_lock_;
    std::mutex raw_scan_fallback_lock_;
    std::unordered_set<std::string> raw_scan_fallback_logged_paths_;

    void initialize_raw_disk_scan_context(bool scan_only, const std::wstring& source_root) {
        raw_disk_scan_context_.reset();
        if (!scan_only || !options_.UseExperimentalRawDiskScan) return;

        std::string reason;
        try {
            auto context = RawDiskScanContext::try_create(source_root, reason);
            if (context) {
                emit_log("Assessment backend: Raw volume direct reads enabled. Sector size " +
                         format_bytes(context->sector_size_bytes()) + ", cluster size " +
                         format_bytes(context->cluster_size_bytes()) + ".");
                raw_disk_scan_context_ = std::move(context);
                return;
            }
        } catch (const std::exception& ex) {
            reason = ex.what();
        }

        if (reason.empty()) reason = "unsupported source or media.";
        emit_log("Assessment backend: Raw volume unavailable (" + reason +
                 "); using standard file reads.");
    }

    void emit_raw_disk_fallback_once(const std::wstring& source_path, const std::string& reason) {
        std::string key = wide_to_utf8(storage::fsutil::get_full_path(source_path));
        if (key.empty()) key = wide_to_utf8(source_path);
        if (key.empty()) key = "(unknown)";
        std::string folded = key;
        for (char& c : folded) {
            if (c >= 'A' && c <= 'Z') c = static_cast<char>(c - 'A' + 'a');
        }
        std::lock_guard<std::mutex> guard(raw_scan_fallback_lock_);
        if (!raw_scan_fallback_logged_paths_.insert(folded).second) return;
        std::string detail = reason.empty() ? "unsupported source layout" : reason;
        if (detail.empty() || detail.back() != '.') detail.push_back('.');
        emit_log("Raw-volume assessment fallback for '" + key + "': " + detail);
    }

    // ---- Options validation ----------------------------------------------

    void validate_options() {
        auto fail = [](const std::string& message) { throw IoError(message); };
        if (is_blank(options_.SourceRoot)) fail("Source path is required.");
        const bool scan_only = options_.OperationMode == models::JobOperationMode::ScanOnly;
        if (!scan_only && is_blank(options_.DestinationRoot)) fail("Destination path is required.");
        if (scan_only && is_blank(options_.DestinationRoot)) options_.DestinationRoot = options_.SourceRoot;

        if (!options_.WaitForMediaAvailability &&
            !directory_exists(utf8_to_wide(options_.SourceRoot))) {
            fail("Source directory not found: " + options_.SourceRoot);
        }

        std::wstring source_norm = storage::fsutil::to_upper_invariant(
            detail::trim_trailing_separators(storage::fsutil::get_full_path(utf8_to_wide(options_.SourceRoot))));
        std::wstring dest_norm = storage::fsutil::to_upper_invariant(
            detail::trim_trailing_separators(storage::fsutil::get_full_path(utf8_to_wide(options_.DestinationRoot))));
        if (!scan_only && source_norm == dest_norm) {
            fail("Source and destination cannot be the same path.");
        }
        if (!scan_only && dest_norm.size() > source_norm.size() &&
            dest_norm.compare(0, source_norm.size(), source_norm) == 0 &&
            dest_norm[source_norm.size()] == L'\\') {
            fail("Destination cannot be inside the source tree; choose a sibling or separate volume.");
        }
        if (options_.AllowJournalRootRemap) {
            fail("AllowJournalRootRemap is disabled because it can apply coverage to unrelated data.");
        }
        if (options_.TreatAccessDeniedAsContention && options_.SalvageUnreadableBlocks) {
            fail("TreatAccessDeniedAsContention cannot be combined with SalvageUnreadableBlocks; "
                 "permission failures must remain visible.");
        }

        if (options_.BufferSizeBytes < CopyJobOptions::MinimumBufferSizeBytes) {
            fail("BufferSizeBytes must be at least 1 MiB.");
        }
        if (options_.BufferSizeBytes > CopyJobOptions::MaximumBufferSizeBytes) {
            fail("BufferSizeBytes cannot exceed 256 MiB.");
        }
        if (options_.MaxRetries < 0) fail("MaxRetries cannot be negative.");
        if (options_.MaxRetries > CopyJobOptions::MaximumRetries) {
            fail("MaxRetries cannot exceed the safety limit of 32.");
        }
        if (options_.OperationTimeout.ticks <= 0) fail("OperationTimeout must be greater than zero.");
        if (options_.OperationTimeout.total_seconds() >
            CopyJobOptions::MaximumOperationTimeoutSeconds) {
            fail("OperationTimeout cannot exceed 3600 seconds.");
        }
        if (options_.PerFileTimeout.ticks < 0) fail("PerFileTimeout cannot be negative.");
        if (options_.PerFileTimeout.total_seconds() >
            CopyJobOptions::MaximumPerFileTimeoutSeconds) {
            fail("PerFileTimeout cannot exceed 86400 seconds.");
        }
        if (options_.InitialRetryDelay.ticks <= 0) {
            fail("InitialRetryDelay must be greater than zero.");
        }
        if (options_.MaxRetryDelay.ticks < options_.InitialRetryDelay.ticks) {
            fail("MaxRetryDelay cannot be shorter than InitialRetryDelay.");
        }
        if (options_.MaxRetryDelay.total_seconds() >
            CopyJobOptions::MaximumRetryDelaySeconds) {
            fail("MaxRetryDelay cannot exceed 86400 seconds.");
        }
        const double lock_probe_ms = options_.LockContentionProbeInterval.total_milliseconds();
        if (lock_probe_ms < CopyJobOptions::MinimumLockProbeIntervalMilliseconds ||
            lock_probe_ms > CopyJobOptions::MaximumLockProbeIntervalMilliseconds) {
            fail("LockContentionProbeInterval must be from 100 to 10000 milliseconds.");
        }
        if (options_.MaxThroughputBytesPerSecond < 0) fail("MaxThroughputBytesPerSecond cannot be negative.");
        if (options_.SampleVerificationChunkBytes <= 0) fail("SampleVerificationChunkBytes must be greater than zero.");
        if (options_.SampleVerificationChunkBytes >
            CopyJobOptions::MaximumSampleVerificationChunkBytes) {
            fail("SampleVerificationChunkBytes cannot exceed 4 MiB.");
        }
        if (options_.SampleVerificationChunkCount <= 0) fail("SampleVerificationChunkCount must be greater than zero.");
        if (options_.SampleVerificationChunkCount >
            CopyJobOptions::MaximumSampleVerificationChunkCount) {
            fail("SampleVerificationChunkCount cannot exceed 64.");
        }
        if (options_.BadRangeMapMaxAgeDays < 0) fail("BadRangeMapMaxAgeDays cannot be negative.");
        if (options_.BadRangeMapMaxAgeDays > CopyJobOptions::MaximumBadRangeMapAgeDays) {
            fail("BadRangeMapMaxAgeDays cannot exceed 3650.");
        }
        if (options_.RescueFastScanChunkBytes < 0 || options_.RescueTrimChunkBytes < 0 ||
            options_.RescueScrapeChunkBytes < 0 || options_.RescueRetryChunkBytes < 0 ||
            options_.RescueSplitMinimumBytes < 0) {
            fail("Rescue chunk/split tuning values cannot be negative.");
        }
        if (options_.RescueFastScanChunkBytes > CopyJobOptions::MaximumRescueChunkBytes ||
            options_.RescueTrimChunkBytes > CopyJobOptions::MaximumRescueChunkBytes ||
            options_.RescueScrapeChunkBytes > CopyJobOptions::MaximumRescueChunkBytes ||
            options_.RescueRetryChunkBytes > CopyJobOptions::MaximumRescueChunkBytes ||
            options_.RescueSplitMinimumBytes > CopyJobOptions::MaximumRescueSplitBytes) {
            fail("Rescue chunk/split tuning exceeds the supported memory bounds.");
        }
        if (options_.RescueFastScanRetries < 0 || options_.RescueTrimRetries < 0 ||
            options_.RescueScrapeRetries < 0) {
            fail("Rescue pass retries cannot be negative.");
        }
        if (options_.RescueFastScanRetries > 32 || options_.RescueTrimRetries > 32 ||
            options_.RescueScrapeRetries > 32) {
            fail("Rescue pass retries cannot exceed the safety limit of 32.");
        }
        if (options_.SalvageUnreadableBlocks &&
            options_.SalvageFillPatternValue == models::SalvageFillPattern::Random) {
            fail("Random salvage fill is disabled because non-deterministic bytes obscure recovery boundaries. Use zero or 0xFF fill.");
        }
        if (options_.ParallelSmallFileWorkers < 0) fail("ParallelSmallFileWorkers cannot be negative.");
        if (options_.ParallelScanWorkers < 0) fail("ParallelScanWorkers cannot be negative.");
        if (options_.ParallelSmallFileWorkers > CopyJobOptions::MaximumParallelWorkers ||
            options_.ParallelScanWorkers > CopyJobOptions::MaximumParallelWorkers) {
            fail("Parallel worker counts cannot exceed 64.");
        }
        if (options_.SmallFileThresholdBytes < MinimumRescueBlockSize) fail("SmallFileThresholdBytes must be at least 4096.");
        if (options_.SmallFileThresholdBytes >
            CopyJobOptions::MaximumSmallFileThresholdBytes) {
            fail("SmallFileThresholdBytes cannot exceed 1 GiB.");
        }
        if (options_.FragileFailureWindowSeconds <= 0) fail("FragileFailureWindowSeconds must be greater than zero.");
        if (options_.FragileFailureThreshold <= 0) fail("FragileFailureThreshold must be greater than zero.");
        if (options_.FragileCooldownSeconds < 0) fail("FragileCooldownSeconds cannot be negative.");
        if (options_.FragileFailureWindowSeconds >
                CopyJobOptions::MaximumFragileFailureWindowSeconds ||
            options_.FragileFailureThreshold >
                CopyJobOptions::MaximumFragileFailureThreshold ||
            options_.FragileCooldownSeconds >
                CopyJobOptions::MaximumFragileCooldownSeconds) {
            fail("Fragile-media guard values exceed the supported bounds.");
        }

        if (options_.MaxRetries > CopyJobOptions::MaximumNormalRetries) {
            emit_log("WARNING: application retry count exceeds 3; each attempt is additional to any device or Windows storage-stack retries.");
        }
        if (options_.ParallelScanWorkers > 8 && !options_.FragileMediaMode) {
            emit_log("WARNING: more than 8 scan workers can sharply increase seek pressure on unstable media.");
        }
        if (options_.SalvageUnreadableBlocks && options_.AllowRecoveredOverwriteExisting) {
            emit_log("WARNING: expert override allows synthetic recovery bytes to replace an existing destination.");
        }
        if (!scan_only && resolve_verification_mode() == models::VerificationMode::None) {
            emit_log("WARNING: destination content verification is disabled.");
        }
    }

    static bool is_blank(const std::string& text) {
        for (char c : text) {
            if (c != ' ' && c != '\t' && c != '\r' && c != '\n') return false;
        }
        return true;
    }

    // ---- Journal handling -------------------------------------------------

    // ASCII a-z/A-Z fold, matching models::detail::equals_ignore_case, for
    // O(1) relative-path set membership.
    static std::string fold_relative_path(std::string_view path) {
        std::string out(path);
        for (char& c : out) {
            if (c >= 'A' && c <= 'Z') c = static_cast<char>(c - 'A' + 'a');
        }
        return out;
    }

    JobJournal merge_journal(std::optional<JobJournal>& existing,
                             const std::vector<SourceFileDescriptor>& source_files,
                             const std::string& source_root, const std::string& destination_root,
                             const std::string& job_id, const CancelContext& cancel) {
        bool can_reuse = existing.has_value() &&
                         roots_equivalent(existing->SourceRoot, source_root) &&
                         roots_equivalent(existing->DestinationRoot, destination_root);
        if (!can_reuse && existing.has_value() && options_.AllowJournalRootRemap) {
            can_reuse = true;
            emit_log("Journal root remap enabled. Reusing existing journal rooted at source='" +
                     existing->SourceRoot + "', destination='" + existing->DestinationRoot + "'.");
        }

        JobJournal journal;
        if (can_reuse) {
            journal = std::move(*existing);
        } else {
            journal.CreatedUtc = time::DateTimeOffset::now_utc();
            journal.UpdatedUtc = journal.CreatedUtc;
        }
        journal.JobId = job_id;
        journal.SourceRoot = source_root;
        journal.DestinationRoot = destination_root;
        models::CopyJobOptions persisted_options = options_;
        // A path hint is a transport detail for locating this journal, not part
        // of the task's behavior. Everything else is required to reconstruct
        // the exact copy/assessment after a restart.
        persisted_options.ResumeJournalPathHint.clear();
        journal.RunOptions = std::move(persisted_options);

        ULONGLONG last_prepare_log_tick = GetTickCount64();
        auto pump_prepare = [&](std::size_t index, std::string_view phase) {
            cancel.throw_if_cancelled();
            ULONGLONG now = GetTickCount64();
            if (now - last_prepare_log_tick >= 1000 || index + 1 >= source_files.size()) {
                last_prepare_log_tick = now;
                emit_log("Preparing journal " + std::string(phase) + ": " +
                         std::to_string(std::min(index + 1, source_files.size())) + "/" +
                         std::to_string(source_files.size()) + ".");
            }
        };

        for (std::size_t index = 0; index < source_files.size(); ++index) {
            const auto& descriptor = source_files[index];
            JournalFileEntry* entry = journal.Files.find(descriptor.relative_path);
            if (entry == nullptr) {
                JournalFileEntry fresh;
                fresh.RelativePath = descriptor.relative_path;
                fresh.SourceLength = descriptor.length;
                fresh.SourceLastWriteUtcTicks = descriptor.last_write_utc_ticks;
                journal.Files.set(descriptor.relative_path, std::move(fresh));
                continue;
            }

            bool source_changed = entry->SourceLength != descriptor.length ||
                                  entry->SourceLastWriteUtcTicks != descriptor.last_write_utc_ticks;
            bool source_identity_changed = false;
            // Fresh journal entries are seeded from the already collected
            // enumeration metadata. Opening a handle for every file here made
            // a new scan pay an unnecessary per-file identity lookup cost. An
            // identity is required only when an existing entry actually has
            // one to compare; copy paths record the identity they capture
            // before publishing, so future resume runs retain this protection.
            const bool identity_required = entry->SourceFileIndex != 0 ||
                                            entry->SourceVolumeSerial != 0;
            const bool stability_required = entry->SourceChangeUtcTicks != 0;
            FileIdentity source_identity;
            bool source_identity_valid = false;
            FileStabilitySnapshot source_stability;
            bool source_stability_valid = false;
            if (stability_required) {
                pump_prepare(index, "source change times");
                source_stability_valid =
                    try_get_file_stability(descriptor.full_path, source_stability);
                source_identity_valid = source_stability_valid;
                if (source_stability_valid) source_identity = source_stability.identity;
                source_identity_changed =
                    !source_stability_valid ||
                    entry->SourceFileIndex !=
                        static_cast<std::int64_t>(source_stability.identity.file_index) ||
                    entry->SourceVolumeSerial !=
                        static_cast<std::int64_t>(source_stability.identity.volume_serial) ||
                    entry->SourceChangeUtcTicks != source_stability.change_utc_ticks ||
                    source_stability.length != descriptor.length ||
                    source_stability.last_write_utc_ticks != descriptor.last_write_utc_ticks;
                source_changed = source_changed || source_identity_changed;
            } else if (identity_required) {
                pump_prepare(index, "source identities");
                source_identity_valid = try_get_file_identity(descriptor.full_path, source_identity);
                source_identity_changed = !source_identity_valid ||
                                          entry->SourceFileIndex != static_cast<std::int64_t>(source_identity.file_index) ||
                                          entry->SourceVolumeSerial != static_cast<std::int64_t>(source_identity.volume_serial);
                source_changed = source_changed || source_identity_changed;
            }
            entry->SourceLength = descriptor.length;
            entry->SourceLastWriteUtcTicks = descriptor.last_write_utc_ticks;
            if (source_identity_valid) {
                entry->SourceFileIndex = static_cast<std::int64_t>(source_identity.file_index);
                entry->SourceVolumeSerial = static_cast<std::int64_t>(source_identity.volume_serial);
            }
            if (source_stability_valid) {
                entry->SourceChangeUtcTicks = source_stability.change_utc_ticks;
            }

            if (source_changed) {
                if (source_identity_changed) {
                    emit_log("Journal source binding changed; resetting coverage for " +
                             descriptor.relative_path + ".");
                }
                entry->BytesCopied = 0;
                entry->State = FileCopyState::Pending;
                entry->LastError.clear();
                entry->DoNotRetry = false;
                entry->RecoveredRanges.clear();
                entry->RescueRanges.clear();
                entry->LastRescuePass.clear();
                continue;
            }

            if (entry->State == FileCopyState::InProgress) {
                if (options_.FragileMediaMode && options_.SkipFileOnFirstReadError) {
                    entry->State = FileCopyState::Failed;
                    entry->LastError = "Fragile media guard: previous attempt stalled while this file was in progress.";
                    entry->DoNotRetry = true;
                    emit_log("[Fragile] Promoted stale in-progress entry to non-retry: " + descriptor.relative_path);
                } else {
                    entry->State = FileCopyState::Pending;
                    if (is_blank(entry->LastError)) {
                        entry->LastError = "Previous attempt ended while file was in progress.";
                    }
                    entry->DoNotRetry = false;
                }
            }

            if (options_.OperationMode == models::JobOperationMode::ScanOnly &&
                options_.ResumeFromJournal &&
                (entry->BytesCopied > 0 || !entry->RescueRanges.empty()) &&
                (entry->SourceChangeUtcTicks == 0 || entry->SourceFileIndex == 0 ||
                 entry->SourceVolumeSerial == 0)) {
                // Pre-binding journals can still be read for diagnostics, but
                // partial coverage cannot safely survive a same-size,
                // same-timestamp source replacement.
                entry->BytesCopied = 0;
                entry->State = FileCopyState::Pending;
                entry->RecoveredRanges.clear();
                entry->RescueRanges.clear();
                entry->LastRescuePass.clear();
                entry->LastError =
                    "Legacy scan coverage lacked a stable source binding and was reset.";
            }
        }

        if (options_.OperationMode != models::JobOperationMode::ScanOnly && options_.ResumeFromJournal) {
            // Partial rescue ranges describe work in a temporary destination
            // that may never have been published. Reusing those ranges against
            // the final file can skip corrupt bytes after a crash, so only a
            // fully completed entry that passes the later full-hash gate may
            // retain coverage.
            std::size_t index = 0;
            for (auto& item : journal.Files.entries) {
                pump_prepare(index++, "resume coverage");
                JournalFileEntry& entry = item.second;
                if (entry.State == FileCopyState::CompletedWithRecovery) {
                    entry.State = FileCopyState::Pending;
                    entry.LastError = "Previous recovered coverage requires a fresh exact copy.";
                    entry.DoNotRetry = false;
                }
                if (entry.State == FileCopyState::Completed) continue;
                entry.BytesCopied = 0;
                entry.RecoveredRanges.clear();
                entry.RescueRanges.clear();
                entry.LastRescuePass.clear();
            }
        }

        // Drop journal entries for files no longer present in the source. Use a
        // folded-key set for O(n) membership instead of an O(n*m) nested scan
        // (a whole-drive journal has 100k+ entries and files).
        std::unordered_set<std::string> source_keys;
        source_keys.reserve(source_files.size());
        for (std::size_t index = 0; index < source_files.size(); ++index) {
            const auto& descriptor = source_files[index];
            pump_prepare(index, "source index");
            source_keys.insert(fold_relative_path(descriptor.relative_path));
        }
        std::vector<std::pair<std::string, JournalFileEntry>> retained;
        retained.reserve(journal.Files.size());
        std::size_t retained_index = 0;
        for (auto& [key, entry] : journal.Files.entries) {
            pump_prepare(retained_index++, "journal entries");
            if (source_keys.count(fold_relative_path(key)) != 0) {
                retained.emplace_back(key, std::move(entry));
            }
        }
        journal.Files.replace_entries(std::move(retained));
        return journal;
    }

    static bool roots_equivalent(const std::string& left, const std::string& right) {
        std::wstring left_full = storage::fsutil::get_full_path(utf8_to_wide(left));
        return models::detail::equals_ignore_case(wide_to_utf8(left_full), right) ||
               models::detail::equals_ignore_case(left, right);
    }

    std::vector<std::wstring> journal_load_candidates(const std::wstring& default_path,
                                                      const std::wstring& active_path) const {
        std::vector<std::wstring> candidates;
        std::wstring hinted = utf8_to_wide(options_.ResumeJournalPathHint);
        auto add = [&candidates](const std::wstring& value) {
            if (value.empty()) return;
            for (const auto& existing : candidates) {
                if (storage::fsutil::to_upper_invariant(existing) ==
                    storage::fsutil::to_upper_invariant(value)) {
                    return;
                }
            }
            candidates.push_back(value);
        };
        add(detail::trim_trailing_separators(hinted).empty() ? std::wstring() : hinted);
        add(active_path);
        add(default_path);
        return candidates;
    }

    std::wstring resolve_writable_journal_path(const std::string& job_id,
                                               const std::wstring& default_journal_path) {
        if (is_journal_path_writable(default_journal_path)) return default_journal_path;

        wchar_t temp_root[MAX_PATH];
        GetTempPathW(MAX_PATH, temp_root);
        std::wstring fallback_dir = std::wstring(temp_root) + L"XactCopy\\journals";
        std::wstring fallback = fallback_dir + L"\\job-" + utf8_to_wide(job_id) + L".json";
        if (is_journal_path_writable(fallback)) {
            emit_log("Journal path fallback active: " + wide_to_utf8(fallback));
            return fallback;
        }
        throw IoError("Unable to write journal at default path '" + wide_to_utf8(default_journal_path) +
                      "' or fallback '" + wide_to_utf8(fallback) + "'.");
    }

    static bool is_journal_path_writable(const std::wstring& journal_path) {
        if (journal_path.empty()) return false;
        std::wstring directory = storage::fsutil::get_directory_name(journal_path);
        if (directory.empty()) return false;
        storage::fsutil::create_directories(directory);
        std::wstring probe = directory + L"\\.xactcopy-probe-" +
                             storage::fsutil::random_temp_suffix() + L".tmp";
        if (!storage::fsutil::write_file_raw(probe, reinterpret_cast<const unsigned char*>("ok"), 2,
                                             false, true)) {
            return false;
        }
        DeleteFileW(probe.c_str());
        return true;
    }

    // Saves unconditionally and records what it cost, so the throttle below can
    // price the next flush. Used for the job's own boundaries (first save,
    // phase end, failure exits) where the snapshot must land regardless.
    void save_journal_now(const std::wstring& journal_path, JobJournal& journal) {
        ULONGLONG started = GetTickCount64();
        if (!journal_store_.save(journal_path, journal)) {
            emit_log("WARNING: Journal checkpoint was committed with reduced snapshot/ledger redundancy.");
        }
        last_journal_save_cost_ms_ = GetTickCount64() - started;
        last_journal_flush_tick_ = GetTickCount64();
    }

    // A journal save is O(journal size): the whole snapshot is re-serialized,
    // hashed, rotated, and mirrored. On a whole-drive job that snapshot is
    // 100 MB+, so a fixed 500 ms cadence spends more I/O on journaling than on
    // the copy itself — and the throttle silently degrades into "save as fast
    // as the disk allows", which is what makes a large job crawl.
    //
    // Instead of a fixed cadence, bound journaling's share of wall time: after
    // a save costing N ms the next flush waits at least N * (Divisor - 1), so
    // journal I/O settles near 1/Divisor of the run whatever the journal size.
    // Small journals save in ~1 ms and keep the original sub-second cadence;
    // only journals big enough to hurt back off. The ceiling keeps even a huge
    // journal checkpointing regularly.
    //
    // The cost is resume granularity: a large job may redo up to one interval
    // of work after a crash. That is a good trade against spending half the run
    // rewriting a snapshot, and it does not affect stall detection — the
    // supervisor watches heartbeats and progress events, not journal writes.
    static constexpr ULONGLONG JournalFlushDutyDivisor = 20;
    static constexpr ULONGLONG MaximumJournalFlushIntervalMs = 5 * 60 * 1000;

    void flush_journal(const std::wstring& journal_path, JobJournal& journal, bool force) {
        ULONGLONG now = GetTickCount64();
        ULONGLONG interval_ms = force ? 500 : 1000;
        ULONGLONG budget_ms = last_journal_save_cost_ms_ * (JournalFlushDutyDivisor - 1);
        if (budget_ms > interval_ms) {
            interval_ms = std::min(budget_ms, MaximumJournalFlushIntervalMs);
        }
        if (last_journal_flush_tick_ != 0 && now - last_journal_flush_tick_ < interval_ms) return;
        save_journal_now(journal_path, journal);
    }

    // ---- Per-file gates ---------------------------------------------------

    bool is_already_completed(JournalFileEntry& entry, const SourceFileDescriptor& descriptor,
                              const std::wstring& destination_root,
                              const CancelContext& cancel) {
        if (entry.State != FileCopyState::Completed &&
            entry.State != FileCopyState::CompletedWithRecovery) {
            return false;
        }
        if (entry.State == FileCopyState::CompletedWithRecovery) {
            reset_journal_coverage_for_retry(entry);
            return false;
        }
        FileStabilitySnapshot source_stability;
        FileIdentity source_identity;
        if (try_get_file_stability(descriptor.full_path, source_stability)) {
            source_identity = source_stability.identity;
            note_hard_link_topology(source_stability);
        } else if (!try_get_file_identity(descriptor.full_path, source_identity)) {
            reset_journal_coverage_for_retry(entry);
            return false;
        }
        const std::wstring destination_path = destination_path_for(destination_root, descriptor);
        ExistingFileMetadata metadata;
        if (!try_get_existing_file_metadata(destination_path, metadata)) {
            reset_journal_coverage_for_retry(entry);
            return false;
        }
        if (metadata.length != descriptor.length) {
            reset_journal_coverage_for_retry(entry);
            return false;
        }
        FileIdentity destination_identity;
        if (!try_get_file_identity(destination_path, destination_identity)) {
            reset_journal_coverage_for_retry(entry);
            return false;
        }

        // A journal entry is not proof that the destination bytes survived a
        // crash or were not replaced by another file. Only a full source/
        // destination hash comparison permits the engine to skip a completed
        // file; otherwise the staged copy path below reconstructs it safely.
        if (resolve_verification_mode() != models::VerificationMode::Full) {
            reset_journal_coverage_for_retry(entry);
            return false;
        }
        try {
            verify_file_with_retries(descriptor.full_path,
                                     destination_path,
                                     descriptor.relative_path,
                                     models::VerificationMode::Full, cancel);
            ensure_source_identity_unchanged(descriptor, source_identity, true);
            FileIdentity current_destination_identity;
            if (!try_get_file_identity(destination_path, current_destination_identity) ||
                current_destination_identity.file_index != destination_identity.file_index ||
                current_destination_identity.volume_serial != destination_identity.volume_serial) {
                throw IoError("Destination file identity changed during journal validation: " +
                              descriptor.relative_path + ".");
            }
            return true;
        } catch (const std::exception& ex) {
            emit_log("Journal completion rejected for " + descriptor.relative_path + ": " +
                     ex.what() + "; copying again.");
            reset_journal_coverage_for_retry(entry);
            return false;
        }
    }

    static void validate_resolved_root_relationship(const std::wstring& source_root,
                                                    const std::wstring& destination_root,
                                                    bool scan_only) {
        if (scan_only) return;
        auto trim_separators = [](std::wstring value) {
            while (value.size() > 3 && !value.empty() &&
                   (value.back() == L'\\' || value.back() == L'/')) {
                value.pop_back();
            }
            return value;
        };
        std::wstring source = trim_separators(storage::fsutil::resolve_final_path(source_root));
        std::wstring destination =
            trim_separators(storage::fsutil::resolve_final_path(destination_root));
        if (source.empty() || destination.empty()) return;
        source = storage::fsutil::to_upper_invariant(source);
        destination = storage::fsutil::to_upper_invariant(destination);
        if (source == destination) {
            throw IoError("Source and destination resolve to the same physical path.");
        }
        if (destination.size() > source.size() &&
            destination.compare(0, source.size(), source) == 0 &&
            destination[source.size()] == L'\\') {
            throw IoError("Destination resolves inside the source tree through a junction, mount point, or path alias.");
        }
    }

    static void reset_journal_coverage_for_retry(JournalFileEntry& entry) {
        entry.BytesCopied = 0;
        entry.State = FileCopyState::Pending;
        entry.LastError.clear();
        entry.DoNotRetry = false;
        entry.RecoveredRanges.clear();
        entry.RescueRanges.clear();
        entry.LastRescuePass.clear();
    }

    bool should_skip_failed_entry_for_fragile_resume(const JournalFileEntry& entry) const {
        return entry.State == FileCopyState::Failed && entry.DoNotRetry &&
               (options_.PersistFragileSkipAcrossResume || options_.FragileMediaMode);
    }

    enum class ExistingDestinationDecision : std::uint8_t { Copy, Skip, Reject };

    ExistingDestinationDecision decide_overwrite_policy(const SourceFileDescriptor& descriptor,
                                                        const std::wstring& destination_root,
                                                        std::string& reason) const {
        reason.clear();
        ExistingFileMetadata metadata;
        if (!try_get_existing_file_metadata(destination_path_for(destination_root, descriptor), metadata)) {
            return ExistingDestinationDecision::Copy;
        }
        switch (options_.OverwritePolicyValue) {
            case models::OverwritePolicy::SkipExisting:
                reason = "destination already exists";
                return ExistingDestinationDecision::Skip;
            case models::OverwritePolicy::OverwriteIfSourceNewer:
                if (descriptor.last_write_utc_ticks <= metadata.last_write_utc_ticks) {
                    reason = "destination is newer or same age";
                    return ExistingDestinationDecision::Skip;
                }
                return ExistingDestinationDecision::Copy;
            case models::OverwritePolicy::Ask:
                reason = "overwrite policy 'Ask' requires explicit confirmation; refusing to overwrite an existing destination";
                return ExistingDestinationDecision::Reject;
            default:
                return ExistingDestinationDecision::Copy;
        }
    }

    static std::wstring destination_path_for(const std::wstring& destination_root,
                                             const SourceFileDescriptor& descriptor) {
        return detail::trim_trailing_separators(destination_root) + L"\\" +
               utf8_to_wide(descriptor.relative_path);
    }

    std::int64_t record_fragile_read_skip(const SourceFileDescriptor& descriptor,
                                          JournalFileEntry& entry,
                                          const FragileReadSkip& failure) {
        entry.State = FileCopyState::Failed;
        entry.LastError = failure.what();
        entry.DoNotRetry = options_.PersistFragileSkipAcrossResume;

        if (!failure.PrimaryDataRange || failure.Offset < 0 || failure.Length <= 0 ||
            failure.Offset >= descriptor.length) {
            return 0;
        }
        const std::int64_t observed_length = std::min<std::int64_t>(
            failure.Length, descriptor.length - failure.Offset);
        if (entry.RescueRanges.empty()) {
            append_range(entry.RescueRanges, 0, descriptor.length,
                         RescueRangeState::Pending);
        }
        set_range_state(entry.RescueRanges, failure.Offset, observed_length,
                        RescueRangeState::Bad);
        entry.LastRescuePass = "FragileFirstRead";
        return observed_length;
    }

    void handle_fragile_read_skip(const SourceFileDescriptor& descriptor, JournalFileEntry& entry,
                                  ProgressAccumulator& progress, JobJournal& journal,
                                  const std::wstring& journal_path,
                                  const FragileReadSkip& failure, const CancelContext& cancel,
                                  const std::string& operation_label = "copy",
                                  bool register_failure = true,
                                  bool force_journal_flush = true) {
        const std::string message = failure.what();
        std::int64_t already = entry.BytesCopied > 0 ? entry.BytesCopied : 0;
        std::int64_t remaining = descriptor.length - already;
        progress.skipped_files += 1;
        const std::int64_t skipped = remaining > 0 ? remaining : 0;
        progress.total_bytes_copied += skipped;
        progress.bytes_skipped += skipped;

        const std::int64_t observed_length =
            record_fragile_read_skip(descriptor, entry, failure);
        if (observed_length > 0) {
            sync_entry_bytes_copied(entry);
            emit_log("[Fragile] Recorded first unreadable read range for " +
                     descriptor.relative_path + " at " + format_bytes(failure.Offset) +
                     " (" + format_bytes(observed_length) +
                     "; diagnostic until a later matching observation confirms it).");
        } else if (!failure.PrimaryDataRange) {
            emit_log("[Fragile] The failed read was outside the primary file data; no primary-data "
                     "bad-range hint was created.");
        }

        emit_log("Skipped " + operation_label + ": " + descriptor.relative_path + " (" + message + ")");
        emit_progress(progress, descriptor.relative_path, descriptor.length, descriptor.length, 0,
                      resolve_buffer_size_for_file(descriptor.length), entry.LastRescuePass,
                      unreadable_region_count(entry.RescueRanges),
                      rescue_bytes(entry.RescueRanges,
                                   {RescueRangeState::Bad, RescueRangeState::KnownBad}));
        if (register_failure) {
            register_fragile_failure_and_maybe_cooldown(descriptor.relative_path, message, cancel);
        }
        try_persist_bad_range_map_entry(descriptor, entry, false);
        flush_journal(journal_path, journal, force_journal_flush);
    }

    void register_fragile_failure_and_maybe_cooldown(const std::string& relative_path,
                                                     const std::string& reason,
                                                     const CancelContext& cancel) {
        if (!options_.FragileMediaMode) return;

        ULONGLONG now = GetTickCount64();
        std::int64_t window_ms = std::max(1, options_.FragileFailureWindowSeconds) * 1000LL;
        std::int32_t threshold = std::max(1, options_.FragileFailureThreshold);
        std::int32_t cooldown_seconds = std::max(0, options_.FragileCooldownSeconds);

        fragile_failure_timestamps_.push_back(now);
        while (!fragile_failure_timestamps_.empty() &&
               fragile_failure_timestamps_.front() + static_cast<ULONGLONG>(window_ms) < now) {
            fragile_failure_timestamps_.pop_front();
        }
        if (static_cast<std::int32_t>(fragile_failure_timestamps_.size()) < threshold) return;

        std::string path_label = relative_path.empty() ? "?" : relative_path;
        std::string reason_suffix = reason.empty() ? std::string() : (" (" + reason + ")");
        emit_log("[Fragile] Failure threshold reached after " + path_label + reason_suffix + ".");
        fragile_failure_timestamps_.clear();
        if (cooldown_seconds <= 0) return;
        emit_log("[Fragile] Cooling down for " + std::to_string(cooldown_seconds) +
                 " second(s) to reduce stress on unstable media.");
        cancel.delay(cooldown_seconds * 1000LL);
    }

    // ---- Single-file copy pipeline ---------------------------------------

    static bool stream_names_equal(const std::wstring& left, const std::wstring& right) {
        return CompareStringOrdinal(left.c_str(), static_cast<int>(left.size()),
                                    right.c_str(), static_cast<int>(right.size()), TRUE) ==
               CSTR_EQUAL;
    }

    static const NamedStreamDescriptor* find_stream(
        const std::vector<NamedStreamDescriptor>& streams, const std::wstring& name) {
        for (const auto& stream : streams) {
            if (stream_names_equal(stream.name, name)) return &stream;
        }
        return nullptr;
    }

    static bool path_is_encrypted(const std::wstring& path) {
        DWORD attributes = GetFileAttributesW(path.c_str());
        return attributes != INVALID_FILE_ATTRIBUTES &&
               (attributes & FILE_ATTRIBUTE_ENCRYPTED) != 0;
    }

    void prepare_managed_encryption(const SourceFileDescriptor& descriptor,
                                    const std::wstring& working_path) {
        if ((descriptor.attributes & FILE_ATTRIBUTE_ENCRYPTED) == 0 ||
            options_.AllowDecryptedDestination || path_is_encrypted(working_path)) {
            return;
        }
        if (!EncryptFileW(working_path.c_str()) || !path_is_encrypted(working_path)) {
            DWORD error = GetLastError();
            throw IoError::from_win32(
                "Unable to preserve EFS encryption on the staged destination. Enable the explicit plaintext override only when decrypted output is intended.",
                error);
        }
    }

    void enforce_encryption_result(const SourceFileDescriptor& descriptor,
                                   const std::wstring& working_path) {
        if ((descriptor.attributes & FILE_ATTRIBUTE_ENCRYPTED) == 0 ||
            path_is_encrypted(working_path)) {
            return;
        }
        if (!options_.AllowDecryptedDestination) {
            throw IoError("Encrypted source would be published as plaintext: " +
                          descriptor.relative_path + ".");
        }
        add_integrity_warning(
            "EFS encryption was explicitly removed from one or more destination files.");
    }

    void add_integrity_warning(const std::string& warning) {
        if (warning.empty()) return;
        std::lock_guard<std::mutex> guard(metadata_warning_lock_);
        if (integrity_warning_.find(warning) != std::string::npos) return;
        if (!integrity_warning_.empty()) integrity_warning_ += " ";
        integrity_warning_ += warning;
    }

    void add_metadata_notice(const std::string& notice) {
        if (notice.empty()) return;
        std::lock_guard<std::mutex> guard(metadata_warning_lock_);
        if (metadata_notice_.find(notice) != std::string::npos) return;
        if (!metadata_notice_.empty()) metadata_notice_ += " ";
        metadata_notice_ += notice;
    }

    void note_hard_link_topology(const FileStabilitySnapshot& snapshot) {
        if (snapshot.link_count > 1) {
            add_metadata_notice(
                "NTFS hard-link topology was not preserved; linked source names were copied as independent destination files.");
        }
    }

    void copy_named_streams(const SourceFileDescriptor& descriptor,
                            const std::wstring& working_path,
                            const NamedStreamEnumeration& source_streams,
                            const CancelContext& cancel,
                            bool verify_stream_content) {
        NamedStreamEnumeration destination_streams = enumerate_named_streams(working_path);
        if (source_streams.supported && !destination_streams.supported &&
            !source_streams.streams.empty()) {
            throw IoError("Destination filesystem does not support alternate data streams for " +
                          descriptor.relative_path + ".");
        }

        // The stage may have been cloned from an older destination. Remove
        // streams absent from the current source so stale Zone.Identifier or
        // application metadata cannot survive a successful replacement.
        for (const auto& destination_stream : destination_streams.streams) {
            if (find_stream(source_streams.streams, destination_stream.name) != nullptr) continue;
            std::wstring stale_path = working_path + destination_stream.name;
            if (!DeleteFileW(stale_path.c_str())) {
                throw IoError::from_win32("Unable to remove stale destination data stream.",
                                          GetLastError());
            }
        }

        if (!source_streams.supported) return;
        std::int32_t stream_buffer_size = std::min<std::int32_t>(
            1024 * 1024, std::max<std::int32_t>(MinimumRescueBlockSize, options_.BufferSizeBytes));
        std::vector<unsigned char> buffer(static_cast<std::size_t>(stream_buffer_size));

        for (const auto& source_stream : source_streams.streams) {
            cancel.throw_if_cancelled();
            std::wstring source_stream_path = descriptor.full_path + source_stream.name;
            std::wstring destination_stream_path = working_path + source_stream.name;
            HANDLE create = CreateFileW(destination_stream_path.c_str(), GENERIC_WRITE,
                                         FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE,
                                         nullptr, CREATE_ALWAYS,
                                         FILE_ATTRIBUTE_NORMAL | FILE_FLAG_OVERLAPPED, nullptr);
            if (create == INVALID_HANDLE_VALUE) {
                throw IoError::from_win32("Unable to create destination alternate data stream.",
                                          GetLastError());
            }
            CloseHandle(create);

            emit_log("Copying alternate stream " + descriptor.relative_path +
                     " [" + wide_to_utf8(source_stream.name) + "].");
            FileTransferSession session(source_stream_path, destination_stream_path);
            std::int64_t offset = 0;
            std::string stream_label = descriptor.relative_path + " " +
                                       wide_to_utf8(source_stream.name);
            while (offset < source_stream.length) {
                cancel.throw_if_cancelled();
                std::int32_t count = static_cast<std::int32_t>(std::min<std::int64_t>(
                    stream_buffer_size, source_stream.length - offset));
                std::int32_t read = read_chunk_with_retries(
                    source_stream_path, stream_label, offset, count, session, buffer.data(),
                    cancel, std::max(0, options_.MaxRetries), false,
                    /*primary_data_range*/ false);
                if (read != count) {
                    throw IoError("Short read while copying alternate data stream " + stream_label + ".");
                }
                write_chunk_with_retries(stream_label, offset, buffer.data(), count, session, cancel);
                offset += count;
            }
            session.invalidate_source();
            session.invalidate_destination();
        }

        NamedStreamEnumeration after = enumerate_named_streams(descriptor.full_path);
        if (after.supported != source_streams.supported ||
            after.streams.size() != source_streams.streams.size()) {
            throw IoError("Source alternate data streams changed during copy of " +
                          descriptor.relative_path + ".");
        }
        for (const auto& before : source_streams.streams) {
            const auto* current = find_stream(after.streams, before.name);
            if (current == nullptr || current->length != before.length) {
                throw IoError("Source alternate data streams changed during copy of " +
                              descriptor.relative_path + ".");
            }
        }
        if (verify_stream_content) {
            for (const auto& stream : source_streams.streams) {
                cancel.throw_if_cancelled();
                emit_log("Verifying alternate stream " + descriptor.relative_path +
                         " [" + wide_to_utf8(stream.name) + "].");
                std::wstring source_stream_path = descriptor.full_path + stream.name;
                std::wstring destination_stream_path = working_path + stream.name;
                if (compute_file_hash(source_stream_path, cancel) !=
                    compute_file_hash(destination_stream_path, cancel)) {
                    throw IoError("Alternate data stream verification mismatch: " +
                                  descriptor.relative_path + " [" +
                                  wide_to_utf8(stream.name) + "].");
                }
            }
        }
    }

    void preserve_file_metadata(const SourceFileDescriptor& descriptor,
                                const std::wstring& working_path,
                                const FileSecuritySnapshot& source_security,
                                const NamedStreamEnumeration& source_streams,
                                const CancelContext& cancel,
                                models::VerificationMode verification_mode) {
        cancel.throw_if_cancelled();
        if ((descriptor.attributes & ~(CopyableFileAttributes | FILE_ATTRIBUTE_ENCRYPTED)) != 0) {
            add_metadata_notice(
                "Filesystem-specific attributes (compression, sparse, or reparse) "
                "were not fully preserved for one or more files.");
        }
        enforce_encryption_result(descriptor, working_path);
        copy_named_streams(descriptor, working_path, source_streams, cancel,
                           verification_mode == models::VerificationMode::Full);

        FileSecurityApplyResult security_result =
            apply_file_security(working_path, source_security);
        if (source_security.sacl_omitted || !security_result.sacl_preserved) {
            add_metadata_notice("Security audit entries (SACLs) were not fully preserved for one or more files.");
        }
        if (!security_result.owner_group_preserved) {
            add_metadata_notice("Security owner/group were not fully preserved for one or more files; "
                                "the source DACL was applied.");
        }

        if (options_.PreserveTimestamps &&
            !set_file_times_utc(working_path, descriptor.creation_time,
                                descriptor.last_access_time, descriptor.last_write_time)) {
            throw IoError::from_win32("Unable to preserve source file timestamps.", GetLastError());
        }

        // Apply READONLY last so all stream, security, timestamp, and durability
        // operations above still have write access to the stage.
        if (!set_copyable_file_attributes(working_path, descriptor.attributes)) {
            throw IoError::from_win32("Unable to preserve source file attributes.", GetLastError());
        }
    }

    void write_recovery_manifest(const std::wstring& manifest_path,
                                 const std::wstring& recovered_path,
                                 const DestinationStage& stage,
                                 const SourceFileDescriptor& descriptor,
                                 const JournalFileEntry* entry) {
        json::Writer writer(/*indented*/ true);
        writer.begin_object();
        writer.key("SchemaVersion"); writer.value(1);
        writer.key("CreatedUtc"); writer.value_literal_string(time::DateTimeOffset::now_utc().to_string());
        writer.key("SourcePath"); writer.value(wide_to_utf8(descriptor.full_path));
        writer.key("SourceRelativePath"); writer.value(descriptor.relative_path);
        writer.key("SourceMediaIdentity"); writer.value(expected_media_identity(true));
        writer.key("IntendedDestinationPath"); writer.value(wide_to_utf8(stage.final_path()));
        writer.key("RecoveredOutputPath"); writer.value(wide_to_utf8(recovered_path));
        writer.key("SourceLength"); writer.value(descriptor.length);
        writer.key("FillPattern"); writer.value(describe_salvage_fill());
        writer.key("SyntheticRanges");
        writer.begin_array();
        if (entry != nullptr) {
            for (const RescueRange& range : entry->RescueRanges) {
                if (range.State != RescueRangeState::Recovered || range.Length <= 0) continue;
                writer.begin_object();
                writer.key("Offset"); writer.value(range.Offset);
                writer.key("Length"); writer.value(range.Length);
                writer.end_object();
            }
        }
        writer.end_array();
        writer.end_object();
        if (!storage::fsutil::write_atomic_bytes(manifest_path, writer.take())) {
            throw IoError("Unable to write recovered-output manifest.");
        }
    }

    void remove_owned_recovery_manifest(const std::wstring& recovered_path) {
        const std::wstring manifest_path = recovered_path + L".recovery.json";
        auto bytes = storage::fsutil::read_all_bytes(manifest_path, 4ULL * 1024 * 1024);
        if (!bytes.has_value()) return;
        bool owned = false;
        try {
            const std::string text(reinterpret_cast<const char*>(bytes->data()), bytes->size());
            json::Value value = json::parse(text);
            const json::Object* object = value.as_object();
            std::string recorded_path;
            if (object != nullptr) {
                models::detail::read(object, "RecoveredOutputPath", recorded_path);
                owned = models::detail::equals_ignore_case(
                    recorded_path, wide_to_utf8(recovered_path));
            }
        } catch (const std::exception&) {
            return; // Never delete an unrecognized adjacent user file.
        }
        if (owned && !DeleteFileW(manifest_path.c_str())) {
            DWORD error = GetLastError();
            if (error != ERROR_FILE_NOT_FOUND) {
                add_metadata_notice(
                    "An obsolete XactCopy recovery manifest could not be removed after an exact replacement.");
            }
        }
    }

    void publish_stage(DestinationStage& stage, const SourceFileDescriptor& descriptor,
                       bool contains_synthetic_bytes,
                       const JournalFileEntry* entry = nullptr) {
        if (contains_synthetic_bytes && !options_.AllowRecoveredOverwriteExisting) {
            const bool existing_destination_preserved = stage.final_exists_now();
            std::wstring sidecar = stage.commit_recovered_sidecar();
            std::wstring manifest = sidecar + L".recovery.json";
            bool manifest_written = false;
            try {
                write_recovery_manifest(manifest, sidecar, stage, descriptor, entry);
                manifest_written = true;
            } catch (const std::exception& ex) {
                std::string notice =
                    "Recovered bytes were preserved, but their recovery manifest could not be written: " +
                    std::string(ex.what());
                add_metadata_notice(notice);
                emit_log("WARNING: " + notice);
            }
            emit_log("RECOVERED OUTPUT: " +
                     std::string(existing_destination_preserved
                                     ? "existing destination preserved for "
                                     : "non-exact bytes withheld from the normal destination for ") +
                     descriptor.relative_path + "; recovery bytes published to " +
                     wide_to_utf8(sidecar) +
                     (manifest_written ? "; manifest " + wide_to_utf8(manifest) : std::string()) + ".");
        } else {
            stage.commit();
            if (contains_synthetic_bytes) {
                std::wstring manifest = stage.published_path() + L".recovery.json";
                try {
                    write_recovery_manifest(manifest, stage.published_path(), stage,
                                            descriptor, entry);
                    emit_log("RECOVERED OUTPUT OVERRIDE: synthetic bytes replaced the normal "
                             "destination for " + descriptor.relative_path + "; manifest " +
                             wide_to_utf8(manifest) + ".");
                } catch (const std::exception& ex) {
                    std::string notice =
                        "Recovered destination bytes were published, but their recovery manifest could not be written: " +
                        std::string(ex.what());
                    add_metadata_notice(notice);
                    emit_log("WARNING: " + notice);
                }
            } else {
                remove_owned_recovery_manifest(stage.published_path());
            }
        }
        if (stage.post_publish_attribute_error() != ERROR_SUCCESS) {
            std::string warning =
                "Published file attributes could not be fully restored for " +
                descriptor.relative_path + " (Win32 " +
                std::to_string(stage.post_publish_attribute_error()) + ").";
            add_metadata_notice(warning);
            emit_log("WARNING: " + warning);
        }
    }

    bool copy_single_file(const SourceFileDescriptor& descriptor, JournalFileEntry& entry,
                          const std::wstring& destination_root, ProgressAccumulator& progress,
                          JobJournal& journal, const std::wstring& journal_path,
                          const CancelContext& cancel) {
        std::wstring destination_path = destination_path_for(destination_root, descriptor);
        std::wstring destination_directory = storage::fsutil::get_directory_name(destination_path);
        if (!destination_directory.empty() && add_created_directory(destination_directory)) {
            storage::fsutil::create_directories(destination_directory);
        }

        FileStabilitySnapshot source_stability;
        const bool source_stability_valid =
            try_get_file_stability(descriptor.full_path, source_stability);
        if (source_stability_valid) note_hard_link_topology(source_stability);
        if (source_stability_valid &&
            (source_stability.length != descriptor.length ||
             source_stability.last_write_utc_ticks != descriptor.last_write_utc_ticks)) {
            throw IoError("Source changed after enumeration: " + descriptor.relative_path + ".");
        }
        FileIdentity source_identity = source_stability.identity;
        const bool source_identity_valid = source_stability_valid ||
                                           try_get_file_identity(descriptor.full_path, source_identity);
        if (source_identity_valid) {
            entry.SourceFileIndex = static_cast<std::int64_t>(source_identity.file_index);
            entry.SourceVolumeSerial = static_cast<std::int64_t>(source_identity.volume_serial);
        }
        if (source_stability_valid) {
            entry.SourceChangeUtcTicks = source_stability.change_utc_ticks;
        }
        FileSecuritySnapshot source_security = capture_file_security(descriptor.full_path);
        NamedStreamEnumeration source_streams = enumerate_named_streams(descriptor.full_path);

        std::int64_t destination_length = get_existing_file_length(destination_path);
        bool has_persisted_coverage = false;
        for (const auto& range : entry.RescueRanges) {
            if (range.Length <= 0) continue;
            RescueRangeState state = normalize_state(range.State);
            if (state == RescueRangeState::Good || state == RescueRangeState::Recovered) {
                has_persisted_coverage = true;
                break;
            }
        }
        bool preserve_existing = options_.ResumeFromJournal && destination_length > 0 &&
                                 (entry.BytesCopied > 0 || has_persisted_coverage);

        DestinationStage stage(destination_path, cancel);
        if (preserve_existing) stage.clone_existing_if_present();
        const std::wstring& working_path = stage.working_path();
        prepare_destination_file(working_path, descriptor.length, destination_length, preserve_existing);
        destination_length = get_existing_file_length(working_path);

        // Native CopyFileEx fast path.
        std::int64_t native_baseline = progress.total_bytes_copied;
        NativeAttempt native = try_copy_with_native_fast_path(
            descriptor, entry, working_path, destination_length, has_persisted_coverage,
            progress, journal, journal_path, cancel);
        if (native.attempted) {
            if (native.succeeded) {
                ensure_source_identity_unchanged(descriptor, source_identity, source_identity_valid);
                ensure_media_identity_integrity(utf8_to_wide(options_.SourceRoot), destination_root,
                                                true, true);
                stage.make_writable();
                if (resolve_verification_mode() != models::VerificationMode::None) {
                    verify_file_with_retries(descriptor.full_path, working_path,
                                             descriptor.relative_path, resolve_verification_mode(), cancel);
                }
                preserve_file_metadata(descriptor, working_path, source_security, source_streams,
                                       cancel, resolve_verification_mode());
                ensure_source_stability_unchanged(descriptor, source_stability,
                                                  source_stability_valid);
                publish_stage(stage, descriptor, false);
                return false;
            }
            progress.total_bytes_copied = native_baseline;
            entry.BytesCopied = 0;
            entry.LastRescuePass = "Init";
            if (!native.fallback_reason.empty()) {
                emit_log("Native fast-path fallback on " + descriptor.relative_path + ": " +
                         native.fallback_reason);
            }
            prepare_destination_file(working_path, descriptor.length,
                                     get_existing_file_length(working_path), false);
            destination_length = get_existing_file_length(working_path);
        }

        prepare_managed_encryption(descriptor, working_path);

        AdaptiveBufferController controller =
            create_buffer_controller(descriptor.length, BufferPurpose::Copy);
        const std::int32_t active_buffer_size = controller.current_size();
        std::vector<unsigned char> io_buffer(
            static_cast<std::size_t>(std::max(controller.maximum_size(), MinimumRescueBlockSize)));

        FileTransferSession session(descriptor.full_path, working_path);
        entry.RescueRanges = build_rescue_ranges(entry, descriptor.length, destination_length);
        std::int64_t mapped_unreadable = apply_known_bad_ranges_from_map(descriptor, entry);
        entry.LastRescuePass = "Init";

        std::int64_t already_satisfied = rescue_bytes(
            entry.RescueRanges, {RescueRangeState::Good, RescueRangeState::Recovered});
        std::optional<CopyVerificationTracker> verification_tracker;
        if (resolve_verification_mode() == models::VerificationMode::Full) {
            verification_tracker.emplace(use_sha512());
            if (preserve_existing || already_satisfied > 0) verification_tracker->invalidate();
        }
        entry.BytesCopied = already_satisfied;
        progress.total_bytes_copied += already_satisfied;
        emit_progress(progress, descriptor.relative_path,
                      display_progress_bytes(entry, descriptor.length, already_satisfied),
                      descriptor.length, 0, active_buffer_size, entry.LastRescuePass,
                      unreadable_region_count(entry.RescueRanges),
                      rescue_bytes(entry.RescueRanges, {RescueRangeState::Pending, RescueRangeState::Bad, RescueRangeState::KnownBad}));

        bool recovered_any = false;
        bool use_small_fast = should_use_small_file_fast_path(descriptor.length, has_persisted_coverage) &&
                              mapped_unreadable <= 0;
        if (use_small_fast) {
            entry.LastRescuePass = "SmallFileFast";
            recovered_any = copy_small_file_fast(descriptor, entry, working_path, session,
                                                 io_buffer, controller, progress, journal,
                                                 journal_path, cancel,
                                                 verification_tracker ? &*verification_tracker : nullptr);
        } else {
            std::vector<RescuePassDefinition> pass_plan = build_rescue_pass_plan(active_buffer_size);
            double bad_density = rescue_bad_density(entry.RescueRanges, descriptor.length);
            for (const auto& base_pass : pass_plan) {
                RescuePassDefinition pass = adapt_pass_for_density(base_pass, bad_density, active_buffer_size);
                if (region_count(entry.RescueRanges, {pass.target_state}) <= 0) continue;
                entry.LastRescuePass = pass.name;

                RescuePassOutcome outcome = execute_rescue_pass(
                    pass, descriptor, entry, session, io_buffer, progress, journal, journal_path,
                    cancel, /*write_recovered*/ true, /*count_failed_as_processed*/ false,
                    &controller, verification_tracker ? &*verification_tracker : nullptr);

                std::int32_t remaining_bad = unreadable_region_count(entry.RescueRanges);
                std::int64_t remaining_bad_bytes = rescue_bytes(
                    entry.RescueRanges, {RescueRangeState::Bad, RescueRangeState::KnownBad});
                bad_density = rescue_bad_density(entry.RescueRanges, descriptor.length);
                bool log_summary = models::detail::equals_ignore_case(pass.name, "FastScan")
                                       ? outcome.failed_segments > 0
                                       : (outcome.attempted_segments > 0 || outcome.recovered_bytes > 0 ||
                                          remaining_bad > 0);
                if (log_summary) {
                    char density_text[16];
                    std::snprintf(density_text, sizeof(density_text), "%.1f", bad_density * 100.0);
                    emit_log("[Rescue Engine] " + pass.name + " " + descriptor.relative_path +
                             ": attempts " + std::to_string(outcome.attempted_segments) +
                             ", recovered " + format_bytes(outcome.recovered_bytes) +
                             ", failed " + std::to_string(outcome.failed_segments) +
                             ", remaining bad " + std::to_string(remaining_bad) +
                             " (" + format_bytes(remaining_bad_bytes) + "), density " +
                             density_text + " %.");
                }
                sync_entry_bytes_copied(entry);
                flush_journal(journal_path, journal, true);
            }

            recovered_any = region_count(entry.RescueRanges, {RescueRangeState::Recovered}) > 0;
            std::int64_t remaining_bad_bytes = rescue_bytes(
                entry.RescueRanges, {RescueRangeState::Bad, RescueRangeState::KnownBad});
            if (remaining_bad_bytes > 0) {
                if (!options_.SalvageUnreadableBlocks) {
                    throw IoError("[Rescue Engine] Unrecoverable regions remain on " +
                                  descriptor.relative_path + ": " + format_bytes(remaining_bad_bytes) +
                                  " in " + std::to_string(unreadable_region_count(entry.RescueRanges)) +
                                  " block(s).");
                }
                entry.LastRescuePass = "SalvageFill";
                emit_log("[Rescue Engine] Salvaging remaining unreadable regions on " +
                         descriptor.relative_path + ".");
                bool salvaged = salvage_remaining_ranges(descriptor, entry, session, io_buffer,
                                                         controller.current_size(), progress,
                                                         journal, journal_path, cancel);
                recovered_any = salvaged || recovered_any;
                if (verification_tracker) verification_tracker->invalidate();
                sync_entry_bytes_copied(entry);
                entry.RecoveredRanges = merge_byte_ranges(entry.RecoveredRanges);
                flush_journal(journal_path, journal, true);
            }
        }

        ensure_source_identity_unchanged(descriptor, source_identity, source_identity_valid);
        session.invalidate_destination(); // release the staged write handle before metadata/commit

        models::VerificationMode verification = resolve_verification_mode();
        if (verification != models::VerificationMode::None) {
            if (recovered_any) {
                emit_log("Verification skipped for recovered file: " + descriptor.relative_path);
            } else {
                session.invalidate_source();
                session.invalidate_destination();
                std::optional<std::vector<unsigned char>> copied_source_hash;
                if (verification_tracker) {
                    copied_source_hash = verification_tracker->finish(descriptor.length);
                }
                verify_file_with_retries(descriptor.full_path, working_path,
                                         descriptor.relative_path, verification, cancel,
                                         copied_source_hash);
            }
        }

        preserve_file_metadata(descriptor, working_path, source_security, source_streams, cancel,
                               recovered_any ? models::VerificationMode::None : verification);
        ensure_source_stability_unchanged(descriptor, source_stability,
                                          source_stability_valid);

        ensure_media_identity_integrity(utf8_to_wide(options_.SourceRoot), destination_root,
                                        true, true);
        publish_stage(stage, descriptor, recovered_any, &entry);
        managed_copy_files_ += 1;
        return recovered_any;
    }

    static void ensure_source_identity_unchanged(const SourceFileDescriptor& descriptor,
                                                 const FileIdentity& expected,
                                                 bool expected_valid) {
        if (!expected_valid) return;
        FileIdentity current;
        if (!try_get_file_identity(descriptor.full_path, current) ||
            current.file_index != expected.file_index ||
            current.volume_serial != expected.volume_serial) {
            throw IoError("Source file identity changed during copy: " + descriptor.relative_path + ".");
        }
    }

    static void ensure_source_stability_unchanged(
        const SourceFileDescriptor& descriptor, const FileStabilitySnapshot& expected,
        bool expected_valid) {
        if (!expected_valid) return;
        FileStabilitySnapshot current;
        if (!try_get_file_stability(descriptor.full_path, current) ||
            !file_stability_matches(expected, current)) {
            throw IoError("Source content or metadata changed during copy: " +
                          descriptor.relative_path + ".");
        }
    }

    bool add_created_directory(const std::wstring& directory) {
        std::wstring upper = storage::fsutil::to_upper_invariant(directory);
        for (const auto& existing : created_directories_) {
            if (existing == upper) return false;
        }
        created_directories_.push_back(upper);
        return true;
    }

    static void prepare_destination_file(const std::wstring& destination_path,
                                         std::int64_t expected_length, std::int64_t existing_length,
                                         bool preserve_existing) {
        if (existing_length < 0) existing_length = 0;
        if (existing_length == 0 || !preserve_existing) {
            HANDLE handle = CreateFileW(destination_path.c_str(), GENERIC_WRITE, FILE_SHARE_READ,
                                        nullptr, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, nullptr);
            if (handle == INVALID_HANDLE_VALUE) {
                throw IoError::from_win32("Unable to create destination file.", GetLastError());
            }
            CloseHandle(handle);
            return;
        }
        if (existing_length > expected_length) {
            HANDLE handle = CreateFileW(destination_path.c_str(), GENERIC_WRITE, FILE_SHARE_READ,
                                        nullptr, OPEN_ALWAYS, FILE_ATTRIBUTE_NORMAL, nullptr);
            if (handle == INVALID_HANDLE_VALUE) {
                throw IoError::from_win32("Unable to open destination file for trim.", GetLastError());
            }
            LARGE_INTEGER position;
            position.QuadPart = expected_length;
            SetFilePointerEx(handle, position, nullptr, FILE_BEGIN);
            SetEndOfFile(handle);
            CloseHandle(handle);
        }
    }

    static std::int32_t normalize_io_buffer_size(std::int32_t value) {
        return std::min(MaximumIoBufferSize, std::max(4096, value));
    }

    bool should_use_small_file_fast_path(std::int64_t file_length, bool has_persisted_coverage) const {
        std::int32_t threshold = std::max(MinimumRescueBlockSize, options_.SmallFileThresholdBytes);
        if (threshold <= 0 || has_persisted_coverage) return false;
        return file_length > 0 && file_length <= threshold;
    }

    bool is_native_acceleration_allowed() const {
        if (options_.TransferEnginePolicyValue == models::TransferEnginePolicy::ManagedRescue) return false;
        if (fault_injector_.has_value()) return false;
        // Auto must honor the configured managed retry/rescue policy. Native
        // acceleration remains available as an explicit NativeFast choice.
        if (options_.TransferEnginePolicyValue == models::TransferEnginePolicy::Auto &&
            (options_.MaxRetries > 0 || options_.SalvageUnreadableBlocks)) {
            return false;
        }
        // CopyFileExW is a single opaque operation: it cannot participate in
        // the managed media-identity, lock-wait, fragile-media, or source-
        // mutation policies. Those options must force the retry-aware path.
        if (options_.FragileMediaMode || options_.WaitForMediaAvailability ||
            options_.WaitForFileLockRelease || options_.TreatAccessDeniedAsContention ||
            options_.SourceMutationPolicyValue != models::SourceMutationPolicy::FailFile) {
            return false;
        }
        return true;
    }

    std::string preserve_directory_metadata(
        const std::wstring& source_root, const std::wstring& destination_root,
        const std::vector<std::string>& relative_directories,
        const CancelContext& cancel) {
        std::vector<std::string> deepest_first = relative_directories;
        std::stable_sort(deepest_first.begin(), deepest_first.end(),
                         [](const std::string& left, const std::string& right) {
                             const auto depth = [](const std::string& path) {
                                 return static_cast<std::size_t>(
                                     std::count(path.begin(), path.end(), '\\'));
                             };
                             const std::size_t left_depth = depth(left);
                             const std::size_t right_depth = depth(right);
                             if (left_depth != right_depth) return left_depth > right_depth;
                             return left.size() > right.size();
                         });

        for (const auto& relative : deepest_first) {
            try {
                cancel.throw_if_cancelled();
                if (control_ != nullptr) control_->wait_if_paused(cancel);
                const std::wstring source_path =
                    detail::trim_trailing_separators(source_root) + L"\\" + utf8_to_wide(relative);
                const std::wstring destination_path =
                    detail::trim_trailing_separators(destination_root) + L"\\" + utf8_to_wide(relative);

                WIN32_FILE_ATTRIBUTE_DATA source_data{};
                if (!GetFileAttributesExW(source_path.c_str(), GetFileExInfoStandard,
                                          &source_data) ||
                    (source_data.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) == 0) {
                    throw IoError::from_win32(
                        "Unable to inspect source directory metadata.", GetLastError());
                }
                NamedStreamEnumeration directory_streams =
                    enumerate_named_streams(source_path);
                if (directory_streams.supported && !directory_streams.streams.empty()) {
                    add_metadata_notice(
                        "Alternate data streams attached to directories were not copied.");
                }

                if (options_.PreserveTimestamps &&
                    !set_file_times_utc(destination_path, source_data.ftCreationTime,
                                        source_data.ftLastAccessTime,
                                        source_data.ftLastWriteTime)) {
                    throw IoError::from_win32(
                        "Unable to preserve source directory timestamps.", GetLastError());
                }
                if (!set_copyable_file_attributes(destination_path,
                                                  source_data.dwFileAttributes)) {
                    throw IoError::from_win32(
                        "Unable to preserve source directory attributes.", GetLastError());
                }
                // Apply permissions only after timestamp work. A restrictive
                // source DACL can legitimately deny further metadata writes.
                FileSecuritySnapshot security = capture_file_security(source_path);
                FileSecurityApplyResult security_result =
                    apply_file_security(destination_path, security);
                if (security.sacl_omitted || !security_result.sacl_preserved) {
                    add_metadata_notice(
                        "Security audit entries (SACLs) were not fully preserved for one or more directories.");
                }
                if (!security_result.owner_group_preserved) {
                    add_metadata_notice(
                        "Security owner/group were not fully preserved for one or more directories; the source DACL was applied.");
                }
            } catch (const OperationCanceled&) {
                throw;
            } catch (const std::exception& ex) {
                return "Directory metadata failed for " + relative + ": " + ex.what();
            }
        }
        return std::string();
    }

    bool should_use_native_copy_fast_path(std::int64_t file_length, std::int64_t destination_length,
                                          bool has_persisted_coverage) const {
        if (!is_native_acceleration_allowed()) return false;
        if (file_length <= 0) return false;
        if (destination_length > 0 || has_persisted_coverage) return false;
        if (options_.MaxThroughputBytesPerSecond > 0) return false;
        if (options_.UseBadRangeMap && options_.SkipKnownBadRanges) return false;
        return options_.OperationMode == models::JobOperationMode::Copy;
    }

    // ---- Native fast path -------------------------------------------------

    struct NativeAttempt {
        bool attempted = false;
        bool succeeded = false;
        std::string fallback_reason;
    };

    struct NativeCallbackState {
        ResilientCopyEngine* engine;
        const SourceFileDescriptor* descriptor;
        JournalFileEntry* entry;
        ProgressAccumulator* progress;
        const CancelContext* cancel;
        std::int64_t baseline_bytes;
        std::int64_t last_transferred;
        bool cancelled_for_pause;
    };

    static DWORD CALLBACK native_progress_routine(
        LARGE_INTEGER /*total*/, LARGE_INTEGER transferred, LARGE_INTEGER, LARGE_INTEGER,
        DWORD, DWORD, HANDLE, HANDLE, LPVOID context) {
        auto* state = static_cast<NativeCallbackState*>(context);
        if (state->cancel->is_cancelled() || state->cancel->deadline_passed()) {
            return PROGRESS_CANCEL;
        }
        if (state->engine->control_ != nullptr && state->engine->control_->is_paused()) {
            state->cancelled_for_pause = true;
            return PROGRESS_CANCEL;
        }

        std::int64_t bounded = std::clamp<std::int64_t>(transferred.QuadPart, 0,
                                                        state->descriptor->length);
        std::int64_t chunk = bounded - state->last_transferred;
        if (chunk < 0) chunk = 0;
        state->last_transferred = bounded;
        if (chunk > 0) {
            // CopyFileEx reports bytes as it completes them. Count those real
            // source reads/stage writes even if the opaque native attempt later
            // fails and falls back to Managed Rescue.
            state->engine->bytes_read_.fetch_add(chunk);
            state->engine->bytes_written_.fetch_add(chunk);
        }
        state->entry->BytesCopied = bounded;
        state->progress->total_bytes_copied = state->baseline_bytes + bounded;
        state->engine->emit_progress(
            *state->progress, state->descriptor->relative_path, bounded, state->descriptor->length,
            static_cast<std::int32_t>(std::min<std::int64_t>(INT32_MAX, chunk)),
            state->engine->resolve_buffer_size_for_file(state->descriptor->length),
            "NativeFastPath", 0, 0);
        return PROGRESS_CONTINUE;
    }

    NativeAttempt try_copy_with_native_fast_path(
        const SourceFileDescriptor& descriptor, JournalFileEntry& entry,
        const std::wstring& destination_path, std::int64_t destination_length,
        bool has_persisted_coverage, ProgressAccumulator& progress, JobJournal& journal,
        const std::wstring& journal_path, const CancelContext& cancel) {

        NativeAttempt result;
        if (!should_use_native_copy_fast_path(descriptor.length, destination_length, has_persisted_coverage)) {
            return result;
        }

        result.attempted = true;
        NativeCallbackState state{this, &descriptor, &entry, &progress, &cancel,
                                  progress.total_bytes_copied, 0, false};

        DWORD flags = options_.AllowDecryptedDestination
                          ? COPY_FILE_ALLOW_DECRYPTED_DESTINATION
                          : 0;
        if (descriptor.length >= 16LL * 1024 * 1024) flags |= COPY_FILE_NO_BUFFERING;

        BOOL cancel_flag = FALSE;
        BOOL ok = CopyFileExW(descriptor.full_path.c_str(), destination_path.c_str(),
                              native_progress_routine, &state, &cancel_flag, flags);
        DWORD last_error = ok ? 0 : GetLastError();

        if (ok) {
            result.succeeded = true;
            const std::int64_t unreported =
                std::max<std::int64_t>(0, descriptor.length - state.last_transferred);
            if (unreported > 0) {
                bytes_read_.fetch_add(unreported);
                bytes_written_.fetch_add(unreported);
            }
            progress.total_bytes_copied = state.baseline_bytes + descriptor.length;
            entry.BytesCopied = descriptor.length;
            entry.LastError.clear();
            entry.LastRescuePass = "NativeFastPath";
            entry.RecoveredRanges.clear();
            entry.RescueRanges.clear();
            append_range(entry.RescueRanges, 0, descriptor.length, RescueRangeState::Good);
            sync_entry_bytes_copied(entry);
            emit_progress(progress, descriptor.relative_path,
                          display_progress_bytes(entry, descriptor.length, entry.BytesCopied),
                          descriptor.length,
                          static_cast<std::int32_t>(std::min<std::int64_t>(
                              INT32_MAX, std::max<std::int64_t>(0, descriptor.length - state.last_transferred))),
                          resolve_buffer_size_for_file(descriptor.length), "NativeFastPath", 0, 0);
            flush_journal(journal_path, journal, false);
            native_fast_path_files_ += 1;
            return result;
        }

        progress.total_bytes_copied = state.baseline_bytes;
        entry.BytesCopied = 0;
        native_fallback_files_ += 1;

        cancel.throw_if_cancelled();
        if (state.cancelled_for_pause) {
            result.fallback_reason = "paused by user during native fast-path copy.";
            return result;
        }
        result.fallback_reason = "CopyFileEx failed (" + std::to_string(last_error) + ").";
        return result;
    }

    // ---- Rescue-range bookkeeping ----------------------------------------

    static RescueRangeState normalize_state(RescueRangeState value) {
        switch (value) {
            case RescueRangeState::Pending:
            case RescueRangeState::Good:
            case RescueRangeState::Bad:
            case RescueRangeState::Recovered:
            case RescueRangeState::KnownBad:
                return value;
            default:
                return RescueRangeState::Pending;
        }
    }

    static void append_range(std::vector<RescueRange>& target, std::int64_t offset,
                             std::int64_t length, RescueRangeState state) {
        if (length <= 0) return;
        target.push_back(RescueRange{offset, length, normalize_state(state)});
    }

    static std::vector<RescueRange> merge_rescue_ranges(const std::vector<RescueRange>& ranges) {
        std::vector<RescueRange> sorted;
        for (const auto& item : ranges) {
            if (item.Length > 0) sorted.push_back(item);
        }
        std::stable_sort(sorted.begin(), sorted.end(),
                         [](const RescueRange& a, const RescueRange& b) { return a.Offset < b.Offset; });

        std::vector<RescueRange> merged;
        for (const auto& item : sorted) {
            RescueRangeState state = normalize_state(item.State);
            std::int64_t start = item.Offset;
            std::int64_t end = start + item.Length;
            if (end <= start) continue;

            if (merged.empty()) {
                merged.push_back(RescueRange{start, item.Length, state});
                continue;
            }
            RescueRange& previous = merged.back();
            std::int64_t previous_end = previous.Offset + previous.Length;
            if (previous.State == state && previous_end == start) {
                previous.Length += item.Length;
            } else if (previous_end > start) {
                std::int64_t adjusted_start = previous_end;
                if (end > adjusted_start) {
                    append_range(merged, adjusted_start, end - adjusted_start, state);
                }
            } else {
                merged.push_back(RescueRange{start, item.Length, state});
            }
        }
        return merged;
    }

    static std::vector<RescueRange> normalize_rescue_ranges(const std::vector<RescueRange>& ranges,
                                                            std::int64_t file_length) {
        std::vector<RescueRange> normalized;
        if (file_length <= 0) return normalized;

        std::vector<RescueRange> ordered;
        for (const auto& item : ranges) {
            if (item.Length <= 0) continue;
            std::int64_t start = std::max<std::int64_t>(0, item.Offset);
            std::int64_t end = std::min(file_length, start + item.Length);
            if (end <= start) continue;
            ordered.push_back(RescueRange{start, end - start, normalize_state(item.State)});
        }
        std::stable_sort(ordered.begin(), ordered.end(),
                         [](const RescueRange& a, const RescueRange& b) { return a.Offset < b.Offset; });

        std::int64_t cursor = 0;
        for (const auto& item : ordered) {
            std::int64_t item_start = std::max(cursor, item.Offset);
            std::int64_t item_end = std::min(file_length, item.Offset + item.Length);
            if (item_end <= item_start) continue;
            if (item_start > cursor) {
                append_range(normalized, cursor, item_start - cursor, RescueRangeState::Pending);
            }
            append_range(normalized, item_start, item_end - item_start, item.State);
            cursor = item_end;
        }
        if (cursor < file_length) {
            append_range(normalized, cursor, file_length - cursor, RescueRangeState::Pending);
        }
        return merge_rescue_ranges(normalized);
    }

    static std::vector<RescueRange> downgrade_coverage_beyond_destination(
        const std::vector<RescueRange>& ranges, std::int64_t destination_length,
        std::int64_t file_length) {
        if (file_length <= 0) return {};
        std::int64_t effective = std::max<std::int64_t>(0, std::min(destination_length, file_length));
        std::vector<RescueRange> adjusted;
        for (const auto& item : ranges) {
            std::int64_t item_start = std::max<std::int64_t>(0, item.Offset);
            std::int64_t item_end = std::min(file_length, item_start + item.Length);
            if (item_end <= item_start) continue;
            RescueRangeState state = normalize_state(item.State);
            bool satisfied = state == RescueRangeState::Good || state == RescueRangeState::Recovered;

            if (satisfied && item_start >= effective) {
                append_range(adjusted, item_start, item_end - item_start, RescueRangeState::Pending);
                continue;
            }
            if (satisfied && item_end > effective) {
                if (effective > item_start) {
                    append_range(adjusted, item_start, effective - item_start, state);
                }
                append_range(adjusted, effective, item_end - effective, RescueRangeState::Pending);
                continue;
            }
            append_range(adjusted, item_start, item_end - item_start, state);
        }
        return normalize_rescue_ranges(adjusted, file_length);
    }

    std::vector<RescueRange> build_rescue_ranges(const JournalFileEntry& entry,
                                                 std::int64_t file_length,
                                                 std::int64_t destination_length) const {
        if (file_length <= 0) return {};

        std::vector<RescueRange> ranges;
        if (!entry.RescueRanges.empty()) {
            for (const auto& range : entry.RescueRanges) {
                ranges.push_back(RescueRange{range.Offset, range.Length, normalize_state(range.State)});
            }
        } else {
            std::int64_t legacy = 0;
            if (options_.ResumeFromJournal) {
                legacy = std::max<std::int64_t>(
                    0, std::min(file_length, std::min(destination_length, entry.BytesCopied)));
            }
            if (legacy > 0) append_range(ranges, 0, legacy, RescueRangeState::Good);
            if (legacy < file_length) {
                append_range(ranges, legacy, file_length - legacy, RescueRangeState::Pending);
            }
        }
        return downgrade_coverage_beyond_destination(
            normalize_rescue_ranges(ranges, file_length), destination_length, file_length);
    }

    static std::int64_t rescue_bytes(const std::vector<RescueRange>& ranges,
                                     std::initializer_list<RescueRangeState> states) {
        std::int64_t total = 0;
        for (const auto& item : ranges) {
            if (item.Length <= 0) continue;
            RescueRangeState state = normalize_state(item.State);
            for (RescueRangeState allowed : states) {
                if (state == allowed) {
                    total += item.Length;
                    break;
                }
            }
        }
        return total;
    }

    static std::int32_t region_count(const std::vector<RescueRange>& ranges,
                                     std::initializer_list<RescueRangeState> states) {
        std::int32_t count = 0;
        for (const auto& item : ranges) {
            if (item.Length <= 0) continue;
            RescueRangeState state = normalize_state(item.State);
            for (RescueRangeState allowed : states) {
                if (state == allowed) {
                    ++count;
                    break;
                }
            }
        }
        return count;
    }

    static std::int32_t unreadable_region_count(const std::vector<RescueRange>& ranges) {
        return region_count(ranges, {RescueRangeState::Bad, RescueRangeState::KnownBad});
    }

    static std::vector<ByteRange> snapshot_by_state(const std::vector<RescueRange>& ranges,
                                                    std::initializer_list<RescueRangeState> states) {
        std::vector<ByteRange> snapshots;
        for (const auto& item : ranges) {
            if (item.Length <= 0) continue;
            RescueRangeState state = normalize_state(item.State);
            for (RescueRangeState allowed : states) {
                if (state == allowed) {
                    snapshots.push_back(ByteRange{item.Offset, item.Length});
                    break;
                }
            }
        }
        return snapshots;
    }

    static void set_range_state(std::vector<RescueRange>& ranges, std::int64_t offset,
                                std::int64_t length, RescueRangeState new_state) {
        if (ranges.empty() || length <= 0) return;
        std::int64_t target_start = offset;
        std::int64_t target_end = offset + length;
        if (target_end <= target_start) return;
        RescueRangeState normalized = normalize_state(new_state);

        std::vector<RescueRange> rebuilt;
        rebuilt.reserve(ranges.size() + 2);
        bool affected = false;
        for (const auto& item : ranges) {
            if (item.Length <= 0) continue;
            std::int64_t item_start = item.Offset;
            std::int64_t item_end = item_start + item.Length;
            if (item_end <= target_start || item_start >= target_end) {
                rebuilt.push_back(item);
                continue;
            }
            affected = true;
            if (item_start < target_start) {
                append_range(rebuilt, item_start, target_start - item_start, item.State);
            }
            std::int64_t overlap_start = std::max(item_start, target_start);
            std::int64_t overlap_end = std::min(item_end, target_end);
            if (overlap_end > overlap_start) {
                append_range(rebuilt, overlap_start, overlap_end - overlap_start, normalized);
            }
            if (item_end > target_end) {
                append_range(rebuilt, target_end, item_end - target_end, item.State);
            }
        }
        if (!affected) return;
        ranges = merge_rescue_ranges(rebuilt);
    }

    static std::int64_t display_progress_bytes(const JournalFileEntry& entry, std::int64_t file_length,
                                               std::int64_t fallback_bytes) {
        std::int64_t safe_length = std::max<std::int64_t>(0, file_length);
        std::int64_t display = std::max<std::int64_t>(0, fallback_bytes);
        if (!entry.RescueRanges.empty()) {
            std::int64_t accounted = rescue_bytes(
                entry.RescueRanges,
                {RescueRangeState::Good, RescueRangeState::Recovered, RescueRangeState::Bad,
                 RescueRangeState::KnownBad});
            display = std::max(display, accounted);
        }
        return std::max<std::int64_t>(0, std::min(safe_length, display));
    }

    void sync_entry_bytes_copied(JournalFileEntry& entry) const {
        if (options_.OperationMode == models::JobOperationMode::ScanOnly) {
            entry.BytesCopied = rescue_bytes(
                entry.RescueRanges, {RescueRangeState::Good, RescueRangeState::Recovered,
                                     RescueRangeState::Bad, RescueRangeState::KnownBad});
            return;
        }
        entry.BytesCopied = rescue_bytes(entry.RescueRanges,
                                         {RescueRangeState::Good, RescueRangeState::Recovered});
    }

    static std::vector<ByteRange> merge_byte_ranges(const std::vector<ByteRange>& ranges) {
        std::vector<ByteRange> sorted;
        for (const auto& item : ranges) {
            if (item.Length > 0) sorted.push_back(item);
        }
        std::stable_sort(sorted.begin(), sorted.end(),
                         [](const ByteRange& a, const ByteRange& b) { return a.Offset < b.Offset; });

        std::vector<ByteRange> merged;
        for (const auto& item : sorted) {
            std::int64_t start = item.Offset;
            std::int64_t end = start + item.Length;
            if (end <= start) continue;
            if (merged.empty()) {
                merged.push_back(item);
                continue;
            }
            ByteRange& last = merged.back();
            std::int64_t last_end = last.Offset + last.Length;
            if (start <= last_end && end > last_end) {
                last.Length = end - last.Offset;
            } else if (start > last_end) {
                merged.push_back(item);
            }
        }
        return merged;
    }

    // ---- Rescue pass plan -------------------------------------------------

    static std::int32_t align_to_block(std::int32_t value, std::int32_t block_size) {
        std::int32_t normalized_block = std::max(MinimumRescueBlockSize, block_size);
        std::int32_t aligned = (value / normalized_block) * normalized_block;
        return aligned <= 0 ? normalized_block : aligned;
    }

    std::vector<RescuePassDefinition> build_rescue_pass_plan(std::int32_t active_buffer_size) const {
        std::int32_t fast_auto = normalize_io_buffer_size(active_buffer_size);
        std::int32_t trim_auto = normalize_io_buffer_size(
            std::max(MinimumRescueBlockSize * 16, fast_auto / 8));
        std::int32_t scrape_auto = normalize_io_buffer_size(
            std::max(MinimumRescueBlockSize * 2, trim_auto / 4));
        std::int32_t retry_auto = normalize_io_buffer_size(MinimumRescueBlockSize);

        auto chunk = [](std::int32_t configured, std::int32_t auto_value) {
            return configured <= 0 ? normalize_io_buffer_size(auto_value)
                                   : normalize_io_buffer_size(configured);
        };
        std::int32_t fast_chunk = chunk(options_.RescueFastScanChunkBytes, fast_auto);
        std::int32_t trim_chunk = chunk(options_.RescueTrimChunkBytes, trim_auto);
        std::int32_t scrape_chunk = chunk(options_.RescueScrapeChunkBytes, scrape_auto);
        std::int32_t retry_chunk = chunk(options_.RescueRetryChunkBytes, retry_auto);

        std::int32_t split_minimum = options_.RescueSplitMinimumBytes <= 0
                                         ? normalize_io_buffer_size(retry_chunk)
                                         : normalize_io_buffer_size(options_.RescueSplitMinimumBytes);
        auto cap_split = [](std::int32_t split, std::int32_t pass_chunk) {
            std::int32_t half = std::max(MinimumRescueBlockSize, pass_chunk / 2);
            return std::max(MinimumRescueBlockSize, std::min(split, half));
        };
        std::int32_t trim_split = cap_split(split_minimum, trim_chunk);
        std::int32_t scrape_split = cap_split(split_minimum, scrape_chunk);

        auto retries = [](std::int32_t configured, std::int32_t fallback) {
            return configured < 0 ? std::max(0, fallback) : configured;
        };
        std::int32_t fast_retries = retries(options_.RescueFastScanRetries, 0);
        std::int32_t trim_retries = retries(options_.RescueTrimRetries, 1);
        std::int32_t scrape_retries = retries(options_.RescueScrapeRetries, 2);

        return {
            {"FastScan", RescueRangeState::Pending, fast_chunk, fast_retries, false, trim_split, false},
            {"TrimSweep", RescueRangeState::Bad, trim_chunk, trim_retries, true, trim_split, false},
            {"TrimSweepReverse", RescueRangeState::Bad, trim_chunk,
             std::max(trim_retries, scrape_retries), true, trim_split, true},
            {"Scrape", RescueRangeState::Bad, scrape_chunk, scrape_retries, true, scrape_split, false},
            {"RetryBad", RescueRangeState::Bad, retry_chunk, std::max(0, options_.MaxRetries), false,
             retry_chunk, true}};
    }

    static double rescue_bad_density(const std::vector<RescueRange>& ranges, std::int64_t total_bytes) {
        if (total_bytes <= 0) return 0.0;
        std::int64_t bad = rescue_bytes(ranges, {RescueRangeState::Bad, RescueRangeState::KnownBad});
        double density = static_cast<double>(bad) / static_cast<double>(total_bytes);
        return std::clamp(density, 0.0, 1.0);
    }

    RescuePassDefinition adapt_pass_for_density(const RescuePassDefinition& pass, double bad_density,
                                                std::int32_t active_buffer_size) const {
        if (pass.target_state != RescueRangeState::Bad) return pass;

        std::int32_t tuned_chunk = pass.chunk_size_bytes;
        std::int32_t tuned_retries = pass.max_read_retries;
        if (bad_density >= 0.25) {
            tuned_chunk = std::max(
                pass.minimum_split_bytes,
                align_to_block(std::max(MinimumRescueBlockSize, pass.chunk_size_bytes / 2),
                               pass.minimum_split_bytes));
            tuned_retries = std::min(32, pass.max_read_retries + 2);
        } else if (bad_density <= 0.02) {
            std::int32_t ceiling = normalize_io_buffer_size(
                std::max(MinimumRescueBlockSize, active_buffer_size));
            std::int64_t boosted64 = static_cast<std::int64_t>(pass.chunk_size_bytes) * 2;
            std::int32_t boosted = normalize_io_buffer_size(
                static_cast<std::int32_t>(std::min<std::int64_t>(INT32_MAX, boosted64)));
            tuned_chunk = std::min(ceiling, boosted);
        }
        if (tuned_chunk == pass.chunk_size_bytes && tuned_retries == pass.max_read_retries) return pass;

        RescuePassDefinition tuned = pass;
        tuned.chunk_size_bytes = tuned_chunk;
        tuned.max_read_retries = tuned_retries;
        return tuned;
    }

    // ---- Rescue execution -------------------------------------------------

    static std::int32_t resolve_rescue_pass_chunk_size(const RescuePassDefinition& pass,
                                                       AdaptiveBufferController* controller) {
        if (controller != nullptr && controller->is_dynamic() &&
            models::detail::equals_ignore_case(pass.name, "FastScan")) {
            return controller->current_size();
        }
        return pass.chunk_size_bytes;
    }

    void report_pass_buffer_decision(const std::string& relative_path,
                                     const RescuePassDefinition& pass, const BufferDecision& decision) {
        // Only shrink decisions are logged; routine growths stay silent (the
        // .NET worker batches those into a periodic summary).
        if (!decision.changed) return;
        if (decision.current_size >= decision.previous_size) return;
        emit_log("[AdaptiveBuffer] " + pass.name + " " + relative_path + ": " +
                 format_bytes(decision.previous_size) + " -> " + format_bytes(decision.current_size) +
                 " (" + decision.reason + ").");
    }

    RescuePassOutcome execute_rescue_pass(const RescuePassDefinition& pass,
                                          const SourceFileDescriptor& descriptor,
                                          JournalFileEntry& entry, FileTransferSession& session,
                                          std::vector<unsigned char>& io_buffer,
                                          ProgressAccumulator& progress, JobJournal& journal,
                                          const std::wstring& journal_path,
                                          const CancelContext& cancel,
                                           bool write_recovered = true,
                                           bool count_failed_as_processed = false,
                                           AdaptiveBufferController* controller = nullptr,
                                           CopyVerificationTracker* verification_tracker = nullptr) {
        RescuePassOutcome outcome;
        const bool adaptive_pass = models::detail::equals_ignore_case(pass.name, "FastScan");
        std::vector<ByteRange> targets = snapshot_by_state(entry.RescueRanges, {pass.target_state});
        if (pass.process_descending) std::reverse(targets.begin(), targets.end());

        std::wstring source_root = utf8_to_wide(options_.SourceRoot);
        std::wstring destination_root = utf8_to_wide(options_.DestinationRoot);

        for (const auto& target : targets) {
            cancel.throw_if_cancelled();
            if (control_ != nullptr) control_->wait_if_paused(cancel);
            wait_for_media_availability(source_root, destination_root, cancel);

            std::vector<ByteRange> work;
            work.push_back(target);

            while (!work.empty()) {
                cancel.throw_if_cancelled();
                if (control_ != nullptr) control_->wait_if_paused(cancel);
                wait_for_media_availability(source_root, destination_root, cancel);

                ByteRange segment = work.back();
                work.pop_back();
                if (segment.Length <= 0) continue;

                std::int64_t segment_length = segment.Length;
                std::int32_t effective_chunk = resolve_rescue_pass_chunk_size(pass, controller);
                if (segment_length > effective_chunk) {
                    // Take exactly one chunk; push the continuation first so the
                    // chunk pops next (LIFO), preserving directional order.
                    std::int64_t chunk_length = effective_chunk;
                    if (pass.process_descending) {
                        std::int64_t chunk_offset = segment.Offset + segment_length - chunk_length;
                        if (segment_length - chunk_length > 0) {
                            work.push_back(ByteRange{segment.Offset, segment_length - chunk_length});
                        }
                        work.push_back(ByteRange{chunk_offset, chunk_length});
                    } else {
                        if (segment_length - chunk_length > 0) {
                            work.push_back(ByteRange{segment.Offset + chunk_length,
                                                     segment_length - chunk_length});
                        }
                        work.push_back(ByteRange{segment.Offset, chunk_length});
                    }
                    continue;
                }

                outcome.attempted_segments += 1;
                std::int32_t read_length = static_cast<std::int32_t>(segment_length);
                ULONGLONG chunk_started = GetTickCount64();
                std::int32_t bytes_read = read_chunk_with_retries(
                    descriptor.full_path, descriptor.relative_path, segment.Offset, read_length,
                    session, io_buffer.data(), cancel, pass.max_read_retries, false);

                if (bytes_read > 0) {
                    if (write_recovered) {
                        if (verification_tracker != nullptr) {
                            verification_tracker->observe(segment.Offset, io_buffer.data(), bytes_read);
                        }
                        write_chunk_with_retries(descriptor.relative_path, segment.Offset,
                                                 io_buffer.data(), bytes_read, session, cancel);
                    }
                    if (adaptive_pass && controller != nullptr) {
                        report_pass_buffer_decision(
                            descriptor.relative_path, pass,
                            controller->report_success(
                                bytes_read, static_cast<double>(GetTickCount64() - chunk_started)));
                    }
                    apply_throughput_throttle(bytes_read, cancel);

                    set_range_state(entry.RescueRanges, segment.Offset, bytes_read, RescueRangeState::Good);
                    progress.total_bytes_copied += bytes_read;
                    sync_entry_bytes_copied(entry);
                    emit_progress(progress, descriptor.relative_path,
                                  display_progress_bytes(entry, descriptor.length, entry.BytesCopied),
                                  descriptor.length, bytes_read, effective_chunk, pass.name,
                                  unreadable_region_count(entry.RescueRanges),
                                  rescue_bytes(entry.RescueRanges,
                                               {RescueRangeState::Pending, RescueRangeState::Bad,
                                                RescueRangeState::KnownBad}));
                    flush_journal(journal_path, journal, false);
                    outcome.recovered_bytes += bytes_read;
                    continue;
                }

                if (adaptive_pass && controller != nullptr) {
                    report_pass_buffer_decision(descriptor.relative_path, pass,
                                                controller->report_failure());
                }

                if (pass.split_on_failure &&
                    segment_length >= static_cast<std::int64_t>(pass.minimum_split_bytes) * 2) {
                    std::int64_t half = std::max<std::int64_t>(
                        pass.minimum_split_bytes,
                        align_to_block(static_cast<std::int32_t>(std::min<std::int64_t>(
                                           INT32_MAX, segment_length / 2)),
                                       pass.minimum_split_bytes));
                    half = std::min(segment_length - pass.minimum_split_bytes, half);
                    if (half > 0 && half < segment_length) {
                        std::int64_t second = segment_length - half;
                        if (pass.process_descending) {
                            work.push_back(ByteRange{segment.Offset, half});
                            work.push_back(ByteRange{segment.Offset + half, second});
                        } else {
                            work.push_back(ByteRange{segment.Offset + half, second});
                            work.push_back(ByteRange{segment.Offset, half});
                        }
                        continue;
                    }
                }

                std::int32_t failed_length = static_cast<std::int32_t>(segment_length);
                set_range_state(entry.RescueRanges, segment.Offset, failed_length, RescueRangeState::Bad);
                if (count_failed_as_processed) {
                    progress.total_bytes_copied += failed_length;
                }
                sync_entry_bytes_copied(entry);
                emit_progress(progress, descriptor.relative_path,
                              display_progress_bytes(entry, descriptor.length, entry.BytesCopied),
                              descriptor.length, 0, effective_chunk, pass.name,
                              unreadable_region_count(entry.RescueRanges),
                              rescue_bytes(entry.RescueRanges,
                                           {RescueRangeState::Pending, RescueRangeState::Bad,
                                            RescueRangeState::KnownBad}));
                flush_journal(journal_path, journal, false);
                outcome.failed_segments += 1;
            }
        }
        return outcome;
    }

    bool salvage_remaining_ranges(const SourceFileDescriptor& descriptor, JournalFileEntry& entry,
                                  FileTransferSession& session, std::vector<unsigned char>& io_buffer,
                                  std::int32_t io_buffer_size, ProgressAccumulator& progress,
                                  JobJournal& journal, const std::wstring& journal_path,
                                  const CancelContext& cancel) {
        bool recovered_any = false;
        std::wstring source_root = utf8_to_wide(options_.SourceRoot);
        std::wstring destination_root = utf8_to_wide(options_.DestinationRoot);
        std::vector<ByteRange> bad = snapshot_by_state(
            entry.RescueRanges, {RescueRangeState::Bad, RescueRangeState::KnownBad});
        for (const auto& bad_range : bad) {
            std::int64_t offset = bad_range.Offset;
            std::int64_t remaining = bad_range.Length;
            while (remaining > 0) {
                cancel.throw_if_cancelled();
                if (control_ != nullptr) control_->wait_if_paused(cancel);
                wait_for_media_availability(source_root, destination_root, cancel);

                std::int32_t chunk = static_cast<std::int32_t>(
                    std::min<std::int64_t>(io_buffer_size, remaining));
                fill_salvage_buffer(io_buffer.data(), chunk);
                write_chunk_with_retries(descriptor.relative_path, offset, io_buffer.data(), chunk,
                                         session, cancel);
                apply_throughput_throttle(chunk, cancel);

                set_range_state(entry.RescueRanges, offset, chunk, RescueRangeState::Recovered);
                entry.RecoveredRanges.push_back(ByteRange{offset, chunk});
                progress.total_bytes_copied += chunk;
                sync_entry_bytes_copied(entry);
                emit_progress(progress, descriptor.relative_path,
                              display_progress_bytes(entry, descriptor.length, entry.BytesCopied),
                              descriptor.length, chunk, io_buffer_size, "SalvageFill",
                              unreadable_region_count(entry.RescueRanges),
                              rescue_bytes(entry.RescueRanges,
                                           {RescueRangeState::Pending, RescueRangeState::Bad,
                                            RescueRangeState::KnownBad}));
                flush_journal(journal_path, journal, false);

                offset += chunk;
                remaining -= chunk;
                recovered_any = true;
            }
        }
        return recovered_any;
    }

    bool copy_small_file_fast(const SourceFileDescriptor& descriptor, JournalFileEntry& entry,
                              const std::wstring& destination_path, FileTransferSession& session,
                              std::vector<unsigned char>& io_buffer,
                              AdaptiveBufferController& controller, ProgressAccumulator& progress,
                              JobJournal& journal, const std::wstring& journal_path,
                              const CancelContext& cancel,
                              CopyVerificationTracker* verification_tracker = nullptr) {
        (void)destination_path;
        entry.RescueRanges.clear();
        append_range(entry.RescueRanges, 0, descriptor.length, RescueRangeState::Pending);
        entry.RecoveredRanges.clear();
        entry.LastRescuePass = "SmallFileFast";
        entry.BytesCopied = 0;

        std::wstring source_root = utf8_to_wide(options_.SourceRoot);
        std::wstring destination_root = utf8_to_wide(options_.DestinationRoot);

        bool recovered_any = false;
        std::int64_t offset = 0;
        std::int64_t remaining = descriptor.length;
        while (remaining > 0) {
            cancel.throw_if_cancelled();
            if (control_ != nullptr) control_->wait_if_paused(cancel);
            wait_for_media_availability(source_root, destination_root, cancel);

            std::int32_t chunk_length = controller.next_chunk_length(remaining);
            ULONGLONG chunk_started = GetTickCount64();
            std::int32_t bytes_read = read_chunk_with_retries(
                descriptor.full_path, descriptor.relative_path, offset, chunk_length, session,
                io_buffer.data(), cancel, std::max(0, options_.MaxRetries), false);

            RescueRangeState segment_state = RescueRangeState::Good;
            bool read_succeeded = bytes_read > 0;
            if (bytes_read <= 0) {
                controller.report_failure();
                if (verification_tracker != nullptr) verification_tracker->invalidate();
                if (!options_.SalvageUnreadableBlocks) {
                    throw IoError("Read failed on " + descriptor.relative_path + " at offset " +
                                  std::to_string(offset) + ".");
                }
                bytes_read = chunk_length;
                fill_salvage_buffer(io_buffer.data(), bytes_read);
                segment_state = RescueRangeState::Recovered;
                recovered_any = true;
                entry.RecoveredRanges.push_back(ByteRange{offset, bytes_read});
                emit_log("Recovered unreadable block on " + descriptor.relative_path + " at " +
                         format_bytes(offset) + " (" + std::to_string(bytes_read) + " bytes " +
                         describe_salvage_fill() + "-filled).");
            }

            if (read_succeeded && verification_tracker != nullptr) {
                verification_tracker->observe(offset, io_buffer.data(), bytes_read);
            }
            write_chunk_with_retries(descriptor.relative_path, offset, io_buffer.data(), bytes_read,
                                     session, cancel);
            if (read_succeeded) {
                controller.report_success(bytes_read,
                                          static_cast<double>(GetTickCount64() - chunk_started));
            }
            apply_throughput_throttle(bytes_read, cancel);

            set_range_state(entry.RescueRanges, offset, bytes_read, segment_state);
            progress.total_bytes_copied += bytes_read;
            sync_entry_bytes_copied(entry);
            emit_progress(progress, descriptor.relative_path,
                          display_progress_bytes(entry, descriptor.length, entry.BytesCopied),
                          descriptor.length, bytes_read, chunk_length, entry.LastRescuePass,
                          unreadable_region_count(entry.RescueRanges),
                          rescue_bytes(entry.RescueRanges,
                                       {RescueRangeState::Pending, RescueRangeState::Bad,
                                        RescueRangeState::KnownBad}));
            offset += bytes_read;
            remaining -= bytes_read;
        }

        entry.RecoveredRanges = merge_byte_ranges(entry.RecoveredRanges);
        sync_entry_bytes_copied(entry);
        flush_journal(journal_path, journal, true);
        return recovered_any;
    }

    // ---- Chunk I/O with retries ------------------------------------------

    std::int64_t operation_timeout_ms() const {
        return options_.OperationTimeout.ticks / time::TicksPerMillisecond;
    }

    // Returns bytes read (== length), or salvage-filled length, or -1.
    std::int32_t read_chunk_with_retries(const std::wstring& source_path,
                                         const std::string& relative_path, std::int64_t offset,
                                         std::int32_t length, FileTransferSession& session,
                                         unsigned char* buffer, const CancelContext& cancel,
                                         std::int32_t max_retries, bool allow_salvage,
                                         bool primary_data_range = true) {
        std::int32_t attempt = 0;
        std::optional<IoError> last_error;
        std::int32_t effective_max = max_retries >= 0 ? max_retries : options_.MaxRetries;

        while (attempt <= effective_max) {
            cancel.throw_if_cancelled();
            throw_if_media_identity_mismatch_throttled(source_path, true);

            try {
                if (fault_injector_.has_value()) {
                    if (auto fault = fault_injector_->create_read_fault(relative_path, offset, length)) {
                        throw *fault;
                    }
                }

                std::int32_t bytes_read = 0;
                if (raw_disk_scan_context_) {
                    RawDiskReadResult raw = raw_disk_scan_context_->read_chunk(
                        source_path, relative_path, offset, length, buffer,
                        operation_timeout_ms(), cancel);
                    if (raw.state == RawDiskReadResult::State::Success) {
                        bytes_read = raw.BytesRead;
                    } else if (raw.state == RawDiskReadResult::State::Failure && raw.Error.has_value()) {
                        throw *raw.Error;
                    } else {
                        emit_raw_disk_fallback_once(source_path, raw.FallbackReason);
                        bytes_read = timed_io::read_at(
                            session.source_handle(), offset, buffer, length,
                            operation_timeout_ms(), cancel);
                    }
                } else {
                    bytes_read = timed_io::read_at(
                        session.source_handle(), offset, buffer, length, operation_timeout_ms(), cancel);
                }
                if (bytes_read != length) {
                    throw IoError("Short read at offset " + std::to_string(offset) + ". Expected " +
                                  std::to_string(length) + ", got " + std::to_string(bytes_read) + ".");
                }
                bytes_read_.fetch_add(bytes_read);
                return bytes_read;
            } catch (const IoError& ex) {
                if (ex.kind == IoErrorKind::Timeout) {
                    last_error = IoError("Read timeout at offset " + std::to_string(offset) + ".",
                                         0, IoErrorKind::Timeout);
                } else {
                    last_error = ex;
                }
            }

            if (is_source_file_missing(*last_error)) {
                switch (options_.SourceMutationPolicyValue) {
                    case models::SourceMutationPolicy::SkipFile:
                        throw SourceMutationSkipped("Source file disappeared during copy: " +
                                                    relative_path + ".");
                    case models::SourceMutationPolicy::WaitForReappearance:
                        session.invalidate_source();
                        emit_log("Source file disappeared during copy of " + relative_path +
                                 ". Waiting for it to reappear.");
                        wait_for_source_file(source_path, cancel, true);
                        attempt = 0;
                        continue;
                    default:
                        throw IoError("Source file disappeared during copy: " + relative_path + ".");
                }
            }

            if (is_read_contention(*last_error) && options_.WaitForFileLockRelease) {
                session.invalidate_source();
                emit_log("Source contention detected on " + relative_path + "; waiting for lock release.");
                wait_for_source_read_access(source_path, cancel);
                attempt = 0;
                continue;
            }

            if (is_fatal_read(*last_error)) {
                if (options_.FragileMediaMode && options_.SkipFileOnFirstReadError &&
                    !is_access_denied(*last_error)) {
                    session.invalidate_source();
                    throw FragileReadSkip(
                        "Fragile mode: first read failure at " + format_bytes(offset) +
                            " (" + last_error->what() + ").",
                        offset, length, primary_data_range);
                }
                break;
            }

            if (options_.WaitForMediaAvailability && is_availability_related(*last_error, false)) {
                session.invalidate_source();
                emit_log("Source unavailable during read of " + relative_path +
                         ". Waiting for media to return.");
                wait_for_source_file(source_path, cancel, false);
                attempt = 0;
                continue;
            }

            // A missing/locked file, access denial, or unavailable device is
            // not evidence that this byte range is physically unreadable.
            // Classify those conditions first so fragile mode records only an
            // actual data-read failure (including a read timeout/short read).
            const bool non_media_read_failure =
                is_read_contention(*last_error) ||
                is_availability_related(*last_error, true);
            if (options_.FragileMediaMode && options_.SkipFileOnFirstReadError &&
                !non_media_read_failure) {
                session.invalidate_source();
                const std::string error_text = last_error->what();
                throw FragileReadSkip(
                    "Fragile mode: first read failure at " + format_bytes(offset) +
                        " (" + error_text + ").",
                    offset, length, primary_data_range);
            }

            session.invalidate_source();
            if (is_read_contention(*last_error)) {
                emit_log("Source file contention detected on " + relative_path + "; retrying.");
            }

            attempt += 1;
            if (attempt > effective_max) break;
            emit_log("Read retry " + std::to_string(attempt) + "/" + std::to_string(effective_max) +
                     " on " + relative_path + " at " + format_bytes(offset) + ": " + last_error->what());
            delay_for_retry(attempt, cancel);
        }

        if (last_error.has_value() && is_fatal_read(*last_error)) {
            throw IoError("Read failed on " + relative_path + " at offset " + std::to_string(offset) + ".");
        }
        if (last_error.has_value() && is_availability_related(*last_error, true)) {
            throw IoError("Read failed on " + relative_path + " at offset " + std::to_string(offset) +
                          " because source is unavailable.");
        }
        if (allow_salvage && options_.SalvageUnreadableBlocks) {
            fill_salvage_buffer(buffer, length);
            emit_log("Recovered unreadable block on " + relative_path + " at " + format_bytes(offset) +
                     " (" + std::to_string(length) + " bytes " + describe_salvage_fill() + "-filled).");
            return length;
        }
        return -1;
    }

    void write_chunk_with_retries(const std::string& relative_path, std::int64_t offset,
                                  const unsigned char* buffer, std::int32_t count,
                                  FileTransferSession& session, const CancelContext& cancel) {
        std::int32_t attempt = 0;
        std::optional<IoError> last_error;
        std::wstring destination_root = utf8_to_wide(options_.DestinationRoot);

        while (attempt <= options_.MaxRetries) {
            cancel.throw_if_cancelled();
            throw_if_media_identity_mismatch_throttled(destination_root, false);

            try {
                if (fault_injector_.has_value()) {
                    if (auto fault = fault_injector_->create_write_fault(relative_path, offset, count)) {
                        throw *fault;
                    }
                }
                timed_io::write_at(session.destination_handle(), offset, buffer, count,
                                   operation_timeout_ms(), cancel);
                bytes_written_.fetch_add(count);
                return;
            } catch (const IoError& ex) {
                if (ex.kind == IoErrorKind::Timeout) {
                    last_error = IoError("Write timeout at offset " + std::to_string(offset) + ".",
                                         0, IoErrorKind::Timeout);
                } else {
                    last_error = ex;
                }
            }

            if (is_disk_full(*last_error)) {
                session.invalidate_destination();
                throw IoError("Destination is out of free space while writing " + relative_path +
                              " at offset " + std::to_string(offset) + ".");
            }
            if (is_access_denied(*last_error) && !options_.TreatAccessDeniedAsContention) {
                session.invalidate_destination();
                throw IoError("Destination access denied while writing " + relative_path +
                              " at offset " + std::to_string(offset) + ".");
            }
            if (is_write_contention(*last_error) && options_.WaitForFileLockRelease) {
                session.invalidate_destination();
                emit_log("Destination contention detected on " + relative_path +
                         "; waiting for lock release.");
                wait_for_destination_write_access(
                    destination_path_from_relative(relative_path), cancel);
                attempt = 0;
                continue;
            }
            if (options_.WaitForMediaAvailability && is_availability_related(*last_error, true)) {
                session.invalidate_destination();
                emit_log("Destination unavailable during write of " + relative_path +
                         ". Waiting for media to return.");
                wait_for_destination_path(destination_path_from_relative(relative_path), cancel);
                attempt = 0;
                continue;
            }

            session.invalidate_destination();
            if (is_write_contention(*last_error)) {
                emit_log("Destination file contention detected on " + relative_path + "; retrying.");
            }

            attempt += 1;
            if (attempt > options_.MaxRetries) break;
            emit_log("Write retry " + std::to_string(attempt) + "/" +
                     std::to_string(options_.MaxRetries) + " on " + relative_path + " at " +
                     format_bytes(offset) + ": " + last_error->what());
            delay_for_retry(attempt, cancel);
        }

        throw IoError("Write failed on " + relative_path + " at offset " + std::to_string(offset) + ".");
    }

    std::wstring destination_path_from_relative(const std::string& relative_path) const {
        return detail::trim_trailing_separators(utf8_to_wide(options_.DestinationRoot)) + L"\\" +
               utf8_to_wide(relative_path);
    }

    void delay_for_retry(std::int32_t attempt, const CancelContext& cancel) {
        std::int32_t bounded = std::min(attempt, 20);
        double base_ms = options_.InitialRetryDelay.total_milliseconds() *
                         std::pow(2.0, bounded - 1);
        double capped = std::min(options_.MaxRetryDelay.total_milliseconds(), base_ms);
        std::uint32_t jitter_seed = 0;
        {
            std::lock_guard<std::mutex> guard(retry_jitter_lock_);
            jitter_seed = retry_jitter_();
        }
        double jitter_unit = (static_cast<double>(jitter_seed % 10000) / 5000.0) - 1.0;
        double jitter = capped * 0.25 * jitter_unit;
        std::int64_t delay_ms = static_cast<std::int64_t>(std::max(1.0, capped + jitter));
        cancel.delay(delay_ms);
    }

    // ---- Verification -----------------------------------------------------

    models::VerificationMode resolve_verification_mode() const {
        if (options_.VerificationModeValue != models::VerificationMode::None) {
            return options_.VerificationModeValue;
        }
        return options_.VerifyAfterCopy ? models::VerificationMode::Full
                                        : models::VerificationMode::None;
    }

    bool use_sha512() const {
        return options_.VerificationHashAlgorithmValue == models::VerificationHashAlgorithm::Sha512;
    }

    void verify_file_with_retries(const std::wstring& source_path, const std::wstring& destination_path,
                                  const std::string& relative_path,
                                  models::VerificationMode mode, const CancelContext& cancel,
                                  const std::optional<std::vector<unsigned char>>& copied_source_hash =
                                      std::nullopt) {
        if (mode == models::VerificationMode::Sampled) {
            emit_log("Verifying (sampled): " + relative_path);
            verify_sampled(source_path, destination_path, relative_path, cancel);
            return;
        }
        if (mode == models::VerificationMode::Full) {
            emit_log("Verifying (full hash): " + relative_path);
            auto source_hash = copied_source_hash.has_value()
                                   ? *copied_source_hash
                                   : compute_file_hash(source_path, cancel);
            if (copied_source_hash.has_value()) {
                emit_log("Using source hash captured during copy: " + relative_path);
            }
            auto destination_hash = compute_file_hash(destination_path, cancel);
            if (source_hash != destination_hash) {
                throw IoError("Verification hash mismatch: " + relative_path);
            }
        }
    }

    void verify_sampled(const std::wstring& source_path, const std::wstring& destination_path,
                        const std::string& relative_path, const CancelContext& cancel) {
        std::int64_t source_length = get_existing_file_length(source_path);
        std::int64_t destination_length = get_existing_file_length(destination_path);
        if (source_length != destination_length) {
            throw IoError("Verification length mismatch: " + relative_path);
        }
        if (source_length <= 0) return;

        std::int32_t chunk_size = static_cast<std::int32_t>(std::min<std::int64_t>(
            source_length, std::max(4096, options_.SampleVerificationChunkBytes)));
        std::int32_t chunk_count = std::max(1, options_.SampleVerificationChunkCount);

        std::vector<std::int64_t> offsets;
        if (source_length <= chunk_size || chunk_count == 1) {
            offsets.push_back(0);
        } else {
            std::int64_t max_offset = source_length - chunk_size;
            double step = static_cast<double>(max_offset) / (chunk_count - 1);
            for (std::int32_t index = 0; index < chunk_count; ++index) {
                double raw = step * index;
                std::int64_t value = static_cast<std::int64_t>(raw + (raw >= 0 ? 0.5 : -0.5));
                offsets.push_back(std::clamp<std::int64_t>(value, 0, max_offset));
            }
            std::sort(offsets.begin(), offsets.end());
            offsets.erase(std::unique(offsets.begin(), offsets.end()), offsets.end());
        }

        std::vector<unsigned char> buffer(static_cast<std::size_t>(chunk_size));
        for (std::int64_t sample_offset : offsets) {
            cancel.throw_if_cancelled();
            std::int32_t sample_length = static_cast<std::int32_t>(
                std::min<std::int64_t>(chunk_size, source_length - sample_offset));
            auto source_hash = read_and_hash_chunk(source_path, sample_offset, sample_length,
                                                   buffer, cancel);
            auto destination_hash = read_and_hash_chunk(destination_path, sample_offset, sample_length,
                                                        buffer, cancel);
            if (source_hash != destination_hash) {
                throw IoError("Sample verification mismatch: " + relative_path + " at offset " +
                              std::to_string(sample_offset) + ".");
            }
        }
    }

    std::vector<unsigned char> read_and_hash_chunk(const std::wstring& path, std::int64_t offset,
                                                   std::int32_t count,
                                                   std::vector<unsigned char>& buffer,
                                                   const CancelContext& cancel) {
        HANDLE handle = CreateFileW(path.c_str(), GENERIC_READ,
                                    FILE_SHARE_READ, nullptr,
                                    OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL | FILE_FLAG_OVERLAPPED, nullptr);
        if (handle == INVALID_HANDLE_VALUE) {
            throw IoError::from_win32("Unable to open file for verification.", GetLastError());
        }
        std::int32_t read = 0;
        try {
            read = timed_io::read_at(handle, offset, buffer.data(), count, operation_timeout_ms(), cancel);
        } catch (...) {
            CloseHandle(handle);
            throw;
        }
        CloseHandle(handle);
        if (read != count) {
            throw IoError("Short read while sampling verification chunk. Expected " +
                          std::to_string(count) + ", got " + std::to_string(read) + ".");
        }
        bytes_verified_.fetch_add(read);
        return crypto::hash_buffer(buffer.data(), static_cast<std::size_t>(count), use_sha512());
    }

    std::vector<unsigned char> compute_file_hash(const std::wstring& path, const CancelContext& cancel) {
        HANDLE handle = CreateFileW(path.c_str(), GENERIC_READ,
                                    FILE_SHARE_READ, nullptr,
                                    OPEN_EXISTING,
                                    FILE_ATTRIBUTE_NORMAL | FILE_FLAG_OVERLAPPED | FILE_FLAG_SEQUENTIAL_SCAN,
                                    nullptr);
        if (handle == INVALID_HANDLE_VALUE) {
            throw IoError::from_win32("Unable to open file for hashing.", GetLastError());
        }

        crypto::detail::StreamingHash hasher(use_sha512());
        std::vector<unsigned char> buffer(1 << 20);
        std::int64_t offset = 0;
        try {
            while (true) {
                std::int32_t read = timed_io::read_at(handle, offset, buffer.data(),
                                                      static_cast<std::int32_t>(buffer.size()),
                                                      operation_timeout_ms(), cancel);
                if (read <= 0) break;
                hasher.update(buffer.data(), static_cast<std::size_t>(read));
                bytes_verified_.fetch_add(read);
                offset += read;
                if (read < static_cast<std::int32_t>(buffer.size())) break;
            }
        } catch (...) {
            CloseHandle(handle);
            throw;
        }
        CloseHandle(handle);
        return hasher.finish();
    }

    // ---- Error classification --------------------------------------------

    static bool is_availability_related(const IoError& ex, bool include_file_not_found) {
        if (ex.kind == IoErrorKind::DriveNotFound || ex.kind == IoErrorKind::DirectoryNotFound) return true;
        switch (ex.win32) {
            case ErrPathNotFound:
            case ErrNotReady:
            case ErrBadNetPath:
            case ErrNetNameDeleted:
            case ErrBadNetName:
            case ErrDeviceNotConnected:
                return true;
            case ErrFileNotFound:
                if (include_file_not_found) return true;
                break;
            default:
                break;
        }
        std::string message = to_lower(ex.what());
        if (message.find("device is not ready") != std::string::npos ||
            message.find("network path") != std::string::npos ||
            message.find("network name") != std::string::npos ||
            message.find("specified network") != std::string::npos ||
            message.find("media identity mismatch") != std::string::npos) {
            return true;
        }
        if (include_file_not_found &&
            (message.find("cannot find the path") != std::string::npos ||
             message.find("file not found") != std::string::npos)) {
            return true;
        }
        return false;
    }

    static bool is_disk_full(const IoError& ex) {
        if (ex.win32 == ErrDiskFull || ex.win32 == ErrHandleDiskFull) return true;
        std::string message = to_lower(ex.what());
        return message.find("not enough space") != std::string::npos ||
               message.find("disk full") != std::string::npos;
    }

    static bool is_access_denied(const IoError& ex) {
        return ex.kind == IoErrorKind::UnauthorizedAccess || ex.win32 == ErrAccessDenied;
    }

    static bool is_file_lock(const IoError& ex) {
        return ex.win32 == ErrSharingViolation || ex.win32 == ErrLockViolation;
    }

    static bool is_source_file_missing(const IoError& ex) {
        if (ex.kind == IoErrorKind::FileNotFound) return true;
        return ex.win32 == ErrFileNotFound;
    }

    bool is_read_contention(const IoError& ex) const {
        return is_file_lock(ex) || (options_.TreatAccessDeniedAsContention && is_access_denied(ex));
    }

    bool is_write_contention(const IoError& ex) const {
        return is_file_lock(ex) || (options_.TreatAccessDeniedAsContention && is_access_denied(ex));
    }

    bool is_fatal_read(const IoError& ex) const {
        if (!options_.TreatAccessDeniedAsContention && is_access_denied(ex)) return true;
        return ex.kind == IoErrorKind::NotSupported || ex.kind == IoErrorKind::InvalidArgument;
    }

    static std::string to_lower(std::string text) {
        for (char& c : text) {
            if (c >= 'A' && c <= 'Z') c = static_cast<char>(c - 'A' + 'a');
        }
        return text;
    }

    // ---- Buffer controller factory ---------------------------------------

    AdaptiveBufferController create_buffer_controller(std::int64_t file_length, BufferPurpose purpose,
                                                      std::int32_t worker_count = 1) const {
        return AdaptiveBufferController::create(
            file_length, normalize_io_buffer_size(options_.BufferSizeBytes),
            options_.UseAdaptiveBufferSizing, options_.FragileMediaMode, purpose, worker_count);
    }

    std::int32_t resolve_buffer_size_for_file(std::int64_t file_length) const {
        return create_buffer_controller(file_length, BufferPurpose::Copy).current_size();
    }

    // ---- Bad-range map (InitializeBadRangeMapAsync et al) ----------------

    // GetFullPath + trailing-separator trim that preserves drive roots ("D:\").
    static std::wstring normalize_root_path(const std::wstring& path_value) {
        if (path_value.empty()) return std::wstring();
        std::wstring full = storage::fsutil::get_full_path(storage::fsutil::trim(path_value));
        std::wstring trimmed = detail::trim_trailing_separators(full);
        if (trimmed.empty()) return full;
        if (trimmed.size() == 2 && trimmed[1] == L':') return trimmed + L"\\";
        return trimmed;
    }

    static std::string build_file_fingerprint(const SourceFileDescriptor& descriptor,
                                              std::int64_t change_utc_ticks) {
        char text[64];
        std::snprintf(text, sizeof(text), "%016llX:%016llX:%016llX",
                      static_cast<unsigned long long>(descriptor.length),
                      static_cast<unsigned long long>(descriptor.last_write_utc_ticks),
                      static_cast<unsigned long long>(change_utc_ticks));
        return text;
    }

    bool should_track_bad_range_map() const {
        return options_.UseBadRangeMap || options_.UpdateBadRangeMapFromRun;
    }

    bool is_bad_range_map_fresh(const storage::BadRangeMap& map) const {
        if (options_.BadRangeMapMaxAgeDays <= 0) return true;
        if (map.UpdatedUtc == time::DateTimeOffset::min_value()) return false;
        std::int64_t age_ticks = time::DateTimeOffset::now_utc().utc_ticks() - map.UpdatedUtc.utc_ticks();
        return age_ticks <= static_cast<std::int64_t>(options_.BadRangeMapMaxAgeDays) * time::TicksPerDay;
    }

    void initialize_bad_range_map(const std::wstring& source_root) {
        bad_range_map_.reset();
        bad_range_map_path_.clear();
        bad_range_map_loaded_ = false;
        bad_range_map_read_hints_enabled_ = false;
        if (!should_track_bad_range_map()) return;

        std::wstring normalized_root = normalize_root_path(source_root);
        bad_range_map_path_ = storage::BadRangeMapStore::get_default_map_path(normalized_root);
        if (options_.UseBadRangeMap && options_.SkipKnownBadRanges &&
            options_.BadRangeMapMaxAgeDays <= 0) {
            emit_log("WARNING: bad-range map hints are configured to never expire; "
                     "file identity checks still apply, but review this advanced setting regularly.");
        }

        std::optional<storage::BadRangeMap> loaded;
        try {
            loaded = bad_range_map_store_.load(bad_range_map_path_);
        } catch (const std::exception& ex) {
            emit_log("Bad-range map load failed, starting fresh: " + std::string(ex.what()));
            loaded.reset();
        }

        bool allow_read_hints = false;
        const std::string source_media_identity = expected_media_identity(true);
        if (loaded.has_value()) {
            std::wstring map_source = normalize_root_path(utf8_to_wide(loaded->SourceRoot));
            if (!map_source.empty() &&
                storage::fsutil::to_upper_invariant(map_source) !=
                    storage::fsutil::to_upper_invariant(normalized_root)) {
                emit_log("Bad-range map source root mismatch; ignoring existing map payload.");
                loaded.reset();
            }
            if (loaded.has_value() && !trim_ascii(loaded->SourceIdentity).empty() &&
                (source_media_identity.empty() ||
                 !identities_equivalent(loaded->SourceIdentity, source_media_identity))) {
                emit_log("Bad-range map source media identity mismatch; ignoring existing map payload.");
                loaded.reset();
            }
        }

        if (!loaded.has_value()) {
            storage::BadRangeMap fresh;
            fresh.SchemaVersion = 1;
            fresh.SourceRoot = wide_to_utf8(normalized_root);
            fresh.SourceIdentity = source_media_identity;
            fresh.UpdatedUtc = time::DateTimeOffset::now_utc();
            loaded = std::move(fresh);
            emit_log("Bad-range map: initialized new map for this source.");
        } else {
            loaded->SourceRoot = wide_to_utf8(normalized_root);
            if (loaded->SchemaVersion <= 0) loaded->SchemaVersion = 1;
            if (!source_media_identity.empty()) loaded->SourceIdentity = source_media_identity;
            allow_read_hints = is_bad_range_map_fresh(*loaded);
            if (allow_read_hints) {
                std::size_t count = loaded->Files.size();
                emit_log("Bad-range map loaded: " + std::to_string(count) + " file entr" +
                         (count == 1 ? "y" : "ies") + " available.");
            } else if (options_.UseBadRangeMap && options_.SkipKnownBadRanges) {
                emit_log("Bad-range map is stale (>" + std::to_string(options_.BadRangeMapMaxAgeDays) +
                         " days). Skip hints disabled for this run.");
            }
        }

        bad_range_map_ = std::move(loaded);
        bad_range_map_loaded_ = true;
        bad_range_map_read_hints_enabled_ =
            options_.UseBadRangeMap && options_.SkipKnownBadRanges && allow_read_hints;
    }

    bool is_bad_range_map_entry_compatible(const SourceFileDescriptor& descriptor,
                                           const storage::BadRangeMapFileEntry& map_entry) const {
        if (map_entry.SourceLength > 0 && map_entry.SourceLength != descriptor.length) return false;
        if (map_entry.LastWriteUtcTicks > 0 &&
            map_entry.LastWriteUtcTicks != descriptor.last_write_utc_ticks) {
            return false;
        }
        // Legacy entries do not carry a file identity. Their size/time and
        // fingerprint can still match a replacement, so they are readable
        // but are never trusted as skip hints after identity hardening.
        if (map_entry.SourceFileIndex == 0 || map_entry.SourceVolumeSerial == 0) return false;
        FileStabilitySnapshot current;
        if (!try_get_file_stability(descriptor.full_path, current) ||
            current.length != descriptor.length ||
            current.last_write_utc_ticks != descriptor.last_write_utc_ticks) {
            return false;
        }
        if (map_entry.SourceFileIndex !=
                static_cast<std::int64_t>(current.identity.file_index) ||
            map_entry.SourceVolumeSerial !=
                static_cast<std::int64_t>(current.identity.volume_serial)) {
            return false;
        }
        std::string fingerprint = trim_ascii(map_entry.FileFingerprint);
        if (fingerprint.empty()) return false;
        return models::detail::equals_ignore_case(
            fingerprint, build_file_fingerprint(descriptor, current.change_utc_ticks));
    }

    std::optional<storage::BadRangeMapFileEntry> try_get_bad_range_map_entry(
        const SourceFileDescriptor& descriptor) {
        std::lock_guard<std::mutex> guard(map_lock_);
        if (!bad_range_map_.has_value()) return std::nullopt;

        std::string relative = detail::normalize_relative_path(descriptor.relative_path);
        if (relative.empty()) return std::nullopt;

        const storage::BadRangeMapFileEntry* candidate = bad_range_map_->Files.find(relative);
        if (candidate == nullptr || candidate->BadRanges.empty() ||
            candidate->ConfirmationCount < 2) {
            return std::nullopt;
        }
        if (!is_bad_range_map_entry_compatible(descriptor, *candidate)) return std::nullopt;
        return *candidate;
    }

    // Marks map-known bad ranges as KnownBad in the entry's rescue ranges so
    // rescue passes skip re-reading them. Returns total mapped bytes.
    std::int64_t apply_known_bad_ranges_from_map(const SourceFileDescriptor& descriptor,
                                                 JournalFileEntry& entry) {
        if (!bad_range_map_read_hints_enabled_) return 0;
        std::optional<storage::BadRangeMapFileEntry> map_entry =
            try_get_bad_range_map_entry(descriptor);
        if (!map_entry.has_value()) return 0;

        std::int64_t total_mapped = 0;
        for (const auto& bad_range : map_entry->BadRanges) {
            if (bad_range.Length <= 0) continue;
            std::int64_t range_start = std::max<std::int64_t>(0, bad_range.Offset);
            if (range_start >= descriptor.length) continue;
            std::int64_t range_length = std::min(bad_range.Length, descriptor.length - range_start);
            if (range_length <= 0) continue;
            set_range_state(entry.RescueRanges, range_start, range_length, RescueRangeState::KnownBad);
            total_mapped += range_length;
        }
        if (total_mapped > 0) {
            emit_log("[Rescue Engine] Applied bad-range map hints to " + descriptor.relative_path +
                     ": " + format_bytes(total_mapped) + ".");
        }
        return total_mapped;
    }

    void try_persist_bad_range_map_entry(const SourceFileDescriptor& descriptor,
                                         const JournalFileEntry& entry, bool force_save) {
        if (!bad_range_map_loaded_ || bad_range_map_path_.empty()) return;
        if (!options_.UpdateBadRangeMapFromRun) return;

        std::lock_guard<std::mutex> guard(map_lock_);
        if (!bad_range_map_.has_value()) return;

        std::string relative = detail::normalize_relative_path(descriptor.relative_path);
        if (relative.empty()) return;

        // Synthetic salvage output is deliberately not promoted to a future
        // skip hint. A later pass may be able to read those bytes after the
        // device cools or is reconnected.
        std::vector<ByteRange> unreadable = merge_byte_ranges(snapshot_by_state(
            entry.RescueRanges, {RescueRangeState::Bad, RescueRangeState::KnownBad}));

        if (unreadable.empty()) {
            bad_range_map_->Files.remove(relative);
        } else {
            const storage::BadRangeMapFileEntry* previous =
                bad_range_map_->Files.find(relative);
            storage::BadRangeMapFileEntry map_entry;
            map_entry.RelativePath = relative;
            map_entry.SourceLength = descriptor.length;
            map_entry.LastWriteUtcTicks = descriptor.last_write_utc_ticks;
            FileStabilitySnapshot source_stability;
            if (try_get_file_stability(descriptor.full_path, source_stability) &&
                source_stability.length == descriptor.length &&
                source_stability.last_write_utc_ticks == descriptor.last_write_utc_ticks) {
                map_entry.SourceFileIndex =
                    static_cast<std::int64_t>(source_stability.identity.file_index);
                map_entry.SourceVolumeSerial =
                    static_cast<std::int64_t>(source_stability.identity.volume_serial);
                map_entry.FileFingerprint = build_file_fingerprint(
                    descriptor, source_stability.change_utc_ticks);
            }
            map_entry.BadRanges = std::move(unreadable);
            auto ranges_equal = [](const std::vector<ByteRange>& left,
                                   const std::vector<ByteRange>& right) {
                if (left.size() != right.size()) return false;
                for (std::size_t index = 0; index < left.size(); ++index) {
                    if (left[index].Offset != right[index].Offset ||
                        left[index].Length != right[index].Length) {
                        return false;
                    }
                }
                return true;
            };
            const bool repeated_observation =
                previous != nullptr && previous->SourceFileIndex == map_entry.SourceFileIndex &&
                previous->SourceVolumeSerial == map_entry.SourceVolumeSerial &&
                models::detail::equals_ignore_case(previous->FileFingerprint,
                                                   map_entry.FileFingerprint) &&
                ranges_equal(previous->BadRanges, map_entry.BadRanges);
            map_entry.ConfirmationCount = repeated_observation
                                              ? std::min(1000, std::max(1, previous->ConfirmationCount) + 1)
                                              : 1;
            map_entry.LastScanUtc = time::DateTimeOffset::now_utc();
            map_entry.LastError = entry.LastError;
            bad_range_map_->Files.set(relative, std::move(map_entry));
        }

        bad_range_map_->SourceRoot =
            wide_to_utf8(normalize_root_path(utf8_to_wide(options_.SourceRoot)));
        bad_range_map_->SourceIdentity = expected_media_identity(true);
        bad_range_map_->UpdatedUtc = time::DateTimeOffset::now_utc();
        if (bad_range_map_->SchemaVersion <= 0) bad_range_map_->SchemaVersion = 1;

        bool should_save = force_save;
        if (!should_save) {
            ULONGLONG now = GetTickCount64();
            should_save = last_bad_range_map_flush_tick_ == 0 ||
                          now - last_bad_range_map_flush_tick_ >= 2000;
        }
        if (!should_save) return;

        try {
            if (!bad_range_map_store_.save(bad_range_map_path_, *bad_range_map_)) {
                emit_log("WARNING: Bad-range map was saved with reduced snapshot redundancy.");
            }
            last_bad_range_map_flush_tick_ = GetTickCount64();
        } catch (const std::exception& ex) {
            emit_log("Bad-range map save failed: " + std::string(ex.what()));
        }
    }

    void flush_bad_range_map() {
        if (!bad_range_map_loaded_ || bad_range_map_path_.empty()) return;
        if (!options_.UpdateBadRangeMapFromRun) return;
        std::lock_guard<std::mutex> guard(map_lock_);
        if (!bad_range_map_.has_value()) return;
        try {
            bad_range_map_->UpdatedUtc = time::DateTimeOffset::now_utc();
            if (!bad_range_map_store_.save(bad_range_map_path_, *bad_range_map_)) {
                emit_log("WARNING: Final bad-range map was saved with reduced snapshot redundancy.");
            }
            last_bad_range_map_flush_tick_ = GetTickCount64();
        } catch (const std::exception& ex) {
            emit_log("Bad-range map final flush failed: " + std::string(ex.what()));
        }
    }

    // ---- Scan mode --------------------------------------------------------

    models::ScanPerformanceProfile resolve_scan_performance_profile() const {
        if (options_.OperationMode != models::JobOperationMode::ScanOnly) {
            return models::ScanPerformanceProfile::Precise;
        }
        if (options_.FragileMediaMode) return models::ScanPerformanceProfile::Precise;
        switch (options_.ScanPerformanceProfileValue) {
            case models::ScanPerformanceProfile::Precise:
                return models::ScanPerformanceProfile::Precise;
            default:
                return models::ScanPerformanceProfile::Fast;
        }
    }

    static std::optional<bool> query_seek_penalty(const std::string& root_utf8) {
        std::wstring full = storage::fsutil::get_full_path(utf8_to_wide(root_utf8));
        if (full.size() < 2 || full[1] != L':') return std::nullopt;
        std::wstring volume = L"\\\\.\\" + full.substr(0, 2);
        HANDLE handle = CreateFileW(
            volume.c_str(), 0, FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE,
            nullptr, OPEN_EXISTING, 0, nullptr);
        if (handle == INVALID_HANDLE_VALUE) return std::nullopt;

        STORAGE_PROPERTY_QUERY query{};
        query.PropertyId = StorageDeviceSeekPenaltyProperty;
        query.QueryType = PropertyStandardQuery;
        DEVICE_SEEK_PENALTY_DESCRIPTOR descriptor{};
        DWORD returned = 0;
        bool available = DeviceIoControl(
            handle, IOCTL_STORAGE_QUERY_PROPERTY, &query, sizeof(query), &descriptor,
            sizeof(descriptor), &returned, nullptr) != FALSE &&
                         returned >= sizeof(descriptor) &&
                         descriptor.Size >= sizeof(descriptor);
        CloseHandle(handle);
        if (!available) return std::nullopt;
        return descriptor.IncursSeekPenalty != FALSE;
    }

    std::int32_t resolve_parallel_scan_workers(std::int32_t file_count) const {
        if (options_.OperationMode != models::JobOperationMode::ScanOnly || file_count <= 1) return 1;
        if (options_.FragileMediaMode) return 1;
        if (options_.ParallelScanWorkers > 0) {
            return std::max(1, std::min(64, std::min(file_count, options_.ParallelScanWorkers)));
        }
        if (query_seek_penalty(options_.SourceRoot).value_or(false)) return 1;
        SYSTEM_INFO info{};
        GetSystemInfo(&info);
        std::int32_t processors = std::max<std::int32_t>(
            1, static_cast<std::int32_t>(info.dwNumberOfProcessors));
        return std::max(1, std::min(file_count, std::min(8, processors)));
    }

    std::int32_t resolve_parallel_small_file_workers(std::int32_t file_count) const {
        if (options_.OperationMode != models::JobOperationMode::Copy || file_count <= 1) return 1;
        if (options_.ParallelSmallFileWorkers > 0) {
            return std::max(1, std::min(64, std::min(file_count, options_.ParallelSmallFileWorkers)));
        }
        if (query_seek_penalty(options_.SourceRoot).value_or(false) ||
            query_seek_penalty(options_.DestinationRoot).value_or(false)) {
            return 1;
        }
        SYSTEM_INFO info{};
        GetSystemInfo(&info);
        std::int32_t processors = std::max<std::int32_t>(
            1, static_cast<std::int32_t>(info.dwNumberOfProcessors));
        return std::max(1, std::min(file_count, std::min(8, processors)));
    }

    static bool is_completed_scan_journal_entry(const JournalFileEntry& entry,
                                                const SourceFileDescriptor& descriptor) {
        if (entry.State != FileCopyState::Completed &&
            entry.State != FileCopyState::CompletedWithRecovery) {
            return false;
        }
        return entry.SourceChangeUtcTicks != 0 &&
               entry.SourceFileIndex != 0 && entry.SourceVolumeSerial != 0 &&
               entry.SourceLength == descriptor.length &&
               entry.SourceLastWriteUtcTicks == descriptor.last_write_utc_ticks &&
               entry.BytesCopied >= descriptor.length;
    }

    static void seed_fast_scan_fallback_ranges(JournalFileEntry& entry, std::int64_t file_length,
                                               std::int64_t bytes_already_scanned) {
        std::int64_t safe_length = std::max<std::int64_t>(0, file_length);
        std::int64_t scanned = std::max<std::int64_t>(0, std::min(safe_length, bytes_already_scanned));
        entry.RescueRanges.clear();
        if (scanned > 0) append_range(entry.RescueRanges, 0, scanned, RescueRangeState::Good);
        if (scanned < safe_length) {
            append_range(entry.RescueRanges, scanned, safe_length - scanned, RescueRangeState::Pending);
        }
        entry.BytesCopied = scanned;
    }

    struct FastScanOutcome {
        bool succeeded = false;
        std::int64_t bytes_read = 0;
        std::int32_t buffer_size_bytes = 0;
        std::string error_message;
    };

    struct FastFragileFailure {
        SourceFileDescriptor descriptor;
        std::string message;
        std::int64_t offset = -1;
        std::int32_t length = 0;
        bool primary_data_range = false;
    };

    // Shared coordination state for the fast-scan worker pool.
    struct FastScanShared {
        std::mutex queue_lock;
        std::vector<SourceFileDescriptor> queue;
        std::size_t next_index = 0;
        std::vector<SourceFileDescriptor> fallback;
        std::vector<FastFragileFailure> fragile_failures;
        std::vector<std::string> started_keys;
        std::vector<std::string> active_files;
        std::mutex fatal_lock;
        std::string first_fatal;
        std::atomic<bool> cancelled_user{false};
        std::atomic<bool> cancelled_timeout{false};
        std::int32_t worker_count = 1;
    };

    FastScanOutcome scan_single_file_fast_health(const SourceFileDescriptor& descriptor,
                                                 JournalFileEntry& entry,
                                                 ProgressAccumulator& progress,
                                                 FastScanShared& shared,
                                                 const CancelContext& cancel) {
        AdaptiveBufferController controller =
            create_buffer_controller(descriptor.length, BufferPurpose::FastHealthScan,
                                     shared.worker_count);
        FastScanOutcome outcome;
        outcome.buffer_size_bytes = controller.current_size();
        std::vector<unsigned char> io_buffer(
            static_cast<std::size_t>(std::max(controller.maximum_size(), MinimumRescueBlockSize)));
        FileTransferSession session(descriptor.full_path, std::wstring());
        FileStabilitySnapshot scan_source_stability;
        if (!try_get_file_stability(session.source_handle(), scan_source_stability) ||
            scan_source_stability.length != descriptor.length ||
            scan_source_stability.last_write_utc_ticks != descriptor.last_write_utc_ticks) {
            throw IoError("Source changed after scan enumeration: " +
                          descriptor.relative_path + ".");
        }

        std::int64_t initial_offset = 0;
        {
            std::lock_guard<std::mutex> guard(journal_lock_);
            entry.SourceFileIndex = static_cast<std::int64_t>(
                scan_source_stability.identity.file_index);
            entry.SourceVolumeSerial = static_cast<std::int64_t>(
                scan_source_stability.identity.volume_serial);
            entry.SourceChangeUtcTicks = scan_source_stability.change_utc_ticks;
            initial_offset = std::clamp<std::int64_t>(entry.BytesCopied, 0,
                                                       descriptor.length);
        }
        if (initial_offset > 0) {
            std::lock_guard<std::mutex> guard(progress_lock_);
            progress.total_bytes_copied += initial_offset;
            progress.bytes_reused += initial_offset;
        }

        try {
            std::int64_t offset = initial_offset;
            while (offset < descriptor.length) {
                cancel.throw_if_cancelled();
                if (control_ != nullptr) control_->wait_if_paused(cancel);

                std::int32_t chunk_length = controller.next_chunk_length(descriptor.length - offset);
                ULONGLONG chunk_started = GetTickCount64();
                std::int32_t bytes_read = read_chunk_with_retries(
                    descriptor.full_path, descriptor.relative_path, offset, chunk_length, session,
                    io_buffer.data(), cancel, 0, false);

                if (bytes_read != chunk_length) {
                    controller.report_failure();
                    outcome.succeeded = false;
                    outcome.bytes_read = offset;
                    outcome.buffer_size_bytes = controller.current_size();
                    outcome.error_message = "short or failed read";
                    return outcome;
                }

                controller.report_success(bytes_read,
                                          static_cast<double>(GetTickCount64() - chunk_started));
                offset += bytes_read;
                {
                    std::lock_guard<std::mutex> guard(journal_lock_);
                    entry.BytesCopied = offset;
                }
                {
                    std::lock_guard<std::mutex> guard(progress_lock_);
                    progress.total_bytes_copied += bytes_read;
                    std::vector<std::string> active_snapshot;
                    {
                        std::lock_guard<std::mutex> queue_guard(shared.queue_lock);
                        active_snapshot = shared.active_files;
                    }
                    emit_progress(progress, descriptor.relative_path, offset, descriptor.length,
                                  bytes_read, chunk_length, "FastHealthScan", 0, 0,
                                  static_cast<std::int32_t>(active_snapshot.size()),
                                  shared.worker_count, &active_snapshot);
                }
            }
            FileStabilitySnapshot completed_scan_stability;
            if (!try_get_file_stability(session.source_handle(), completed_scan_stability) ||
                !file_stability_matches(scan_source_stability,
                                        completed_scan_stability)) {
                throw IoError("Source content or metadata changed during scan: " +
                              descriptor.relative_path + ".");
            }
            {
                std::lock_guard<std::mutex> guard(journal_lock_);
                entry.SourceFileIndex = static_cast<std::int64_t>(
                    completed_scan_stability.identity.file_index);
                entry.SourceVolumeSerial = static_cast<std::int64_t>(
                    completed_scan_stability.identity.volume_serial);
                entry.SourceChangeUtcTicks = completed_scan_stability.change_utc_ticks;
            }
            outcome.succeeded = true;
            outcome.bytes_read = offset;
            outcome.buffer_size_bytes = controller.current_size();
            return outcome;
        } catch (const FragileReadSkip&) {
            throw;
        } catch (const SourceMutationSkipped&) {
            throw;
        } catch (const OperationCanceled&) {
            throw;
        } catch (const std::exception& ex) {
            outcome.succeeded = false;
            {
                std::lock_guard<std::mutex> guard(journal_lock_);
                outcome.bytes_read = std::max<std::int64_t>(0, entry.BytesCopied);
            }
            outcome.buffer_size_bytes = controller.current_size();
            outcome.error_message = ex.what();
            return outcome;
        }
    }

    void scan_fast_worker(FastScanShared& shared, ProgressAccumulator& progress, JobJournal& journal,
                          const CancelContext& cancel) {
        std::wstring source_root = utf8_to_wide(options_.SourceRoot);
        std::wstring destination_root = utf8_to_wide(options_.DestinationRoot);

        try {
            while (true) {
                SourceFileDescriptor descriptor;
                {
                    std::lock_guard<std::mutex> guard(shared.queue_lock);
                    if (shared.next_index >= shared.queue.size()) return;
                    descriptor = shared.queue[shared.next_index++];
                }

                cancel.throw_if_cancelled();
                if (control_ != nullptr) control_->wait_if_paused(cancel);
                wait_for_media_availability(source_root, destination_root, cancel);
                ensure_media_identity_integrity(source_root, destination_root, false, false);

                std::string active_key = detail::normalize_relative_path(descriptor.relative_path);
                if (active_key.empty()) active_key = wide_to_utf8(descriptor.full_path);

                {
                    std::lock_guard<std::mutex> guard(shared.queue_lock);
                    bool already_started = false;
                    for (const auto& started : shared.started_keys) {
                        if (models::detail::equals_ignore_case(started, active_key)) {
                            already_started = true;
                            break;
                        }
                    }
                    if (already_started) {
                        emit_log("Fast scan duplicate suppressed: " + active_key);
                        continue;
                    }
                    shared.started_keys.push_back(active_key);
                    shared.active_files.push_back(active_key);
                }

                struct ActiveGuard {
                    FastScanShared& shared;
                    std::string key;
                    ~ActiveGuard() {
                        std::lock_guard<std::mutex> guard(shared.queue_lock);
                        for (std::size_t i = 0; i < shared.active_files.size(); ++i) {
                            if (shared.active_files[i] == key) {
                                shared.active_files.erase(
                                    shared.active_files.begin() + static_cast<std::ptrdiff_t>(i));
                                break;
                            }
                        }
                    }
                } active_guard{shared, active_key};

                try {
                    JournalFileEntry* entry = nullptr;
                    bool already_completed = false;
                    {
                        std::lock_guard<std::mutex> guard(journal_lock_);
                        entry = journal.Files.find(descriptor.relative_path);
                        if (entry == nullptr) continue;
                        if (is_completed_scan_journal_entry(*entry, descriptor)) {
                            already_completed = true;
                        } else {
                            if (entry->State == FileCopyState::Failed && !entry->DoNotRetry) {
                                entry->State = FileCopyState::Pending;
                                entry->LastError.clear();
                            }
                            if (!should_skip_failed_entry_for_fragile_resume(*entry)) {
                                entry->State = FileCopyState::InProgress;
                                entry->LastError.clear();
                                entry->DoNotRetry = false;
                                entry->LastRescuePass = "FastHealthScan";
                                entry->RecoveredRanges.clear();
                                entry->RescueRanges.clear();
                            }
                        }
                    }

                    if (already_completed) {
                        std::lock_guard<std::mutex> guard(progress_lock_);
                        progress.completed_files += 1;
                        if (entry->State == FileCopyState::CompletedWithRecovery) {
                            progress.recovered_files += 1;
                        }
                        progress.total_bytes_copied += descriptor.length;
                        progress.bytes_reused += descriptor.length;
                        emit_progress(progress, descriptor.relative_path, descriptor.length,
                                      descriptor.length, 0,
                                      resolve_buffer_size_for_file(descriptor.length),
                                      "FastHealthScanResume", 0, 0, 0, shared.worker_count);
                        continue;
                    }

                    if (should_skip_failed_entry_for_fragile_resume(*entry)) {
                        std::lock_guard<std::mutex> guard(progress_lock_);
                        progress.skipped_files += 1;
                        std::int64_t already = entry->BytesCopied > 0 ? entry->BytesCopied : 0;
                        std::int64_t remaining = descriptor.length - already;
                        const std::int64_t skipped = remaining > 0 ? remaining : 0;
                        progress.total_bytes_copied += skipped;
                        progress.bytes_skipped += skipped;
                        emit_progress(progress, descriptor.relative_path, descriptor.length,
                                      descriptor.length, 0,
                                      resolve_buffer_size_for_file(descriptor.length),
                                      "FastHealthScanSkipped", 0, 0, 0, shared.worker_count);
                        continue;
                    }

                    if (options_.UseBadRangeMap && options_.SkipKnownBadRanges &&
                        try_get_bad_range_map_entry(descriptor).has_value()) {
                        {
                            std::lock_guard<std::mutex> guard(journal_lock_);
                            seed_fast_scan_fallback_ranges(*entry, descriptor.length, 0);
                            entry->LastRescuePass = "FastMappedFallback";
                        }
                        {
                            std::lock_guard<std::mutex> guard(shared.queue_lock);
                            shared.fallback.push_back(descriptor);
                        }
                        emit_log("Fast scan fallback queued: " + descriptor.relative_path +
                                 " has known bad-range map hints.");
                        continue;
                    }

                    FastScanOutcome outcome =
                        scan_single_file_fast_health(descriptor, *entry, progress, shared, cancel);

                    if (outcome.succeeded) {
                        {
                            std::lock_guard<std::mutex> guard(journal_lock_);
                            entry->State = FileCopyState::Completed;
                            entry->BytesCopied = descriptor.length;
                            entry->LastError.clear();
                            entry->DoNotRetry = false;
                            entry->LastRescuePass = "FastHealthScan";
                            entry->RescueRanges.clear();
                            entry->RecoveredRanges.clear();
                        }
                        std::lock_guard<std::mutex> guard(progress_lock_);
                        progress.completed_files += 1;
                        emit_progress(progress, descriptor.relative_path, descriptor.length,
                                      descriptor.length, 0, outcome.buffer_size_bytes,
                                      "FastHealthScan", 0, 0, 0, shared.worker_count);
                    } else {
                        {
                            std::lock_guard<std::mutex> guard(journal_lock_);
                            seed_fast_scan_fallback_ranges(*entry, descriptor.length,
                                                           outcome.bytes_read);
                            entry->LastRescuePass = "FastFallback";
                        }
                        {
                            std::lock_guard<std::mutex> guard(shared.queue_lock);
                            shared.fallback.push_back(descriptor);
                        }
                        emit_log("Fast scan fallback queued: " + descriptor.relative_path + " at " +
                                 format_bytes(outcome.bytes_read) + " (" + outcome.error_message +
                                 ").");
                    }
                } catch (const FragileReadSkip& ex) {
                    {
                        std::lock_guard<std::mutex> guard(journal_lock_);
                        JournalFileEntry* failed_entry =
                            journal.Files.find(descriptor.relative_path);
                        if (failed_entry != nullptr) {
                            // Capture the finding before any cooldown or process
                            // interruption. The joined phase persists the map and
                            // accounts progress without issuing another read.
                            record_fragile_read_skip(descriptor, *failed_entry, ex);
                        }
                    }
                    {
                        std::lock_guard<std::mutex> guard(shared.queue_lock);
                        shared.fragile_failures.push_back(FastFragileFailure{
                            descriptor, ex.what(), ex.Offset, ex.Length,
                            ex.PrimaryDataRange});
                    }
                    emit_log("Fast scan fragile stop captured: " +
                             descriptor.relative_path + " (no precise reread queued).");
                    register_fragile_failure_and_maybe_cooldown(
                        descriptor.relative_path, ex.what(), cancel);
                } catch (const SourceMutationSkipped&) {
                    std::lock_guard<std::mutex> guard(progress_lock_);
                    progress.skipped_files += 1;
                    progress.total_bytes_copied += descriptor.length;
                    progress.bytes_skipped += descriptor.length;
                    emit_progress(progress, descriptor.relative_path, descriptor.length,
                                  descriptor.length, 0,
                                  resolve_buffer_size_for_file(descriptor.length),
                                  "FastHealthScanSkipped", 0, 0, 0, shared.worker_count);
                } catch (const OperationCanceled&) {
                    throw;
                } catch (const std::exception& ex) {
                    {
                        std::lock_guard<std::mutex> guard(journal_lock_);
                        JournalFileEntry* entry = journal.Files.find(descriptor.relative_path);
                        if (entry != nullptr) {
                            entry->State = FileCopyState::Pending;
                            entry->LastError = ex.what();
                            entry->LastRescuePass = "FastFallback";
                            seed_fast_scan_fallback_ranges(
                                *entry, descriptor.length,
                                std::max<std::int64_t>(0, entry->BytesCopied));
                        }
                    }
                    {
                        std::lock_guard<std::mutex> guard(shared.queue_lock);
                        shared.fallback.push_back(descriptor);
                    }
                    emit_log("Fast scan fallback queued: " + descriptor.relative_path + " (" +
                             std::string(ex.what()) + ").");
                }
            }
        } catch (const OperationCanceled& oc) {
            if (oc.user_requested) shared.cancelled_user.store(true);
            else shared.cancelled_timeout.store(true);
        } catch (const std::exception& ex) {
            std::lock_guard<std::mutex> guard(shared.fatal_lock);
            if (shared.first_fatal.empty()) shared.first_fatal = ex.what();
        }
    }

    // Returns the first fatal error message ("" = clean).
    std::string run_fast_scan_mode(const std::vector<SourceFileDescriptor>& source_files,
                                   const std::wstring& source_root, ProgressAccumulator& progress,
                                   JobJournal& journal, const std::wstring& journal_path,
                                   const CancelContext& cancel) {
        if (source_files.empty()) return std::string();

        FastScanShared shared;
        shared.queue = source_files;
        shared.worker_count = resolve_parallel_scan_workers(
            static_cast<std::int32_t>(source_files.size()));
        emit_log("Fast scan engine: " + std::to_string(source_files.size()) + " file(s), " +
                 std::to_string(shared.worker_count) + " worker(s).");

        HANDLE checkpoint_stop = CreateEventW(nullptr, TRUE, FALSE, nullptr);
        std::mutex checkpoint_error_lock;
        std::string checkpoint_error;
        auto save_checkpoint = [&]() {
            JobJournal snapshot;
            {
                std::lock_guard<std::mutex> guard(journal_lock_);
                snapshot = journal;
            }
            try {
                if (!journal_store_.save(journal_path, std::move(snapshot))) {
                    std::lock_guard<std::mutex> guard(checkpoint_error_lock);
                    checkpoint_error =
                        "live checkpoint has reduced snapshot/ledger redundancy";
                }
            } catch (const std::exception& ex) {
                std::lock_guard<std::mutex> guard(checkpoint_error_lock);
                checkpoint_error = ex.what();
            }
        };

        // Fast workers update different entries concurrently and previously
        // wrote nothing until every worker and fallback pass had finished.
        // Snapshot under journal_lock_ at an adaptive cadence; cancellation
        // performs one final synchronous snapshot below.
        std::thread checkpoint_thread;
        if (checkpoint_stop != nullptr) {
            checkpoint_thread = std::thread([&]() {
                DWORD interval_ms = 5000;
                while (WaitForSingleObject(checkpoint_stop, interval_ms) == WAIT_TIMEOUT) {
                    const ULONGLONG started = GetTickCount64();
                    save_checkpoint();
                    const ULONGLONG cost = std::max<ULONGLONG>(1, GetTickCount64() - started);
                    interval_ms = static_cast<DWORD>(std::clamp<ULONGLONG>(
                        cost * (JournalFlushDutyDivisor - 1), 5000,
                        MaximumJournalFlushIntervalMs));
                }
            });
        }

        std::vector<std::thread> workers;
        for (std::int32_t i = 0; i < shared.worker_count; ++i) {
            workers.emplace_back([this, &shared, &progress, &journal, &cancel]() {
                scan_fast_worker(shared, progress, journal, cancel);
            });
        }
        for (auto& worker : workers) worker.join();

        if (checkpoint_stop != nullptr) SetEvent(checkpoint_stop);
        if (checkpoint_thread.joinable()) checkpoint_thread.join();

        // Fragile mode is deliberately single-worker. Preserve each first-read
        // failure now that the checkpoint writer is joined; sending it through
        // precise fallback would contradict "skip on first read error" by
        // reading the same failing range a second time.
        for (const auto& captured : shared.fragile_failures) {
            JournalFileEntry* entry =
                journal.Files.find(captured.descriptor.relative_path);
            if (entry == nullptr) continue;
            FragileReadSkip failure(captured.message, captured.offset, captured.length,
                                    captured.primary_data_range);
            handle_fragile_read_skip(captured.descriptor, *entry, progress, journal,
                                     journal_path, failure, cancel, "scan",
                                     /*register_failure*/ false,
                                     /*force_journal_flush*/ false);
        }

        if (shared.cancelled_user.load() || shared.cancelled_timeout.load()) {
            // Preserve every completed file and the latest bound offset for the
            // handful of files active at cancellation before unwinding.
            save_checkpoint();
        }
        if (checkpoint_stop != nullptr) CloseHandle(checkpoint_stop);

        {
            std::lock_guard<std::mutex> guard(checkpoint_error_lock);
            if (!checkpoint_error.empty()) {
                emit_log("Fast scan checkpoint warning: " + checkpoint_error + ".");
            }
        }

        if (shared.cancelled_user.load()) throw OperationCanceled{true};
        if (shared.cancelled_timeout.load()) throw OperationCanceled{false};

        std::string first_fatal;
        {
            std::lock_guard<std::mutex> guard(shared.fatal_lock);
            first_fatal = shared.first_fatal;
        }
        if (!first_fatal.empty() && !options_.ContinueOnFileError) return first_fatal;
        if (shared.fallback.empty()) return first_fatal;

        emit_log("Fast scan fallback: " + std::to_string(shared.fallback.size()) +
                 " file(s) require precise bad-range localization.");
        std::wstring destination_root = utf8_to_wide(options_.DestinationRoot);
        for (const auto& descriptor : shared.fallback) {
            cancel.throw_if_cancelled();
            if (control_ != nullptr) control_->wait_if_paused(cancel);
            wait_for_media_availability(source_root, destination_root, cancel);

            JournalFileEntry* entry = journal.Files.find(descriptor.relative_path);
            if (entry == nullptr) continue;

            std::string error = process_precise_scan_file(descriptor, *entry, progress, journal,
                                                          journal_path, true, false, cancel);
            if (!error.empty()) {
                first_fatal = error;
                if (!options_.ContinueOnFileError) return first_fatal;
            }
        }
        return first_fatal;
    }

    void scan_small_file_fast(const SourceFileDescriptor& descriptor, JournalFileEntry& entry,
                              FileTransferSession& session, std::vector<unsigned char>& io_buffer,
                              AdaptiveBufferController& controller, ProgressAccumulator& progress,
                              JobJournal& journal, const std::wstring& journal_path,
                              const CancelContext& cancel) {
        std::wstring source_root = utf8_to_wide(options_.SourceRoot);
        std::wstring destination_root = utf8_to_wide(options_.DestinationRoot);
        std::vector<ByteRange> pending = snapshot_by_state(entry.RescueRanges,
                                                           {RescueRangeState::Pending});
        for (const auto& pending_range : pending) {
            std::int64_t offset = pending_range.Offset;
            std::int64_t remaining = pending_range.Length;
            while (remaining > 0) {
                cancel.throw_if_cancelled();
                if (control_ != nullptr) control_->wait_if_paused(cancel);
                wait_for_media_availability(source_root, destination_root, cancel);

                std::int32_t chunk_length = controller.next_chunk_length(remaining);
                ULONGLONG chunk_started = GetTickCount64();
                std::int32_t bytes_read = read_chunk_with_retries(
                    descriptor.full_path, descriptor.relative_path, offset, chunk_length, session,
                    io_buffer.data(), cancel, std::max(0, options_.MaxRetries), false);

                if (bytes_read > 0) {
                    controller.report_success(bytes_read,
                                              static_cast<double>(GetTickCount64() - chunk_started));
                    set_range_state(entry.RescueRanges, offset, bytes_read, RescueRangeState::Good);
                    progress.total_bytes_copied += bytes_read;
                } else {
                    controller.report_failure();
                    set_range_state(entry.RescueRanges, offset, chunk_length, RescueRangeState::Bad);
                    progress.total_bytes_copied += chunk_length;
                }

                sync_entry_bytes_copied(entry);
                emit_progress(progress, descriptor.relative_path,
                              display_progress_bytes(entry, descriptor.length, entry.BytesCopied),
                              descriptor.length, std::max(0, bytes_read), chunk_length,
                              "ScanSmallFast", unreadable_region_count(entry.RescueRanges),
                              rescue_bytes(entry.RescueRanges,
                                           {RescueRangeState::Pending, RescueRangeState::Bad,
                                            RescueRangeState::KnownBad}));
                flush_journal(journal_path, journal, false);

                offset += chunk_length;
                remaining -= chunk_length;
            }
        }
    }

    // Returns true when unreadable ranges were detected.
    bool scan_single_file_for_bad_ranges(const SourceFileDescriptor& descriptor,
                                         JournalFileEntry& entry, ProgressAccumulator& progress,
                                         JobJournal& journal, const std::wstring& journal_path,
                                         const CancelContext& cancel, bool preserve_existing_coverage,
                                         bool seed_progress_from_existing) {
        AdaptiveBufferController controller =
            create_buffer_controller(descriptor.length, BufferPurpose::PreciseScan);
        std::int32_t active_buffer_size = controller.current_size();
        std::vector<unsigned char> io_buffer(
            static_cast<std::size_t>(std::max(controller.maximum_size(), MinimumRescueBlockSize)));
        FileTransferSession session(descriptor.full_path, std::wstring());
        FileStabilitySnapshot scan_source_stability;
        if (!try_get_file_stability(session.source_handle(), scan_source_stability) ||
            scan_source_stability.length != descriptor.length ||
            scan_source_stability.last_write_utc_ticks != descriptor.last_write_utc_ticks) {
            throw IoError("Source changed after scan enumeration: " +
                          descriptor.relative_path + ".");
        }
        // Persist the binding before the first read. If power or the process is
        // lost mid-file, the checkpointed ranges can be reused only against
        // this exact file generation.
        entry.SourceFileIndex = static_cast<std::int64_t>(
            scan_source_stability.identity.file_index);
        entry.SourceVolumeSerial = static_cast<std::int64_t>(
            scan_source_stability.identity.volume_serial);
        entry.SourceChangeUtcTicks = scan_source_stability.change_utc_ticks;

        if (!preserve_existing_coverage || entry.RescueRanges.empty()) {
            entry.RescueRanges.clear();
            append_range(entry.RescueRanges, 0, descriptor.length, RescueRangeState::Pending);
        }
        entry.RecoveredRanges.clear();

        std::int64_t mapped = apply_known_bad_ranges_from_map(descriptor, entry);
        entry.LastRescuePass = "Init";
        sync_entry_bytes_copied(entry);
        if (seed_progress_from_existing) {
            progress.total_bytes_copied += entry.BytesCopied;
            progress.bytes_reused += entry.BytesCopied;
        }

        emit_progress(progress, descriptor.relative_path,
                      display_progress_bytes(entry, descriptor.length, entry.BytesCopied),
                      descriptor.length, 0, active_buffer_size, entry.LastRescuePass,
                      unreadable_region_count(entry.RescueRanges),
                      rescue_bytes(entry.RescueRanges,
                                   {RescueRangeState::Pending, RescueRangeState::Bad,
                                    RescueRangeState::KnownBad}));
        if (mapped > 0) {
            emit_log("[Rescue Engine] Scan seeded with map hints: " + format_bytes(mapped) +
                     " known unreadable on " + descriptor.relative_path + ".");
        }

        bool use_small_fast = should_use_small_file_fast_path(
            descriptor.length, preserve_existing_coverage && entry.BytesCopied > 0);
        if (use_small_fast) {
            entry.LastRescuePass = "ScanSmallFast";
            scan_small_file_fast(descriptor, entry, session, io_buffer, controller, progress,
                                 journal, journal_path, cancel);
        } else {
            std::vector<RescuePassDefinition> pass_plan = build_rescue_pass_plan(active_buffer_size);
            double bad_density = rescue_bad_density(entry.RescueRanges, descriptor.length);
            for (const auto& base_pass : pass_plan) {
                RescuePassDefinition pass =
                    adapt_pass_for_density(base_pass, bad_density, active_buffer_size);
                if (region_count(entry.RescueRanges, {pass.target_state}) <= 0) continue;
                entry.LastRescuePass = pass.name;

                RescuePassOutcome outcome = execute_rescue_pass(
                    pass, descriptor, entry, session, io_buffer, progress, journal, journal_path,
                    cancel, /*write_recovered*/ false, /*count_failed_as_processed*/ true,
                    &controller);

                std::int32_t remaining_regions = unreadable_region_count(entry.RescueRanges);
                std::int64_t remaining_bytes = rescue_bytes(
                    entry.RescueRanges, {RescueRangeState::Bad, RescueRangeState::KnownBad});
                bad_density = rescue_bad_density(entry.RescueRanges, descriptor.length);
                bool log_summary = models::detail::equals_ignore_case(pass.name, "FastScan")
                                       ? outcome.failed_segments > 0
                                       : (outcome.attempted_segments > 0 ||
                                          outcome.recovered_bytes > 0 || remaining_regions > 0);
                if (log_summary) {
                    char density_text[16];
                    std::snprintf(density_text, sizeof(density_text), "%.1f", bad_density * 100.0);
                    emit_log("[Rescue Engine] " + pass.name + " " + descriptor.relative_path +
                             ": attempts " + std::to_string(outcome.attempted_segments) +
                             ", recovered " + format_bytes(outcome.recovered_bytes) + ", failed " +
                             std::to_string(outcome.failed_segments) + ", remaining bad " +
                             std::to_string(remaining_regions) + " (" +
                             format_bytes(remaining_bytes) + "), density " + density_text + " %.");
                }
                sync_entry_bytes_copied(entry);
                flush_journal(journal_path, journal, true);
            }
        }

        std::int64_t final_unreadable = rescue_bytes(
            entry.RescueRanges, {RescueRangeState::Bad, RescueRangeState::KnownBad});
        std::int32_t final_regions = unreadable_region_count(entry.RescueRanges);
        entry.RecoveredRanges = merge_byte_ranges(snapshot_by_state(
            entry.RescueRanges, {RescueRangeState::Bad, RescueRangeState::KnownBad}));
        sync_entry_bytes_copied(entry);

        FileStabilitySnapshot completed_scan_stability;
        if (!try_get_file_stability(session.source_handle(), completed_scan_stability) ||
            !file_stability_matches(scan_source_stability, completed_scan_stability)) {
            throw IoError("Source content or metadata changed during scan: " +
                          descriptor.relative_path + ".");
        }
        entry.SourceFileIndex = static_cast<std::int64_t>(
            completed_scan_stability.identity.file_index);
        entry.SourceVolumeSerial = static_cast<std::int64_t>(
            completed_scan_stability.identity.volume_serial);
        entry.SourceChangeUtcTicks = completed_scan_stability.change_utc_ticks;

        emit_progress(progress, descriptor.relative_path,
                      display_progress_bytes(entry, descriptor.length, entry.BytesCopied),
                      descriptor.length, 0, active_buffer_size, "ScanComplete", final_regions,
                      final_unreadable);
        if (final_unreadable > 0) {
            emit_log("[Rescue Engine] Scan detected unreadable ranges on " +
                     descriptor.relative_path + ": " + format_bytes(final_unreadable) + " in " +
                     std::to_string(final_regions) + " segment(s).");
        }
        return final_unreadable > 0;
    }

    // Returns "" on success/skip, or the error message on scan failure.
    std::string process_precise_scan_file(const SourceFileDescriptor& descriptor,
                                          JournalFileEntry& entry, ProgressAccumulator& progress,
                                          JobJournal& journal, const std::wstring& journal_path,
                                          bool preserve_existing_coverage,
                                          bool seed_progress_from_existing,
                                          const CancelContext& cancel) {
        std::optional<std::string> scan_error;
        bool detected = false;

        CancelContext file_cancel = cancel;
        if (options_.PerFileTimeout.ticks > 0) {
            file_cancel = cancel.with_deadline(options_.PerFileTimeout.ticks / time::TicksPerMillisecond);
        }

        try {
            entry.State = FileCopyState::InProgress;
            entry.LastError.clear();
            entry.DoNotRetry = false;
            flush_journal(journal_path, journal, false);

            detected = scan_single_file_for_bad_ranges(descriptor, entry, progress, journal,
                                                       journal_path, file_cancel,
                                                       preserve_existing_coverage,
                                                       seed_progress_from_existing);
        } catch (const SourceMutationSkipped& ex) {
            progress.skipped_files += 1;
            progress.total_bytes_copied += descriptor.length;
            progress.bytes_skipped += descriptor.length;
            entry.State = FileCopyState::Pending;
            entry.LastError = ex.what();
            entry.DoNotRetry = false;
            emit_log("Skipped scan: " + descriptor.relative_path + " (" + ex.what() + ")");
            flush_journal(journal_path, journal, true);
            return std::string();
        } catch (const FragileReadSkip& ex) {
            handle_fragile_read_skip(descriptor, entry, progress, journal, journal_path,
                                     ex, cancel, "scan");
            return std::string();
        } catch (const OperationCanceled& oc) {
            if (oc.user_requested) throw;
            std::int64_t timeout_seconds = options_.PerFileTimeout.ticks / time::TicksPerSecond;
            scan_error = "Per-file timeout (" + std::to_string(timeout_seconds) +
                         " sec) reached while scanning " + descriptor.relative_path + ".";
        } catch (const std::exception& ex) {
            scan_error = ex.what();
        }

        if (!scan_error.has_value()) {
            progress.completed_files += 1;
            if (detected) progress.recovered_files += 1;
            entry.State = detected ? FileCopyState::CompletedWithRecovery : FileCopyState::Completed;
            entry.BytesCopied = descriptor.length;
            entry.LastError.clear();
            entry.DoNotRetry = false;
            emit_progress(progress, descriptor.relative_path, descriptor.length, descriptor.length,
                          0, resolve_buffer_size_for_file(descriptor.length), entry.LastRescuePass,
                          unreadable_region_count(entry.RescueRanges),
                          rescue_bytes(entry.RescueRanges,
                                       {RescueRangeState::Pending, RescueRangeState::Bad,
                                        RescueRangeState::KnownBad}));
            try_persist_bad_range_map_entry(descriptor, entry, false);
            flush_journal(journal_path, journal, false);
            return std::string();
        }

        progress.failed_files += 1;
        entry.State = FileCopyState::Failed;
        entry.LastError = *scan_error;
        entry.DoNotRetry = false;
        emit_log("Scan failed: " + descriptor.relative_path + " (" + *scan_error + ")");
        register_fragile_failure_and_maybe_cooldown(descriptor.relative_path, *scan_error, cancel);
        try_persist_bad_range_map_entry(descriptor, entry, false);
        flush_journal(journal_path, journal, false);
        return *scan_error;
    }

    // ---- Parallel small-file copy phase (CopyFileEx workers) ---------------

    std::vector<std::string> copy_small_files_parallel(
        const std::vector<SourceFileDescriptor>& files, const std::wstring& destination_root,
        JobJournal& journal, const std::wstring& journal_path, const CancelContext& cancel,
        std::int32_t worker_count) {
        std::vector<std::string> completed_paths;
        if (files.empty()) return completed_paths;

        emit_log("Parallel small file phase: " + std::to_string(files.size()) + " file(s) using " +
                 std::to_string(worker_count) + " worker(s).");

        for (const auto& descriptor : files) {
            std::wstring destination_path = destination_path_for(destination_root, descriptor);
            std::wstring destination_directory = storage::fsutil::get_directory_name(destination_path);
            if (!destination_directory.empty() && add_created_directory(destination_directory)) {
                storage::fsutil::create_directories(destination_directory);
            }
        }

        std::int32_t concurrency = std::max(
            1, std::min<std::int32_t>(worker_count,
                                      static_cast<std::int32_t>(files.size())));
        std::mutex work_lock;
        std::size_t next_index = 0;
        std::atomic<std::int32_t> completed_count{0};
        std::atomic<std::int32_t> failed_count{0};
        std::atomic<bool> cancelled{false};

        HANDLE checkpoint_stop = CreateEventW(nullptr, TRUE, FALSE, nullptr);
        std::thread checkpoint_thread;
        if (checkpoint_stop != nullptr) {
            checkpoint_thread = std::thread([this, &journal, &journal_path,
                                              checkpoint_stop]() {
                DWORD interval_ms = 5000;
                while (WaitForSingleObject(checkpoint_stop, interval_ms) == WAIT_TIMEOUT) {
                    JobJournal snapshot;
                    {
                        std::lock_guard<std::mutex> guard(journal_lock_);
                        snapshot = journal;
                    }
                    const ULONGLONG started = GetTickCount64();
                    try {
                        (void)journal_store_.save(journal_path, std::move(snapshot));
                    } catch (const std::exception&) {
                        // The joined phase performs a final synchronous save;
                        // keep copying if an intermediate redundancy write fails.
                    }
                    const ULONGLONG cost = std::max<ULONGLONG>(1, GetTickCount64() - started);
                    interval_ms = static_cast<DWORD>(std::clamp<ULONGLONG>(
                        cost * (JournalFlushDutyDivisor - 1), 5000,
                        MaximumJournalFlushIntervalMs));
                }
            });
        }

        struct QuietCallbackState {
            const CancelContext* cancel;
            ExecutionControl* control;
        };
        auto quiet_routine = [](LARGE_INTEGER, LARGE_INTEGER, LARGE_INTEGER, LARGE_INTEGER, DWORD,
                                DWORD, HANDLE, HANDLE, LPVOID context) -> DWORD {
            auto* state = static_cast<QuietCallbackState*>(context);
            if (state->cancel->is_cancelled() ||
                (state->control != nullptr && state->control->is_paused())) {
                return PROGRESS_CANCEL;
            }
            return PROGRESS_QUIET;
        };

        std::vector<std::thread> workers;
        for (std::int32_t worker_index = 0; worker_index < concurrency; ++worker_index) {
            workers.emplace_back([&]() {
                while (true) {
                    SourceFileDescriptor descriptor;
                    {
                        std::lock_guard<std::mutex> guard(work_lock);
                        if (next_index >= files.size()) return;
                        descriptor = files[next_index++];
                    }
                    if (cancel.is_cancelled()) {
                        cancelled.store(true);
                        return;
                    }
                    try {
                        if (control_ != nullptr) control_->wait_if_paused(cancel);
                    } catch (const OperationCanceled&) {
                        cancelled.store(true);
                        return;
                    }

                    try {
                        std::wstring destination_path = destination_path_for(destination_root, descriptor);
                        FileStabilitySnapshot source_stability;
                        const bool source_stability_valid =
                            try_get_file_stability(descriptor.full_path, source_stability);
                        if (source_stability_valid) note_hard_link_topology(source_stability);
                        if (source_stability_valid &&
                            (source_stability.length != descriptor.length ||
                             source_stability.last_write_utc_ticks !=
                                 descriptor.last_write_utc_ticks)) {
                            failed_count += 1;
                            continue;
                        }
                        FileIdentity source_identity = source_stability.identity;
                        const bool source_identity_valid = source_stability_valid ||
                            try_get_file_identity(descriptor.full_path, source_identity);
                        if (source_identity_valid) {
                            std::lock_guard<std::mutex> guard(journal_lock_);
                            JournalFileEntry* entry = journal.Files.find(descriptor.relative_path);
                            if (entry != nullptr) {
                                entry->SourceFileIndex = static_cast<std::int64_t>(source_identity.file_index);
                                entry->SourceVolumeSerial = static_cast<std::int64_t>(source_identity.volume_serial);
                                if (source_stability_valid) {
                                    entry->SourceChangeUtcTicks =
                                        source_stability.change_utc_ticks;
                                }
                            }
                        }
                        FileSecuritySnapshot source_security = capture_file_security(descriptor.full_path);
                        NamedStreamEnumeration source_streams = enumerate_named_streams(descriptor.full_path);
                        DestinationStage stage(destination_path, cancel);
                        const std::wstring& working_path = stage.working_path();
                        HANDLE reset = CreateFileW(working_path.c_str(), GENERIC_WRITE,
                                                   FILE_SHARE_READ, nullptr, CREATE_ALWAYS,
                                                   FILE_ATTRIBUTE_NORMAL, nullptr);
                        if (reset == INVALID_HANDLE_VALUE) {
                            failed_count += 1;
                            continue;
                        }
                        CloseHandle(reset);

                        DWORD flags = options_.AllowDecryptedDestination
                                          ? COPY_FILE_ALLOW_DECRYPTED_DESTINATION
                                          : 0;
                        if (descriptor.length >= 16LL * 1024 * 1024) flags |= COPY_FILE_NO_BUFFERING;
                        QuietCallbackState state{&cancel, control_};
                        BOOL cancel_flag = FALSE;
                        BOOL ok = CopyFileExW(descriptor.full_path.c_str(), working_path.c_str(),
                                              quiet_routine, &state, &cancel_flag, flags);
                        if (!ok) {
                            failed_count += 1;
                            continue;
                        }
                        bytes_read_.fetch_add(descriptor.length);
                        bytes_written_.fetch_add(descriptor.length);

                        stage.make_writable();
                        ensure_source_identity_unchanged(descriptor, source_identity, source_identity_valid);
                        verify_file_with_retries(descriptor.full_path, working_path,
                                                 descriptor.relative_path,
                                                 models::VerificationMode::Full, cancel);
                        preserve_file_metadata(descriptor, working_path, source_security,
                                               source_streams, cancel,
                                               models::VerificationMode::Full);
                        ensure_source_stability_unchanged(descriptor, source_stability,
                                                          source_stability_valid);
                        ensure_media_identity_integrity(utf8_to_wide(options_.SourceRoot), destination_root,
                                                        true, true);
                        publish_stage(stage, descriptor, false);
                        {
                            std::lock_guard<std::mutex> guard(journal_lock_);
                            JournalFileEntry* entry = journal.Files.find(descriptor.relative_path);
                            if (entry != nullptr) {
                                entry->State = FileCopyState::Completed;
                                entry->BytesCopied = descriptor.length;
                                entry->LastError.clear();
                                entry->DoNotRetry = false;
                                entry->LastRescuePass = "ParallelNativeFast";
                                entry->RecoveredRanges.clear();
                                entry->RescueRanges.clear();
                                append_range(entry->RescueRanges, 0, descriptor.length,
                                             RescueRangeState::Good);
                            }
                        }
                        parallel_native_fast_path_files_ += 1;
                        {
                            std::lock_guard<std::mutex> guard(work_lock);
                            completed_paths.push_back(descriptor.relative_path);
                        }
                        completed_count += 1;
                    } catch (const OperationCanceled&) {
                        cancelled.store(true);
                        return;
                    } catch (const std::exception&) {
                        failed_count += 1;
                    }
                }
            });
        }
        for (auto& worker : workers) worker.join();

        if (checkpoint_stop != nullptr) SetEvent(checkpoint_stop);
        if (checkpoint_thread.joinable()) checkpoint_thread.join();
        if (checkpoint_stop != nullptr) CloseHandle(checkpoint_stop);

        // Save completed staged publications before cancellation unwinds. The
        // previous order threw first and discarded every journal update from a
        // partially completed parallel phase.
        if (completed_count.load() > 0) {
            save_journal_now(journal_path, journal);
        }

        if (cancelled.load() || cancel.is_cancelled()) throw OperationCanceled{true};

        emit_log("Parallel small file phase complete: " + std::to_string(completed_count.load()) +
                 " succeeded, " + std::to_string(failed_count.load()) +
                 " deferred to sequential path.");
        return completed_paths;
    }

    // ---- Media identity + availability -----------------------------------

    void initialize_media_identity_expectations(const std::wstring& source_root,
                                                const std::wstring& destination_root) {
        std::string source_identity = trim_ascii(options_.ExpectedSourceIdentity);
        std::string destination_identity = trim_ascii(options_.ExpectedDestinationIdentity);
        bool captured_source = false;
        bool captured_destination = false;

        if (source_identity.empty()) {
            source_identity = resolve_media_identity(source_root);
            captured_source = !source_identity.empty();
        }
        if (destination_identity.empty()) {
            destination_identity = resolve_media_identity(destination_root);
            captured_destination = !destination_identity.empty();
        }

        {
            std::lock_guard<std::mutex> guard(media_identity_lock_);
            expected_source_identity_ = source_identity;
            expected_destination_identity_ = destination_identity;
            if (captured_source) options_.ExpectedSourceIdentity = source_identity;
            if (captured_destination) options_.ExpectedDestinationIdentity = destination_identity;
        }
        source_media_identity_accepted_.store(true, std::memory_order_release);
        destination_media_identity_accepted_.store(true, std::memory_order_release);
        if (captured_source) emit_log("Source media identity baseline captured.");
        if (captured_destination) emit_log("Destination media identity baseline captured.");
    }

    static std::string trim_ascii(const std::string& text) {
        std::size_t begin = 0, end = text.size();
        auto is_space = [](char c) { return c == ' ' || c == '\t' || c == '\r' || c == '\n'; };
        while (begin < end && is_space(text[begin])) ++begin;
        while (end > begin && is_space(text[end - 1])) --end;
        return text.substr(begin, end - begin);
    }

    static std::string resolve_media_identity(const std::wstring& path_value) {
        return storage::fsutil::resolve_media_identity(path_value);
    }

    static bool identities_equivalent(const std::string& expected, const std::string& current) {
        std::string e = trim_ascii(expected);
        std::string c = trim_ascii(current);
        if (e.empty() || c.empty()) return e.size() == c.size();
        if (models::detail::equals_ignore_case(e, c)) return true;
        auto legacy_unc_base = [](const std::string& identity) {
            if (!detail::starts_with_ignore_case(identity, "unc:")) return std::string();
            const std::size_t separator = identity.find('|');
            return identity.substr(0, separator);
        };
        const std::string expected_unc = legacy_unc_base(e);
        const std::string current_unc = legacy_unc_base(c);
        if (!expected_unc.empty() && !current_unc.empty() &&
            (e.find('|') == std::string::npos) !=
                (c.find('|') == std::string::npos) &&
            models::detail::equals_ignore_case(expected_unc, current_unc)) {
            return true;
        }
        // A new identity is vol:<volume-guid>:<serial>. Permit a serial-only
        // comparison only when exactly one side is a legacy vol:<serial>
        // value; two new identities with different GUIDs are not equivalent.
        auto extract_serial = [](const std::string& identity, std::string& serial) {
            if (!detail::starts_with_ignore_case(identity, "vol:")) return false;
            std::size_t position = identity.find_last_of(':');
            serial = identity.substr(position + 1);
            return !serial.empty();
        };
        auto is_legacy_volume_identity = [](const std::string& identity) {
            if (!detail::starts_with_ignore_case(identity, "vol:")) return false;
            return identity.find(':', 4) == std::string::npos;
        };
        std::string es, cs;
        if (is_legacy_volume_identity(e) != is_legacy_volume_identity(c) &&
            extract_serial(e, es) && extract_serial(c, cs)) {
            return models::detail::equals_ignore_case(es, cs);
        }
        return false;
    }

    bool should_probe_media_identity(bool is_source) {
        ULONGLONG now = GetTickCount64();
        std::atomic<ULONGLONG>& last = is_source ? last_source_identity_probe_tick_
                                                 : last_destination_identity_probe_tick_;
        ULONGLONG previous = last.load(std::memory_order_relaxed);
        while (previous == 0 || now - previous >= 1000) {
            if (last.compare_exchange_weak(previous, now, std::memory_order_acq_rel,
                                           std::memory_order_relaxed)) {
                return true;
            }
        }
        return false;
    }

    std::string expected_media_identity(bool is_source) {
        std::lock_guard<std::mutex> guard(media_identity_lock_);
        return is_source ? expected_source_identity_ : expected_destination_identity_;
    }

    void throw_if_media_identity_mismatch(const std::wstring& path_value,
                                          const std::string& expected_identity, bool is_source) {
        std::string expected = trim_ascii(expected_identity);
        if (expected.empty()) return;
        std::string current = resolve_media_identity(path_value);
        if (current.empty()) return;
        if (identities_equivalent(expected, current)) return;
        throw IoError(std::string(is_source ? "Source" : "Destination") +
                      " media identity mismatch. Expected '" + expected + "', found '" + current + "'.");
    }

    void throw_if_media_identity_mismatch_throttled(const std::wstring& path_value,
                                                    bool is_source, bool force = false) {
        if (!force && !should_probe_media_identity(is_source)) return;
        throw_if_media_identity_mismatch(path_value, expected_media_identity(is_source), is_source);
    }

    void ensure_media_identity_integrity(const std::wstring& source_root,
                                         const std::wstring& destination_root,
                                         bool include_destination, bool force) {
        throw_if_media_identity_mismatch_throttled(source_root, true, force);
        if (!include_destination) return;
        throw_if_media_identity_mismatch_throttled(destination_root, false, force);
    }

    bool is_media_identity_accepted(const std::wstring& path_value, bool is_source) {
        std::string current = resolve_media_identity(path_value);
        std::string expected_identity;
        bool captured = false;
        {
            // Parallel scan workers all pass through this path. Keep the shared
            // identity strings behind one lock: assigning even an unchanged
            // std::string concurrently can double-free its heap allocation.
            std::lock_guard<std::mutex> guard(media_identity_lock_);
            std::string& stored = is_source ? expected_source_identity_
                                            : expected_destination_identity_;
            expected_identity = stored;
            if (!current.empty() && expected_identity.empty()) {
                stored = current;
                expected_identity = current;
                if (is_source) options_.ExpectedSourceIdentity = current;
                else options_.ExpectedDestinationIdentity = current;
                captured = true;
            }
        }
        if (current.empty()) return expected_identity.empty();
        if (captured) {
            emit_log(std::string(is_source ? "Source" : "Destination") +
                     " media identity baseline captured.");
            return true;
        }
        if (identities_equivalent(expected_identity, current)) return true;

        ULONGLONG now = GetTickCount64();
        std::atomic<ULONGLONG>& last_log = is_source ? last_source_mismatch_log_tick_
                                                     : last_destination_mismatch_log_tick_;
        ULONGLONG previous_log = last_log.load(std::memory_order_relaxed);
        while (previous_log == 0 || now - previous_log >= 5000) {
            if (last_log.compare_exchange_weak(previous_log, now, std::memory_order_acq_rel,
                                               std::memory_order_relaxed)) {
                emit_log(std::string(is_source ? "Source" : "Destination") +
                         " media identity mismatch on '" + wide_to_utf8(path_value) +
                         "'. Expected " + expected_identity + ", found " + current + ".");
                break;
            }
        }
        return false;
    }

    bool cached_media_identity_accepted(const std::wstring& path_value, bool is_source) {
        std::atomic<bool>& accepted = is_source ? source_media_identity_accepted_
                                                : destination_media_identity_accepted_;
        if (should_probe_media_identity(is_source)) {
            accepted.store(is_media_identity_accepted(path_value, is_source),
                           std::memory_order_release);
        }
        return accepted.load(std::memory_order_acquire);
    }

    bool is_source_available(const std::wstring& source_root) {
        if (source_root.empty() || !directory_exists(source_root)) return false;
        return cached_media_identity_accepted(source_root, true);
    }

    bool is_destination_available(const std::wstring& target_path) {
        if (target_path.empty()) return false;
        std::wstring full = storage::fsutil::get_full_path(target_path);
        if (full.size() < 2) return false;
        std::wstring root = full.substr(0, 2) + L"\\";
        if (full[1] == L':' && !directory_exists(root)) return false;

        std::wstring probe = full;
        bool existing_path_found = false;
        while (!probe.empty()) {
            if (storage::fsutil::file_exists(probe) || directory_exists(probe)) {
                existing_path_found = true;
                break;
            }
            std::wstring parent = storage::fsutil::get_directory_name(probe);
            if (parent == probe) break;
            probe = parent;
        }
        if (!existing_path_found) return false;
        return cached_media_identity_accepted(full, false);
    }

    void wait_for_media_availability(const std::wstring& source_root,
                                     const std::wstring& destination_root,
                                     const CancelContext& cancel) {
        if (!options_.WaitForMediaAvailability) return;
        ULONGLONG last_log = 0;
        while (true) {
            cancel.throw_if_cancelled();
            if (control_ != nullptr) control_->wait_if_paused(cancel);

            bool source_ok = is_source_available(source_root);
            bool destination_ok = is_destination_available(destination_root);
            if (source_ok && destination_ok) return;

            ULONGLONG now = GetTickCount64();
            if (last_log == 0 || now - last_log >= 5000) {
                emit_log(std::string("Waiting for media availability. Source=") +
                         (source_ok ? "online" : "offline") + " Destination=" +
                         (destination_ok ? "online" : "offline"));
                last_log = now;
            }
            cancel.delay(1000);
        }
    }

    void wait_for_source_file(const std::wstring& source_path, const CancelContext& cancel,
                              bool allow_without_media_wait) {
        if (!options_.WaitForMediaAvailability && !allow_without_media_wait) return;
        ULONGLONG last_log = 0;
        while (true) {
            cancel.throw_if_cancelled();
            if (control_ != nullptr) control_->wait_if_paused(cancel);

            bool file_ok = storage::fsutil::file_exists(source_path);
            bool identity_ok = file_ok && cached_media_identity_accepted(source_path, true);
            if (file_ok && identity_ok) return;

            ULONGLONG now = GetTickCount64();
            if (last_log == 0 || now - last_log >= 5000) {
                emit_log("Waiting for source path: " + wide_to_utf8(source_path));
                last_log = now;
            }
            cancel.delay(1000);
        }
    }

    void wait_for_destination_path(const std::wstring& destination_path, const CancelContext& cancel) {
        if (!options_.WaitForMediaAvailability) return;
        std::wstring directory = storage::fsutil::get_directory_name(destination_path);
        std::wstring probe = directory.empty() ? destination_path : directory;
        ULONGLONG last_log = 0;
        while (true) {
            cancel.throw_if_cancelled();
            if (control_ != nullptr) control_->wait_if_paused(cancel);
            if (is_destination_available(probe)) return;
            ULONGLONG now = GetTickCount64();
            if (last_log == 0 || now - last_log >= 5000) {
                emit_log("Waiting for destination path: " + wide_to_utf8(probe));
                last_log = now;
            }
            cancel.delay(1000);
        }
    }

    void wait_for_source_read_access(const std::wstring& source_path, const CancelContext& cancel) {
        wait_for_path_access(
            [&source_path]() {
                if (!storage::fsutil::file_exists(source_path)) return false;
                HANDLE handle = CreateFileW(source_path.c_str(), GENERIC_READ,
                                            FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE,
                                            nullptr, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, nullptr);
                if (handle == INVALID_HANDLE_VALUE) return false;
                CloseHandle(handle);
                return true;
            },
            "Waiting for source lock/access release: " + wide_to_utf8(source_path), cancel);
    }

    void wait_for_destination_write_access(const std::wstring& destination_path,
                                           const CancelContext& cancel) {
        wait_for_path_access(
            [&destination_path]() {
                std::wstring directory = storage::fsutil::get_directory_name(destination_path);
                if (!directory.empty()) storage::fsutil::create_directories(directory);
                HANDLE handle = CreateFileW(destination_path.c_str(), GENERIC_WRITE, FILE_SHARE_READ,
                                            nullptr, OPEN_ALWAYS, FILE_ATTRIBUTE_NORMAL, nullptr);
                if (handle == INVALID_HANDLE_VALUE) return false;
                CloseHandle(handle);
                return true;
            },
            "Waiting for destination lock/access release: " + wide_to_utf8(destination_path), cancel);
    }

    void wait_for_path_access(const std::function<bool()>& probe, const std::string& log_text,
                              const CancelContext& cancel) {
        ULONGLONG last_log = 0;
        std::int64_t interval_ms = options_.LockContentionProbeInterval.ticks / time::TicksPerMillisecond;
        if (interval_ms <= 0) interval_ms = 500;

        while (true) {
            cancel.throw_if_cancelled();
            if (control_ != nullptr) control_->wait_if_paused(cancel);

            bool ready = false;
            try {
                ready = probe();
            } catch (...) {
                ready = false;
            }
            if (ready) return;

            if (options_.WaitForMediaAvailability) {
                wait_for_media_availability(utf8_to_wide(options_.SourceRoot),
                                            utf8_to_wide(options_.DestinationRoot), cancel);
            }

            ULONGLONG now = GetTickCount64();
            if (last_log == 0 || now - last_log >= 5000) {
                emit_log(log_text);
                last_log = now;
            }
            cancel.delay(interval_ms);
        }
    }

    // ---- Throttle, salvage fill, capacity ---------------------------------

    void apply_throughput_throttle(std::int32_t bytes_transferred, const CancelContext& cancel) {
        if (options_.MaxThroughputBytesPerSecond <= 0 || bytes_transferred <= 0) return;

        ULONGLONG now = GetTickCount64();
        if (throttle_window_start_tick_ == 0) {
            throttle_window_start_tick_ = now;
            throttle_window_bytes_ = 0;
        }
        throttle_window_bytes_ += bytes_transferred;

        double expected_seconds = static_cast<double>(throttle_window_bytes_) /
                                  static_cast<double>(options_.MaxThroughputBytesPerSecond);
        double actual_seconds = static_cast<double>(now - throttle_window_start_tick_) / 1000.0;
        if (expected_seconds > actual_seconds) {
            std::int64_t delay_ms = static_cast<std::int64_t>(
                (expected_seconds - actual_seconds) * 1000.0 + 0.999);
            if (delay_ms > 0) cancel.delay(delay_ms);
        }
        if (static_cast<double>(GetTickCount64() - throttle_window_start_tick_) / 1000.0 >= 5.0) {
            throttle_window_start_tick_ = GetTickCount64();
            throttle_window_bytes_ = 0;
        }
    }

    void fill_salvage_buffer(unsigned char* buffer, std::int32_t count) {
        switch (options_.SalvageFillPatternValue) {
            case models::SalvageFillPattern::Ones:
                std::fill_n(buffer, count, static_cast<unsigned char>(0xFF));
                break;
            case models::SalvageFillPattern::Random: {
                std::vector<unsigned char> random_bytes(static_cast<std::size_t>(count));
                crypto::fill_random(random_bytes);
                std::copy(random_bytes.begin(), random_bytes.end(), buffer);
                break;
            }
            default:
                std::fill_n(buffer, count, static_cast<unsigned char>(0));
                break;
        }
    }

    std::string describe_salvage_fill() const {
        switch (options_.SalvageFillPatternValue) {
            case models::SalvageFillPattern::Ones: return "0xFF";
            case models::SalvageFillPattern::Random: return "random";
            default: return "zero";
        }
    }

    void emit_destination_capacity_warning(const std::wstring& destination_root,
                                           std::int64_t total_bytes) {
        if (total_bytes <= 0 || destination_root.empty()) return;
        std::wstring full = storage::fsutil::get_full_path(destination_root);
        if (full.size() < 2 || full[1] != L':') return;
        std::wstring root = full.substr(0, 2) + L"\\";
        ULARGE_INTEGER available{}, total{}, free_total{};
        if (!GetDiskFreeSpaceExW(root.c_str(), &available, &total, &free_total)) return;
        if (static_cast<std::int64_t>(available.QuadPart) < total_bytes) {
            emit_log("Destination free-space warning: estimated " + format_bytes(total_bytes) +
                     " required, " + format_bytes(static_cast<std::int64_t>(available.QuadPart)) +
                     " available.");
        }
    }

    // ---- Progress, logging, result ----------------------------------------

    void emit_progress(const ProgressAccumulator& progress, const std::string& current_file,
                       std::int64_t current_file_bytes, std::int64_t current_file_total,
                       std::int32_t last_chunk, std::int32_t buffer_size_bytes,
                       const std::string& rescue_pass = std::string(),
                       std::int32_t rescue_bad_regions = 0, std::int64_t rescue_remaining = 0,
                       std::int32_t active_file_count = 0, std::int32_t scan_worker_count = 0,
                       const std::vector<std::string>* active_files = nullptr) {
        if (!progress_callback_) return;

        std::int64_t safe_total = std::max<std::int64_t>(0, current_file_total);
        std::int64_t safe_bytes = std::max<std::int64_t>(0, current_file_bytes);
        if (safe_total > 0) safe_bytes = std::min(safe_bytes, safe_total);
        std::int64_t safe_all = std::max<std::int64_t>(0, progress.total_bytes);
        std::int64_t safe_copied = std::max<std::int64_t>(0, progress.total_bytes_copied);
        if (safe_all > 0) safe_copied = std::min(safe_copied, safe_all);

        CopyProgressSnapshot snapshot;
        snapshot.SourceMediaIdentity = expected_media_identity(true);
        snapshot.DestinationMediaIdentity = expected_media_identity(false);
        snapshot.CurrentFile = current_file;
        snapshot.CurrentFileBytesCopied = safe_bytes;
        snapshot.CurrentFileBytesTotal = safe_total;
        snapshot.TotalBytesCopied = safe_copied;
        snapshot.WorkBytesCompleted = safe_copied;
        snapshot.BytesRead = std::max<std::int64_t>(0, bytes_read_.load());
        snapshot.BytesWritten = std::max<std::int64_t>(0, bytes_written_.load());
        snapshot.BytesVerified = std::max<std::int64_t>(0, bytes_verified_.load());
        snapshot.BytesSkipped = std::max<std::int64_t>(0, progress.bytes_skipped);
        snapshot.BytesReused = std::max<std::int64_t>(0, progress.bytes_reused);
        snapshot.TotalBytes = safe_all;
        snapshot.LastChunkBytesTransferred = std::max(0, last_chunk);
        snapshot.BufferSizeBytes = std::max(0, buffer_size_bytes);
        snapshot.CompletedFiles = progress.completed_files;
        snapshot.FailedFiles = progress.failed_files;
        snapshot.RecoveredFiles = progress.recovered_files;
        snapshot.SkippedFiles = progress.skipped_files;
        snapshot.TotalFiles = progress.total_files;
        snapshot.RescuePass = rescue_pass;
        snapshot.RescueBadRegionCount = std::max(0, rescue_bad_regions);
        snapshot.RescueRemainingBytes = std::max<std::int64_t>(0, rescue_remaining);
        snapshot.ActiveFileCount = std::max(0, active_file_count);
        snapshot.ScanWorkerCount = std::max(0, scan_worker_count);
        if (active_files != nullptr) {
            for (const auto& active_file : *active_files) {
                if (snapshot.ActiveFiles.size() >= 8) break;
                if (!active_file.empty()) snapshot.ActiveFiles.push_back(active_file);
            }
        }
        progress_callback_(snapshot);
    }

    void emit_media_identity_snapshot() {
        if (!progress_callback_) return;
        CopyProgressSnapshot snapshot;
        snapshot.SourceMediaIdentity = expected_media_identity(true);
        snapshot.DestinationMediaIdentity = expected_media_identity(false);
        // Keep the initialization event at zero progress instead of briefly
        // rendering an empty job as 100 percent complete.
        snapshot.TotalFiles = 1;
        progress_callback_(snapshot);
    }

    void emit_log(const std::string& message) {
        if (!log_callback_) return;
        SYSTEMTIME local;
        GetLocalTime(&local);
        char prefix[16];
        std::snprintf(prefix, sizeof(prefix), "[%02u:%02u:%02u] ", local.wHour, local.wMinute,
                      local.wSecond);
        log_callback_(prefix + message);
    }

    void emit_run_summary(const CopyJobResult& result) {
        std::string speed = format_bytes(static_cast<std::int64_t>(
                                std::max(0.0, result.AverageBytesPerSecond))) + "/s";
        const bool scan = options_.OperationMode == models::JobOperationMode::ScanOnly;
        const std::int64_t primary_bytes = scan ? result.BytesRead : result.BytesWritten;
        emit_log(std::string("Run summary: ") +
                 (result.Succeeded ? "succeeded" : "completed with failures") + "; " +
                 (scan ? "read " : "wrote ") + format_bytes(primary_bytes) + " in " +
                 std::to_string(result.ElapsedMilliseconds) + " ms (" + speed + ").");
        if (result.NativeFastPathFiles > 0 || result.ParallelNativeFastPathFiles > 0 ||
            result.ManagedCopyFiles > 0 || result.NativeFallbackFiles > 0) {
            emit_log("Engine summary: native=" + std::to_string(result.NativeFastPathFiles) +
                     ", parallel-native=" + std::to_string(result.ParallelNativeFastPathFiles) +
                     ", managed=" + std::to_string(result.ManagedCopyFiles) +
                     ", native-fallback=" + std::to_string(result.NativeFallbackFiles) + ".");
        }
    }

    CopyJobResult create_result(const ProgressAccumulator& progress, const std::wstring& journal_path,
                                bool succeeded, bool cancelled, const std::string& error_message) {
        std::int64_t elapsed_ms = run_started_tick_ == 0
                                      ? 0
                                      : static_cast<std::int64_t>(GetTickCount64() - run_started_tick_);
        const std::int64_t actual_bytes =
            options_.OperationMode == models::JobOperationMode::ScanOnly
                ? bytes_read_.load()
                : bytes_written_.load();
        double average = elapsed_ms > 0 ? static_cast<double>(actual_bytes) /
                                              (static_cast<double>(elapsed_ms) / 1000.0)
                                        : 0.0;

        CopyJobResult result;
        // A copy is successful only when every source file was copied exactly.
        // Recovered ranges are synthetic bytes and skipped files are missing
        // data, so neither condition may be reported as a clean success. Scan
        // mode is different: recovered files are the scan's findings, not
        // substituted destination bytes.
        const bool incomplete = source_enumeration_incomplete_ ||
                                progress.failed_files > 0 || progress.skipped_files > 0 ||
                                (options_.OperationMode == models::JobOperationMode::Copy &&
                                 progress.recovered_files > 0);
        result.Succeeded = succeeded && !cancelled && !incomplete;
        result.Cancelled = cancelled;
        result.TotalFiles = progress.total_files;
        result.CompletedFiles = progress.completed_files;
        result.FailedFiles = progress.failed_files;
        result.RecoveredFiles = progress.recovered_files;
        result.SkippedFiles = progress.skipped_files;
        result.TotalBytes = progress.total_bytes;
        result.CopiedBytes = options_.OperationMode == models::JobOperationMode::Copy
                                 ? std::max<std::int64_t>(0, bytes_written_.load())
                                 : 0;
        result.WorkBytesCompleted = progress.total_bytes_copied;
        result.BytesRead = std::max<std::int64_t>(0, bytes_read_.load());
        result.BytesWritten = std::max<std::int64_t>(0, bytes_written_.load());
        result.BytesVerified = std::max<std::int64_t>(0, bytes_verified_.load());
        result.BytesSkipped = std::max<std::int64_t>(0, progress.bytes_skipped);
        result.BytesReused = std::max<std::int64_t>(0, progress.bytes_reused);
        result.TransferEnginePolicyValue = options_.TransferEnginePolicyValue;
        result.ElapsedMilliseconds = elapsed_ms;
        result.AverageBytesPerSecond = average;
        result.NativeFastPathFiles = native_fast_path_files_;
        result.ParallelNativeFastPathFiles = parallel_native_fast_path_files_;
        result.ManagedCopyFiles = managed_copy_files_;
        result.NativeFallbackFiles = native_fallback_files_;
        result.JournalPath = wide_to_utf8(journal_path);
        result.ErrorMessage = error_message.empty() ? source_enumeration_error_ : error_message;
        {
            std::lock_guard<std::mutex> guard(metadata_warning_lock_);
            result.IntegrityNotice = integrity_warning_;
            result.MetadataNotice = metadata_notice_;
        }
        return result;
    }
};

} // namespace xact::engine
