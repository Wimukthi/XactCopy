// -----------------------------------------------------------------------------
// File: src\worker\engine.h
// Purpose: Native port of ResilientCopyService — the XactCopy copy engine:
//          journal-merged resumable copies, rescue-pass pipeline (FastScan /
//          TrimSweep / TrimSweepReverse / Scrape / RetryBad), salvage fill,
//          retry/backoff with contention and availability policies, media
//          identity guard, native CopyFileEx fast path, and verification.
//          Phase 3a scope: Copy mode. ScanOnly, bad-range-map hints, raw disk
//          scan, parallel small-file phase, and adaptive buffer sizing are
//          declared below and refused/degraded explicitly until ported.
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
        throttle_window_start_tick_ = 0;
        throttle_window_bytes_ = 0;

        if (fault_injector_.has_value()) {
            emit_log("[DevFault] Enabled: " + fault_injector_->description());
        }

        const bool scan_only = options_.OperationMode == models::JobOperationMode::ScanOnly;
        std::wstring source_root = storage::fsutil::get_full_path(utf8_to_wide(options_.SourceRoot));
        std::wstring destination_root = storage::fsutil::get_full_path(utf8_to_wide(options_.DestinationRoot));

        initialize_media_identity_expectations(source_root, destination_root);
        wait_for_media_availability(source_root, destination_root, cancel);
        ensure_media_identity_integrity(source_root, destination_root, !scan_only, true);
        if (scan_only && options_.UseExperimentalRawDiskScan) {
            // Raw-disk reads are not ported; the .NET worker uses the same
            // standard-file-read fallback when raw access is unavailable.
            emit_log("Scan backend: Raw disk unavailable (not ported in native build); using standard file reads.");
        }
        if (!scan_only) {
            storage::fsutil::create_directories(destination_root);
        }

        initialize_bad_range_map(source_root);

        emit_log("Scanning source: " + wide_to_utf8(source_root));
        auto log_fn = [this](const std::string& text) { emit_log(text); };
        SourceScanResult scan = scan_source(source_root, options_.SelectedRelativePaths,
                                            options_.SymlinkHandling, options_.CopyEmptyDirectories, log_fn,
                                            [&cancel] { return cancel.is_cancelled(); });
        std::int64_t total_bytes = 0;
        for (const auto& file : scan.files) total_bytes += file.length;
        emit_log("Scan complete. " + std::to_string(scan.files.size()) + " file(s), " +
                 format_bytes(total_bytes) + " total.");
        if (scan_only) {
            emit_log("Operation mode: Scan only (read-only bad-block detection).");
            emit_log("Scan performance profile: " +
                     std::string(models::to_string(resolve_scan_performance_profile())) +
                     "; workers=" +
                     std::to_string(resolve_parallel_scan_workers(
                         static_cast<std::int32_t>(scan.files.size()))) + ".");
        } else {
            emit_destination_capacity_warning(destination_root, total_bytes);
            emit_log("Transfer engine policy: " + std::string(models::to_string(options_.TransferEnginePolicyValue)) + ".");
            emit_log("Small-file acceleration: workers=" + std::to_string(options_.ParallelSmallFileWorkers) +
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

        JobJournal journal = merge_journal(existing, scan.files,
                                           wide_to_utf8(source_root), wide_to_utf8(destination_root), job_id);
        // Priming save: also measures what a save of this journal costs, so the
        // very first throttled flush is already priced correctly.
        save_journal_now(journal_path, journal);

        ProgressAccumulator progress;
        progress.total_files = static_cast<std::int32_t>(scan.files.size());
        progress.total_bytes = total_bytes;

        // Fast scan profile: parallel healthy-file reads with precise fallback.
        if (scan_only &&
            resolve_scan_performance_profile() == models::ScanPerformanceProfile::Fast) {
            std::string fast_error = run_fast_scan_mode(scan.files, source_root, progress, journal,
                                                        journal_path, cancel);
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
        if (!scan_only && is_native_acceleration_allowed() && options_.ParallelSmallFileWorkers > 1 &&
            !options_.FragileMediaMode && !options_.WaitForFileLockRelease &&
            options_.MaxThroughputBytesPerSecond <= 0 &&
            resolve_verification_mode() == models::VerificationMode::None &&
            !(options_.UseBadRangeMap && options_.SkipKnownBadRanges)) {

            std::vector<SourceFileDescriptor> eligible;
            std::int64_t threshold = std::max(MinimumRescueBlockSize, options_.SmallFileThresholdBytes);
            for (const auto& candidate : scan.files) {
                if (candidate.length <= 0 || candidate.length > threshold) continue;
                JournalFileEntry* candidate_entry = journal.Files.find(candidate.relative_path);
                if (candidate_entry == nullptr) continue;
                if (is_already_completed(*candidate_entry, candidate, destination_root)) continue;
                if (should_skip_failed_entry_for_fragile_resume(*candidate_entry)) continue;
                std::string parallel_skip_reason;
                if (should_skip_by_overwrite_policy(candidate, destination_root, parallel_skip_reason)) {
                    continue;
                }
                eligible.push_back(candidate);
            }
            if (eligible.size() > 1) {
                parallel_completed = copy_small_files_parallel(eligible, destination_root, journal,
                                                               journal_path, cancel);
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
                std::string skip_reason;
                if (should_skip_by_overwrite_policy(descriptor, destination_root, skip_reason)) {
                    progress.skipped_files += 1;
                    progress.total_bytes_copied += descriptor.length;
                    emit_log("Skipped: " + descriptor.relative_path + " (" + skip_reason + ")");
                    emit_progress(progress, descriptor.relative_path, descriptor.length,
                                  descriptor.length, 0,
                                  resolve_buffer_size_for_file(descriptor.length));
                    continue;
                }

                if (is_already_completed(*entry, descriptor, destination_root)) {
                    if (take_parallel_completed(descriptor.relative_path)) {
                        progress.completed_files += 1;
                    } else {
                        progress.skipped_files += 1;
                    }
                    progress.total_bytes_copied += descriptor.length;
                    if (entry->State == FileCopyState::CompletedWithRecovery) {
                        progress.recovered_files += 1;
                    }
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
                if (should_skip_failed_entry_for_fragile_resume(*entry)) {
                    progress.skipped_files += 1;
                    std::int64_t scan_already = entry->BytesCopied > 0 ? entry->BytesCopied : 0;
                    std::int64_t scan_remaining = descriptor.length - scan_already;
                    progress.total_bytes_copied += scan_remaining > 0 ? scan_remaining : 0;
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

                std::string scan_error = process_precise_scan_file(
                    descriptor, *entry, progress, journal, journal_path,
                    /*preserve_existing_coverage*/ false, /*seed_progress*/ true, cancel);
                if (!scan_error.empty() && !options_.ContinueOnFileError) {
                    save_journal_now(journal_path, journal);
                    return create_result(progress, journal_path, false, false, scan_error);
                }
                continue;
            }

            if (should_skip_failed_entry_for_fragile_resume(*entry)) {
                progress.skipped_files += 1;
                std::int64_t remaining = descriptor.length - (entry->BytesCopied > 0 ? entry->BytesCopied : 0);
                progress.total_bytes_copied += remaining > 0 ? remaining : 0;
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
            bool fragile_skip = false;
            std::string fragile_message;

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
                entry->State = FileCopyState::Pending;
                entry->LastError = ex.what();
                entry->DoNotRetry = false;
                emit_log("Skipped: " + descriptor.relative_path + " (" + ex.what() + ")");
                flush_journal(journal_path, journal, true);
                continue;
            } catch (const FragileReadSkip& ex) {
                fragile_skip = true;
                fragile_message = ex.what();
            } catch (const OperationCanceled& oc) {
                if (oc.user_requested) throw;
                std::int64_t timeout_seconds =
                    (options_.PerFileTimeout.ticks / time::TicksPerSecond);
                copy_error = "Per-file timeout (" + std::to_string(timeout_seconds) +
                             " sec) reached while copying " + descriptor.relative_path + ".";
            } catch (const std::exception& ex) {
                copy_error = ex.what();
            }

            if (fragile_skip) {
                handle_fragile_read_skip(descriptor, *entry, progress, journal, journal_path,
                                         fragile_message, cancel);
                continue;
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

        flush_bad_range_map();
        save_journal_now(journal_path, journal);
        CopyJobResult final_result = create_result(progress, journal_path,
                                                   progress.failed_files == 0, false, std::string());
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
    storage::JobJournalStore journal_store_;
    storage::BadRangeMapStore bad_range_map_store_;

    ULONGLONG run_started_tick_ = 0;
    ULONGLONG last_journal_flush_tick_ = 0;
    ULONGLONG last_journal_save_cost_ms_ = 0;
    ULONGLONG throttle_window_start_tick_ = 0;
    std::int64_t throttle_window_bytes_ = 0;
    std::string expected_source_identity_;
    std::string expected_destination_identity_;
    std::atomic<ULONGLONG> last_source_identity_probe_tick_{0};
    std::atomic<ULONGLONG> last_destination_identity_probe_tick_{0};
    std::atomic<ULONGLONG> last_source_mismatch_log_tick_{0};
    std::atomic<ULONGLONG> last_destination_mismatch_log_tick_{0};
    std::vector<std::wstring> created_directories_;
    std::deque<ULONGLONG> fragile_failure_timestamps_;
    std::atomic<std::int32_t> native_fast_path_files_{0};
    std::atomic<std::int32_t> parallel_native_fast_path_files_{0};
    std::atomic<std::int32_t> managed_copy_files_{0};
    std::atomic<std::int32_t> native_fallback_files_{0};
    std::mt19937 retry_jitter_{0xC0FFEE};

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

        if (options_.BufferSizeBytes < 4096) fail("BufferSizeBytes must be at least 4096.");
        if (options_.MaxRetries < 0) fail("MaxRetries cannot be negative.");
        if (options_.OperationTimeout.ticks <= 0) fail("OperationTimeout must be greater than zero.");
        if (options_.PerFileTimeout.ticks < 0) fail("PerFileTimeout cannot be negative.");
        if (options_.LockContentionProbeInterval.ticks <= 0) fail("LockContentionProbeInterval must be greater than zero.");
        if (options_.MaxThroughputBytesPerSecond < 0) fail("MaxThroughputBytesPerSecond cannot be negative.");
        if (options_.SampleVerificationChunkBytes <= 0) fail("SampleVerificationChunkBytes must be greater than zero.");
        if (options_.SampleVerificationChunkCount <= 0) fail("SampleVerificationChunkCount must be greater than zero.");
        if (options_.BadRangeMapMaxAgeDays < 0) fail("BadRangeMapMaxAgeDays cannot be negative.");
        if (options_.RescueFastScanChunkBytes < 0 || options_.RescueTrimChunkBytes < 0 ||
            options_.RescueScrapeChunkBytes < 0 || options_.RescueRetryChunkBytes < 0 ||
            options_.RescueSplitMinimumBytes < 0) {
            fail("Rescue chunk/split tuning values cannot be negative.");
        }
        if (options_.RescueFastScanRetries < 0 || options_.RescueTrimRetries < 0 ||
            options_.RescueScrapeRetries < 0) {
            fail("Rescue pass retries cannot be negative.");
        }
        if (options_.ParallelSmallFileWorkers < 1) fail("ParallelSmallFileWorkers must be at least 1.");
        if (options_.ParallelScanWorkers < 0) fail("ParallelScanWorkers cannot be negative.");
        if (options_.SmallFileThresholdBytes < MinimumRescueBlockSize) fail("SmallFileThresholdBytes must be at least 4096.");
        if (options_.FragileFailureWindowSeconds <= 0) fail("FragileFailureWindowSeconds must be greater than zero.");
        if (options_.FragileFailureThreshold <= 0) fail("FragileFailureThreshold must be greater than zero.");
        if (options_.FragileCooldownSeconds < 0) fail("FragileCooldownSeconds cannot be negative.");
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
                             const std::string& job_id) {
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

        for (const auto& descriptor : source_files) {
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
            entry->SourceLength = descriptor.length;
            entry->SourceLastWriteUtcTicks = descriptor.last_write_utc_ticks;

            if (source_changed) {
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
        }

        // Drop journal entries for files no longer present in the source. Use a
        // folded-key set for O(n) membership instead of an O(n*m) nested scan
        // (a whole-drive journal has 100k+ entries and files).
        std::unordered_set<std::string> source_keys;
        source_keys.reserve(source_files.size());
        for (const auto& descriptor : source_files) {
            source_keys.insert(fold_relative_path(descriptor.relative_path));
        }
        std::vector<std::pair<std::string, JournalFileEntry>> retained;
        retained.reserve(journal.Files.size());
        for (auto& [key, entry] : journal.Files.entries) {
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
        journal_store_.save(journal_path, journal);
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

    bool is_already_completed(const JournalFileEntry& entry, const SourceFileDescriptor& descriptor,
                              const std::wstring& destination_root) const {
        if (entry.State != FileCopyState::Completed &&
            entry.State != FileCopyState::CompletedWithRecovery) {
            return false;
        }
        ExistingFileMetadata metadata;
        if (!try_get_existing_file_metadata(destination_path_for(destination_root, descriptor), metadata)) {
            return false;
        }
        return metadata.length == descriptor.length;
    }

    bool should_skip_failed_entry_for_fragile_resume(const JournalFileEntry& entry) const {
        return entry.State == FileCopyState::Failed && entry.DoNotRetry &&
               (options_.PersistFragileSkipAcrossResume || options_.FragileMediaMode);
    }

    bool should_skip_by_overwrite_policy(const SourceFileDescriptor& descriptor,
                                         const std::wstring& destination_root,
                                         std::string& reason) const {
        reason.clear();
        ExistingFileMetadata metadata;
        if (!try_get_existing_file_metadata(destination_path_for(destination_root, descriptor), metadata)) {
            return false;
        }
        switch (options_.OverwritePolicyValue) {
            case models::OverwritePolicy::SkipExisting:
                reason = "destination already exists";
                return true;
            case models::OverwritePolicy::OverwriteIfSourceNewer:
                if (descriptor.last_write_utc_ticks <= metadata.last_write_utc_ticks) {
                    reason = "destination is newer or same age";
                    return true;
                }
                return false;
            default:
                return false;
        }
    }

    static std::wstring destination_path_for(const std::wstring& destination_root,
                                             const SourceFileDescriptor& descriptor) {
        return detail::trim_trailing_separators(destination_root) + L"\\" +
               utf8_to_wide(descriptor.relative_path);
    }

    void handle_fragile_read_skip(const SourceFileDescriptor& descriptor, JournalFileEntry& entry,
                                  ProgressAccumulator& progress, JobJournal& journal,
                                  const std::wstring& journal_path, const std::string& message,
                                  const CancelContext& cancel,
                                  const std::string& operation_label = "copy") {
        std::int64_t already = entry.BytesCopied > 0 ? entry.BytesCopied : 0;
        std::int64_t remaining = descriptor.length - already;
        progress.skipped_files += 1;
        progress.total_bytes_copied += remaining > 0 ? remaining : 0;

        entry.State = FileCopyState::Failed;
        entry.LastError = message;
        entry.DoNotRetry = options_.PersistFragileSkipAcrossResume;

        emit_log("Skipped " + operation_label + ": " + descriptor.relative_path + " (" + message + ")");
        emit_progress(progress, descriptor.relative_path, descriptor.length, descriptor.length, 0,
                      resolve_buffer_size_for_file(descriptor.length));
        register_fragile_failure_and_maybe_cooldown(descriptor.relative_path, message, cancel);
        try_persist_bad_range_map_entry(descriptor, entry, false);
        flush_journal(journal_path, journal, true);
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

    bool copy_single_file(const SourceFileDescriptor& descriptor, JournalFileEntry& entry,
                          const std::wstring& destination_root, ProgressAccumulator& progress,
                          JobJournal& journal, const std::wstring& journal_path,
                          const CancelContext& cancel) {
        std::wstring destination_path = destination_path_for(destination_root, descriptor);
        std::wstring destination_directory = storage::fsutil::get_directory_name(destination_path);
        if (!destination_directory.empty() && add_created_directory(destination_directory)) {
            storage::fsutil::create_directories(destination_directory);
        }

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

        prepare_destination_file(destination_path, descriptor.length, destination_length, preserve_existing);
        destination_length = get_existing_file_length(destination_path);

        // Native CopyFileEx fast path.
        std::int64_t native_baseline = progress.total_bytes_copied;
        NativeAttempt native = try_copy_with_native_fast_path(
            descriptor, entry, destination_path, destination_length, has_persisted_coverage,
            progress, journal, journal_path, cancel);
        if (native.attempted) {
            if (native.succeeded) {
                if (options_.PreserveTimestamps) {
                    set_last_write_time_utc(destination_path, descriptor.last_write_utc_ticks);
                }
                if (resolve_verification_mode() != models::VerificationMode::None) {
                    verify_file_with_retries(descriptor.full_path, destination_path,
                                             descriptor.relative_path, resolve_verification_mode(), cancel);
                }
                return false;
            }
            progress.total_bytes_copied = native_baseline;
            entry.BytesCopied = 0;
            entry.LastRescuePass = "Init";
            if (!native.fallback_reason.empty()) {
                emit_log("Native fast-path fallback on " + descriptor.relative_path + ": " +
                         native.fallback_reason);
            }
            prepare_destination_file(destination_path, descriptor.length,
                                     get_existing_file_length(destination_path), false);
            destination_length = get_existing_file_length(destination_path);
        }

        AdaptiveBufferController controller =
            create_buffer_controller(descriptor.length, BufferPurpose::Copy);
        const std::int32_t active_buffer_size = controller.current_size();
        std::vector<unsigned char> io_buffer(
            static_cast<std::size_t>(std::max(controller.maximum_size(), MinimumRescueBlockSize)));

        FileTransferSession session(descriptor.full_path, destination_path);
        entry.RescueRanges = build_rescue_ranges(entry, descriptor.length, destination_length);
        std::int64_t mapped_unreadable = apply_known_bad_ranges_from_map(descriptor, entry);
        entry.LastRescuePass = "Init";

        std::int64_t already_satisfied = rescue_bytes(
            entry.RescueRanges, {RescueRangeState::Good, RescueRangeState::Recovered});
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
            recovered_any = copy_small_file_fast(descriptor, entry, destination_path, session,
                                                 io_buffer, controller, progress, journal,
                                                 journal_path, cancel);
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
                    &controller);

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
                sync_entry_bytes_copied(entry);
                entry.RecoveredRanges = merge_byte_ranges(entry.RecoveredRanges);
                flush_journal(journal_path, journal, true);
            }
        }

        if (options_.PreserveTimestamps) {
            session.invalidate_destination(); // release write handle before touching times
            set_last_write_time_utc(destination_path, descriptor.last_write_utc_ticks);
        }

        models::VerificationMode verification = resolve_verification_mode();
        if (verification != models::VerificationMode::None) {
            if (recovered_any) {
                emit_log("Verification skipped for recovered file: " + descriptor.relative_path);
            } else {
                session.invalidate_source();
                session.invalidate_destination();
                verify_file_with_retries(descriptor.full_path, destination_path,
                                         descriptor.relative_path, verification, cancel);
            }
        }

        managed_copy_files_ += 1;
        return recovered_any;
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
        return true;
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

        DWORD flags = COPY_FILE_ALLOW_DECRYPTED_DESTINATION;
        if (descriptor.length >= 16LL * 1024 * 1024) flags |= COPY_FILE_NO_BUFFERING;

        BOOL cancel_flag = FALSE;
        BOOL ok = CopyFileExW(descriptor.full_path.c_str(), destination_path.c_str(),
                              native_progress_routine, &state, &cancel_flag, flags);
        DWORD last_error = ok ? 0 : GetLastError();

        if (ok) {
            result.succeeded = true;
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
                                          AdaptiveBufferController* controller = nullptr) {
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
                              const CancelContext& cancel) {
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
                                         std::int32_t max_retries, bool allow_salvage) {
        std::int32_t attempt = 0;
        std::optional<IoError> last_error;
        std::int32_t effective_max = max_retries >= 0 ? max_retries : options_.MaxRetries;

        while (attempt <= effective_max) {
            cancel.throw_if_cancelled();
            throw_if_media_identity_mismatch_throttled(source_path, expected_source_identity_, true);

            try {
                if (fault_injector_.has_value()) {
                    if (auto fault = fault_injector_->create_read_fault(relative_path, offset, length)) {
                        throw *fault;
                    }
                }
                std::int32_t bytes_read = timed_io::read_at(
                    session.source_handle(), offset, buffer, length, operation_timeout_ms(), cancel);
                if (bytes_read != length) {
                    throw IoError("Short read at offset " + std::to_string(offset) + ". Expected " +
                                  std::to_string(length) + ", got " + std::to_string(bytes_read) + ".");
                }
                return bytes_read;
            } catch (const IoError& ex) {
                if (ex.kind == IoErrorKind::Timeout) {
                    last_error = IoError("Read timeout at offset " + std::to_string(offset) + ".",
                                         0, IoErrorKind::Timeout);
                } else {
                    last_error = ex;
                }
            }

            if (options_.FragileMediaMode && options_.SkipFileOnFirstReadError) {
                session.invalidate_source();
                std::string error_text = last_error.has_value() ? last_error->what() : "unknown read error";
                throw FragileReadSkip("Fragile mode: first read failure at " + format_bytes(offset) +
                                      " (" + error_text + ").");
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

            if (is_fatal_read(*last_error)) break;

            if (options_.WaitForMediaAvailability && is_availability_related(*last_error, false)) {
                session.invalidate_source();
                emit_log("Source unavailable during read of " + relative_path +
                         ". Waiting for media to return.");
                wait_for_source_file(source_path, cancel, false);
                attempt = 0;
                continue;
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
            throw_if_media_identity_mismatch_throttled(destination_root, expected_destination_identity_, false);

            try {
                if (fault_injector_.has_value()) {
                    if (auto fault = fault_injector_->create_write_fault(relative_path, offset, count)) {
                        throw *fault;
                    }
                }
                timed_io::write_at(session.destination_handle(), offset, buffer, count,
                                   operation_timeout_ms(), cancel);
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
        double jitter_unit = (static_cast<double>(retry_jitter_() % 10000) / 5000.0) - 1.0;
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
                                  models::VerificationMode mode, const CancelContext& cancel) {
        if (mode == models::VerificationMode::Sampled) {
            emit_log("Verifying (sampled): " + relative_path);
            verify_sampled(source_path, destination_path, relative_path, cancel);
            return;
        }
        if (mode == models::VerificationMode::Full) {
            emit_log("Verifying (full hash): " + relative_path);
            auto source_hash = compute_file_hash(source_path, cancel);
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
                                    FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE, nullptr,
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
        return crypto::hash_buffer(buffer.data(), static_cast<std::size_t>(count), use_sha512());
    }

    std::vector<unsigned char> compute_file_hash(const std::wstring& path, const CancelContext& cancel) {
        HANDLE handle = CreateFileW(path.c_str(), GENERIC_READ,
                                    FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE, nullptr,
                                    OPEN_EXISTING,
                                    FILE_ATTRIBUTE_NORMAL | FILE_FLAG_OVERLAPPED | FILE_FLAG_SEQUENTIAL_SCAN,
                                    nullptr);
        if (handle == INVALID_HANDLE_VALUE) {
            throw IoError::from_win32("Unable to open file for hashing.", GetLastError());
        }

        // Dynamic timeout like EstimateHashTimeout: assume >= 8 MB/s + 10 s pad.
        std::int64_t length = get_existing_file_length(path);
        std::int64_t estimated_seconds = length / (8LL * 1024 * 1024) + 1;
        std::int64_t timeout_seconds = std::clamp<std::int64_t>(estimated_seconds + 10, 30, 3600);
        CancelContext hash_cancel = cancel.with_deadline(timeout_seconds * 1000);

        crypto::detail::StreamingHash hasher(use_sha512());
        std::vector<unsigned char> buffer(1 << 20);
        std::int64_t offset = 0;
        try {
            while (true) {
                std::int32_t read = timed_io::read_at(handle, offset, buffer.data(),
                                                      static_cast<std::int32_t>(buffer.size()),
                                                      timeout_seconds * 1000, hash_cancel);
                if (read <= 0) break;
                hasher.update(buffer.data(), static_cast<std::size_t>(read));
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

    static std::string build_file_fingerprint(const SourceFileDescriptor& descriptor) {
        char text[40];
        std::snprintf(text, sizeof(text), "%016llX:%016llX",
                      static_cast<unsigned long long>(descriptor.length),
                      static_cast<unsigned long long>(descriptor.last_write_utc_ticks));
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

        std::optional<storage::BadRangeMap> loaded;
        try {
            loaded = bad_range_map_store_.load(bad_range_map_path_);
        } catch (const std::exception& ex) {
            emit_log("Bad-range map load failed, starting fresh: " + std::string(ex.what()));
            loaded.reset();
        }

        bool allow_read_hints = false;
        if (loaded.has_value()) {
            std::wstring map_source = normalize_root_path(utf8_to_wide(loaded->SourceRoot));
            if (!map_source.empty() &&
                storage::fsutil::to_upper_invariant(map_source) !=
                    storage::fsutil::to_upper_invariant(normalized_root)) {
                emit_log("Bad-range map source root mismatch; ignoring existing map payload.");
                loaded.reset();
            }
        }

        if (!loaded.has_value()) {
            storage::BadRangeMap fresh;
            fresh.SchemaVersion = 1;
            fresh.SourceRoot = wide_to_utf8(normalized_root);
            fresh.SourceIdentity = expected_source_identity_;
            fresh.UpdatedUtc = time::DateTimeOffset::now_utc();
            loaded = std::move(fresh);
            emit_log("Bad-range map: initialized new map for this source.");
        } else {
            loaded->SourceRoot = wide_to_utf8(normalized_root);
            if (loaded->SchemaVersion <= 0) loaded->SchemaVersion = 1;
            if (!expected_source_identity_.empty()) loaded->SourceIdentity = expected_source_identity_;
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
        std::string fingerprint = trim_ascii(map_entry.FileFingerprint);
        if (fingerprint.empty()) return true;
        return models::detail::equals_ignore_case(fingerprint, build_file_fingerprint(descriptor));
    }

    const storage::BadRangeMapFileEntry* try_get_bad_range_map_entry(
        const SourceFileDescriptor& descriptor) {
        std::lock_guard<std::mutex> guard(map_lock_);
        if (!bad_range_map_.has_value()) return nullptr;

        std::string relative = detail::normalize_relative_path(descriptor.relative_path);
        if (relative.empty()) return nullptr;

        const storage::BadRangeMapFileEntry* candidate = bad_range_map_->Files.find(relative);
        if (candidate == nullptr || candidate->BadRanges.empty()) return nullptr;
        if (!is_bad_range_map_entry_compatible(descriptor, *candidate)) return nullptr;
        return candidate;
    }

    // Marks map-known bad ranges as KnownBad in the entry's rescue ranges so
    // rescue passes skip re-reading them. Returns total mapped bytes.
    std::int64_t apply_known_bad_ranges_from_map(const SourceFileDescriptor& descriptor,
                                                 JournalFileEntry& entry) {
        if (!bad_range_map_read_hints_enabled_) return 0;
        const storage::BadRangeMapFileEntry* map_entry = try_get_bad_range_map_entry(descriptor);
        if (map_entry == nullptr) return 0;

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
        if (!options_.UpdateBadRangeMapFromRun &&
            options_.OperationMode != models::JobOperationMode::ScanOnly) {
            return;
        }

        std::lock_guard<std::mutex> guard(map_lock_);
        if (!bad_range_map_.has_value()) return;

        std::string relative = detail::normalize_relative_path(descriptor.relative_path);
        if (relative.empty()) return;

        std::vector<ByteRange> unreadable = merge_byte_ranges(snapshot_by_state(
            entry.RescueRanges,
            {RescueRangeState::Bad, RescueRangeState::KnownBad, RescueRangeState::Recovered}));

        if (unreadable.empty()) {
            bad_range_map_->Files.remove(relative);
        } else {
            storage::BadRangeMapFileEntry map_entry;
            map_entry.RelativePath = relative;
            map_entry.SourceLength = descriptor.length;
            map_entry.LastWriteUtcTicks = descriptor.last_write_utc_ticks;
            map_entry.FileFingerprint = build_file_fingerprint(descriptor);
            map_entry.BadRanges = std::move(unreadable);
            map_entry.LastScanUtc = time::DateTimeOffset::now_utc();
            map_entry.LastError = entry.LastError;
            bad_range_map_->Files.set(relative, std::move(map_entry));
        }

        bad_range_map_->SourceRoot =
            wide_to_utf8(normalize_root_path(utf8_to_wide(options_.SourceRoot)));
        bad_range_map_->SourceIdentity = expected_source_identity_;
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
            bad_range_map_store_.save(bad_range_map_path_, *bad_range_map_);
            last_bad_range_map_flush_tick_ = GetTickCount64();
        } catch (const std::exception& ex) {
            emit_log("Bad-range map save failed: " + std::string(ex.what()));
        }
    }

    void flush_bad_range_map() {
        if (!bad_range_map_loaded_ || bad_range_map_path_.empty()) return;
        if (!options_.UpdateBadRangeMapFromRun &&
            options_.OperationMode != models::JobOperationMode::ScanOnly) {
            return;
        }
        std::lock_guard<std::mutex> guard(map_lock_);
        if (!bad_range_map_.has_value()) return;
        try {
            bad_range_map_->UpdatedUtc = time::DateTimeOffset::now_utc();
            bad_range_map_store_.save(bad_range_map_path_, *bad_range_map_);
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

    std::int32_t resolve_parallel_scan_workers(std::int32_t file_count) const {
        if (options_.OperationMode != models::JobOperationMode::ScanOnly || file_count <= 1) return 1;
        if (options_.FragileMediaMode) return 1;
        if (options_.ParallelScanWorkers > 0) {
            return std::max(1, std::min(64, std::min(file_count, options_.ParallelScanWorkers)));
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
        return entry.SourceLength == descriptor.length &&
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

    // Shared coordination state for the fast-scan worker pool.
    struct FastScanShared {
        std::mutex queue_lock;
        std::vector<SourceFileDescriptor> queue;
        std::size_t next_index = 0;
        std::vector<SourceFileDescriptor> fallback;
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
        if (descriptor.length <= 0) {
            outcome.succeeded = true;
            return outcome;
        }

        std::vector<unsigned char> io_buffer(
            static_cast<std::size_t>(std::max(controller.maximum_size(), MinimumRescueBlockSize)));
        FileTransferSession session(descriptor.full_path, std::wstring());

        try {
            std::int64_t offset = 0;
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
                entry.BytesCopied = offset;
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
            outcome.bytes_read = std::max<std::int64_t>(0, entry.BytesCopied);
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
                        progress.total_bytes_copied += remaining > 0 ? remaining : 0;
                        emit_progress(progress, descriptor.relative_path, descriptor.length,
                                      descriptor.length, 0,
                                      resolve_buffer_size_for_file(descriptor.length),
                                      "FastHealthScanSkipped", 0, 0, 0, shared.worker_count);
                        continue;
                    }

                    if (options_.UseBadRangeMap && options_.SkipKnownBadRanges &&
                        try_get_bad_range_map_entry(descriptor) != nullptr) {
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
                } catch (const SourceMutationSkipped&) {
                    std::lock_guard<std::mutex> guard(progress_lock_);
                    progress.skipped_files += 1;
                    progress.total_bytes_copied += descriptor.length;
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

        std::vector<std::thread> workers;
        for (std::int32_t i = 0; i < shared.worker_count; ++i) {
            workers.emplace_back([this, &shared, &progress, &journal, &cancel]() {
                scan_fast_worker(shared, progress, journal, cancel);
            });
        }
        for (auto& worker : workers) worker.join();

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
        bool fragile_skip = false;
        std::string fragile_message;

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
            entry.State = FileCopyState::Pending;
            entry.LastError = ex.what();
            entry.DoNotRetry = false;
            emit_log("Skipped scan: " + descriptor.relative_path + " (" + ex.what() + ")");
            flush_journal(journal_path, journal, true);
            return std::string();
        } catch (const FragileReadSkip& ex) {
            fragile_skip = true;
            fragile_message = ex.what();
        } catch (const OperationCanceled& oc) {
            if (oc.user_requested) throw;
            std::int64_t timeout_seconds = options_.PerFileTimeout.ticks / time::TicksPerSecond;
            scan_error = "Per-file timeout (" + std::to_string(timeout_seconds) +
                         " sec) reached while scanning " + descriptor.relative_path + ".";
        } catch (const std::exception& ex) {
            scan_error = ex.what();
        }

        if (fragile_skip) {
            handle_fragile_read_skip(descriptor, entry, progress, journal, journal_path,
                                     fragile_message, cancel, "scan");
            return std::string();
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
        JobJournal& journal, const std::wstring& journal_path, const CancelContext& cancel) {
        std::vector<std::string> completed_paths;
        if (files.empty()) return completed_paths;

        emit_log("Parallel small file phase: " + std::to_string(files.size()) + " file(s) using " +
                 std::to_string(options_.ParallelSmallFileWorkers) + " worker(s).");

        for (const auto& descriptor : files) {
            std::wstring destination_path = destination_path_for(destination_root, descriptor);
            std::wstring destination_directory = storage::fsutil::get_directory_name(destination_path);
            if (!destination_directory.empty() && add_created_directory(destination_directory)) {
                storage::fsutil::create_directories(destination_directory);
            }
        }

        std::int32_t concurrency = std::max(
            1, std::min<std::int32_t>(options_.ParallelSmallFileWorkers,
                                      static_cast<std::int32_t>(files.size())));
        std::mutex work_lock;
        std::size_t next_index = 0;
        std::atomic<std::int32_t> completed_count{0};
        std::atomic<std::int32_t> failed_count{0};
        std::atomic<bool> cancelled{false};

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

                    std::wstring destination_path = destination_path_for(destination_root, descriptor);
                    HANDLE reset = CreateFileW(destination_path.c_str(), GENERIC_WRITE,
                                               FILE_SHARE_READ, nullptr, CREATE_ALWAYS,
                                               FILE_ATTRIBUTE_NORMAL, nullptr);
                    if (reset == INVALID_HANDLE_VALUE) {
                        failed_count += 1;
                        continue;
                    }
                    CloseHandle(reset);

                    DWORD flags = COPY_FILE_ALLOW_DECRYPTED_DESTINATION;
                    if (descriptor.length >= 16LL * 1024 * 1024) flags |= COPY_FILE_NO_BUFFERING;
                    QuietCallbackState state{&cancel, control_};
                    BOOL cancel_flag = FALSE;
                    BOOL ok = CopyFileExW(descriptor.full_path.c_str(), destination_path.c_str(),
                                          quiet_routine, &state, &cancel_flag, flags);
                    if (!ok) {
                        failed_count += 1;
                        continue;
                    }

                    if (options_.PreserveTimestamps) {
                        set_last_write_time_utc(destination_path, descriptor.last_write_utc_ticks);
                    }
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
                }
            });
        }
        for (auto& worker : workers) worker.join();

        if (cancelled.load() || cancel.is_cancelled()) throw OperationCanceled{true};

        emit_log("Parallel small file phase complete: " + std::to_string(completed_count.load()) +
                 " succeeded, " + std::to_string(failed_count.load()) +
                 " deferred to sequential path.");
        if (completed_count.load() > 0) {
            flush_journal(journal_path, journal, true);
        }
        return completed_paths;
    }

    // ---- Media identity + availability -----------------------------------

    void initialize_media_identity_expectations(const std::wstring& source_root,
                                                const std::wstring& destination_root) {
        expected_source_identity_ = trim_ascii(options_.ExpectedSourceIdentity);
        expected_destination_identity_ = trim_ascii(options_.ExpectedDestinationIdentity);

        if (expected_source_identity_.empty()) {
            std::string identity = resolve_media_identity(source_root);
            if (!identity.empty()) {
                expected_source_identity_ = identity;
                options_.ExpectedSourceIdentity = identity;
                emit_log("Source media identity baseline captured.");
            }
        }
        if (expected_destination_identity_.empty()) {
            std::string identity = resolve_media_identity(destination_root);
            if (!identity.empty()) {
                expected_destination_identity_ = identity;
                options_.ExpectedDestinationIdentity = identity;
                emit_log("Destination media identity baseline captured.");
            }
        }
    }

    static std::string trim_ascii(const std::string& text) {
        std::size_t begin = 0, end = text.size();
        auto is_space = [](char c) { return c == ' ' || c == '\t' || c == '\r' || c == '\n'; };
        while (begin < end && is_space(text[begin])) ++begin;
        while (end > begin && is_space(text[end - 1])) --end;
        return text.substr(begin, end - begin);
    }

    static std::string resolve_media_identity(const std::wstring& path_value) {
        if (path_value.empty()) return std::string();
        std::wstring full = storage::fsutil::get_full_path(path_value);
        if (full.size() >= 2 && full[0] == L'\\' && full[1] == L'\\') {
            // UNC \\server\share
            std::vector<std::wstring> parts;
            std::wstring current;
            for (wchar_t c : full.substr(2)) {
                if (c == L'\\' || c == L'/') {
                    if (!current.empty()) parts.push_back(current);
                    current.clear();
                    if (parts.size() >= 2) break;
                } else {
                    current.push_back(c);
                }
            }
            if (!current.empty() && parts.size() < 2) parts.push_back(current);
            if (parts.size() < 2) return std::string();
            std::wstring share = L"\\\\" + parts[0] + L"\\" + parts[1];
            return "unc:" + wide_to_utf8(storage::fsutil::to_upper_invariant(share));
        }

        if (full.size() < 2 || full[1] != L':') return std::string();
        std::wstring root = full.substr(0, 2) + L"\\";
        DWORD serial = 0, max_component = 0, flags = 0;
        if (!GetVolumeInformationW(root.c_str(), nullptr, 0, &serial, &max_component, &flags,
                                   nullptr, 0)) {
            return std::string();
        }
        char text[16];
        std::snprintf(text, sizeof(text), "vol:%08X", static_cast<unsigned int>(serial));
        return text;
    }

    static bool identities_equivalent(const std::string& expected, const std::string& current) {
        std::string e = trim_ascii(expected);
        std::string c = trim_ascii(current);
        if (e.empty() || c.empty()) return e.size() == c.size();
        if (models::detail::equals_ignore_case(e, c)) return true;
        // vol:XXXXXXXX serial equivalence.
        auto extract_serial = [](const std::string& identity, std::string& serial) {
            if (!detail::starts_with_ignore_case(identity, "vol:")) return false;
            std::size_t position = identity.find_last_of(':');
            serial = identity.substr(position + 1);
            return !serial.empty();
        };
        std::string es, cs;
        if (extract_serial(e, es) && extract_serial(c, cs)) {
            return models::detail::equals_ignore_case(es, cs);
        }
        return false;
    }

    bool should_probe_media_identity(bool is_source) {
        ULONGLONG now = GetTickCount64();
        std::atomic<ULONGLONG>& last = is_source ? last_source_identity_probe_tick_
                                                 : last_destination_identity_probe_tick_;
        ULONGLONG previous = last.load(std::memory_order_relaxed);
        if (previous == 0 || now - previous >= 1000) {
            last.store(now, std::memory_order_relaxed);
            return true;
        }
        return false;
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
                                                    const std::string& expected_identity,
                                                    bool is_source, bool force = false) {
        if (!force && !should_probe_media_identity(is_source)) return;
        throw_if_media_identity_mismatch(path_value, expected_identity, is_source);
    }

    void ensure_media_identity_integrity(const std::wstring& source_root,
                                         const std::wstring& destination_root,
                                         bool include_destination, bool force) {
        throw_if_media_identity_mismatch_throttled(source_root, expected_source_identity_, true, force);
        if (!include_destination) return;
        throw_if_media_identity_mismatch_throttled(destination_root, expected_destination_identity_,
                                                   false, force);
    }

    bool is_media_identity_accepted(const std::wstring& path_value, std::string& expected_identity,
                                    bool is_source) {
        expected_identity = trim_ascii(expected_identity);
        std::string current = resolve_media_identity(path_value);
        if (current.empty()) return expected_identity.empty();
        if (expected_identity.empty()) {
            expected_identity = current;
            if (is_source) {
                expected_source_identity_ = current;
                options_.ExpectedSourceIdentity = current;
            } else {
                expected_destination_identity_ = current;
                options_.ExpectedDestinationIdentity = current;
            }
            emit_log(std::string(is_source ? "Source" : "Destination") +
                     " media identity baseline captured.");
            return true;
        }
        if (identities_equivalent(expected_identity, current)) return true;

        ULONGLONG now = GetTickCount64();
        std::atomic<ULONGLONG>& last_log = is_source ? last_source_mismatch_log_tick_
                                                     : last_destination_mismatch_log_tick_;
        ULONGLONG previous_log = last_log.load(std::memory_order_relaxed);
        if (previous_log == 0 || now - previous_log >= 5000) {
            last_log.store(now, std::memory_order_relaxed);
            emit_log(std::string(is_source ? "Source" : "Destination") +
                     " media identity mismatch on '" + wide_to_utf8(path_value) + "'. Expected " +
                     expected_identity + ", found " + current + ".");
        }
        return false;
    }

    bool is_source_available(const std::wstring& source_root) {
        if (source_root.empty() || !directory_exists(source_root)) return false;
        return is_media_identity_accepted(source_root, expected_source_identity_, true);
    }

    bool is_destination_available(const std::wstring& target_path) {
        if (target_path.empty()) return false;
        std::wstring full = storage::fsutil::get_full_path(target_path);
        if (full.size() < 2) return false;
        std::wstring root = full.substr(0, 2) + L"\\";
        if (full[1] == L':' && !directory_exists(root)) return false;
        if (!is_media_identity_accepted(full, expected_destination_identity_, false)) return false;

        std::wstring probe = full;
        while (!probe.empty()) {
            if (storage::fsutil::file_exists(probe) || directory_exists(probe)) return true;
            std::wstring parent = storage::fsutil::get_directory_name(probe);
            if (parent == probe) break;
            probe = parent;
        }
        return false;
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
            bool identity_ok = is_media_identity_accepted(source_path, expected_source_identity_, true);
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
        snapshot.CurrentFile = current_file;
        snapshot.CurrentFileBytesCopied = safe_bytes;
        snapshot.CurrentFileBytesTotal = safe_total;
        snapshot.TotalBytesCopied = safe_copied;
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
        emit_log(std::string("Run summary: ") +
                 (result.Succeeded ? "succeeded" : "completed with failures") + "; copied " +
                 format_bytes(result.CopiedBytes) + " in " +
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
        double average = elapsed_ms > 0 ? static_cast<double>(progress.total_bytes_copied) /
                                              (static_cast<double>(elapsed_ms) / 1000.0)
                                        : 0.0;

        CopyJobResult result;
        result.Succeeded = succeeded;
        result.Cancelled = cancelled;
        result.TotalFiles = progress.total_files;
        result.CompletedFiles = progress.completed_files;
        result.FailedFiles = progress.failed_files;
        result.RecoveredFiles = progress.recovered_files;
        result.SkippedFiles = progress.skipped_files;
        result.TotalBytes = progress.total_bytes;
        result.CopiedBytes = progress.total_bytes_copied;
        result.TransferEnginePolicyValue = options_.TransferEnginePolicyValue;
        result.ElapsedMilliseconds = elapsed_ms;
        result.AverageBytesPerSecond = average;
        result.NativeFastPathFiles = native_fast_path_files_;
        result.ParallelNativeFastPathFiles = parallel_native_fast_path_files_;
        result.ManagedCopyFiles = managed_copy_files_;
        result.NativeFallbackFiles = native_fallback_files_;
        result.JournalPath = wide_to_utf8(journal_path);
        result.ErrorMessage = error_message;
        return result;
    }
};

} // namespace xact::engine
