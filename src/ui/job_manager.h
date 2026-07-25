// JobManagerService (port of XactCopy.UI Services\JobManagerService.vb):
// saved-job catalog, run queue, and run history over JobCatalogStore.
#pragma once

#include "../storage/job_catalog.h"

#include <algorithm>
#include <cstdio>
#include <mutex>
#include <optional>
#include <string>
#include <vector>

namespace xact::ui {

using time::DateTimeOffset;

enum class QueueMoveDirection : std::int32_t {
    Up = 0,
    Down = 1,
    Top = 2,
    Bottom = 3,
};

struct JobQueueEntryView {
    std::string QueueEntryId;
    std::string JobId;
    std::int32_t Position = 0;
    std::string JobName;
    std::string SourceRoot;
    std::string DestinationRoot;
    DateTimeOffset EnqueuedUtc = DateTimeOffset::min_value();
    DateTimeOffset LastUpdatedUtc = DateTimeOffset::min_value();
    std::string Trigger;
    std::string EnqueuedBy;
    std::int32_t AttemptCount = 0;
    std::optional<DateTimeOffset> LastAttemptUtc;
    std::string LastErrorMessage;
};

struct QueuedJobWorkItem {
    std::string QueueEntryId;
    std::string Trigger = "queued";
    std::string EnqueuedBy;
    DateTimeOffset EnqueuedUtc = DateTimeOffset::min_value();
    std::int32_t Attempt = 0;
    storage::ManagedJob Job;
};

class JobManagerService {
public:
    static constexpr std::int32_t MaximumRunHistory = 1000;

    explicit JobManagerService(std::wstring catalog_path = std::wstring())
        : store_(std::move(catalog_path)) {
        catalog_ = store_.load();
    }

    const storage::JobCatalogStore& store() const noexcept { return store_; }

    std::vector<storage::ManagedJob> get_jobs() {
        std::lock_guard<std::mutex> lock(mutex_);
        std::vector<storage::ManagedJob> jobs = catalog_.Jobs;
        std::sort(jobs.begin(), jobs.end(), [](const storage::ManagedJob& left, const storage::ManagedJob& right) {
            return compare_ordinal_ignore_case(left.Name, right.Name) < 0;
        });
        return jobs;
    }

    std::vector<storage::ManagedJobRun> get_recent_runs(std::int32_t take = 200) {
        std::lock_guard<std::mutex> lock(mutex_);
        std::size_t bounded = static_cast<std::size_t>(std::max(1, take));
        std::vector<storage::ManagedJobRun> runs;
        runs.reserve(std::min(bounded, catalog_.Runs.size()));
        for (const auto& run : catalog_.Runs) {
            if (runs.size() >= bounded) break;
            runs.push_back(run);
        }
        return runs;
    }

    std::vector<storage::ManagedJob> get_queued_jobs() {
        std::lock_guard<std::mutex> lock(mutex_);
        std::vector<storage::ManagedJob> queued;
        for (const auto& entry : build_sanitized_queue_locked()) {
            if (const auto* job = find_job_by_id_locked(entry.JobId)) queued.push_back(*job);
        }
        return queued;
    }

    std::vector<JobQueueEntryView> get_queue_entries() {
        std::lock_guard<std::mutex> lock(mutex_);
        std::vector<JobQueueEntryView> results;
        std::int32_t position = 1;
        for (const auto& entry : build_sanitized_queue_locked()) {
            const auto* job = find_job_by_id_locked(entry.JobId);
            if (job == nullptr) continue;
            JobQueueEntryView view;
            view.QueueEntryId = entry.QueueEntryId;
            view.JobId = job->JobId;
            view.Position = position;
            view.JobName = job->Name;
            view.SourceRoot = job->Options.SourceRoot;
            view.DestinationRoot = job->Options.DestinationRoot;
            view.EnqueuedUtc = entry.EnqueuedUtc;
            view.LastUpdatedUtc = entry.LastUpdatedUtc;
            view.Trigger = entry.Trigger;
            view.EnqueuedBy = entry.EnqueuedBy;
            view.AttemptCount = entry.AttemptCount;
            view.LastAttemptUtc = entry.LastAttemptUtc;
            view.LastErrorMessage = entry.LastErrorMessage;
            results.push_back(std::move(view));
            ++position;
        }
        return results;
    }

    std::optional<storage::ManagedJob> get_job_by_id(const std::string& job_id) {
        if (storage::catalog_detail::is_blank(job_id)) return std::nullopt;
        std::lock_guard<std::mutex> lock(mutex_);
        if (const auto* job = find_job_by_id_locked(job_id)) return *job;
        return std::nullopt;
    }

    std::optional<storage::ManagedJob> save_job(const std::string& name, const models::CopyJobOptions& options,
                                                const std::string& existing_job_id = std::string()) {
        if (storage::catalog_detail::is_blank(name)) return std::nullopt;

        std::lock_guard<std::mutex> lock(mutex_);
        DateTimeOffset now_utc = DateTimeOffset::now_utc();
        std::string trimmed_name = storage::catalog_detail::trim_copy(name);

        storage::ManagedJob* job = nullptr;
        if (!storage::catalog_detail::is_blank(existing_job_id)) {
            job = find_job_by_id_locked(existing_job_id);
        }

        if (job == nullptr) {
            storage::ManagedJob created;
            created.JobId = storage::detail::new_guid_n();
            created.CreatedUtc = now_utc;
            catalog_.Jobs.push_back(std::move(created));
            job = &catalog_.Jobs.back();
        }

        job->Name = trimmed_name;
        models::CopyJobOptions template_options = options;
        template_options.ExpectedSourceIdentity.clear();
        template_options.ExpectedDestinationIdentity.clear();
        template_options.ResumeJournalPathHint.clear();
        template_options.AllowJournalRootRemap = false;
        job->Options = std::move(template_options);
        job->UpdatedUtc = now_utc;

        save_catalog_locked();
        return *job;
    }

    bool rename_job(const std::string& job_id, const std::string& new_name) {
        if (storage::catalog_detail::is_blank(job_id) || storage::catalog_detail::is_blank(new_name)) return false;

        std::lock_guard<std::mutex> lock(mutex_);
        storage::ManagedJob* job = find_job_by_id_locked(job_id);
        if (job == nullptr) return false;

        job->Name = storage::catalog_detail::trim_copy(new_name);
        job->UpdatedUtc = DateTimeOffset::now_utc();
        save_catalog_locked();
        return true;
    }

    std::optional<storage::ManagedJob> duplicate_job(const std::string& job_id, const std::string& new_name) {
        if (storage::catalog_detail::is_blank(job_id) || storage::catalog_detail::is_blank(new_name)) return std::nullopt;

        std::lock_guard<std::mutex> lock(mutex_);
        const storage::ManagedJob* source_job = find_job_by_id_locked(job_id);
        if (source_job == nullptr) return std::nullopt;

        DateTimeOffset now_utc = DateTimeOffset::now_utc();
        storage::ManagedJob duplicate;
        duplicate.JobId = storage::detail::new_guid_n();
        duplicate.Name = storage::catalog_detail::trim_copy(new_name);
        duplicate.Options = source_job->Options;
        duplicate.CreatedUtc = now_utc;
        duplicate.UpdatedUtc = now_utc;

        catalog_.Jobs.push_back(duplicate);
        save_catalog_locked();
        return duplicate;
    }

    bool delete_job(const std::string& job_id) {
        if (storage::catalog_detail::is_blank(job_id)) return false;

        std::lock_guard<std::mutex> lock(mutex_);
        std::size_t before = catalog_.Jobs.size();
        std::erase_if(catalog_.Jobs, [&](const storage::ManagedJob& job) {
            return models::detail::equals_ignore_case(job.JobId, job_id);
        });
        if (catalog_.Jobs.size() == before) return false;

        std::erase_if(catalog_.QueueEntries, [&](const storage::ManagedJobQueueEntry& entry) {
            return models::detail::equals_ignore_case(entry.JobId, job_id);
        });
        std::erase_if(catalog_.QueuedJobIds, [&](const std::string& queued_id) {
            return models::detail::equals_ignore_case(queued_id, job_id);
        });

        save_catalog_locked();
        return true;
    }

    bool queue_job(const std::string& job_id, bool allow_duplicate = false,
                   const std::string& trigger = "manual", const std::string& enqueued_by = std::string()) {
        if (storage::catalog_detail::is_blank(job_id)) return false;

        std::lock_guard<std::mutex> lock(mutex_);
        const storage::ManagedJob* job = find_job_by_id_locked(job_id);
        if (job == nullptr) return false;

        build_sanitized_queue_locked();

        if (!allow_duplicate) {
            for (const auto& existing : catalog_.QueueEntries) {
                if (models::detail::equals_ignore_case(existing.JobId, job_id)) return false;
            }
        }

        DateTimeOffset now_utc = DateTimeOffset::now_utc();
        storage::ManagedJobQueueEntry entry;
        entry.QueueEntryId = storage::detail::new_guid_n();
        entry.JobId = job->JobId;
        entry.EnqueuedUtc = now_utc;
        entry.LastUpdatedUtc = now_utc;
        entry.Trigger = storage::catalog_detail::is_blank(trigger) ? std::string("manual")
                                                                   : storage::catalog_detail::trim_copy(trigger);
        entry.EnqueuedBy = enqueued_by;
        entry.AttemptCount = 0;
        entry.LastErrorMessage.clear();

        catalog_.QueueEntries.push_back(std::move(entry));
        save_catalog_locked();
        return true;
    }

    bool remove_queued_job(const std::string& queue_entry_id_or_job_id) {
        if (storage::catalog_detail::is_blank(queue_entry_id_or_job_id)) return false;

        std::lock_guard<std::mutex> lock(mutex_);
        std::string key = storage::catalog_detail::trim_copy(queue_entry_id_or_job_id);

        std::size_t before = catalog_.QueueEntries.size();
        std::erase_if(catalog_.QueueEntries, [&](const storage::ManagedJobQueueEntry& existing) {
            return models::detail::equals_ignore_case(existing.QueueEntryId, key);
        });
        bool removed = catalog_.QueueEntries.size() != before;

        if (!removed) {
            before = catalog_.QueueEntries.size();
            std::erase_if(catalog_.QueueEntries, [&](const storage::ManagedJobQueueEntry& existing) {
                return models::detail::equals_ignore_case(existing.JobId, key);
            });
            removed = catalog_.QueueEntries.size() != before;
        }

        if (removed) save_catalog_locked();
        return removed;
    }

    bool move_queue_entry(const std::string& queue_entry_id, QueueMoveDirection direction) {
        if (storage::catalog_detail::is_blank(queue_entry_id)) return false;

        std::lock_guard<std::mutex> lock(mutex_);
        build_sanitized_queue_locked();
        if (catalog_.QueueEntries.size() <= 1) return false;

        std::ptrdiff_t current_index = -1;
        for (std::size_t i = 0; i < catalog_.QueueEntries.size(); ++i) {
            if (models::detail::equals_ignore_case(catalog_.QueueEntries[i].QueueEntryId, queue_entry_id)) {
                current_index = static_cast<std::ptrdiff_t>(i);
                break;
            }
        }
        if (current_index < 0) return false;

        std::ptrdiff_t last = static_cast<std::ptrdiff_t>(catalog_.QueueEntries.size()) - 1;
        std::ptrdiff_t target_index = current_index;
        switch (direction) {
            case QueueMoveDirection::Up: target_index = std::max<std::ptrdiff_t>(0, current_index - 1); break;
            case QueueMoveDirection::Down: target_index = std::min(last, current_index + 1); break;
            case QueueMoveDirection::Top: target_index = 0; break;
            case QueueMoveDirection::Bottom: target_index = last; break;
        }
        if (target_index == current_index) return false;

        storage::ManagedJobQueueEntry entry = std::move(catalog_.QueueEntries[current_index]);
        catalog_.QueueEntries.erase(catalog_.QueueEntries.begin() + current_index);
        entry.LastUpdatedUtc = DateTimeOffset::now_utc();
        catalog_.QueueEntries.insert(catalog_.QueueEntries.begin() + target_index, std::move(entry));
        save_catalog_locked();
        return true;
    }

    std::int32_t clear_queue() {
        std::lock_guard<std::mutex> lock(mutex_);
        std::int32_t removed_count = static_cast<std::int32_t>(catalog_.QueueEntries.size());
        if (removed_count == 0) return 0;
        catalog_.QueueEntries.clear();
        save_catalog_locked();
        return removed_count;
    }

    bool try_dequeue_next_job(QueuedJobWorkItem& work_item) {
        return try_dequeue_internal(std::string(), work_item);
    }

    bool try_dequeue_queued_entry(const std::string& queue_entry_id, QueuedJobWorkItem& work_item) {
        if (storage::catalog_detail::is_blank(queue_entry_id)) return false;
        return try_dequeue_internal(storage::catalog_detail::trim_copy(queue_entry_id), work_item);
    }

    std::optional<storage::ManagedJobRun> create_run_for_job(const std::string& job_id, const std::string& trigger,
                                                             const std::string& queue_entry_id = std::string(),
                                                             std::int32_t queue_attempt = 0) {
        if (storage::catalog_detail::is_blank(job_id)) return std::nullopt;

        std::lock_guard<std::mutex> lock(mutex_);
        const storage::ManagedJob* job = find_job_by_id_locked(job_id);
        if (job == nullptr) return std::nullopt;

        storage::ManagedJobRun run = create_run_locked(
            job->JobId,
            storage::catalog_detail::is_blank(job->Name) ? std::string("Saved Job") : job->Name,
            job->Options, trigger, queue_entry_id, queue_attempt);
        save_catalog_locked();
        return run;
    }

    storage::ManagedJobRun create_ad_hoc_run(const models::CopyJobOptions& options,
                                             const std::string& display_name, const std::string& trigger) {
        std::lock_guard<std::mutex> lock(mutex_);
        std::string resolved_name = storage::catalog_detail::is_blank(display_name) ? std::string("Manual Copy")
                                                                                    : display_name;
        storage::ManagedJobRun run = create_run_locked(std::string(), resolved_name, options, trigger,
                                                       std::string(), 0);
        save_catalog_locked();
        return run;
    }

    void mark_run_running(const std::string& run_id, const std::string& journal_path) {
        if (storage::catalog_detail::is_blank(run_id)) return;

        std::lock_guard<std::mutex> lock(mutex_);
        storage::ManagedJobRun* run = find_run_by_id_locked(run_id);
        if (run == nullptr) return;

        run->Status = storage::ManagedJobRunStatus::Running;
        run->JournalPath = journal_path;
        run->LastUpdatedUtc = DateTimeOffset::now_utc();
        save_catalog_locked();
    }

    void mark_run_paused(const std::string& run_id) {
        update_run_status(run_id, storage::ManagedJobRunStatus::Paused);
    }

    void mark_run_resumed(const std::string& run_id) {
        update_run_status(run_id, storage::ManagedJobRunStatus::Running);
    }

    void mark_run_completed(const std::string& run_id, const std::optional<models::CopyJobResult>& result) {
        if (storage::catalog_detail::is_blank(run_id)) return;

        std::lock_guard<std::mutex> lock(mutex_);
        storage::ManagedJobRun* run = find_run_by_id_locked(run_id);
        if (run == nullptr) return;

        DateTimeOffset now_utc = DateTimeOffset::now_utc();
        run->Result = result;
        if (result.has_value()) run->JournalPath = result->JournalPath;
        run->ErrorMessage = result.has_value() ? result->ErrorMessage : std::string();
        run->LastUpdatedUtc = now_utc;
        run->FinishedUtc = now_utc;

        if (!result.has_value()) {
            run->Status = storage::ManagedJobRunStatus::Failed;
            run->Summary = "Run finished with unknown result.";
        } else if (result->Cancelled) {
            run->Status = storage::ManagedJobRunStatus::Cancelled;
            run->Summary = "Cancelled by user.";
        } else if (result->Succeeded) {
            run->Status = storage::ManagedJobRunStatus::Completed;
            run->Summary = "Completed: " + std::to_string(result->CompletedFiles) + "/" +
                           std::to_string(result->TotalFiles) + " files" + format_result_speed_suffix(*result) + ".";
        } else {
            run->Status = storage::ManagedJobRunStatus::Failed;
            run->Summary = "Completed with failures: failed " + std::to_string(result->FailedFiles) +
                           ", recovered " + std::to_string(result->RecoveredFiles) +
                           format_result_speed_suffix(*result) + ".";
        }

        save_catalog_locked();
    }

    bool mark_run_interrupted(const std::string& run_id, const std::string& reason) {
        if (storage::catalog_detail::is_blank(run_id)) return false;

        std::lock_guard<std::mutex> lock(mutex_);
        storage::ManagedJobRun* run = find_run_by_id_locked(run_id);
        if (run == nullptr) return false;

        switch (run->Status) {
            case storage::ManagedJobRunStatus::Running:
            case storage::ManagedJobRunStatus::Paused:
            case storage::ManagedJobRunStatus::Queued: {
                run->Status = storage::ManagedJobRunStatus::Interrupted;
                run->LastUpdatedUtc = DateTimeOffset::now_utc();
                run->FinishedUtc = DateTimeOffset::now_utc();
                run->ErrorMessage = storage::catalog_detail::is_blank(reason) ? std::string("Run interrupted.") : reason;
                run->Summary = run->ErrorMessage;
                save_catalog_locked();
                return true;
            }
            default:
                return false;
        }
    }

    std::int32_t mark_any_running_runs_interrupted(const std::string& reason) {
        std::lock_guard<std::mutex> lock(mutex_);
        std::int32_t count = 0;
        std::string message = storage::catalog_detail::is_blank(reason) ? std::string("Run interrupted.") : reason;
        for (auto& run : catalog_.Runs) {
            if (run.Status == storage::ManagedJobRunStatus::Running ||
                run.Status == storage::ManagedJobRunStatus::Paused) {
                run.Status = storage::ManagedJobRunStatus::Interrupted;
                run.LastUpdatedUtc = DateTimeOffset::now_utc();
                run.FinishedUtc = DateTimeOffset::now_utc();
                run.ErrorMessage = message;
                run.Summary = message;
                ++count;
            }
        }
        if (count > 0) save_catalog_locked();
        return count;
    }

    bool delete_run(const std::string& run_id) {
        if (storage::catalog_detail::is_blank(run_id)) return false;

        std::lock_guard<std::mutex> lock(mutex_);
        std::size_t before = catalog_.Runs.size();
        std::erase_if(catalog_.Runs, [&](const storage::ManagedJobRun& run) {
            return models::detail::equals_ignore_case(run.RunId, run_id);
        });
        bool removed = catalog_.Runs.size() != before;
        if (removed) save_catalog_locked();
        return removed;
    }

    std::int32_t clear_run_history(std::int32_t keep_latest = 0) {
        std::lock_guard<std::mutex> lock(mutex_);
        std::size_t safe_keep = static_cast<std::size_t>(std::max(0, keep_latest));
        if (catalog_.Runs.size() <= safe_keep) return 0;

        std::int32_t removed = static_cast<std::int32_t>(catalog_.Runs.size() - safe_keep);
        catalog_.Runs.resize(safe_keep);
        save_catalog_locked();
        return removed;
    }

    static std::string format_bytes(std::int64_t value) {
        if (value < 1024) return std::to_string(value) + " B";
        static constexpr const char* units[] = {"KB", "MB", "GB", "TB"};
        double size = static_cast<double>(value);
        int unit_index = -1;
        do {
            size /= 1024.0;
            ++unit_index;
        } while (size >= 1024.0 && unit_index < 3);
        char buffer[64];
        // .NET "0.##" format: up to two fractional digits, trailing zeros trimmed.
        std::snprintf(buffer, sizeof(buffer), "%.2f", size);
        std::string text = buffer;
        while (!text.empty() && text.back() == '0') text.pop_back();
        if (!text.empty() && text.back() == '.') text.pop_back();
        return text + " " + units[unit_index];
    }

private:
    static int compare_ordinal_ignore_case(const std::string& left, const std::string& right) {
        // StringComparer.OrdinalIgnoreCase: invariant upper-case ordinal compare
        // (ASCII folding is sufficient for the job names we compare here; .NET
        // folds the full Unicode range, which only diverges for non-ASCII pairs).
        std::size_t n = std::min(left.size(), right.size());
        for (std::size_t i = 0; i < n; ++i) {
            unsigned char l = static_cast<unsigned char>(left[i]);
            unsigned char r = static_cast<unsigned char>(right[i]);
            if (l >= 'a' && l <= 'z') l = static_cast<unsigned char>(l - 'a' + 'A');
            if (r >= 'a' && r <= 'z') r = static_cast<unsigned char>(r - 'a' + 'A');
            if (l != r) return l < r ? -1 : 1;
        }
        if (left.size() == right.size()) return 0;
        return left.size() < right.size() ? -1 : 1;
    }

    bool try_dequeue_internal(const std::string& queue_entry_id, QueuedJobWorkItem& work_item) {
        std::lock_guard<std::mutex> lock(mutex_);
        DateTimeOffset now_utc = DateTimeOffset::now_utc();
        bool mutated = false;
        std::size_t entry_index = 0;
        bool specific_entry_requested = !queue_entry_id.empty();

        while (entry_index < catalog_.QueueEntries.size()) {
            storage::ManagedJobQueueEntry& entry = catalog_.QueueEntries[entry_index];
            if (storage::catalog_detail::is_blank(entry.JobId)) {
                catalog_.QueueEntries.erase(catalog_.QueueEntries.begin() + static_cast<std::ptrdiff_t>(entry_index));
                mutated = true;
                continue;
            }

            if (specific_entry_requested &&
                !models::detail::equals_ignore_case(entry.QueueEntryId, queue_entry_id)) {
                ++entry_index;
                continue;
            }

            const storage::ManagedJob* job = find_job_by_id_locked(entry.JobId);
            if (job == nullptr) {
                catalog_.QueueEntries.erase(catalog_.QueueEntries.begin() + static_cast<std::ptrdiff_t>(entry_index));
                mutated = true;
                continue;
            }

            storage::ManagedJobQueueEntry taken = std::move(entry);
            catalog_.QueueEntries.erase(catalog_.QueueEntries.begin() + static_cast<std::ptrdiff_t>(entry_index));

            taken.AttemptCount = std::max(0, taken.AttemptCount) + 1;
            taken.LastAttemptUtc = now_utc;
            taken.LastUpdatedUtc = now_utc;

            save_catalog_locked();

            work_item.QueueEntryId = taken.QueueEntryId;
            work_item.Trigger = storage::catalog_detail::is_blank(taken.Trigger) ? std::string("queued") : taken.Trigger;
            work_item.EnqueuedBy = taken.EnqueuedBy;
            work_item.EnqueuedUtc = taken.EnqueuedUtc;
            work_item.Attempt = taken.AttemptCount;
            work_item.Job = *job;
            return true;
        }

        if (mutated) save_catalog_locked();
        return false;
    }

    void update_run_status(const std::string& run_id, storage::ManagedJobRunStatus status) {
        if (storage::catalog_detail::is_blank(run_id)) return;

        std::lock_guard<std::mutex> lock(mutex_);
        storage::ManagedJobRun* run = find_run_by_id_locked(run_id);
        if (run == nullptr) return;

        run->Status = status;
        run->LastUpdatedUtc = DateTimeOffset::now_utc();
        save_catalog_locked();
    }

    storage::ManagedJobRun create_run_locked(const std::string& source_job_id, const std::string& display_name,
                                             const models::CopyJobOptions& options, const std::string& trigger,
                                             const std::string& queue_entry_id, std::int32_t queue_attempt) {
        DateTimeOffset now_utc = DateTimeOffset::now_utc();
        storage::ManagedJobRun run;
        run.RunId = storage::detail::new_guid_n();
        run.JobId = source_job_id;
        run.DisplayName = display_name;
        run.SourceRoot = options.SourceRoot;
        run.DestinationRoot = options.DestinationRoot;
        run.Trigger = trigger;
        run.QueueEntryId = queue_entry_id;
        run.QueueAttempt = std::max(0, queue_attempt);
        run.Status = storage::ManagedJobRunStatus::Queued;
        run.StartedUtc = now_utc;
        run.LastUpdatedUtc = now_utc;
        run.JournalPath.clear();
        run.Summary = "Queued";

        catalog_.Runs.insert(catalog_.Runs.begin(), run);
        if (catalog_.Runs.size() > static_cast<std::size_t>(MaximumRunHistory)) {
            catalog_.Runs.resize(static_cast<std::size_t>(MaximumRunHistory));
        }
        return run;
    }

    storage::ManagedJob* find_job_by_id_locked(const std::string& job_id) {
        if (storage::catalog_detail::is_blank(job_id)) return nullptr;
        for (auto& job : catalog_.Jobs) {
            if (models::detail::equals_ignore_case(job.JobId, job_id)) return &job;
        }
        return nullptr;
    }

    storage::ManagedJobRun* find_run_by_id_locked(const std::string& run_id) {
        if (storage::catalog_detail::is_blank(run_id)) return nullptr;
        for (auto& run : catalog_.Runs) {
            if (models::detail::equals_ignore_case(run.RunId, run_id)) return &run;
        }
        return nullptr;
    }

    // Drops queue entries with blank/unknown job ids or duplicate entry ids;
    // persists if anything was dropped. Returns the sanitized queue snapshot.
    std::vector<storage::ManagedJobQueueEntry> build_sanitized_queue_locked() {
        std::vector<storage::ManagedJobQueueEntry> sanitized;
        sanitized.reserve(catalog_.QueueEntries.size());
        std::vector<std::string> seen_entry_ids;

        for (auto& entry : catalog_.QueueEntries) {
            std::string entry_id = storage::catalog_detail::trim_copy(entry.QueueEntryId);
            if (entry_id.empty()) {
                entry.QueueEntryId = storage::detail::new_guid_n();
                entry_id = entry.QueueEntryId;
            }

            bool seen = false;
            for (const auto& existing : seen_entry_ids) {
                if (models::detail::equals_ignore_case(existing, entry_id)) {
                    seen = true;
                    break;
                }
            }
            if (seen) continue;
            seen_entry_ids.push_back(entry_id);

            if (storage::catalog_detail::is_blank(entry.JobId)) continue;
            if (find_job_by_id_locked(entry.JobId) == nullptr) continue;
            sanitized.push_back(entry);
        }

        if (catalog_.QueueEntries.size() != sanitized.size()) {
            catalog_.QueueEntries = sanitized;
            save_catalog_locked();
        }

        return sanitized;
    }

    void save_catalog_locked() { store_.save(catalog_); }

    static std::string format_result_speed_suffix(const models::CopyJobResult& result) {
        if (result.AverageBytesPerSecond <= 0) return std::string();
        return " at " + format_bytes(static_cast<std::int64_t>(result.AverageBytesPerSecond)) + "/s";
    }

    std::mutex mutex_;
    storage::JobCatalogStore store_;
    storage::JobCatalog catalog_;
};

} // namespace xact::ui
