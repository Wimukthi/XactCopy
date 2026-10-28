// -----------------------------------------------------------------------------
// File: src\ui\task_model.h
// Purpose: HWND-free model behind the main window's "What to copy" list and
//          the Tasks / Queue / History panes: the user's source items, how they
//          split into worker jobs, task states and labels, the on-exit policy,
//          and the reader for the legacy session-task file. Everything here is
//          pure logic so it can be unit-tested.
// -----------------------------------------------------------------------------

#pragma once

#include <algorithm>
#include <cstdint>
#include <optional>
#include <string>
#include <vector>

#include "../core/json.h"
#include "../core/models.h"
#include "../storage/job_catalog.h"
#include "../storage/stores.h"

namespace xact::ui {

// ---- Sources ----------------------------------------------------------------

// One file or folder in the "What to copy" list. A folder is copied as itself
// (it appears by name inside the destination, as Explorer does) unless
// `contents_only`, which copies what is inside it, the classic "source folder"
// behaviour.
struct SourceItem {
    std::wstring path;      // normalized full path, no trailing separator
    bool is_directory = false;
    bool contents_only = false;
    std::uint32_t batch = 0; // the add action that brought it in; one list row per batch
};

// The items one add action brought in (several folders picked in one dialog,
// one paste, one drop). Each batch is one row in the list and becomes one task.
struct SourceBatch {
    std::uint32_t id = 0;
    std::vector<std::size_t> items; // indexes into the source list
};

// Batches in the order their first item appears.
inline std::vector<SourceBatch> source_batches(const std::vector<SourceItem>& items) {
    std::vector<SourceBatch> batches;
    for (std::size_t i = 0; i < items.size(); ++i) {
        auto found = std::find_if(batches.begin(), batches.end(),
                                  [&](const SourceBatch& batch) { return batch.id == items[i].batch; });
        if (found == batches.end()) {
            batches.push_back(SourceBatch{items[i].batch, {}});
            found = batches.end() - 1;
        }
        found->items.push_back(i);
    }
    return batches;
}

// One worker job's worth of sources: a directory root plus the exact relative
// items beneath it (empty means "everything under root").
struct SourceGroup {
    std::wstring root;
    std::vector<std::string> relative_paths;
    std::vector<std::wstring> item_names; // destination top-level names; empty when unknown
    bool contents_only = false;
};

namespace task_detail {

inline bool same_path(const std::wstring& left, const std::wstring& right) {
    return _wcsicmp(left.c_str(), right.c_str()) == 0;
}

inline std::wstring parent_of(const std::wstring& path) {
    std::wstring parent = storage::fsutil::get_directory_name(path);
    // "C:" is a drive-relative path; the root of a drive is "C:\".
    if (parent.size() == 2 && parent[1] == L':') parent.push_back(L'\\');
    return parent;
}

} // namespace task_detail

// Splits the source list into worker jobs. Items that share a parent folder
// become one job rooted at that parent, so each item keeps its own name in the
// destination. A "contents only" folder is always a job of its own rooted at
// the folder. Order follows the first appearance of each parent.
inline std::vector<SourceGroup> group_sources(const std::vector<SourceItem>& items) {
    std::vector<SourceGroup> groups;
    for (const auto& item : items) {
        if (item.path.empty()) continue;
        if (item.contents_only && item.is_directory) {
            SourceGroup group;
            group.root = item.path;
            group.contents_only = true;
            groups.push_back(std::move(group));
            continue;
        }
        const std::wstring parent = task_detail::parent_of(item.path);
        const std::wstring name = storage::fsutil::get_file_name(item.path);
        if (parent.empty() || name.empty()) continue;
        auto existing = std::find_if(groups.begin(), groups.end(), [&](const SourceGroup& group) {
            return !group.contents_only && task_detail::same_path(group.root, parent);
        });
        if (existing == groups.end()) {
            SourceGroup group;
            group.root = parent;
            groups.push_back(std::move(group));
            existing = groups.end() - 1;
        }
        const std::string relative = storage::fsutil::wide_to_utf8(name);
        bool duplicate = false;
        for (const auto& known : existing->relative_paths) {
            if (models::detail::equals_ignore_case(known, relative)) duplicate = true;
        }
        if (!duplicate) {
            existing->relative_paths.push_back(relative);
            existing->item_names.push_back(name);
        }
    }
    return groups;
}

// Top-level destination names that more than one job would write, e.g. two
// different "Photos" folders. Contents-only jobs are not enumerated here.
inline std::vector<std::wstring> find_destination_collisions(const std::vector<SourceGroup>& groups) {
    std::vector<std::wstring> collisions;
    for (std::size_t a = 0; a < groups.size(); ++a) {
        for (const auto& name : groups[a].item_names) {
            for (std::size_t b = a + 1; b < groups.size(); ++b) {
                for (const auto& other : groups[b].item_names) {
                    if (!task_detail::same_path(name, other)) continue;
                    const bool listed = std::any_of(
                        collisions.begin(), collisions.end(),
                        [&](const std::wstring& known) { return task_detail::same_path(known, name); });
                    if (!listed) collisions.push_back(name);
                }
            }
        }
    }
    return collisions;
}

// Reverse mapping, for restoring a job's options into the sources list.
inline std::vector<SourceItem> items_from_options(const models::CopyJobOptions& options) {
    std::vector<SourceItem> items;
    const std::wstring root = storage::fsutil::utf8_to_wide(options.SourceRoot);
    if (root.empty()) return items;
    if (options.SelectedRelativePaths.empty()) {
        SourceItem item;
        item.path = root;
        item.is_directory = true;
        item.contents_only = true;
        items.push_back(std::move(item));
        return items;
    }
    std::wstring base = root;
    while (base.size() > 3 && (base.back() == L'\\' || base.back() == L'/')) base.pop_back();
    for (const auto& relative : options.SelectedRelativePaths) {
        SourceItem item;
        item.path = (base.back() == L'\\' ? base : base + L"\\") +
                    storage::fsutil::utf8_to_wide(relative);
        items.push_back(std::move(item));
    }
    return items;
}

// ---- Option summary -----------------------------------------------------------

inline std::string summarize_options(const models::CopyJobOptions& options) {
    std::vector<std::string> parts;
    if (options.OperationMode == models::JobOperationMode::ScanOnly) {
        parts.push_back(options.UseExperimentalRawDiskScan ? "direct NTFS reads" : "file reads");
    } else {
        const auto verification = options.VerificationModeValue;
        if (!options.VerifyAfterCopy || verification == models::VerificationMode::None) {
            parts.push_back("no verification");
        } else {
            parts.push_back(verification == models::VerificationMode::Sampled ? "verify sampled"
                                                                              : "verify full");
        }
    }
    parts.push_back(std::to_string(options.MaxRetries) +
                    (options.MaxRetries == 1 ? " retry" : " retries"));
    if (options.ResumeFromJournal) parts.push_back("resume");
    if (options.UseBadRangeMap) {
        parts.push_back(options.SkipKnownBadRanges ? "bad-range map, skip known bad"
                                                   : "bad-range map");
    }
    if (options.OperationMode == models::JobOperationMode::Copy) {
        if (options.SalvageUnreadableBlocks) parts.push_back("salvage");
        if (options.ContinueOnFileError) parts.push_back("continue on error");
    }
    if (options.FragileMediaMode) parts.push_back("fragile media");
    if (options.WaitForMediaAvailability) parts.push_back("wait for media");
    std::string text;
    for (const auto& part : parts) {
        if (!text.empty()) text += " \xC2\xB7 "; // middle dot
        text += part;
    }
    return text;
}

// ---- Tasks -------------------------------------------------------------------
//
// Tasks are the job catalog's saved jobs (JobManagerService): a pinned job
// always stays, an unpinned one was added from the builder and follows the
// on-exit choice. The queue is the catalog's queue, and history is its run
// records. What lives here is the state of the one job that is starting or
// running, and the pure helpers the panes draw from.

enum class TaskState {
    Ready,      // never run
    Queued,
    Starting,
    WaitingForElevation,
    Running,
    Paused,
    Completed,
    CompletedWithRecovery,
    Incomplete,
    Failed,
    Cancelled,
};

inline const char* to_display(TaskState state) {
    switch (state) {
        case TaskState::Ready: return "Ready";
        case TaskState::Queued: return "Queued";
        case TaskState::Starting: return "Starting";
        case TaskState::WaitingForElevation: return "Waiting for administrator approval";
        case TaskState::Running: return "Running";
        case TaskState::Paused: return "Paused";
        case TaskState::Completed: return "Completed";
        case TaskState::CompletedWithRecovery: return "Completed with recovered data";
        case TaskState::Incomplete: return "Incomplete";
        case TaskState::Failed: return "Failed";
        case TaskState::Cancelled: return "Cancelled";
    }
    return "";
}

inline bool is_active(TaskState state) {
    return state == TaskState::Starting || state == TaskState::WaitingForElevation ||
           state == TaskState::Running || state == TaskState::Paused;
}

inline bool is_finished(TaskState state) {
    return state == TaskState::Completed || state == TaskState::CompletedWithRecovery ||
           state == TaskState::Incomplete || state == TaskState::Failed ||
           state == TaskState::Cancelled;
}

// Maps a worker result onto a task outcome. "Incomplete" is a run that made
// progress but cannot claim an exact result (skips, partial listing).
inline TaskState state_from_result(const models::CopyJobResult& result) {
    if (result.Cancelled) return TaskState::Cancelled;
    if (result.Succeeded) {
        if (result.RecoveredFiles > 0) return TaskState::CompletedWithRecovery;
        if (result.SkippedFiles > 0 || result.FailedFiles > 0) return TaskState::Incomplete;
        return TaskState::Completed;
    }
    if (result.CompletedFiles > 0 && result.FailedFiles == 0 && result.ErrorMessage.empty()) {
        return TaskState::Incomplete;
    }
    return TaskState::Failed;
}

// The state a history record shows as (its colour and label in the panes).
inline TaskState state_from_run(const storage::ManagedJobRun& run) {
    using storage::ManagedJobRunStatus;
    switch (run.Status) {
        case ManagedJobRunStatus::Queued: return TaskState::Queued;
        case ManagedJobRunStatus::Running: return TaskState::Running;
        case ManagedJobRunStatus::Paused: return TaskState::Paused;
        case ManagedJobRunStatus::Completed:
            return run.Result.has_value() ? state_from_result(*run.Result) : TaskState::Completed;
        case ManagedJobRunStatus::Failed: return TaskState::Failed;
        case ManagedJobRunStatus::Cancelled: return TaskState::Cancelled;
        case ManagedJobRunStatus::Interrupted:
        case ManagedJobRunStatus::Incomplete: return TaskState::Incomplete;
    }
    return TaskState::Failed;
}

// The job that is starting or running (at most one: the worker runs one job
// at a time). It may come from the queue, a history resume, crash recovery, or
// an elevation handoff, so job_id can be empty.
struct SessionTask {
    std::string job_id;
    models::CopyJobOptions options;
    std::string display_name;   // job name for catalog records
    std::string managed_run_id; // catalog run for this attempt
    bool from_recovery = false;
    TaskState state = TaskState::Starting;
    // Multi-part tasks run their parts back to back as one task.
    std::size_t part_index = 0;
    std::size_t part_count = 1;
    bool continue_parts = false; // started from the queue: run the later parts too
    std::string queue_entry_id;  // the queue entry this attempt came from
    std::int32_t queue_attempt = 0;
    std::string trigger;
    double progress = 0.0;      // 0..1 overall
    double bytes_per_second = 0.0;
    std::int64_t eta_seconds = -1;
};

inline std::string task_title(const models::CopyJobOptions& options) {
    const bool scan = options.OperationMode == models::JobOperationMode::ScanOnly;
    std::string what;
    const std::size_t count = options.SelectedRelativePaths.size();
    if (count == 1) {
        what = storage::fsutil::wide_to_utf8(storage::fsutil::get_file_name(
            storage::fsutil::utf8_to_wide(options.SelectedRelativePaths.front())));
    } else if (count > 1) {
        what = std::to_string(count) + " items";
    } else {
        std::string root = options.SourceRoot;
        while (root.size() > 3 && (root.back() == '\\' || root.back() == '/')) root.pop_back();
        what = "contents of " + root;
    }
    if (scan) {
        return "Assess " + what + (options.UseExperimentalRawDiskScan ? " (direct NTFS)" : "");
    }
    return "Copy " + what + " \xE2\x86\x92 " + options.DestinationRoot; // right arrow
}

// Title for a task that may read several roots (one part per root).
inline std::string task_title_for_parts(const models::CopyJobOptions& options, std::size_t item_count,
                                        std::size_t part_count) {
    if (part_count <= 1) return task_title(options);
    const std::string what = std::to_string(item_count) + " items from " + std::to_string(part_count) +
                             " locations";
    if (options.OperationMode == models::JobOperationMode::ScanOnly) {
        return "Assess " + what + (options.UseExperimentalRawDiskScan ? " (direct NTFS)" : "");
    }
    return "Copy " + what + " \xE2\x86\x92 " + options.DestinationRoot;
}

// Newest run per job id (runs are stored newest first), as indexes into `runs`.
inline std::vector<std::pair<std::string, std::size_t>> latest_run_per_job(
    const std::vector<storage::ManagedJobRun>& runs) {
    std::vector<std::pair<std::string, std::size_t>> latest;
    for (std::size_t i = 0; i < runs.size(); ++i) {
        if (runs[i].JobId.empty()) continue;
        const bool seen = std::any_of(latest.begin(), latest.end(), [&](const auto& entry) {
            return models::detail::equals_ignore_case(entry.first, runs[i].JobId);
        });
        if (!seen) latest.emplace_back(runs[i].JobId, i);
    }
    return latest;
}

// What to do with unpinned tasks when XactCopy closes (TaskListOnExit).
enum class TaskExitPolicy { Ask, Keep, Discard };

inline TaskExitPolicy parse_task_exit_policy(const std::string& text) {
    if (models::detail::equals_ignore_case(text, "keep")) return TaskExitPolicy::Keep;
    if (models::detail::equals_ignore_case(text, "discard")) return TaskExitPolicy::Discard;
    return TaskExitPolicy::Ask;
}

inline const char* to_setting(TaskExitPolicy policy) {
    switch (policy) {
        case TaskExitPolicy::Keep: return "keep";
        case TaskExitPolicy::Discard: return "discard";
        case TaskExitPolicy::Ask: break;
    }
    return "ask";
}

// ---- Legacy session-task file ---------------------------------------------------

// Before tasks moved into the job catalog, tasks that never started were kept
// in session-tasks.json. It is read once at startup to migrate them, then
// deleted. Nothing in it runs by itself.
inline std::wstring pending_tasks_path() {
    return storage::fsutil::local_app_data() + L"\\XactCopy\\session-tasks.json";
}

struct PendingTask {
    std::string display_name;
    models::CopyJobOptions options;
};

inline std::vector<PendingTask> pending_tasks_from_json(std::string_view text) {
    std::vector<PendingTask> tasks;
    try {
        const json::Value parsed = json::parse(text);
        const json::Object* root = parsed.as_object();
        if (root == nullptr) return tasks;
        const json::Value* schema = root->find("SchemaVersion");
        if (schema == nullptr || schema->as_int64() != 1) return tasks;
        const json::Value* list = root->find("Tasks");
        if (list == nullptr || list->as_array() == nullptr) return tasks;
        for (const auto& entry : *list->as_array()) {
            const json::Object* object = entry.as_object();
            if (object == nullptr) continue;
            const json::Value* options = object->find("Options");
            if (options == nullptr || options->as_object() == nullptr) continue;
            PendingTask task;
            task.options = models::CopyJobOptions::from_json(*options);
            if (task.options.SourceRoot.empty()) continue;
            const json::Value* name = object->find("DisplayName");
            if (name != nullptr && name->is_string()) task.display_name = name->as_string();
            tasks.push_back(std::move(task));
        }
    } catch (const std::exception&) {
        tasks.clear();
    }
    return tasks;
}

inline std::vector<PendingTask> load_pending_tasks(const std::wstring& path = pending_tasks_path()) {
    auto bytes = storage::fsutil::read_all_bytes(path, 64ULL * 1024ULL * 1024ULL);
    if (!bytes.has_value() || bytes->empty()) return {};
    return pending_tasks_from_json(std::string(bytes->begin(), bytes->end()));
}

} // namespace xact::ui
