// -----------------------------------------------------------------------------
// File: src\ui\task_model.h
// Purpose: HWND-free model behind the main window's "What to copy" and "Tasks"
//          panels: the user's source items, how they split into worker jobs,
//          the session task list and its states, and the small pending-task
//          file that lets unstarted tasks survive a restart or an elevation
//          handoff. Everything here is pure logic so it can be unit-tested.
// -----------------------------------------------------------------------------

#pragma once

#include <algorithm>
#include <condition_variable>
#include <cstdint>
#include <mutex>
#include <optional>
#include <string>
#include <thread>
#include <vector>

#include "../core/json.h"
#include "../core/models.h"
#include "../storage/stores.h"

namespace xact::ui {

// ---- Sources ----------------------------------------------------------------

// One row of the "What to copy" list. A folder is copied as itself (it appears
// by name inside the destination, as Explorer does) unless `contents_only`,
// which copies what is inside it, the classic "source folder" behaviour.
struct SourceItem {
    std::wstring path;      // normalized full path, no trailing separator
    bool is_directory = false;
    bool contents_only = false;
};

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

enum class TaskState {
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

struct SessionTask {
    std::uint64_t id = 0;
    models::CopyJobOptions options;
    std::string display_name;   // job name for catalog records ("Manual run", a saved job name)
    std::string managed_run_id; // Job Manager run already created for this task, if any
    bool from_recovery = false;
    TaskState state = TaskState::Queued;
    double progress = 0.0;      // 0..1 overall
    double bytes_per_second = 0.0;
    std::int64_t eta_seconds = -1;
    std::string outcome;        // short result text for finished tasks
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

class TaskList {
public:
    std::uint64_t add(SessionTask task) {
        task.id = next_id_++;
        tasks_.push_back(std::move(task));
        return tasks_.back().id;
    }

    // Removes a task that is not running. Returns false for an active task.
    bool remove(std::uint64_t id) {
        auto found = locate(id);
        if (found == tasks_.end() || is_active(found->state)) return false;
        tasks_.erase(found);
        return true;
    }

    // Moves a queued task among the queued tasks (delta -1 up, +1 down).
    bool move(std::uint64_t id, int delta) {
        auto found = locate(id);
        if (found == tasks_.end() || found->state != TaskState::Queued || delta == 0) return false;
        std::size_t index = static_cast<std::size_t>(found - tasks_.begin());
        std::size_t target = index;
        for (;;) {
            if (delta < 0 && target == 0) return false;
            if (delta > 0 && target + 1 >= tasks_.size()) return false;
            target = delta < 0 ? target - 1 : target + 1;
            if (tasks_[target].state == TaskState::Queued) break;
        }
        std::swap(tasks_[index], tasks_[target]);
        return true;
    }

    // Moves a queued task ahead of every other queued task.
    bool move_to_front(std::uint64_t id) {
        auto found = locate(id);
        if (found == tasks_.end() || found->state != TaskState::Queued) return false;
        SessionTask task = std::move(*found);
        tasks_.erase(found);
        auto first_queued = std::find_if(tasks_.begin(), tasks_.end(), [](const SessionTask& t) {
            return t.state == TaskState::Queued;
        });
        tasks_.insert(first_queued, std::move(task));
        return true;
    }

    SessionTask* find(std::uint64_t id) {
        auto found = locate(id);
        return found == tasks_.end() ? nullptr : &*found;
    }

    SessionTask* next_queued() {
        auto found = std::find_if(tasks_.begin(), tasks_.end(), [](const SessionTask& task) {
            return task.state == TaskState::Queued;
        });
        return found == tasks_.end() ? nullptr : &*found;
    }

    SessionTask* active() {
        auto found = std::find_if(tasks_.begin(), tasks_.end(), [](const SessionTask& task) {
            return is_active(task.state);
        });
        return found == tasks_.end() ? nullptr : &*found;
    }

    void clear_finished() {
        tasks_.erase(std::remove_if(tasks_.begin(), tasks_.end(),
                                    [](const SessionTask& task) { return is_finished(task.state); }),
                     tasks_.end());
    }

    struct Counts {
        int queued = 0;
        int active = 0;
        int finished = 0;
    };

    Counts counts() const {
        Counts counts;
        for (const auto& task : tasks_) {
            if (task.state == TaskState::Queued) ++counts.queued;
            else if (is_active(task.state)) ++counts.active;
            else ++counts.finished;
        }
        return counts;
    }

    std::size_t size() const noexcept { return tasks_.size(); }
    SessionTask& at(std::size_t index) { return tasks_[index]; }
    const SessionTask& at(std::size_t index) const { return tasks_[index]; }

    std::optional<std::size_t> index_of(std::uint64_t id) const {
        for (std::size_t i = 0; i < tasks_.size(); ++i) {
            if (tasks_[i].id == id) return i;
        }
        return std::nullopt;
    }

    std::vector<const SessionTask*> queued() const {
        std::vector<const SessionTask*> result;
        for (const auto& task : tasks_) {
            if (task.state == TaskState::Queued) result.push_back(&task);
        }
        return result;
    }

private:
    std::vector<SessionTask> tasks_;
    std::uint64_t next_id_ = 1;

    std::vector<SessionTask>::iterator locate(std::uint64_t id) {
        return std::find_if(tasks_.begin(), tasks_.end(),
                            [id](const SessionTask& task) { return task.id == id; });
    }
};

// ---- Pending-task persistence --------------------------------------------------

// Only tasks that never started are written: options and display name. The
// file is a convenience, not a trust anchor; nothing in it runs until the user
// starts the task list (or, for an elevation handoff, the unelevated window
// that just wrote it asked for that run).
inline std::wstring pending_tasks_path() {
    return storage::fsutil::local_app_data() + L"\\XactCopy\\session-tasks.json";
}

struct PendingTasks {
    std::vector<SessionTask> tasks;
    bool list_running = false; // the user had started the list; continue after a handoff
};

inline std::string pending_tasks_to_json(const std::vector<const SessionTask*>& tasks,
                                         bool list_running) {
    json::Writer w(true);
    w.begin_object();
    w.key("SchemaVersion"); w.value(static_cast<std::int32_t>(1));
    w.key("ListRunning"); w.value(list_running);
    w.key("Tasks");
    w.begin_array();
    for (const SessionTask* task : tasks) {
        w.begin_object();
        w.key("DisplayName"); w.value(task->display_name);
        w.key("Options"); task->options.to_json(w);
        w.end_object();
    }
    w.end_array();
    w.end_object();
    return w.take();
}

inline PendingTasks pending_tasks_from_json(std::string_view text) {
    PendingTasks pending;
    std::vector<SessionTask>& tasks = pending.tasks;
    try {
        const json::Value parsed = json::parse(text);
        const json::Object* root = parsed.as_object();
        if (root == nullptr) return pending;
        const json::Value* schema = root->find("SchemaVersion");
        if (schema == nullptr || schema->as_int64() != 1) return pending;
        const json::Value* running = root->find("ListRunning");
        pending.list_running = running != nullptr && running->as_bool();
        const json::Value* list = root->find("Tasks");
        if (list == nullptr || list->as_array() == nullptr) return pending;
        for (const auto& entry : *list->as_array()) {
            const json::Object* object = entry.as_object();
            if (object == nullptr) continue;
            const json::Value* options = object->find("Options");
            if (options == nullptr || options->as_object() == nullptr) continue;
            SessionTask task;
            task.options = models::CopyJobOptions::from_json(*options);
            if (task.options.SourceRoot.empty()) continue;
            const json::Value* name = object->find("DisplayName");
            if (name != nullptr && name->is_string()) task.display_name = name->as_string();
            tasks.push_back(std::move(task));
        }
    } catch (const std::exception&) {
        tasks.clear();
        pending.list_running = false;
    }
    return pending;
}

inline PendingTasks load_pending_tasks(const std::wstring& path = pending_tasks_path()) {
    auto bytes = storage::fsutil::read_all_bytes(path, 64ULL * 1024ULL * 1024ULL);
    if (!bytes.has_value() || bytes->empty()) return {};
    return pending_tasks_from_json(std::string(bytes->begin(), bytes->end()));
}

// Writes the pending-task file off the UI thread. The latest payload wins;
// flush() writes synchronously on the caller's thread for shutdown.
class PendingTaskWriter {
public:
    explicit PendingTaskWriter(std::wstring path = pending_tasks_path())
        : path_(std::move(path)), thread_([this] { loop(); }) {}

    ~PendingTaskWriter() {
        {
            std::lock_guard<std::mutex> guard(lock_);
            stopping_ = true;
        }
        ready_.notify_one();
        if (thread_.joinable()) thread_.join();
    }

    PendingTaskWriter(const PendingTaskWriter&) = delete;
    PendingTaskWriter& operator=(const PendingTaskWriter&) = delete;

    void post(std::string payload) {
        {
            std::lock_guard<std::mutex> guard(lock_);
            pending_ = std::move(payload);
        }
        ready_.notify_one();
    }

    // Waits for any in-flight write, then writes `payload` now.
    void flush(const std::string& payload) {
        std::lock_guard<std::mutex> write_guard(write_lock_);
        {
            std::lock_guard<std::mutex> guard(lock_);
            pending_.reset();
        }
        write(payload);
    }

private:
    std::wstring path_;
    std::mutex lock_;
    std::mutex write_lock_;
    std::condition_variable ready_;
    std::optional<std::string> pending_;
    bool stopping_ = false;
    std::thread thread_;

    void write(const std::string& payload) {
        storage::fsutil::create_directories(storage::fsutil::get_directory_name(path_));
        (void)storage::fsutil::write_atomic_bytes(path_, payload);
    }

    void loop() {
        for (;;) {
            std::string payload;
            {
                std::unique_lock<std::mutex> guard(lock_);
                ready_.wait(guard, [this] { return stopping_ || pending_.has_value(); });
                if (!pending_.has_value()) return;
                payload = std::move(*pending_);
                pending_.reset();
            }
            std::lock_guard<std::mutex> write_guard(write_lock_);
            write(payload);
        }
    }
};

} // namespace xact::ui
