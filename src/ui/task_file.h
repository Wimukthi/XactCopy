// -----------------------------------------------------------------------------
// File: src\ui\task_file.h
// Purpose: The tasks file (*.xctasks): a readable JSON list of tasks that
//          XactCopy exports and imports, for moving tasks between machines or
//          writing them by hand. Pure logic so it can be unit-tested.
//
//   {
//     "Format": "XactCopy.Tasks",
//     "Version": 1,
//     "Tasks": [
//       {
//         "Name": "Photos to backup",            optional; made from the items
//         "Action": "Copy",                       "Copy" or "Assess"; default Copy
//         "Sources": ["D:\\Photos", "E:\\DCIM"],  full paths, at least one; a source
//                                                 can also be {"Path": "D:\\Photos",
//                                                 "ContentsOnly": true}: copy what
//                                                 is in the folder, not the folder
//         "ContentsOnly": false,                  the same, for a one-source task
//         "Destination": "F:\\Backup",            required for Copy
//         "Pinned": false,                        optional
//         "Options": { ... }                      optional: any task option, by
//       }                                         its saved name; the rest come
//     ]                                           from your defaults
//   }
//
// Paths, media identities, and journal locations are machine-specific, so
// they are carried by Sources / Destination only; Options never sets them.
// -----------------------------------------------------------------------------

#pragma once

#include <string>
#include <string_view>
#include <vector>

#include "../core/json.h"
#include "../core/models.h"
#include "../storage/job_catalog.h"
#include "task_model.h"

namespace xact::ui {

inline constexpr const char* TaskFileFormat = "XactCopy.Tasks";
inline constexpr std::int32_t TaskFileVersion = 1;
inline constexpr std::size_t TaskFileMaximumTasks = 10000;

struct TaskFileEntry {
    std::string name;
    models::CopyJobOptions options;
    std::vector<storage::ManagedJobPart> parts;
    bool pinned = false;
};

struct TaskFileContents {
    std::vector<TaskFileEntry> tasks;
    std::vector<std::string> problems; // tasks that were skipped, and why
    std::string error;                 // the file as a whole could not be read
};

struct TaskFileSource {
    std::string path;
    bool contents_only = false; // copy what is in the folder, not the folder
};

namespace task_file_detail {

// Options members that are machine-specific or come from Action, Sources,
// and Destination.
inline bool is_machine_specific(std::string_view key) {
    for (const char* name : {"SourceRoot", "SelectedRelativePaths", "DestinationRoot", "ExpectedSourceIdentity",
                             "ExpectedDestinationIdentity", "ResumeJournalPathHint", "JournalDirectory",
                             "OperationMode", "AllowJournalRootRemap"}) {
        if (key == name) return true;
    }
    return false;
}

inline void write_value(json::Writer& w, const json::Value& value) {
    switch (value.kind()) {
        case json::Kind::Null: w.value(nullptr); break;
        case json::Kind::Boolean: w.value(value.as_bool()); break;
        case json::Kind::Int64: w.value(value.as_int64()); break;
        case json::Kind::Double: w.value(value.as_double()); break;
        case json::Kind::String: w.value(std::string_view(value.as_string())); break;
        case json::Kind::ArrayKind:
            w.begin_array();
            for (const auto& item : *value.as_array()) write_value(w, item);
            w.end_array();
            break;
        case json::Kind::ObjectKind:
            w.begin_object();
            for (const auto& [key, member] : value.as_object()->members) {
                w.key(key);
                write_value(w, member);
            }
            w.end_object();
            break;
    }
}

inline json::Value options_value(const models::CopyJobOptions& options) {
    json::Writer writer;
    options.to_json(writer);
    return json::parse(writer.take());
}

inline std::string join_path(const std::string& root, const std::string& relative) {
    if (root.empty() || root.back() == '\\' || root.back() == '/') return root + relative;
    return root + "\\" + relative;
}

} // namespace task_file_detail

// A tasks file, by its extension (a drop imports it instead of copying it).
inline bool is_task_file_path(const std::wstring& path) {
    const std::wstring extension = L".xctasks";
    return path.size() > extension.size() &&
           _wcsicmp(path.c_str() + (path.size() - extension.size()), extension.c_str()) == 0;
}

// The full paths a task copies, in order.
inline std::vector<TaskFileSource> task_sources(const storage::ManagedJob& job) {
    std::vector<TaskFileSource> sources;
    auto add_part = [&](const std::string& root, const std::vector<std::string>& items) {
        if (items.empty()) {
            sources.push_back({root, true});
            return;
        }
        for (const auto& item : items) sources.push_back({task_file_detail::join_path(root, item), false});
    };
    add_part(job.Options.SourceRoot, job.Options.SelectedRelativePaths);
    for (const auto& part : job.Parts) add_part(part.SourceRoot, part.SelectedRelativePaths);
    return sources;
}

inline std::string write_task_file(const std::vector<storage::ManagedJob>& jobs) {
    json::Writer w(/*indented*/ true);
    w.begin_object();
    w.key("Format"); w.value(TaskFileFormat);
    w.key("Version"); w.value(TaskFileVersion);
    w.key("Tasks");
    w.begin_array();
    for (const auto& job : jobs) {
        const bool scan = job.Options.OperationMode == models::JobOperationMode::ScanOnly;
        w.begin_object();
        w.key("Name"); w.value(job.Name);
        w.key("Action"); w.value(scan ? "Assess" : "Copy");
        w.key("Sources");
        w.begin_array();
        for (const auto& source : task_sources(job)) {
            // A drive root is always its contents, so it stays a plain path.
            if (source.contents_only && source.path.size() > 3) {
                w.begin_object();
                w.key("Path"); w.value(source.path);
                w.key("ContentsOnly"); w.value(true);
                w.end_object();
            } else {
                w.value(source.path);
            }
        }
        w.end_array();
        if (!scan) {
            w.key("Destination"); w.value(job.Options.DestinationRoot);
        }
        w.key("Pinned"); w.value(job.Pinned);
        w.key("Options");
        w.begin_object();
        const json::Value options = task_file_detail::options_value(job.Options);
        for (const auto& [key, member] : options.as_object()->members) {
            if (task_file_detail::is_machine_specific(key)) continue;
            w.key(key);
            task_file_detail::write_value(w, member);
        }
        w.end_object();
        w.end_object();
    }
    w.end_array();
    w.end_object();
    return w.take();
}

// Reads a tasks file. Each task's options start from `defaults` (the user's
// saved defaults); the file's Options override them member by member. A
// task that cannot be used is skipped and named in `problems`.
inline TaskFileContents read_task_file(std::string_view text, const models::CopyJobOptions& defaults) {
    TaskFileContents contents;
    json::Value root;
    try {
        root = json::parse(text);
    } catch (const std::exception& error) {
        contents.error = std::string("The file is not valid JSON: ") + error.what();
        return contents;
    }
    const json::Object* object = root.as_object();
    const json::Value* format = object != nullptr ? object->find("Format") : nullptr;
    if (format == nullptr || format->as_string() != TaskFileFormat) {
        contents.error = "The file is not an XactCopy tasks file (its \"Format\" is not \"XactCopy.Tasks\").";
        return contents;
    }
    const json::Value* version = object->find("Version");
    if (version == nullptr || version->as_int64() < 1 || version->as_int64() > TaskFileVersion) {
        contents.error = "The file needs a newer XactCopy (tasks file version " +
                         std::to_string(version != nullptr ? version->as_int64() : 0) + ").";
        return contents;
    }
    const json::Value* tasks = object->find("Tasks");
    if (tasks == nullptr || !tasks->is_array()) {
        contents.error = "The file has no \"Tasks\" list.";
        return contents;
    }
    if (tasks->as_array()->size() > TaskFileMaximumTasks) {
        contents.error = "The file lists more than " + std::to_string(TaskFileMaximumTasks) + " tasks.";
        return contents;
    }

    const json::Value default_options = task_file_detail::options_value(defaults);
    std::size_t number = 0;
    for (const auto& entry : *tasks->as_array()) {
        ++number;
        const std::string label = "Task " + std::to_string(number);
        const json::Object* task = entry.as_object();
        if (task == nullptr) {
            contents.problems.push_back(label + ": not an object.");
            continue;
        }
        TaskFileEntry imported;
        models::detail::read(task, "Name", imported.name);
        models::detail::read(task, "Pinned", imported.pinned);
        const std::string named = imported.name.empty() ? label : label + " (" + imported.name + ")";

        // The defaults, with the file's options over them.
        json::Object merged = *default_options.as_object();
        if (const json::Value* overrides = task->find("Options"); overrides != nullptr) {
            if (!overrides->is_object()) {
                contents.problems.push_back(named + ": \"Options\" is not an object.");
                continue;
            }
            for (const auto& [key, member] : overrides->as_object()->members) {
                if (task_file_detail::is_machine_specific(key)) continue;
                bool replaced = false;
                for (auto& existing : merged.members) {
                    if (existing.first == key) {
                        existing.second = member;
                        replaced = true;
                    }
                }
                if (!replaced) merged.add(key, member);
            }
        }
        imported.options = models::CopyJobOptions::from_json(json::Value(std::move(merged)));

        std::string action = "Copy";
        models::detail::read(task, "Action", action);
        const bool scan = models::detail::equals_ignore_case(action, "Assess");
        if (!scan && !models::detail::equals_ignore_case(action, "Copy")) {
            contents.problems.push_back(named + ": \"Action\" must be \"Copy\" or \"Assess\".");
            continue;
        }
        imported.options.OperationMode =
            scan ? models::JobOperationMode::ScanOnly : models::JobOperationMode::Copy;

        std::vector<SourceItem> items;
        if (const json::Value* sources = task->find("Sources"); sources != nullptr && sources->is_array()) {
            for (const auto& source : *sources->as_array()) {
                std::string source_path = source.as_string();
                bool folder_contents = false;
                if (const json::Object* described = source.as_object()) {
                    models::detail::read(described, "Path", source_path);
                    models::detail::read(described, "ContentsOnly", folder_contents);
                }
                std::wstring path = storage::fsutil::trim(storage::fsutil::utf8_to_wide(source_path));
                while (path.size() > 3 && (path.back() == L'\\' || path.back() == L'/')) path.pop_back();
                if (path.empty()) continue;
                SourceItem item;
                item.path = path;
                // A drive root has no parent to be copied "as itself" into.
                item.is_directory = folder_contents || path.size() <= 3;
                item.contents_only = item.is_directory;
                items.push_back(std::move(item));
            }
        }
        if (items.empty()) {
            contents.problems.push_back(named + ": no \"Sources\".");
            continue;
        }
        bool contents_only = false;
        models::detail::read(task, "ContentsOnly", contents_only);
        if (contents_only && items.size() == 1) {
            items[0].is_directory = true;
            items[0].contents_only = true;
        }
        const auto groups = group_sources(items);
        if (groups.empty()) {
            contents.problems.push_back(named + ": its sources are not full paths.");
            continue;
        }
        imported.options.SourceRoot = storage::fsutil::wide_to_utf8(groups[0].root);
        imported.options.SelectedRelativePaths = groups[0].relative_paths;
        for (std::size_t g = 1; g < groups.size(); ++g) {
            imported.parts.push_back({storage::fsutil::wide_to_utf8(groups[g].root), groups[g].relative_paths});
        }

        std::string destination;
        models::detail::read(task, "Destination", destination);
        if (!scan && storage::fsutil::trim(storage::fsutil::utf8_to_wide(destination)).empty()) {
            contents.problems.push_back(named + ": a Copy task needs a \"Destination\".");
            continue;
        }
        imported.options.DestinationRoot = scan ? std::string() : destination;
        if (imported.name.empty()) imported.name = task_title_for_parts(imported.options, items.size(), groups.size());
        contents.tasks.push_back(std::move(imported));
    }
    return contents;
}

} // namespace xact::ui
