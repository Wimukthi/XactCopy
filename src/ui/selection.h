// -----------------------------------------------------------------------------
// File: src\ui\selection.h
// Purpose: SelectionModel — the single source of truth for "what to copy" when
//          the user picks more than one item. Explorer verbs, forwarded launches
//          from a second instance, the Add-Files dialog, and drag-and-drop all
//          funnel into one model, which keeps the exact user-facing selection
//          while resolving the execution root and relative paths the worker's
//          SelectionFilter understands.
// -----------------------------------------------------------------------------

#pragma once

#ifndef NOMINMAX
#define NOMINMAX
#endif
#include <windows.h>

#include <string>
#include <vector>

#include "../core/models.h"
#include "../storage/stores.h"

namespace xact::ui {

class SelectionModel {
public:
    // Adds an existing path. Returns false when it is blank, missing, or already
    // present (case-insensitive), so callers can report what was actually taken.
    bool add(const std::wstring& raw) {
        std::wstring full = storage::fsutil::get_full_path(storage::fsutil::trim(raw));
        if (full.empty()) return false;
        while (full.size() > 3 && (full.back() == L'\\' || full.back() == L'/')) full.pop_back();
        if (GetFileAttributesW(full.c_str()) == INVALID_FILE_ATTRIBUTES) return false;
        for (const auto& existing : items_) {
            if (_wcsicmp(existing.c_str(), full.c_str()) == 0) return false;
        }
        items_.push_back(std::move(full));
        return true;
    }

    void clear() { items_.clear(); }
    bool empty() const { return items_.empty(); }
    std::size_t size() const { return items_.size(); }
    const std::vector<std::wstring>& items() const { return items_; }

    int folder_count() const {
        int folders = 0;
        for (const auto& item : items_) {
            if (is_directory(item)) ++folders;
        }
        return folders;
    }

    // What the Source box should show. A single item remains the exact file or
    // folder the user selected; multiple items show their shared source root.
    std::wstring display_path() const {
        if (items_.empty()) return std::wstring();
        if (items_.size() == 1) return items_[0];
        return common_root();
    }

    // The directory the worker runs from. This is intentionally separate from
    // display_path(): the worker needs a directory root plus a relative filter
    // to copy the selected directory itself rather than flattening its contents.
    std::wstring common_root() const {
        if (items_.empty()) return std::wstring();
        if (items_.size() == 1 && is_directory(items_[0])) {
            std::wstring parent = storage::fsutil::get_directory_name(items_[0]);
            return parent.empty() ? items_[0] : parent;
        }
        std::wstring common = items_[0];
        for (std::size_t i = 1; i < items_.size(); ++i) {
            common = reduce_common(common, items_[i]);
            if (common.empty()) return std::wstring();
        }
        if (common.empty()) return std::wstring();
        if (!is_directory(common) && exists(common)) {
            return storage::fsutil::get_directory_name(common);
        }
        return common;
    }

    // Paths relative to `root`, as the worker's SelectionFilter expects. Empty
    // when the selection is exactly the root (meaning "copy everything").
    std::vector<std::string> relative_paths(const std::wstring& root) const {
        std::vector<std::string> relative;
        for (const auto& item : items_) {
            std::string entry = to_relative(root, item);
            if (entry.empty()) continue;
            bool duplicate = false;
            for (const auto& existing : relative) {
                if (models::detail::equals_ignore_case(existing, entry)) {
                    duplicate = true;
                    break;
                }
            }
            if (!duplicate) relative.push_back(std::move(entry));
        }
        return relative;
    }

    // "3 items selected (2 files, 1 folder)"
    std::wstring summary() const {
        if (items_.empty()) return std::wstring();
        const int folders = folder_count();
        const int files = static_cast<int>(items_.size()) - folders;
        std::wstring text = std::to_wstring(items_.size()) +
                            (items_.size() == 1 ? L" item selected" : L" items selected");
        if (files > 0 && folders > 0) {
            text += L" (" + std::to_wstring(files) + (files == 1 ? L" file, " : L" files, ") +
                    std::to_wstring(folders) + (folders == 1 ? L" folder)" : L" folders)");
        } else if (folders > 0) {
            text += folders == 1 ? L" (folder)" : L" (folders)";
        } else {
            text += files == 1 ? L" (file)" : L" (files)";
        }
        return text;
    }

    static bool is_directory(const std::wstring& path) {
        DWORD attributes = GetFileAttributesW(path.c_str());
        return attributes != INVALID_FILE_ATTRIBUTES &&
               (attributes & FILE_ATTRIBUTE_DIRECTORY) != 0;
    }

    static bool exists(const std::wstring& path) {
        return GetFileAttributesW(path.c_str()) != INVALID_FILE_ATTRIBUTES;
    }

private:
    // Longest shared directory prefix of two normalized full paths.
    static std::wstring reduce_common(const std::wstring& left, const std::wstring& right) {
        std::wstring candidate = left;
        while (!candidate.empty()) {
            if (_wcsicmp(candidate.c_str(), right.c_str()) == 0) return candidate;
            std::wstring prefix = candidate + L"\\";
            if (right.size() >= prefix.size() &&
                _wcsnicmp(right.c_str(), prefix.c_str(), prefix.size()) == 0) {
                return candidate;
            }
            std::wstring parent = storage::fsutil::get_directory_name(candidate);
            if (parent == candidate) break;
            candidate = parent;
        }
        return std::wstring();
    }

    static std::string to_relative(const std::wstring& root, const std::wstring& item) {
        if (item.size() <= root.size()) return std::string();
        std::wstring relative = item.substr(root.size());
        while (!relative.empty() && (relative.front() == L'\\' || relative.front() == L'/')) {
            relative.erase(relative.begin());
        }
        while (!relative.empty() && (relative.back() == L'\\' || relative.back() == L'/')) {
            relative.pop_back();
        }
        for (auto& c : relative) {
            if (c == L'/') c = L'\\';
        }
        return storage::fsutil::wide_to_utf8(relative);
    }

    std::vector<std::wstring> items_;
};

} // namespace xact::ui
