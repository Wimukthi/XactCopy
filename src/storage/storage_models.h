// -----------------------------------------------------------------------------
// File: src\storage\storage_models.h
// Purpose: C++ port of the XactCopy persistence models: job journals and
//          bad-range maps, with JSON output matching the .NET indented
//          serializer field-for-field (dictionaries keep document order and
//          look up keys case-insensitively, like Dictionary(OrdinalIgnoreCase)).
// -----------------------------------------------------------------------------

#pragma once

#include <algorithm>
#include <cstdint>
#include <string>
#include <string_view>
#include <unordered_map>
#include <utility>
#include <vector>

#include "../core/dotnet_time.h"
#include "../core/json.h"
#include "../core/models.h"

namespace xact::storage {

using time::DateTimeOffset;

// ---------------------------------------------------------------------------
// Enums
// ---------------------------------------------------------------------------

enum class FileCopyState : std::int32_t {
    Pending = 0, InProgress = 1, Completed = 2, CompletedWithRecovery = 3, Failed = 4
};

enum class RescueRangeState : std::int32_t {
    Pending = 0, Good = 1, Bad = 2, Recovered = 3, KnownBad = 4
};

inline constexpr models::detail::EnumEntry<FileCopyState> FileCopyState_Table[] = {
    {FileCopyState::Pending, "Pending"},
    {FileCopyState::InProgress, "InProgress"},
    {FileCopyState::Completed, "Completed"},
    {FileCopyState::CompletedWithRecovery, "CompletedWithRecovery"},
    {FileCopyState::Failed, "Failed"}};

inline constexpr models::detail::EnumEntry<RescueRangeState> RescueRangeState_Table[] = {
    {RescueRangeState::Pending, "Pending"},
    {RescueRangeState::Good, "Good"},
    {RescueRangeState::Bad, "Bad"},
    {RescueRangeState::Recovered, "Recovered"},
    {RescueRangeState::KnownBad, "KnownBad"}};

inline std::string_view to_string(FileCopyState v) {
    return models::detail::enum_name(FileCopyState_Table, v);
}
inline std::string_view to_string(RescueRangeState v) {
    return models::detail::enum_name(RescueRangeState_Table, v);
}

// ---------------------------------------------------------------------------
// Ordered case-insensitive map (models .NET Dictionary(StringComparer.
// OrdinalIgnoreCase) as serialized: JSON object, document order preserved).
// ---------------------------------------------------------------------------

template <typename TValue>
class OrderedFileMap {
public:
    // Ordered storage preserves the .NET Dictionary insertion order used for
    // byte-exact serialization; `index_` (folded key -> position) keeps
    // find/set/remove O(1) so building a large map is O(n), not O(n^2). This
    // matters: a whole-drive journal can hold 100k+ entries.
    std::vector<std::pair<std::string, TValue>> entries;

    TValue* find(std::string_view key) noexcept {
        auto it = index_.find(fold_key(key));
        return it == index_.end() ? nullptr : &entries[it->second].second;
    }

    const TValue* find(std::string_view key) const noexcept {
        auto it = index_.find(fold_key(key));
        return it == index_.end() ? nullptr : &entries[it->second].second;
    }

    // Dictionary-style upsert: replaces the value under an existing key
    // (keeping its position) or appends a new pair.
    void set(std::string key, TValue value) {
        std::string folded = fold_key(key);
        auto it = index_.find(folded);
        if (it != index_.end()) {
            entries[it->second].second = std::move(value);
            return;
        }
        index_.emplace(std::move(folded), entries.size());
        entries.emplace_back(std::move(key), std::move(value));
    }

    bool remove(std::string_view key) {
        auto it = index_.find(fold_key(key));
        if (it == index_.end()) return false;
        entries.erase(entries.begin() + static_cast<std::ptrdiff_t>(it->second));
        rebuild_index(); // positions after the erased entry shifted; removals are rare
        return true;
    }

    // Replace the whole backing vector (e.g. after external pruning) and
    // resynchronize the index. Callers that mutate `entries` directly must
    // follow with rebuild_index() or use this.
    void replace_entries(std::vector<std::pair<std::string, TValue>> new_entries) {
        entries = std::move(new_entries);
        rebuild_index();
    }

    void rebuild_index() {
        index_.clear();
        index_.reserve(entries.size());
        for (std::size_t i = 0; i < entries.size(); ++i) {
            index_.emplace(fold_key(entries[i].first), i);
        }
    }

    std::size_t size() const noexcept { return entries.size(); }
    bool empty() const noexcept { return entries.empty(); }

private:
    std::unordered_map<std::string, std::size_t> index_;

    // Matches models::detail::equals_ignore_case (ASCII a-z/A-Z fold) so keys
    // that compare equal there land in the same bucket here.
    static std::string fold_key(std::string_view key) {
        std::string out(key);
        for (char& c : out) {
            if (c >= 'A' && c <= 'Z') c = static_cast<char>(c - 'A' + 'a');
        }
        return out;
    }
};

// ---------------------------------------------------------------------------
// ByteRange / RescueRange
// ---------------------------------------------------------------------------

struct ByteRange {
    std::int64_t Offset = 0;
    std::int64_t Length = 0;

    void to_json(json::Writer& w) const {
        w.begin_object();
        w.key("Offset"); w.value(Offset);
        w.key("Length"); w.value(Length);
        w.end_object();
    }

    static ByteRange from_json(const json::Value& value) {
        ByteRange r;
        const auto* obj = value.as_object();
        if (obj == nullptr) return r;
        models::detail::read(obj, "Offset", r.Offset);
        models::detail::read(obj, "Length", r.Length);
        return r;
    }
};

struct RescueRange {
    std::int64_t Offset = 0;
    std::int64_t Length = 0;
    RescueRangeState State = RescueRangeState::Pending;

    void to_json(json::Writer& w) const {
        w.begin_object();
        w.key("Offset"); w.value(Offset);
        w.key("Length"); w.value(Length);
        w.key("State"); w.value(to_string(State));
        w.end_object();
    }

    static RescueRange from_json(const json::Value& value) {
        RescueRange r;
        const auto* obj = value.as_object();
        if (obj == nullptr) return r;
        models::detail::read(obj, "Offset", r.Offset);
        models::detail::read(obj, "Length", r.Length);
        r.State = models::detail::enum_parse(RescueRangeState_Table, obj->find("State"), r.State);
        return r;
    }
};

namespace detail {

template <typename TRange>
void write_range_array(json::Writer& w, std::string_view name, const std::vector<TRange>& ranges) {
    w.key(name);
    w.begin_array();
    for (const auto& range : ranges) range.to_json(w);
    w.end_array();
}

template <typename TRange>
void read_range_array(const json::Object* obj, std::string_view name, std::vector<TRange>& target) {
    if (const auto* v = obj->find(name); v != nullptr && v->is_array()) {
        target.clear();
        for (const auto& item : *v->as_array()) {
            target.push_back(TRange::from_json(item));
        }
    }
}

} // namespace detail

// ---------------------------------------------------------------------------
// JournalFileEntry / JobJournal
// ---------------------------------------------------------------------------

struct JournalFileEntry {
    std::string RelativePath;
    std::int64_t SourceLength = 0;
    std::int64_t SourceLastWriteUtcTicks = 0;
    std::int64_t SourceChangeUtcTicks = 0;
    std::int64_t SourceFileIndex = 0;
    std::int64_t SourceVolumeSerial = 0;
    std::int64_t BytesCopied = 0;
    FileCopyState State = FileCopyState::Pending;
    std::string LastError;
    bool DoNotRetry = false;
    std::vector<ByteRange> RecoveredRanges;
    std::vector<RescueRange> RescueRanges;
    std::string LastRescuePass;

    // Default-valued members are omitted, and RelativePath is omitted when it
    // just repeats the key this entry is stored under (JobJournal passes that
    // key in; standalone callers pass nothing and get the full shape).
    //
    // This is a writer-side change only: every models::detail::read overload
    // leaves its target at the member initializer when the key is absent, and
    // System.Text.Json does the same, so both the native and .NET readers
    // reconstruct an identical object. normalize_journal already restores
    // RelativePath from the key. It matters because a whole-drive journal is
    // dominated by defaults — on a real 196k-entry, 100 MB journal the omitted
    // fields plus the duplicated path account for 42% of the file.
    void to_json(json::Writer& w, std::string_view map_key = {}) const {
        w.begin_object();
        if (map_key.empty() || RelativePath != map_key) {
            w.key("RelativePath"); w.value(RelativePath);
        }
        w.key("SourceLength"); w.value(SourceLength);
        w.key("SourceLastWriteUtcTicks"); w.value(SourceLastWriteUtcTicks);
        if (SourceChangeUtcTicks != 0) {
            w.key("SourceChangeUtcTicks"); w.value(SourceChangeUtcTicks);
        }
        if (SourceFileIndex != 0) {
            w.key("SourceFileIndex"); w.value(SourceFileIndex);
        }
        if (SourceVolumeSerial != 0) {
            w.key("SourceVolumeSerial"); w.value(SourceVolumeSerial);
        }
        w.key("BytesCopied"); w.value(BytesCopied);
        if (State != FileCopyState::Pending) {
            w.key("State"); w.value(to_string(State));
        }
        if (!LastError.empty()) {
            w.key("LastError"); w.value(LastError);
        }
        if (DoNotRetry) {
            w.key("DoNotRetry"); w.value(DoNotRetry);
        }
        if (!RecoveredRanges.empty()) {
            detail::write_range_array(w, "RecoveredRanges", RecoveredRanges);
        }
        if (!RescueRanges.empty()) {
            detail::write_range_array(w, "RescueRanges", RescueRanges);
        }
        if (!LastRescuePass.empty()) {
            w.key("LastRescuePass"); w.value(LastRescuePass);
        }
        w.end_object();
    }

    static JournalFileEntry from_json(const json::Value& value) {
        JournalFileEntry e;
        const auto* obj = value.as_object();
        if (obj == nullptr) return e;
        models::detail::read(obj, "RelativePath", e.RelativePath);
        models::detail::read(obj, "SourceLength", e.SourceLength);
        models::detail::read(obj, "SourceLastWriteUtcTicks", e.SourceLastWriteUtcTicks);
        models::detail::read(obj, "SourceChangeUtcTicks", e.SourceChangeUtcTicks);
        models::detail::read(obj, "SourceFileIndex", e.SourceFileIndex);
        models::detail::read(obj, "SourceVolumeSerial", e.SourceVolumeSerial);
        models::detail::read(obj, "BytesCopied", e.BytesCopied);
        e.State = models::detail::enum_parse(FileCopyState_Table, obj->find("State"), e.State);
        models::detail::read(obj, "LastError", e.LastError);
        models::detail::read(obj, "DoNotRetry", e.DoNotRetry);
        detail::read_range_array(obj, "RecoveredRanges", e.RecoveredRanges);
        detail::read_range_array(obj, "RescueRanges", e.RescueRanges);
        models::detail::read(obj, "LastRescuePass", e.LastRescuePass);
        return e;
    }
};

struct JobJournal {
    std::string JobId;
    std::string SourceRoot;
    std::string DestinationRoot;
    DateTimeOffset CreatedUtc = DateTimeOffset::now_utc();
    DateTimeOffset UpdatedUtc = DateTimeOffset::now_utc();
    OrderedFileMap<JournalFileEntry> Files;

    void to_json(json::Writer& w) const {
        w.begin_object();
        w.key("JobId"); w.value(JobId);
        w.key("SourceRoot"); w.value(SourceRoot);
        w.key("DestinationRoot"); w.value(DestinationRoot);
        w.key("CreatedUtc"); w.value_literal_string(CreatedUtc.to_string());
        w.key("UpdatedUtc"); w.value_literal_string(UpdatedUtc.to_string());
        w.key("Files");
        w.begin_object();
        for (const auto& [key, entry] : Files.entries) {
            w.key(key);
            entry.to_json(w, key);
        }
        w.end_object();
        w.end_object();
    }

    static JobJournal from_json(const json::Value& value) {
        JobJournal j;
        const auto* obj = value.as_object();
        if (obj == nullptr) return j;
        models::detail::read(obj, "JobId", j.JobId);
        models::detail::read(obj, "SourceRoot", j.SourceRoot);
        models::detail::read(obj, "DestinationRoot", j.DestinationRoot);
        models::detail::read(obj, "CreatedUtc", j.CreatedUtc);
        models::detail::read(obj, "UpdatedUtc", j.UpdatedUtc);
        if (const auto* files = obj->find("Files"); files != nullptr && files->is_object()) {
            for (const auto& [key, entry_value] : files->as_object()->members) {
                j.Files.set(key, JournalFileEntry::from_json(entry_value));
            }
        }
        return j;
    }
};

// ---------------------------------------------------------------------------
// BadRangeMapFileEntry / BadRangeMap
// ---------------------------------------------------------------------------

struct BadRangeMapFileEntry {
    std::string RelativePath;
    std::int64_t SourceLength = 0;
    std::int64_t LastWriteUtcTicks = 0;
    std::int64_t SourceFileIndex = 0;
    std::int64_t SourceVolumeSerial = 0;
    std::string FileFingerprint;
    std::vector<ByteRange> BadRanges;
    std::int32_t ConfirmationCount = 0;
    DateTimeOffset LastScanUtc = DateTimeOffset::now_utc();
    std::string LastError;

    void to_json(json::Writer& w) const {
        w.begin_object();
        w.key("RelativePath"); w.value(RelativePath);
        w.key("SourceLength"); w.value(SourceLength);
        w.key("LastWriteUtcTicks"); w.value(LastWriteUtcTicks);
        if (SourceFileIndex != 0) {
            w.key("SourceFileIndex"); w.value(SourceFileIndex);
        }
        if (SourceVolumeSerial != 0) {
            w.key("SourceVolumeSerial"); w.value(SourceVolumeSerial);
        }
        w.key("FileFingerprint"); w.value(FileFingerprint);
        detail::write_range_array(w, "BadRanges", BadRanges);
        if (ConfirmationCount != 0) {
            w.key("ConfirmationCount"); w.value(ConfirmationCount);
        }
        w.key("LastScanUtc"); w.value_literal_string(LastScanUtc.to_string());
        w.key("LastError"); w.value(LastError);
        w.end_object();
    }

    static BadRangeMapFileEntry from_json(const json::Value& value) {
        BadRangeMapFileEntry e;
        const auto* obj = value.as_object();
        if (obj == nullptr) return e;
        models::detail::read(obj, "RelativePath", e.RelativePath);
        models::detail::read(obj, "SourceLength", e.SourceLength);
        models::detail::read(obj, "LastWriteUtcTicks", e.LastWriteUtcTicks);
        models::detail::read(obj, "SourceFileIndex", e.SourceFileIndex);
        models::detail::read(obj, "SourceVolumeSerial", e.SourceVolumeSerial);
        models::detail::read(obj, "FileFingerprint", e.FileFingerprint);
        detail::read_range_array(obj, "BadRanges", e.BadRanges);
        models::detail::read(obj, "ConfirmationCount", e.ConfirmationCount);
        models::detail::read(obj, "LastScanUtc", e.LastScanUtc);
        models::detail::read(obj, "LastError", e.LastError);
        return e;
    }
};

struct BadRangeMap {
    std::int32_t SchemaVersion = 1;
    std::string SourceRoot;
    std::string SourceIdentity;
    DateTimeOffset UpdatedUtc = DateTimeOffset::now_utc();
    OrderedFileMap<BadRangeMapFileEntry> Files;

    void to_json(json::Writer& w) const {
        w.begin_object();
        w.key("SchemaVersion"); w.value(SchemaVersion);
        w.key("SourceRoot"); w.value(SourceRoot);
        w.key("SourceIdentity"); w.value(SourceIdentity);
        w.key("UpdatedUtc"); w.value_literal_string(UpdatedUtc.to_string());
        w.key("Files");
        w.begin_object();
        for (const auto& [key, entry] : Files.entries) {
            w.key(key);
            entry.to_json(w);
        }
        w.end_object();
        w.end_object();
    }

    static BadRangeMap from_json(const json::Value& value) {
        BadRangeMap m;
        const auto* obj = value.as_object();
        if (obj == nullptr) return m;
        models::detail::read(obj, "SchemaVersion", m.SchemaVersion);
        models::detail::read(obj, "SourceRoot", m.SourceRoot);
        models::detail::read(obj, "SourceIdentity", m.SourceIdentity);
        models::detail::read(obj, "UpdatedUtc", m.UpdatedUtc);
        if (const auto* files = obj->find("Files"); files != nullptr && files->is_object()) {
            for (const auto& [key, entry_value] : files->as_object()->members) {
                m.Files.set(key, BadRangeMapFileEntry::from_json(entry_value));
            }
        }
        return m;
    }
};

} // namespace xact::storage
