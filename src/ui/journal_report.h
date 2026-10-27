#pragma once

#include <algorithm>
#include <cstdint>
#include <limits>
#include <ostream>
#include <string>
#include <string_view>
#include <vector>

#include "../storage/storage_models.h"

namespace xact::ui {

inline void write_csv_field(std::ostream& out, std::string_view value) {
    const bool quoted = value.find_first_of(",\"\r\n") != std::string_view::npos;
    if (quoted) out.put('"');
    for (char c : value) {
        if (c == '"') out.put('"');
        out.put(c);
    }
    if (quoted) out.put('"');
}

inline std::vector<storage::ByteRange> merge_report_ranges(std::vector<storage::ByteRange> ranges,
                                                            std::int64_t file_size) {
    std::vector<storage::ByteRange> valid;
    for (const auto& range : ranges) {
        if (range.Offset < 0 || range.Length <= 0 || range.Offset >= file_size) continue;
        valid.push_back({range.Offset, std::min(range.Length, file_size - range.Offset)});
    }
    std::sort(valid.begin(), valid.end(), [](const auto& a, const auto& b) {
        return a.Offset < b.Offset || (a.Offset == b.Offset && a.Length < b.Length);
    });
    std::vector<storage::ByteRange> merged;
    for (const auto& range : valid) {
        if (!merged.empty() && range.Offset <= merged.back().Offset + merged.back().Length) {
            const auto end = std::max(merged.back().Offset + merged.back().Length,
                                      range.Offset + range.Length);
            merged.back().Length = end - merged.back().Offset;
        } else merged.push_back(range);
    }
    return merged;
}

inline std::int64_t report_range_bytes(const std::vector<storage::ByteRange>& ranges) {
    std::int64_t total = 0;
    for (const auto& range : ranges) total += range.Length;
    return total;
}

inline std::string report_ranges_text(const std::vector<storage::ByteRange>& ranges) {
    std::string text;
    for (const auto& range : ranges) {
        if (!text.empty()) text += "; ";
        text += std::to_string(range.Offset) + ":" + std::to_string(range.Length);
    }
    return text;
}

inline void write_journal_report_csv(std::ostream& out, const storage::JobJournal& journal) {
    out.write("\xEF\xBB\xBF", 3);
    out << "Relative path,State,Size,Bytes copied,Last error,Unreadable bytes,Recovered bytes,"
           "Bad ranges (offset:length),Recovered ranges (offset:length),"
           "Synthetic fill ranges (offset:length),Legacy ranges (offset:length)\r\n";
    const bool scan = journal.RunOptions.has_value() &&
                      journal.RunOptions->OperationMode == models::JobOperationMode::ScanOnly;
    for (const auto& [key, entry] : journal.Files.entries) {
        std::vector<storage::ByteRange> bad, recovered, synthetic, legacy;
        for (const auto& range : entry.RescueRanges) {
            if (range.State == storage::RescueRangeState::Bad ||
                range.State == storage::RescueRangeState::KnownBad)
                bad.push_back({range.Offset, range.Length});
            else if (range.State == storage::RescueRangeState::Recovered)
                recovered.push_back({range.Offset, range.Length});
        }
        // RecoveredRanges means unreadable scan ranges in scans, but synthetic
        // fill in copies. Old journals without options cannot disambiguate it.
        if (journal.RunOptions.has_value()) {
            if (scan) bad.insert(bad.end(), entry.RecoveredRanges.begin(), entry.RecoveredRanges.end());
            else {
                synthetic = entry.RecoveredRanges;
                recovered.insert(recovered.end(), synthetic.begin(), synthetic.end());
            }
        } else legacy = entry.RecoveredRanges;
        const auto size = std::max<std::int64_t>(0, entry.SourceLength);
        bad = merge_report_ranges(std::move(bad), size);
        recovered = merge_report_ranges(std::move(recovered), size);
        synthetic = merge_report_ranges(std::move(synthetic), size);
        legacy = merge_report_ranges(std::move(legacy), size);
        auto unreadable = bad;
        unreadable.insert(unreadable.end(), synthetic.begin(), synthetic.end());
        unreadable = merge_report_ranges(std::move(unreadable), size);
        write_csv_field(out, entry.RelativePath.empty() ? key : entry.RelativePath); out.put(',');
        write_csv_field(out, storage::to_string(entry.State)); out.put(',');
        out << entry.SourceLength << ',' << entry.BytesCopied << ',';
        write_csv_field(out, entry.LastError); out.put(',');
        out << report_range_bytes(unreadable) << ',' << report_range_bytes(recovered) << ',';
        write_csv_field(out, report_ranges_text(bad)); out.put(',');
        write_csv_field(out, report_ranges_text(recovered)); out.put(',');
        write_csv_field(out, report_ranges_text(synthetic)); out.put(',');
        write_csv_field(out, report_ranges_text(legacy)); out << "\r\n";
    }
}

} // namespace xact::ui
