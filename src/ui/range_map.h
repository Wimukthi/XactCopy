#pragma once

#include <algorithm>
#include <cmath>
#include <cstdint>
#include <vector>

#include "../storage/storage_models.h"

namespace xact::ui {

struct RangeMapSpan {
    int left = 0;
    int right = 0;
    storage::RescueRangeState state = storage::RescueRangeState::Pending;
    std::int64_t offset = 0;
    std::int64_t length = 0;
};

inline int range_map_priority(storage::RescueRangeState state) {
    switch (state) {
        case storage::RescueRangeState::Bad: return 5;
        case storage::RescueRangeState::KnownBad: return 4;
        case storage::RescueRangeState::Recovered: return 3;
        case storage::RescueRangeState::Pending: return 2;
        case storage::RescueRangeState::Good: return 1;
        default: return 0;
    }
}

inline std::vector<RangeMapSpan> range_map_spans(
    std::int64_t file_size, const std::vector<storage::RescueRange>& ranges, int width) {
    if (file_size <= 0 || width <= 0) return {};
    struct Pixel {
        storage::RescueRangeState state = storage::RescueRangeState::Pending;
        std::int64_t offset = 0;
        std::int64_t length = 0;
        int priority = 0;
    };
    std::vector<Pixel> pixels(static_cast<std::size_t>(width));
    for (int x = 0; x < width; ++x) {
        const auto first = std::min<std::int64_t>(file_size - 1, static_cast<std::int64_t>(
            static_cast<long double>(file_size) * x / width));
        const auto last = x == width - 1 ? file_size : static_cast<std::int64_t>(
            std::ceil(static_cast<long double>(file_size) * (x + 1) / width));
        pixels[static_cast<std::size_t>(x)].offset = first;
        pixels[static_cast<std::size_t>(x)].length =
            std::min(file_size, std::max(first + 1, last)) - first;
    }
    // Round every nonempty range outward to at least one pixel. When ranges
    // share a pixel, unreadable states take priority over successful coverage.
    for (const auto& range : ranges) {
        if (range.Offset < 0 || range.Length <= 0 || range.Offset >= file_size) continue;
        const auto end = range.Offset + std::min(range.Length, file_size - range.Offset);
        const int first = std::clamp(static_cast<int>(
            static_cast<long double>(range.Offset) * width / file_size), 0, width - 1);
        const int last = std::clamp(static_cast<int>(std::ceil(
            static_cast<long double>(end) * width / file_size)), first + 1, width);
        const int priority = range_map_priority(range.State);
        for (int x = first; x < last; ++x) {
            auto& pixel = pixels[static_cast<std::size_t>(x)];
            if (priority > pixel.priority) {
                pixel = {range.State, range.Offset, end - range.Offset, priority};
            } else if (priority == pixel.priority && pixel.state == range.State &&
                       range.Offset <= pixel.offset + pixel.length && end >= pixel.offset) {
                const auto combined_start = std::min(pixel.offset, range.Offset);
                const auto combined_end = std::max(pixel.offset + pixel.length, end);
                pixel.offset = combined_start;
                pixel.length = combined_end - combined_start;
            }
        }
    }
    std::vector<RangeMapSpan> spans;
    for (int x = 0; x < width; ++x) {
        const auto& pixel = pixels[static_cast<std::size_t>(x)];
        if (!spans.empty() && spans.back().state == pixel.state &&
            pixel.offset <= spans.back().offset + spans.back().length &&
            spans.back().offset <= pixel.offset + pixel.length) {
            auto& span = spans.back();
            const auto end = std::max(span.offset + span.length, pixel.offset + pixel.length);
            span.offset = std::min(span.offset, pixel.offset);
            span.length = end - span.offset;
            span.right = x + 1;
        } else {
            spans.push_back({x, x + 1, pixel.state, pixel.offset, pixel.length});
        }
    }
    return spans;
}

} // namespace xact::ui
