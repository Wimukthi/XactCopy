#pragma once

#include <algorithm>
#include <array>
#include <chrono>
#include <cstdint>
#include <cwctype>
#include <deque>
#include <memory>
#include <mutex>
#include <string>
#include <string_view>
#include <utility>

namespace xact::ui {

enum class LogSeverity { Normal, Error, Warning, Success, Supervisor };
enum class LogFilter { All, WarningsAndErrors, ErrorsOnly };

inline LogSeverity classify_log_line(const std::string& line) {
    std::string lower;
    lower.reserve(line.size());
    for (char c : line) {
        lower.push_back(c >= 'A' && c <= 'Z' ? static_cast<char>(c - 'A' + 'a') : c);
    }
    auto contains = [&lower](const char* fragment) { return lower.find(fragment) != std::string::npos; };
    if (contains("fatal") || contains("failed") || contains("error") || contains("mismatch"))
        return LogSeverity::Error;
    if (contains("recovered") || contains("salvag") || contains("succeeded") || contains("completed:"))
        return LogSeverity::Success;
    if (contains("retry") || contains("warning") || contains("stale") ||
        contains("skipped") || contains("waiting") || contains("fallback"))
        return LogSeverity::Warning;
    if (contains("[supervisor]") || contains("[devfault]") || contains("[workertelemetry]"))
        return LogSeverity::Supervisor;
    return LogSeverity::Normal;
}

struct LogEntry {
    std::wstring text;
    LogSeverity severity = LogSeverity::Normal;
    std::chrono::system_clock::time_point timestamp;
    std::uint64_t sequence = 0;
};

// Three index rings make filter changes and every append/eviction constant time.
class LogModel {
public:
    explicit LogModel(std::size_t capacity = 50000) : capacity_(capacity ? capacity : 1) {}

    struct Snapshot {
        std::uint64_t next = 0;
        std::uint64_t last = 0;
        bool empty = true;
        bool detached = false;
        std::deque<std::shared_ptr<const LogEntry>> evicted;
        std::deque<std::shared_ptr<const LogEntry>> retained;
    };

    void set_capacity(std::size_t capacity) {
        std::lock_guard<std::mutex> guard(mutex_);
        capacity_ = capacity ? capacity : 1;
        while (entries_.size() > capacity_) evict_front();
    }

    void clear() {
        std::lock_guard<std::mutex> guard(mutex_);
        // A reader can keep the old ring without copying it on the GUI thread.
        if (snapshot_ && !snapshot_->detached) {
            snapshot_->retained = std::move(entries_);
            entries_.clear();
            snapshot_->detached = true;
        } else entries_.clear();
        for (auto& view : views_) view.clear();
    }

    void append(std::wstring text, LogSeverity severity,
                std::chrono::system_clock::time_point timestamp = std::chrono::system_clock::now()) {
        std::lock_guard<std::mutex> guard(mutex_);
        if (entries_.size() == capacity_) evict_front();
        const auto id = next_sequence_++;
        entries_.push_back(std::make_shared<const LogEntry>(LogEntry{std::move(text), severity, timestamp, id}));
        views_[0].push_back(id);
        if (severity == LogSeverity::Warning || severity == LogSeverity::Error)
            views_[1].push_back(id);
        if (severity == LogSeverity::Error) views_[2].push_back(id);
    }

    std::size_t size(LogFilter filter = LogFilter::All) const {
        std::lock_guard<std::mutex> guard(mutex_);
        return views_[static_cast<std::size_t>(filter)].size();
    }

    std::shared_ptr<const LogEntry> at(LogFilter filter, std::size_t index) const {
        std::lock_guard<std::mutex> guard(mutex_);
        return at_locked(filter, index);
    }

    std::size_t index_of_sequence(LogFilter filter, std::uint64_t sequence) const {
        std::lock_guard<std::mutex> guard(mutex_);
        const auto& view = views_[static_cast<std::size_t>(filter)];
        const auto found = std::lower_bound(view.begin(), view.end(), sequence);
        return found == view.end() || *found != sequence ? view.size()
                                                          : static_cast<std::size_t>(found - view.begin());
    }

    std::shared_ptr<Snapshot> begin_snapshot() {
        std::lock_guard<std::mutex> guard(mutex_);
        if (snapshot_) return {};
        auto snapshot = std::make_shared<Snapshot>();
        if (!entries_.empty()) {
            snapshot->empty = false;
            snapshot->next = entries_.front()->sequence;
            snapshot->last = entries_.back()->sequence;
        }
        snapshot_ = snapshot;
        return snapshot;
    }

    std::shared_ptr<const LogEntry> next_snapshot(const std::shared_ptr<Snapshot>& snapshot) {
        std::lock_guard<std::mutex> guard(mutex_);
        if (!snapshot || snapshot_ != snapshot || snapshot->empty || snapshot->next > snapshot->last)
            return {};
        const auto sequence = snapshot->next++;
        auto take = [sequence](auto& queue) -> std::shared_ptr<const LogEntry> {
            while (!queue.empty() && queue.front()->sequence < sequence) queue.pop_front();
            if (queue.empty() || queue.front()->sequence != sequence) return {};
            auto result = std::move(queue.front());
            queue.pop_front();
            return result;
        };
        auto entry = take(snapshot->evicted);
        if (!entry && snapshot->detached) entry = take(snapshot->retained);
        if (!entry && !entries_.empty() && sequence >= entries_.front()->sequence &&
            sequence <= entries_.back()->sequence) {
            entry = entries_[static_cast<std::size_t>(sequence - entries_.front()->sequence)];
        }
        return entry;
    }

    void end_snapshot(const std::shared_ptr<Snapshot>& snapshot) {
        std::lock_guard<std::mutex> guard(mutex_);
        if (snapshot_ == snapshot) snapshot_.reset();
    }

    static bool matches_filter(LogSeverity severity, LogFilter filter) {
        return filter == LogFilter::All ||
               (filter == LogFilter::WarningsAndErrors &&
                (severity == LogSeverity::Warning || severity == LogSeverity::Error)) ||
               (filter == LogFilter::ErrorsOnly && severity == LogSeverity::Error);
    }

private:
    std::shared_ptr<const LogEntry> at_locked(LogFilter filter, std::size_t index) const {
        const auto& view = views_[static_cast<std::size_t>(filter)];
        if (index >= view.size() || entries_.empty()) return {};
        return entries_[static_cast<std::size_t>(view[index] - entries_.front()->sequence)];
    }

public:
    std::size_t find_next(LogFilter filter, std::wstring_view needle, std::size_t start) const {
        std::lock_guard<std::mutex> guard(mutex_);
        const auto count = views_[static_cast<std::size_t>(filter)].size();
        if (count == 0 || needle.empty()) return count;
        std::wstring folded(needle);
        for (auto& c : folded) c = static_cast<wchar_t>(std::towlower(c));
        for (std::size_t step = 0; step < count; ++step) {
            const std::size_t index = (start + step) % count;
            const auto& text = at_locked(filter, index)->text;
            auto found = std::search(text.begin(), text.end(), folded.begin(), folded.end(),
                [](wchar_t left, wchar_t right) { return std::towlower(left) == right; });
            if (found != text.end()) return index;
        }
        return count;
    }

private:
    void evict_front() {
        const auto entry = entries_.front();
        const auto id = entry->sequence;
        if (snapshot_ && !snapshot_->detached && !snapshot_->empty &&
            id >= snapshot_->next && id <= snapshot_->last) {
            snapshot_->evicted.push_back(entry);
        }
        for (auto& view : views_) {
            if (!view.empty() && view.front() == id) view.pop_front();
        }
        entries_.pop_front();
    }

    std::size_t capacity_;
    std::uint64_t next_sequence_ = 0;
    std::deque<std::shared_ptr<const LogEntry>> entries_;
    std::array<std::deque<std::uint64_t>, 3> views_;
    mutable std::mutex mutex_;
    std::shared_ptr<Snapshot> snapshot_;
};

} // namespace xact::ui
