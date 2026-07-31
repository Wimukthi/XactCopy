// -----------------------------------------------------------------------------
// File: src\core\dotnet_time.h
// Purpose: DateTimeOffset and TimeSpan value types matching .NET semantics and
//          the System.Text.Json wire formats used by the XactCopy IPC protocol
//          (ISO 8601 round-trip timestamps, constant "c" format durations).
// -----------------------------------------------------------------------------

#pragma once

#include <charconv>
#include <cstdint>
#include <optional>
#include <string>
#include <string_view>

#ifndef NOMINMAX
#define NOMINMAX
#endif
#ifndef WIN32_LEAN_AND_MEAN
#define WIN32_LEAN_AND_MEAN
#endif
#include <windows.h>

namespace xact::time {

inline constexpr std::int64_t TicksPerSecond = 10'000'000LL;
inline constexpr std::int64_t TicksPerMinute = TicksPerSecond * 60;
inline constexpr std::int64_t TicksPerHour = TicksPerMinute * 60;
inline constexpr std::int64_t TicksPerDay = TicksPerHour * 24;
inline constexpr std::int64_t TicksPerMillisecond = 10'000LL;

// Offset between the Windows FILETIME epoch (1601-01-01) and the .NET tick
// epoch (0001-01-01), in 100 ns ticks.
inline constexpr std::int64_t FileTimeEpochTicks = 504'911'232'000'000'000LL;

// ---------------------------------------------------------------------------
// TimeSpan: 100 ns tick count, .NET constant ("c") string format.
// ---------------------------------------------------------------------------

struct TimeSpan {
    std::int64_t ticks = 0;

    constexpr TimeSpan() = default;
    constexpr explicit TimeSpan(std::int64_t tick_count) : ticks(tick_count) {}

    static constexpr TimeSpan zero() { return TimeSpan(0); }
    static constexpr TimeSpan from_milliseconds(std::int64_t ms) { return TimeSpan(ms * TicksPerMillisecond); }
    static constexpr TimeSpan from_seconds(std::int64_t s) { return TimeSpan(s * TicksPerSecond); }
    static constexpr TimeSpan from_minutes(std::int64_t m) { return TimeSpan(m * TicksPerMinute); }

    constexpr double total_seconds() const { return static_cast<double>(ticks) / TicksPerSecond; }
    constexpr double total_milliseconds() const { return static_cast<double>(ticks) / TicksPerMillisecond; }

    constexpr bool operator==(const TimeSpan&) const = default;
    constexpr auto operator<=>(const TimeSpan&) const = default;

    // .NET "c" format: [-][d.]hh:mm:ss[.fffffff] (fraction printed as a full
    // 7-digit field whenever non-zero).
    std::string to_string() const {
        std::int64_t remaining = ticks;
        std::string out;
        if (remaining < 0) {
            out.push_back('-');
            remaining = -remaining;
        }

        const std::int64_t days = remaining / TicksPerDay;
        remaining %= TicksPerDay;
        const std::int64_t hours = remaining / TicksPerHour;
        remaining %= TicksPerHour;
        const std::int64_t minutes = remaining / TicksPerMinute;
        remaining %= TicksPerMinute;
        const std::int64_t seconds = remaining / TicksPerSecond;
        const std::int64_t fraction = remaining % TicksPerSecond;

        if (days != 0) {
            out.append(std::to_string(days));
            out.push_back('.');
        }
        append_padded(out, hours, 2);
        out.push_back(':');
        append_padded(out, minutes, 2);
        out.push_back(':');
        append_padded(out, seconds, 2);
        if (fraction != 0) {
            out.push_back('.');
            append_padded(out, fraction, 7);
        }
        return out;
    }

    // Parses the "c" grammar emitted by both sides: [-][d.]hh:mm:ss[.f{1,7}]
    // ("hh:mm" without seconds is also accepted, as TimeSpan.Parse allows it).
    static std::optional<TimeSpan> parse(std::string_view text) {
        std::size_t pos = 0;
        bool negative = false;
        if (pos < text.size() && text[pos] == '-') {
            negative = true;
            ++pos;
        }

        std::int64_t first = 0;
        if (!read_digits(text, pos, first)) return std::nullopt;

        std::int64_t days = 0;
        std::int64_t hours;
        if (pos < text.size() && text[pos] == '.') {
            ++pos;
            days = first;
            if (!read_digits(text, pos, hours)) return std::nullopt;
        } else {
            hours = first;
        }

        if (pos >= text.size() || text[pos] != ':') return std::nullopt;
        ++pos;
        std::int64_t minutes = 0;
        if (!read_digits(text, pos, minutes)) return std::nullopt;

        std::int64_t seconds = 0;
        std::int64_t fraction_ticks = 0;
        if (pos < text.size() && text[pos] == ':') {
            ++pos;
            if (!read_digits(text, pos, seconds)) return std::nullopt;
            if (pos < text.size() && text[pos] == '.') {
                ++pos;
                int digits = 0;
                while (pos < text.size() && text[pos] >= '0' && text[pos] <= '9') {
                    if (digits < 7) {
                        fraction_ticks = fraction_ticks * 10 + (text[pos] - '0');
                        ++digits;
                    }
                    ++pos;
                }
                if (digits == 0) return std::nullopt;
                for (; digits < 7; ++digits) fraction_ticks *= 10;
            }
        }

        if (pos != text.size()) return std::nullopt;
        if (hours > 23 || minutes > 59 || seconds > 59) return std::nullopt;

        std::int64_t total = days * TicksPerDay + hours * TicksPerHour +
                             minutes * TicksPerMinute + seconds * TicksPerSecond + fraction_ticks;
        return TimeSpan(negative ? -total : total);
    }

private:
    static void append_padded(std::string& out, std::int64_t value, int width) {
        char buffer[24];
        auto result = std::to_chars(buffer, buffer + sizeof(buffer), value);
        int length = static_cast<int>(result.ptr - buffer);
        for (int i = length; i < width; ++i) out.push_back('0');
        out.append(buffer, result.ptr);
    }

    static bool read_digits(std::string_view text, std::size_t& pos, std::int64_t& value) {
        std::size_t start = pos;
        value = 0;
        while (pos < text.size() && text[pos] >= '0' && text[pos] <= '9') {
            value = value * 10 + (text[pos] - '0');
            ++pos;
        }
        return pos > start;
    }
};

// ---------------------------------------------------------------------------
// DateTimeOffset: tick count since 0001-01-01 (local to the stored offset)
// plus the offset itself in minutes, exactly like .NET.
// ---------------------------------------------------------------------------

struct DateTimeOffset {
    // Clock reading in the timestamp's own time zone (offset already applied),
    // in 100 ns ticks since 0001-01-01T00:00:00. UTC instant = ticks - offset.
    std::int64_t ticks = 0;
    std::int32_t offset_minutes = 0;

    constexpr DateTimeOffset() = default;
    constexpr DateTimeOffset(std::int64_t local_ticks, std::int32_t offset_min)
        : ticks(local_ticks), offset_minutes(offset_min) {}

    static constexpr DateTimeOffset min_value() { return DateTimeOffset(0, 0); }

    constexpr std::int64_t utc_ticks() const { return ticks - offset_minutes * TicksPerMinute; }

    // Equality matches .NET DateTimeOffset semantics: same instant, regardless
    // of offset representation.
    constexpr bool operator==(const DateTimeOffset& other) const { return utc_ticks() == other.utc_ticks(); }
    constexpr bool operator<(const DateTimeOffset& other) const { return utc_ticks() < other.utc_ticks(); }

    static DateTimeOffset now_utc() {
        FILETIME ft{};
        GetSystemTimePreciseAsFileTime(&ft);
        ULARGE_INTEGER value;
        value.LowPart = ft.dwLowDateTime;
        value.HighPart = ft.dwHighDateTime;
        return DateTimeOffset(static_cast<std::int64_t>(value.QuadPart) + FileTimeEpochTicks, 0);
    }

    // System.Text.Json DateTimeOffset format: yyyy-MM-ddTHH:mm:ss[.fffffff]±hh:mm
    // with the fractional part omitted when zero and trailing zeros trimmed.
    std::string to_string() const {
        std::int64_t remaining = ticks;
        if (remaining < 0) remaining = 0;

        const std::int64_t day_number = remaining / TicksPerDay;
        std::int64_t time_ticks = remaining % TicksPerDay;

        int year, month, day;
        civil_from_days(day_number, year, month, day);

        const int hours = static_cast<int>(time_ticks / TicksPerHour);
        time_ticks %= TicksPerHour;
        const int minutes = static_cast<int>(time_ticks / TicksPerMinute);
        time_ticks %= TicksPerMinute;
        const int seconds = static_cast<int>(time_ticks / TicksPerSecond);
        std::int64_t fraction = time_ticks % TicksPerSecond;

        std::string out;
        out.reserve(35);
        append_fixed(out, year, 4);
        out.push_back('-');
        append_fixed(out, month, 2);
        out.push_back('-');
        append_fixed(out, day, 2);
        out.push_back('T');
        append_fixed(out, hours, 2);
        out.push_back(':');
        append_fixed(out, minutes, 2);
        out.push_back(':');
        append_fixed(out, seconds, 2);

        if (fraction != 0) {
            char digits[7];
            for (int i = 6; i >= 0; --i) {
                digits[i] = static_cast<char>('0' + fraction % 10);
                fraction /= 10;
            }
            int keep = 7;
            while (keep > 1 && digits[keep - 1] == '0') --keep;
            out.push_back('.');
            out.append(digits, static_cast<std::size_t>(keep));
        }

        int offset = offset_minutes;
        if (offset < 0) {
            out.push_back('-');
            offset = -offset;
        } else {
            out.push_back('+');
        }
        append_fixed(out, offset / 60, 2);
        out.push_back(':');
        append_fixed(out, offset % 60, 2);
        return out;
    }

    // Accepts yyyy-MM-ddTHH:mm:ss[.f+][Z|±hh[:mm]]; absent zone means offset 0
    // (only ever produced for DateTimeOffset.MinValue-style sentinels).
    static std::optional<DateTimeOffset> parse(std::string_view text) {
        std::size_t pos = 0;
        int year, month, day, hours, minutes, seconds;
        if (!read_fixed(text, pos, 4, year)) return std::nullopt;
        if (!consume(text, pos, '-')) return std::nullopt;
        if (!read_fixed(text, pos, 2, month)) return std::nullopt;
        if (!consume(text, pos, '-')) return std::nullopt;
        if (!read_fixed(text, pos, 2, day)) return std::nullopt;
        if (!consume(text, pos, 'T')) return std::nullopt;
        if (!read_fixed(text, pos, 2, hours)) return std::nullopt;
        if (!consume(text, pos, ':')) return std::nullopt;
        if (!read_fixed(text, pos, 2, minutes)) return std::nullopt;
        if (!consume(text, pos, ':')) return std::nullopt;
        if (!read_fixed(text, pos, 2, seconds)) return std::nullopt;

        std::int64_t fraction_ticks = 0;
        if (pos < text.size() && text[pos] == '.') {
            ++pos;
            int digits = 0;
            while (pos < text.size() && text[pos] >= '0' && text[pos] <= '9') {
                if (digits < 7) {
                    fraction_ticks = fraction_ticks * 10 + (text[pos] - '0');
                    ++digits;
                }
                ++pos;
            }
            if (digits == 0) return std::nullopt;
            for (; digits < 7; ++digits) fraction_ticks *= 10;
        }

        std::int32_t offset_min = 0;
        if (pos < text.size()) {
            char zone = text[pos];
            if (zone == 'Z' || zone == 'z') {
                ++pos;
            } else if (zone == '+' || zone == '-') {
                ++pos;
                int offset_hours = 0, offset_minutes_part = 0;
                if (!read_fixed(text, pos, 2, offset_hours)) return std::nullopt;
                if (pos < text.size() && text[pos] == ':') {
                    ++pos;
                    if (!read_fixed(text, pos, 2, offset_minutes_part)) return std::nullopt;
                } else if (pos + 1 < text.size() || pos < text.size()) {
                    if (pos + 1 < text.size() && is_digit(text[pos]) && is_digit(text[pos + 1])) {
                        read_fixed(text, pos, 2, offset_minutes_part);
                    }
                }
                offset_min = offset_hours * 60 + offset_minutes_part;
                if (zone == '-') offset_min = -offset_min;
            } else {
                return std::nullopt;
            }
        }
        if (pos != text.size()) return std::nullopt;

        if (month < 1 || month > 12 || day < 1 || day > 31) return std::nullopt;
        if (hours > 23 || minutes > 59 || seconds > 60) return std::nullopt;
        if (seconds == 60) seconds = 59; // Leap seconds are clamped, as in .NET.

        std::int64_t day_number = days_from_civil(year, month, day);
        std::int64_t local_ticks = day_number * TicksPerDay + hours * TicksPerHour +
                                   minutes * TicksPerMinute + seconds * TicksPerSecond + fraction_ticks;
        return DateTimeOffset(local_ticks, offset_min);
    }

private:
    static bool is_digit(char c) { return c >= '0' && c <= '9'; }

    static bool consume(std::string_view text, std::size_t& pos, char expected) {
        if (pos < text.size() && text[pos] == expected) { ++pos; return true; }
        return false;
    }

    static bool read_fixed(std::string_view text, std::size_t& pos, int width, int& value) {
        if (pos + static_cast<std::size_t>(width) > text.size()) return false;
        value = 0;
        for (int i = 0; i < width; ++i) {
            char c = text[pos + static_cast<std::size_t>(i)];
            if (!is_digit(c)) return false;
            value = value * 10 + (c - '0');
        }
        pos += static_cast<std::size_t>(width);
        return true;
    }

    static void append_fixed(std::string& out, int value, int width) {
        char buffer[12];
        auto result = std::to_chars(buffer, buffer + sizeof(buffer), value);
        int length = static_cast<int>(result.ptr - buffer);
        for (int i = length; i < width; ++i) out.push_back('0');
        out.append(buffer, result.ptr);
    }

    // Proleptic Gregorian calendar conversion (Howard Hinnant's algorithms),
    // shifted so day 0 = 0001-01-01 to match .NET ticks.
    static std::int64_t days_from_civil(int y, int m, int d) {
        y -= m <= 2;
        const std::int64_t era = (y >= 0 ? y : y - 399) / 400;
        const unsigned yoe = static_cast<unsigned>(y - era * 400);
        const unsigned doy = (153u * static_cast<unsigned>(m + (m > 2 ? -3 : 9)) + 2u) / 5u +
                             static_cast<unsigned>(d) - 1u;
        const unsigned doe = yoe * 365u + yoe / 4u - yoe / 100u + doy;
        return era * 146097LL + static_cast<std::int64_t>(doe) - 719468LL + 719162LL;
    }

    static void civil_from_days(std::int64_t days_since_0001, int& y, int& m, int& d) {
        std::int64_t z = days_since_0001 - 719162LL + 719468LL;
        const std::int64_t era = (z >= 0 ? z : z - 146096) / 146097;
        const unsigned doe = static_cast<unsigned>(z - era * 146097);
        const unsigned yoe = (doe - doe / 1460 + doe / 36524 - doe / 146096) / 365;
        const std::int64_t year = static_cast<std::int64_t>(yoe) + era * 400;
        const unsigned doy = doe - (365 * yoe + yoe / 4 - yoe / 100);
        const unsigned mp = (5 * doy + 2) / 153;
        d = static_cast<int>(doy - (153 * mp + 2) / 5 + 1);
        m = static_cast<int>(mp + (mp < 10 ? 3 : -9));
        y = static_cast<int>(year + (m <= 2 ? 1 : 0));
    }
};

} // namespace xact::time
