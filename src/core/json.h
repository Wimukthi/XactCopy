// -----------------------------------------------------------------------------
// File: src\core\json.h
// Purpose: Minimal JSON DOM, parser, and writer compatible with the JSON shape
//          produced by System.Text.Json in the .NET XactCopy build (compact
//          output, PascalCase keys, default JavaScriptEncoder escaping).
// -----------------------------------------------------------------------------

#pragma once

#include <cstdint>
#include <charconv>
#include <cmath>
#include <memory>
#include <limits>
#include <optional>
#include <stdexcept>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

namespace xact::json {

class Value;

using Array = std::vector<Value>;
using Member = std::pair<std::string, Value>;

// Object preserves insertion order (matches System.Text.Json property order)
// and offers case-insensitive lookup (matches PropertyNameCaseInsensitive=true).
class Object {
public:
    std::vector<Member> members;

    const Value* find(std::string_view key) const noexcept;
    void add(std::string key, Value value);
};

enum class Kind : std::uint8_t { Null, Boolean, Int64, Double, String, ArrayKind, ObjectKind };

class Value {
public:
    Value() : kind_(Kind::Null) {}
    explicit Value(std::nullptr_t) : kind_(Kind::Null) {}
    explicit Value(bool v) : kind_(Kind::Boolean), bool_(v) {}
    explicit Value(std::int64_t v) : kind_(Kind::Int64), int_(v) {}
    explicit Value(double v) : kind_(Kind::Double), double_(v) {}
    explicit Value(std::string v) : kind_(Kind::String), string_(std::move(v)) {}
    explicit Value(Array v) : kind_(Kind::ArrayKind), array_(std::make_shared<Array>(std::move(v))) {}
    explicit Value(Object v) : kind_(Kind::ObjectKind), object_(std::make_shared<Object>(std::move(v))) {}

    Kind kind() const noexcept { return kind_; }
    bool is_null() const noexcept { return kind_ == Kind::Null; }
    bool is_object() const noexcept { return kind_ == Kind::ObjectKind; }
    bool is_array() const noexcept { return kind_ == Kind::ArrayKind; }
    bool is_string() const noexcept { return kind_ == Kind::String; }
    bool is_number() const noexcept { return kind_ == Kind::Int64 || kind_ == Kind::Double; }
    bool is_bool() const noexcept { return kind_ == Kind::Boolean; }

    bool as_bool(bool fallback = false) const noexcept {
        return kind_ == Kind::Boolean ? bool_ : fallback;
    }

    std::int64_t as_int64(std::int64_t fallback = 0) const noexcept {
        if (kind_ == Kind::Int64) return int_;
        if (kind_ == Kind::Double) {
            // A direct floating-to-integer cast outside the destination range
            // is undefined behavior. JSON written by XactCopy uses integer
            // tokens for these fields; accept an exactly integral double only
            // when it is safely representable.
            constexpr double Int64Lower = -9223372036854775808.0;
            constexpr double Int64UpperExclusive = 9223372036854775808.0;
            if (!std::isfinite(double_) || double_ < Int64Lower ||
                double_ >= Int64UpperExclusive || std::trunc(double_) != double_) {
                return fallback;
            }
            return static_cast<std::int64_t>(double_);
        }
        return fallback;
    }

    std::int32_t as_int32(std::int32_t fallback = 0) const noexcept {
        const std::int64_t value = as_int64(fallback);
        if (value < std::numeric_limits<std::int32_t>::min() ||
            value > std::numeric_limits<std::int32_t>::max()) {
            return fallback;
        }
        return static_cast<std::int32_t>(value);
    }

    double as_double(double fallback = 0.0) const noexcept {
        if (kind_ == Kind::Double) return double_;
        if (kind_ == Kind::Int64) return static_cast<double>(int_);
        return fallback;
    }

    const std::string& as_string() const noexcept {
        static const std::string empty;
        return kind_ == Kind::String ? string_ : empty;
    }

    const Object* as_object() const noexcept {
        return kind_ == Kind::ObjectKind ? object_.get() : nullptr;
    }

    const Array* as_array() const noexcept {
        return kind_ == Kind::ArrayKind ? array_.get() : nullptr;
    }

private:
    Kind kind_;
    bool bool_ = false;
    std::int64_t int_ = 0;
    double double_ = 0.0;
    std::string string_;
    std::shared_ptr<Array> array_;
    std::shared_ptr<Object> object_;
};

inline const Value* Object::find(std::string_view key) const noexcept {
    // Exact match first (the common case for same-version peers), then the
    // case-insensitive fallback System.Text.Json applies during deserialization.
    for (const auto& m : members) {
        if (m.first == key) return &m.second;
    }
    for (const auto& m : members) {
        if (m.first.size() != key.size()) continue;
        bool equal = true;
        for (std::size_t i = 0; i < key.size(); ++i) {
            char a = m.first[i];
            char b = key[i];
            if (a >= 'A' && a <= 'Z') a = static_cast<char>(a - 'A' + 'a');
            if (b >= 'A' && b <= 'Z') b = static_cast<char>(b - 'A' + 'a');
            if (a != b) { equal = false; break; }
        }
        if (equal) return &m.second;
    }
    return nullptr;
}

inline void Object::add(std::string key, Value value) {
    members.emplace_back(std::move(key), std::move(value));
}

class ParseError : public std::runtime_error {
public:
    explicit ParseError(const std::string& message) : std::runtime_error(message) {}
};

// ---------------------------------------------------------------------------
// Writer
// ---------------------------------------------------------------------------

// Streaming writer that emits JSON matching System.Text.Json output:
// compact by default, or WriteIndented=true layout (2-space indent, "key": v,
// CRLF newlines, empty containers inline) when constructed with indented=true.
// String escaping is equivalent to JavaScriptEncoder.Default (short escapes
// for \b \t \n \f \r and \\, \uXXXX for '"', other control chars,
// '<' '>' '&' '\'' '+' '`', and all non-ASCII code points).
class Writer {
public:
    explicit Writer(bool indented = false, std::string_view newline = "\r\n")
        : indented_(indented), newline_(newline) {}

    std::string take() { return std::move(out_); }
    const std::string& text() const noexcept { return out_; }

    void begin_object() { separator(); out_.push_back('{'); stack_.push_back(true); }
    void end_object() { close_container('}'); }
    void begin_array() { separator(); out_.push_back('['); stack_.push_back(true); }
    void end_array() { close_container(']'); }

    void key(std::string_view name) {
        separator();
        write_string_literal(name);
        out_.push_back(':');
        if (indented_) out_.push_back(' ');
        pending_key_ = true;
    }

    void value(std::nullptr_t) { separator(); out_.append("null"); mark_value_written(); }
    void value(bool v) { separator(); out_.append(v ? "true" : "false"); mark_value_written(); }

    void value(std::int32_t v) { value(static_cast<std::int64_t>(v)); }

    void value(std::int64_t v) {
        separator();
        char buffer[24];
        auto result = std::to_chars(buffer, buffer + sizeof(buffer), v);
        out_.append(buffer, result.ptr);
        mark_value_written();
    }

    void value(std::uint32_t v) { value(static_cast<std::int64_t>(v)); }

    void value(double v) {
        separator();
        if (std::isnan(v) || std::isinf(v)) {
            // System.Text.Json throws on non-finite doubles by default; the
            // XactCopy models never produce them, but never emit invalid JSON.
            out_.append("0");
        } else {
            char buffer[32];
            auto result = std::to_chars(buffer, buffer + sizeof(buffer), v);
            std::string_view text(buffer, static_cast<std::size_t>(result.ptr - buffer));
            // .NET shortest round-trip format uses uppercase 'E' exponents.
            for (char c : text) {
                out_.push_back(c == 'e' ? 'E' : c);
            }
        }
        mark_value_written();
    }

    void value(std::string_view v) {
        separator();
        write_string_literal(v);
        mark_value_written();
    }

    void value(const char* v) { value(std::string_view(v)); }

    // Writes a string value without encoder escaping. System.Text.Json emits
    // DateTimeOffset/TimeSpan values through a formatting fast path that skips
    // the JavaScriptEncoder, so '+' in "+00:00" stays literal. Only safe for
    // content with no quotes, backslashes, or control characters.
    void value_literal_string(std::string_view v) {
        separator();
        out_.push_back('"');
        out_.append(v);
        out_.push_back('"');
        mark_value_written();
    }

    // Emits a pre-serialized JSON fragment verbatim (used for nested payloads).
    void raw(std::string_view fragment) {
        separator();
        out_.append(fragment);
        mark_value_written();
    }

private:
    std::string out_;
    std::vector<bool> stack_;   // true = next element is the first in container
    bool pending_key_ = false;
    bool indented_ = false;
    std::string newline_;

    void separator() {
        if (pending_key_) {
            // Value directly after "key": — no newline or comma.
            pending_key_ = false;
            return;
        }
        if (!stack_.empty()) {
            if (stack_.back()) {
                stack_.back() = false;
            } else {
                out_.push_back(',');
            }
            if (indented_) {
                out_.append(newline_);
                out_.append(stack_.size() * 2, ' ');
            }
        }
    }

    void close_container(char closer) {
        const bool had_content = !stack_.empty() && !stack_.back();
        stack_.pop_back();
        if (indented_ && had_content) {
            out_.append(newline_);
            out_.append(stack_.size() * 2, ' ');
        }
        out_.push_back(closer);
        mark_value_written();
    }

    void mark_value_written() { pending_key_ = false; }

    static void append_u16_escape(std::string& out, std::uint16_t code) {
        static const char hex[] = "0123456789ABCDEF";
        out.append("\\u");
        out.push_back(hex[(code >> 12) & 0xF]);
        out.push_back(hex[(code >> 8) & 0xF]);
        out.push_back(hex[(code >> 4) & 0xF]);
        out.push_back(hex[code & 0xF]);
    }

    static bool needs_escape_ascii(unsigned char c) {
        if (c < 0x20 || c > 0x7E) return true;
        switch (c) {
            case '"': case '\\': case '<': case '>': case '&': case '\'': case '+': case '`':
                return true;
            default:
                return false;
        }
    }

    void write_string_literal(std::string_view v) {
        out_.push_back('"');
        std::size_t i = 0;
        const std::size_t n = v.size();
        while (i < n) {
            unsigned char c = static_cast<unsigned char>(v[i]);
            if (c < 0x80) {
                if (!needs_escape_ascii(c)) {
                    out_.push_back(static_cast<char>(c));
                    ++i;
                    continue;
                }
                switch (c) {
                    case '\b': out_.append("\\b"); break;
                    case '\t': out_.append("\\t"); break;
                    case '\n': out_.append("\\n"); break;
                    case '\f': out_.append("\\f"); break;
                    case '\r': out_.append("\\r"); break;
                    case '"': append_u16_escape(out_, '"'); break; // STJ emits ", not \"
                    case '\\': out_.append("\\\\"); break;
                    default: append_u16_escape(out_, c); break;
                }
                ++i;
                continue;
            }

            // Decode one UTF-8 sequence and emit \uXXXX (surrogate pair when needed).
            std::uint32_t code = 0xFFFD;
            std::size_t length = 1;
            if ((c & 0xE0) == 0xC0 && i + 1 < n && (v[i + 1] & 0xC0) == 0x80) {
                code = (static_cast<std::uint32_t>(c & 0x1F) << 6) |
                       (static_cast<std::uint32_t>(v[i + 1]) & 0x3F);
                length = 2;
            } else if ((c & 0xF0) == 0xE0 && i + 2 < n &&
                       (v[i + 1] & 0xC0) == 0x80 && (v[i + 2] & 0xC0) == 0x80) {
                code = (static_cast<std::uint32_t>(c & 0x0F) << 12) |
                       ((static_cast<std::uint32_t>(v[i + 1]) & 0x3F) << 6) |
                       (static_cast<std::uint32_t>(v[i + 2]) & 0x3F);
                length = 3;
            } else if ((c & 0xF8) == 0xF0 && i + 3 < n &&
                       (v[i + 1] & 0xC0) == 0x80 && (v[i + 2] & 0xC0) == 0x80 &&
                       (v[i + 3] & 0xC0) == 0x80) {
                code = (static_cast<std::uint32_t>(c & 0x07) << 18) |
                       ((static_cast<std::uint32_t>(v[i + 1]) & 0x3F) << 12) |
                       ((static_cast<std::uint32_t>(v[i + 2]) & 0x3F) << 6) |
                       (static_cast<std::uint32_t>(v[i + 3]) & 0x3F);
                length = 4;
            }

            if (code >= 0x10000) {
                std::uint32_t bias = code - 0x10000;
                append_u16_escape(out_, static_cast<std::uint16_t>(0xD800 + (bias >> 10)));
                append_u16_escape(out_, static_cast<std::uint16_t>(0xDC00 + (bias & 0x3FF)));
            } else {
                append_u16_escape(out_, static_cast<std::uint16_t>(code));
            }
            i += length;
        }
        out_.push_back('"');
    }
};

// ---------------------------------------------------------------------------
// Parser
// ---------------------------------------------------------------------------

class Parser {
public:
    static Value parse(std::string_view text) {
        Parser parser(text);
        parser.skip_whitespace();
        Value result = parser.parse_value(0);
        parser.skip_whitespace();
        if (parser.pos_ != parser.text_.size()) {
            throw ParseError("Unexpected trailing characters after JSON value.");
        }
        return result;
    }

private:
    static constexpr int MaxDepth = 64;

    std::string_view text_;
    std::size_t pos_ = 0;

    explicit Parser(std::string_view text) : text_(text) {}

    [[noreturn]] void fail(const char* message) const {
        throw ParseError(std::string(message) + " (offset " + std::to_string(pos_) + ")");
    }

    bool at_end() const noexcept { return pos_ >= text_.size(); }

    char peek() const {
        if (at_end()) throw ParseError("Unexpected end of JSON input.");
        return text_[pos_];
    }

    char take() {
        char c = peek();
        ++pos_;
        return c;
    }

    void skip_whitespace() noexcept {
        while (pos_ < text_.size()) {
            char c = text_[pos_];
            if (c == ' ' || c == '\t' || c == '\n' || c == '\r') ++pos_;
            else break;
        }
    }

    void expect(char c) {
        if (at_end() || text_[pos_] != c) fail("Unexpected character in JSON input.");
        ++pos_;
    }

    bool try_consume(char c) noexcept {
        if (pos_ < text_.size() && text_[pos_] == c) { ++pos_; return true; }
        return false;
    }

    Value parse_value(int depth) {
        if (depth > MaxDepth) fail("JSON nesting depth exceeds limit.");
        skip_whitespace();
        char c = peek();
        switch (c) {
            case '{': return parse_object(depth);
            case '[': return parse_array(depth);
            case '"': return Value(parse_string());
            case 't': consume_literal("true"); return Value(true);
            case 'f': consume_literal("false"); return Value(false);
            case 'n': consume_literal("null"); return Value(nullptr);
            default: return parse_number();
        }
    }

    void consume_literal(std::string_view literal) {
        if (text_.substr(pos_, literal.size()) != literal) fail("Invalid JSON literal.");
        pos_ += literal.size();
    }

    Value parse_object(int depth) {
        expect('{');
        Object object;
        skip_whitespace();
        if (try_consume('}')) return Value(std::move(object));
        while (true) {
            skip_whitespace();
            std::string key = parse_string();
            skip_whitespace();
            expect(':');
            object.add(std::move(key), parse_value(depth + 1));
            skip_whitespace();
            if (try_consume(',')) continue;
            expect('}');
            break;
        }
        return Value(std::move(object));
    }

    Value parse_array(int depth) {
        expect('[');
        Array array;
        skip_whitespace();
        if (try_consume(']')) return Value(std::move(array));
        while (true) {
            array.push_back(parse_value(depth + 1));
            skip_whitespace();
            if (try_consume(',')) continue;
            expect(']');
            break;
        }
        return Value(std::move(array));
    }

    std::uint16_t parse_hex4() {
        std::uint16_t code = 0;
        for (int i = 0; i < 4; ++i) {
            char c = take();
            code = static_cast<std::uint16_t>(code << 4);
            if (c >= '0' && c <= '9') code = static_cast<std::uint16_t>(code + (c - '0'));
            else if (c >= 'a' && c <= 'f') code = static_cast<std::uint16_t>(code + (c - 'a' + 10));
            else if (c >= 'A' && c <= 'F') code = static_cast<std::uint16_t>(code + (c - 'A' + 10));
            else fail("Invalid \\u escape in JSON string.");
        }
        return code;
    }

    static void append_utf8(std::string& out, std::uint32_t code) {
        if (code < 0x80) {
            out.push_back(static_cast<char>(code));
        } else if (code < 0x800) {
            out.push_back(static_cast<char>(0xC0 | (code >> 6)));
            out.push_back(static_cast<char>(0x80 | (code & 0x3F)));
        } else if (code < 0x10000) {
            out.push_back(static_cast<char>(0xE0 | (code >> 12)));
            out.push_back(static_cast<char>(0x80 | ((code >> 6) & 0x3F)));
            out.push_back(static_cast<char>(0x80 | (code & 0x3F)));
        } else {
            out.push_back(static_cast<char>(0xF0 | (code >> 18)));
            out.push_back(static_cast<char>(0x80 | ((code >> 12) & 0x3F)));
            out.push_back(static_cast<char>(0x80 | ((code >> 6) & 0x3F)));
            out.push_back(static_cast<char>(0x80 | (code & 0x3F)));
        }
    }

    std::string parse_string() {
        expect('"');
        std::string out;
        while (true) {
            char c = take();
            if (c == '"') break;
            if (c == '\\') {
                char escape = take();
                switch (escape) {
                    case '"': out.push_back('"'); break;
                    case '\\': out.push_back('\\'); break;
                    case '/': out.push_back('/'); break;
                    case 'b': out.push_back('\b'); break;
                    case 'f': out.push_back('\f'); break;
                    case 'n': out.push_back('\n'); break;
                    case 'r': out.push_back('\r'); break;
                    case 't': out.push_back('\t'); break;
                    case 'u': {
                        std::uint32_t code = parse_hex4();
                        if (code >= 0xD800 && code <= 0xDBFF) {
                            // High surrogate: require a low surrogate escape next.
                            if (pos_ + 1 < text_.size() && text_[pos_] == '\\' && text_[pos_ + 1] == 'u') {
                                pos_ += 2;
                                std::uint32_t low = parse_hex4();
                                if (low >= 0xDC00 && low <= 0xDFFF) {
                                    code = 0x10000 + ((code - 0xD800) << 10) + (low - 0xDC00);
                                } else {
                                    code = 0xFFFD;
                                }
                            } else {
                                code = 0xFFFD;
                            }
                        } else if (code >= 0xDC00 && code <= 0xDFFF) {
                            code = 0xFFFD;
                        }
                        append_utf8(out, code);
                        break;
                    }
                    default:
                        fail("Invalid escape sequence in JSON string.");
                }
                continue;
            }
            if (static_cast<unsigned char>(c) < 0x20) fail("Unescaped control character in JSON string.");
            out.push_back(c);
        }
        return out;
    }

    Value parse_number() {
        std::size_t start = pos_;
        if (try_consume('-')) {}
        if (at_end()) fail("Invalid JSON number.");
        while (!at_end()) {
            char c = text_[pos_];
            if ((c >= '0' && c <= '9') || c == '+' || c == '-' || c == '.' || c == 'e' || c == 'E') ++pos_;
            else break;
        }
        std::string_view token = text_.substr(start, pos_ - start);
        if (token.empty()) fail("Invalid JSON number.");

        bool is_integer = true;
        for (char c : token) {
            if (c == '.' || c == 'e' || c == 'E') { is_integer = false; break; }
        }

        if (is_integer) {
            std::int64_t integer = 0;
            auto result = std::from_chars(token.data(), token.data() + token.size(), integer);
            if (result.ec == std::errc() && result.ptr == token.data() + token.size()) {
                return Value(integer);
            }
        }

        double real = 0.0;
        auto result = std::from_chars(token.data(), token.data() + token.size(), real);
        if (result.ec != std::errc() || result.ptr != token.data() + token.size()) {
            fail("Invalid JSON number.");
        }
        return Value(real);
    }
};

inline Value parse(std::string_view text) { return Parser::parse(text); }

} // namespace xact::json
