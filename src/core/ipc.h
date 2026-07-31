// -----------------------------------------------------------------------------
// File: src\core\ipc.h
// Purpose: C++ port of the XactCopy IPC contract: envelope serialization, the
//          message-type registry, and every command/event payload exchanged
//          between the supervisor (UI) and XactCopyExecutive (worker).
// -----------------------------------------------------------------------------

#pragma once

#include <cstdint>
#include <functional>
#include <optional>
#include <string>
#include <string_view>

#include "dotnet_time.h"
#include "json.h"
#include "models.h"

namespace xact::ipc {

using models::CopyJobOptions;
using models::CopyJobResult;
using models::CopyProgressSnapshot;
using time::DateTimeOffset;

inline constexpr std::int32_t ProtocolVersion = 1;

namespace message_types {
inline constexpr std::string_view StartJobCommand = "StartJobCommand";
inline constexpr std::string_view CancelJobCommand = "CancelJobCommand";
inline constexpr std::string_view PauseJobCommand = "PauseJobCommand";
inline constexpr std::string_view ResumeJobCommand = "ResumeJobCommand";
inline constexpr std::string_view PingCommand = "PingCommand";
inline constexpr std::string_view ShutdownCommand = "ShutdownCommand";

inline constexpr std::string_view WorkerHeartbeatEvent = "WorkerHeartbeatEvent";
inline constexpr std::string_view WorkerProgressEvent = "WorkerProgressEvent";
inline constexpr std::string_view WorkerLogEvent = "WorkerLogEvent";
inline constexpr std::string_view WorkerJobResultEvent = "WorkerJobResultEvent";
inline constexpr std::string_view WorkerFatalEvent = "WorkerFatalEvent";
} // namespace message_types

// ---------------------------------------------------------------------------
// Envelope
// ---------------------------------------------------------------------------

// Serializes {"ProtocolVersion":1,"MessageType":...,"CorrelationId":...,
// "SentUtc":...,"Payload":{...}} exactly like IpcSerializer.SerializeEnvelope.
// The payload is written by a callback so any model type can be nested.
inline std::string serialize_envelope(
    std::string_view message_type,
    const std::function<void(json::Writer&)>& write_payload,
    std::string_view correlation_id = {},
    std::optional<DateTimeOffset> sent_utc = std::nullopt) {

    json::Writer w;
    w.begin_object();
    w.key("ProtocolVersion"); w.value(ProtocolVersion);
    w.key("MessageType"); w.value(message_type);
    w.key("CorrelationId"); w.value(correlation_id);
    w.key("SentUtc"); w.value_literal_string(sent_utc.value_or(DateTimeOffset::now_utc()).to_string());
    w.key("Payload");
    write_payload(w);
    w.end_object();
    return w.take();
}

struct EnvelopeHeader {
    std::int32_t protocol_version = 0;
    std::string message_type;
    std::string correlation_id;
    DateTimeOffset sent_utc;
};

// Parses an envelope and returns the header plus the payload DOM value.
// Throws json::ParseError for malformed JSON and std::runtime_error for
// protocol-version mismatches, mirroring IpcSerializer.DeserializeEnvelope.
inline std::pair<EnvelopeHeader, json::Value> parse_envelope(std::string_view text) {
    json::Value root = json::parse(text);
    const auto* obj = root.as_object();
    if (obj == nullptr) {
        throw std::runtime_error("IPC envelope deserialization returned no object.");
    }

    EnvelopeHeader header;
    if (const auto* v = obj->find("ProtocolVersion"); v != nullptr) header.protocol_version = v->as_int32();
    if (const auto* v = obj->find("MessageType"); v != nullptr) header.message_type = v->as_string();
    if (const auto* v = obj->find("CorrelationId"); v != nullptr) header.correlation_id = v->as_string();
    if (const auto* v = obj->find("SentUtc"); v != nullptr && v->is_string()) {
        if (auto parsed = DateTimeOffset::parse(v->as_string())) header.sent_utc = *parsed;
    }

    if (header.protocol_version != ProtocolVersion) {
        throw std::runtime_error(
            "Unsupported protocol version " + std::to_string(header.protocol_version) +
            ". Expected " + std::to_string(ProtocolVersion) + ".");
    }

    json::Value payload;
    if (const auto* v = obj->find("Payload"); v != nullptr) payload = *v;
    return {std::move(header), std::move(payload)};
}

// Fast message-type probe replicating IpcSerializer.TryReadMessageType: scans
// the raw JSON for the ProtocolVersion and MessageType fields without building
// a DOM, relying on the compact PascalCase envelope shape both sides emit.
inline bool try_read_message_type(std::string_view json_text, std::string& message_type) {
    message_type.clear();
    if (json_text.empty()) return false;

    constexpr std::string_view pv_key = "\"ProtocolVersion\":";
    std::size_t pv_idx = json_text.find(pv_key);
    if (pv_idx == std::string_view::npos) return false;

    std::size_t v_start = pv_idx + pv_key.size();
    std::size_t v_end = v_start;
    while (v_end < json_text.size() && json_text[v_end] >= '0' && json_text[v_end] <= '9') ++v_end;
    if (v_end == v_start) return false;

    std::int32_t version = 0;
    for (std::size_t i = v_start; i < v_end; ++i) {
        version = version * 10 + (json_text[i] - '0');
        if (version > 1'000'000) return false;
    }
    if (version != ProtocolVersion) return false;

    constexpr std::string_view mt_key = "\"MessageType\":\"";
    std::size_t mt_idx = json_text.find(mt_key);
    if (mt_idx == std::string_view::npos) return false;

    std::size_t value_start = mt_idx + mt_key.size();
    std::size_t value_end = json_text.find('"', value_start);
    if (value_end == std::string_view::npos || value_end < value_start) return false;

    message_type.assign(json_text, value_start, value_end - value_start);
    for (char c : message_type) {
        if (c != ' ' && c != '\t') return true;
    }
    return false;
}

// ---------------------------------------------------------------------------
// Command payloads (UI -> worker)
// ---------------------------------------------------------------------------

struct StartJobCommand {
    std::string JobId;
    CopyJobOptions Options;

    void to_json(json::Writer& w) const {
        w.begin_object();
        w.key("JobId"); w.value(JobId);
        w.key("Options");
        Options.to_json(w);
        w.end_object();
    }

    static StartJobCommand from_json(const json::Value& value) {
        StartJobCommand c;
        const auto* obj = value.as_object();
        if (obj == nullptr) return c;
        models::detail::read(obj, "JobId", c.JobId);
        if (const auto* v = obj->find("Options"); v != nullptr) c.Options = CopyJobOptions::from_json(*v);
        return c;
    }
};

// Cancel/Pause/Resume share the {JobId, Reason} shape.
struct JobControlCommand {
    std::string JobId;
    std::string Reason;

    void to_json(json::Writer& w) const {
        w.begin_object();
        w.key("JobId"); w.value(JobId);
        w.key("Reason"); w.value(Reason);
        w.end_object();
    }

    static JobControlCommand from_json(const json::Value& value) {
        JobControlCommand c;
        const auto* obj = value.as_object();
        if (obj == nullptr) return c;
        models::detail::read(obj, "JobId", c.JobId);
        models::detail::read(obj, "Reason", c.Reason);
        return c;
    }
};

using CancelJobCommand = JobControlCommand;
using PauseJobCommand = JobControlCommand;
using ResumeJobCommand = JobControlCommand;

struct PingCommand {
    DateTimeOffset Utc = DateTimeOffset::now_utc();

    void to_json(json::Writer& w) const {
        w.begin_object();
        w.key("Utc"); w.value_literal_string(Utc.to_string());
        w.end_object();
    }

    static PingCommand from_json(const json::Value& value) {
        PingCommand c;
        const auto* obj = value.as_object();
        if (obj == nullptr) return c;
        models::detail::read(obj, "Utc", c.Utc);
        return c;
    }
};

struct ShutdownCommand {
    std::string Reason;

    void to_json(json::Writer& w) const {
        w.begin_object();
        w.key("Reason"); w.value(Reason);
        w.end_object();
    }

    static ShutdownCommand from_json(const json::Value& value) {
        ShutdownCommand c;
        const auto* obj = value.as_object();
        if (obj == nullptr) return c;
        models::detail::read(obj, "Reason", c.Reason);
        return c;
    }
};

// ---------------------------------------------------------------------------
// Event payloads (worker -> UI)
// ---------------------------------------------------------------------------

struct WorkerHeartbeatEvent {
    std::int32_t WorkerProcessId = 0;
    bool IsJobRunning = false;
    bool IsJobPaused = false;
    std::string ActiveJobId;
    DateTimeOffset LastProgressUtc;
    DateTimeOffset TimestampUtc = DateTimeOffset::now_utc();

    void to_json(json::Writer& w) const {
        w.begin_object();
        w.key("WorkerProcessId"); w.value(WorkerProcessId);
        w.key("IsJobRunning"); w.value(IsJobRunning);
        w.key("IsJobPaused"); w.value(IsJobPaused);
        w.key("ActiveJobId"); w.value(ActiveJobId);
        w.key("LastProgressUtc"); w.value_literal_string(LastProgressUtc.to_string());
        w.key("TimestampUtc"); w.value_literal_string(TimestampUtc.to_string());
        w.end_object();
    }

    static WorkerHeartbeatEvent from_json(const json::Value& value) {
        WorkerHeartbeatEvent e;
        const auto* obj = value.as_object();
        if (obj == nullptr) return e;
        models::detail::read(obj, "WorkerProcessId", e.WorkerProcessId);
        models::detail::read(obj, "IsJobRunning", e.IsJobRunning);
        models::detail::read(obj, "IsJobPaused", e.IsJobPaused);
        models::detail::read(obj, "ActiveJobId", e.ActiveJobId);
        models::detail::read(obj, "LastProgressUtc", e.LastProgressUtc);
        models::detail::read(obj, "TimestampUtc", e.TimestampUtc);
        return e;
    }
};

struct WorkerProgressEvent {
    std::string JobId;
    CopyProgressSnapshot Snapshot;
    DateTimeOffset TimestampUtc = DateTimeOffset::now_utc();

    void to_json(json::Writer& w) const {
        w.begin_object();
        w.key("JobId"); w.value(JobId);
        w.key("Snapshot");
        Snapshot.to_json(w);
        w.key("TimestampUtc"); w.value_literal_string(TimestampUtc.to_string());
        w.end_object();
    }

    static WorkerProgressEvent from_json(const json::Value& value) {
        WorkerProgressEvent e;
        const auto* obj = value.as_object();
        if (obj == nullptr) return e;
        models::detail::read(obj, "JobId", e.JobId);
        if (const auto* v = obj->find("Snapshot"); v != nullptr) e.Snapshot = CopyProgressSnapshot::from_json(*v);
        models::detail::read(obj, "TimestampUtc", e.TimestampUtc);
        return e;
    }
};

struct WorkerLogEvent {
    std::string JobId;
    std::string Message;
    DateTimeOffset TimestampUtc = DateTimeOffset::now_utc();

    void to_json(json::Writer& w) const {
        w.begin_object();
        w.key("JobId"); w.value(JobId);
        w.key("Message"); w.value(Message);
        w.key("TimestampUtc"); w.value_literal_string(TimestampUtc.to_string());
        w.end_object();
    }

    static WorkerLogEvent from_json(const json::Value& value) {
        WorkerLogEvent e;
        const auto* obj = value.as_object();
        if (obj == nullptr) return e;
        models::detail::read(obj, "JobId", e.JobId);
        models::detail::read(obj, "Message", e.Message);
        models::detail::read(obj, "TimestampUtc", e.TimestampUtc);
        return e;
    }
};

struct WorkerJobResultEvent {
    std::string JobId;
    CopyJobResult Result;
    DateTimeOffset TimestampUtc = DateTimeOffset::now_utc();

    void to_json(json::Writer& w) const {
        w.begin_object();
        w.key("JobId"); w.value(JobId);
        w.key("Result");
        Result.to_json(w);
        w.key("TimestampUtc"); w.value_literal_string(TimestampUtc.to_string());
        w.end_object();
    }

    static WorkerJobResultEvent from_json(const json::Value& value) {
        WorkerJobResultEvent e;
        const auto* obj = value.as_object();
        if (obj == nullptr) return e;
        models::detail::read(obj, "JobId", e.JobId);
        if (const auto* v = obj->find("Result"); v != nullptr) e.Result = CopyJobResult::from_json(*v);
        models::detail::read(obj, "TimestampUtc", e.TimestampUtc);
        return e;
    }
};

struct WorkerFatalEvent {
    std::string JobId;
    std::string ErrorMessage;
    std::string StackTraceText;
    DateTimeOffset TimestampUtc = DateTimeOffset::now_utc();

    void to_json(json::Writer& w) const {
        w.begin_object();
        w.key("JobId"); w.value(JobId);
        w.key("ErrorMessage"); w.value(ErrorMessage);
        w.key("StackTraceText"); w.value(StackTraceText);
        w.key("TimestampUtc"); w.value_literal_string(TimestampUtc.to_string());
        w.end_object();
    }

    static WorkerFatalEvent from_json(const json::Value& value) {
        WorkerFatalEvent e;
        const auto* obj = value.as_object();
        if (obj == nullptr) return e;
        models::detail::read(obj, "JobId", e.JobId);
        models::detail::read(obj, "ErrorMessage", e.ErrorMessage);
        models::detail::read(obj, "StackTraceText", e.StackTraceText);
        models::detail::read(obj, "TimestampUtc", e.TimestampUtc);
        return e;
    }
};

// Convenience overload: serialize an envelope around any payload model.
template <typename TPayload>
std::string serialize_envelope_for(
    std::string_view message_type,
    const TPayload& payload,
    std::string_view correlation_id = {},
    std::optional<DateTimeOffset> sent_utc = std::nullopt) {
    return serialize_envelope(
        message_type,
        [&payload](json::Writer& w) { payload.to_json(w); },
        correlation_id,
        sent_utc);
}

} // namespace xact::ipc
