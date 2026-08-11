// -----------------------------------------------------------------------------
// File: src\core\models.h
// Purpose: C++ port of XactCopy.Core model types (enums, run options, results,
//          progress snapshots) with JSON serialization that matches the .NET
//          System.Text.Json output field-for-field.
// -----------------------------------------------------------------------------

#pragma once

#include <algorithm>
#include <cstdint>
#include <limits>
#include <optional>
#include <string>
#include <string_view>
#include <vector>

#include "dotnet_time.h"
#include "json.h"

namespace xact::models {

using time::DateTimeOffset;
using time::TimeSpan;

// ---------------------------------------------------------------------------
// Enums (serialized as their .NET member names via JsonStringEnumConverter).
// ---------------------------------------------------------------------------

enum class JobOperationMode : std::int32_t { Copy = 0, ScanOnly = 1 };
enum class OverwritePolicy : std::int32_t { Overwrite = 0, SkipExisting = 1, OverwriteIfSourceNewer = 2, Ask = 3 };
enum class SymlinkHandlingMode : std::int32_t {
    Skip = 0,
    Follow = 1,         // Legacy value: explicitly allow external targets.
    FollowInternal = 2  // Follow only targets that remain beneath SourceRoot.
};
enum class TransferEnginePolicy : std::int32_t { Auto = 0, ManagedRescue = 1, NativeFast = 2 };
enum class SourceMutationPolicy : std::int32_t { FailFile = 0, SkipFile = 1, WaitForReappearance = 2 };
enum class VerificationMode : std::int32_t { None = 0, Sampled = 1, Full = 2 };
enum class VerificationHashAlgorithm : std::int32_t { Sha256 = 0, Sha512 = 1 };
enum class SalvageFillPattern : std::int32_t { Zero = 0, Ones = 1, Random = 2 };
enum class ScanPerformanceProfile : std::int32_t { Auto = 0, Fast = 1, Precise = 2 };
enum class WorkerTelemetryProfile : std::int32_t { Normal = 0, Verbose = 1, Debug = 2 };

namespace detail {

inline bool equals_ignore_case(std::string_view a, std::string_view b) {
    if (a.size() != b.size()) return false;
    for (std::size_t i = 0; i < a.size(); ++i) {
        char x = a[i], y = b[i];
        if (x >= 'A' && x <= 'Z') x = static_cast<char>(x - 'A' + 'a');
        if (y >= 'A' && y <= 'Z') y = static_cast<char>(y - 'A' + 'a');
        if (x != y) return false;
    }
    return true;
}

template <typename TEnum>
struct EnumEntry {
    TEnum value;
    std::string_view name;
};

template <typename TEnum, std::size_t N>
std::string_view enum_name(const EnumEntry<TEnum> (&table)[N], TEnum value) {
    for (const auto& entry : table) {
        if (entry.value == value) return entry.name;
    }
    return table[0].name;
}

// Mirrors JsonStringEnumConverter reads: accepts the member name in any case,
// or the underlying integer.
template <typename TEnum, std::size_t N>
TEnum enum_parse(const EnumEntry<TEnum> (&table)[N], const json::Value* value, TEnum fallback) {
    if (value == nullptr) return fallback;
    if (value->is_number()) {
        const std::int32_t raw = value->as_int32(std::numeric_limits<std::int32_t>::min());
        for (const auto& entry : table) {
            if (static_cast<std::int32_t>(entry.value) == raw) return entry.value;
        }
        return fallback;
    }
    if (!value->is_string()) return fallback;
    const std::string& text = value->as_string();
    for (const auto& entry : table) {
        if (equals_ignore_case(entry.name, text)) return entry.value;
    }
    return fallback;
}

} // namespace detail

#define XACT_ENUM_TABLE(EnumType, ...)                                              \
    inline constexpr detail::EnumEntry<EnumType> EnumType##_Table[] = {__VA_ARGS__}; \
    inline std::string_view to_string(EnumType v) { return detail::enum_name(EnumType##_Table, v); } \
    inline EnumType parse_##EnumType(const json::Value* v, EnumType fallback) {     \
        return detail::enum_parse(EnumType##_Table, v, fallback);                   \
    }

XACT_ENUM_TABLE(JobOperationMode,
    {JobOperationMode::Copy, "Copy"},
    {JobOperationMode::ScanOnly, "ScanOnly"})
XACT_ENUM_TABLE(OverwritePolicy,
    {OverwritePolicy::Overwrite, "Overwrite"},
    {OverwritePolicy::SkipExisting, "SkipExisting"},
    {OverwritePolicy::OverwriteIfSourceNewer, "OverwriteIfSourceNewer"},
    {OverwritePolicy::Ask, "Ask"})
XACT_ENUM_TABLE(SymlinkHandlingMode,
    {SymlinkHandlingMode::Skip, "Skip"},
    {SymlinkHandlingMode::Follow, "Follow"},
    {SymlinkHandlingMode::FollowInternal, "FollowInternal"})
XACT_ENUM_TABLE(TransferEnginePolicy,
    {TransferEnginePolicy::Auto, "Auto"},
    {TransferEnginePolicy::ManagedRescue, "ManagedRescue"},
    {TransferEnginePolicy::NativeFast, "NativeFast"})
XACT_ENUM_TABLE(SourceMutationPolicy,
    {SourceMutationPolicy::FailFile, "FailFile"},
    {SourceMutationPolicy::SkipFile, "SkipFile"},
    {SourceMutationPolicy::WaitForReappearance, "WaitForReappearance"})
XACT_ENUM_TABLE(VerificationMode,
    {VerificationMode::None, "None"},
    {VerificationMode::Sampled, "Sampled"},
    {VerificationMode::Full, "Full"})
XACT_ENUM_TABLE(VerificationHashAlgorithm,
    {VerificationHashAlgorithm::Sha256, "Sha256"},
    {VerificationHashAlgorithm::Sha512, "Sha512"})
XACT_ENUM_TABLE(SalvageFillPattern,
    {SalvageFillPattern::Zero, "Zero"},
    {SalvageFillPattern::Ones, "Ones"},
    {SalvageFillPattern::Random, "Random"})
XACT_ENUM_TABLE(ScanPerformanceProfile,
    {ScanPerformanceProfile::Auto, "Auto"},
    {ScanPerformanceProfile::Fast, "Fast"},
    {ScanPerformanceProfile::Precise, "Precise"})
XACT_ENUM_TABLE(WorkerTelemetryProfile,
    {WorkerTelemetryProfile::Normal, "Normal"},
    {WorkerTelemetryProfile::Verbose, "Verbose"},
    {WorkerTelemetryProfile::Debug, "Debug"})

#undef XACT_ENUM_TABLE

// ---------------------------------------------------------------------------
// Field read helpers (missing fields keep defaults, like STJ deserialization).
// ---------------------------------------------------------------------------

namespace detail {

inline const json::Value* field(const json::Object* obj, std::string_view name) {
    return obj == nullptr ? nullptr : obj->find(name);
}

inline void read(const json::Object* obj, std::string_view name, std::string& target) {
    if (const auto* v = field(obj, name); v != nullptr && v->is_string()) target = v->as_string();
}

inline void read(const json::Object* obj, std::string_view name, bool& target) {
    if (const auto* v = field(obj, name); v != nullptr && v->is_bool()) target = v->as_bool();
}

inline void read(const json::Object* obj, std::string_view name, std::int32_t& target) {
    if (const auto* v = field(obj, name); v != nullptr && v->is_number()) {
        target = v->as_int32(target);
    }
}

inline void read(const json::Object* obj, std::string_view name, std::int64_t& target) {
    if (const auto* v = field(obj, name); v != nullptr && v->is_number()) {
        target = v->as_int64(target);
    }
}

inline void read(const json::Object* obj, std::string_view name, double& target) {
    if (const auto* v = field(obj, name); v != nullptr && v->is_number()) target = v->as_double();
}

inline void read(const json::Object* obj, std::string_view name, TimeSpan& target) {
    if (const auto* v = field(obj, name); v != nullptr && v->is_string()) {
        if (auto parsed = TimeSpan::parse(v->as_string())) target = *parsed;
    }
}

inline void read(const json::Object* obj, std::string_view name, DateTimeOffset& target) {
    if (const auto* v = field(obj, name); v != nullptr && v->is_string()) {
        if (auto parsed = DateTimeOffset::parse(v->as_string())) target = *parsed;
    }
}

inline void read(const json::Object* obj, std::string_view name, std::vector<std::string>& target) {
    if (const auto* v = field(obj, name); v != nullptr && v->is_array()) {
        target.clear();
        for (const auto& item : *v->as_array()) {
            if (item.is_string()) target.push_back(item.as_string());
        }
    }
}

inline void write_string_array(json::Writer& w, std::string_view name, const std::vector<std::string>& items) {
    w.key(name);
    w.begin_array();
    for (const auto& item : items) w.value(item);
    w.end_array();
}

} // namespace detail

// ---------------------------------------------------------------------------
// CopyJobOptions — complete run configuration (field order matches the VB
// declaration order so serialized output is byte-identical to .NET).
// ---------------------------------------------------------------------------

struct CopyJobOptions {
    static constexpr std::int32_t MinimumBufferSizeBytes = 1024 * 1024;
    static constexpr std::int32_t MaximumBufferSizeBytes = 256 * 1024 * 1024;
    static constexpr std::int32_t MaximumRetries = 32;
    static constexpr std::int32_t MaximumNormalRetries = 3;
    static constexpr std::int32_t MinimumOperationTimeoutSeconds = 1;
    static constexpr std::int32_t MaximumOperationTimeoutSeconds = 3600;
    static constexpr std::int32_t MaximumPerFileTimeoutSeconds = 86400;
    static constexpr std::int32_t MaximumRetryDelaySeconds = 86400;
    static constexpr std::int32_t MaximumParallelWorkers = 64;
    static constexpr std::int32_t MaximumSampleVerificationChunkBytes = 4 * 1024 * 1024;
    static constexpr std::int32_t MaximumSampleVerificationChunkCount = 64;
    static constexpr std::int32_t MaximumBadRangeMapAgeDays = 3650;
    static constexpr std::int32_t MaximumSmallFileThresholdBytes = 1024 * 1024 * 1024;
    static constexpr std::int32_t MinimumLockProbeIntervalMilliseconds = 100;
    static constexpr std::int32_t MaximumLockProbeIntervalMilliseconds = 10000;
    static constexpr std::int32_t MaximumFragileFailureWindowSeconds = 3600;
    static constexpr std::int32_t MaximumFragileFailureThreshold = 1000;
    static constexpr std::int32_t MaximumFragileCooldownSeconds = 600;
    static constexpr std::int32_t MaximumRescueChunkBytes = 256 * 1024 * 1024;
    static constexpr std::int32_t MaximumRescueSplitBytes = 64 * 1024 * 1024;

    JobOperationMode OperationMode = JobOperationMode::Copy;
    std::string SourceRoot;
    std::string DestinationRoot;
    std::string ExpectedSourceIdentity;
    std::string ExpectedDestinationIdentity;

    bool UseBadRangeMap = false;
    bool SkipKnownBadRanges = false;
    bool UpdateBadRangeMapFromRun = false;
    std::int32_t BadRangeMapMaxAgeDays = 30;
    bool UseExperimentalRawDiskScan = false;
    ScanPerformanceProfile ScanPerformanceProfileValue = ScanPerformanceProfile::Auto;

    std::string ResumeJournalPathHint;
    bool AllowJournalRootRemap = false;

    std::vector<std::string> SelectedRelativePaths;

    OverwritePolicy OverwritePolicyValue = OverwritePolicy::Overwrite;
    SymlinkHandlingMode SymlinkHandling = SymlinkHandlingMode::Skip;
    bool CopyEmptyDirectories = true;
    // Recovery output is useful, but synthetic bytes must not silently replace
    // a better pre-existing copy. The worker publishes a sidecar unless this
    // explicit expert override is present in the complete job payload.
    bool AllowRecoveredOverwriteExisting = false;
    // EFS reads yield plaintext to normal callers. Keep plaintext publication
    // opt-in instead of treating loss of encryption as a metadata warning.
    bool AllowDecryptedDestination = false;

    std::int32_t BufferSizeBytes = 4 * 1024 * 1024;
    bool UseAdaptiveBufferSizing = false;
    TransferEnginePolicy TransferEnginePolicyValue = TransferEnginePolicy::Auto;
    std::int64_t MaxThroughputBytesPerSecond = 0;
    std::int32_t ParallelSmallFileWorkers = 0;
    std::int32_t SmallFileThresholdBytes = 256 * 1024;
    std::int32_t ParallelScanWorkers = 0;

    bool WaitForMediaAvailability = false;
    bool WaitForFileLockRelease = false;
    bool TreatAccessDeniedAsContention = false;
    TimeSpan LockContentionProbeInterval = TimeSpan::from_milliseconds(500);
    SourceMutationPolicy SourceMutationPolicyValue = SourceMutationPolicy::FailFile;
    bool FragileMediaMode = false;
    bool SkipFileOnFirstReadError = true;
    bool PersistFragileSkipAcrossResume = true;
    std::int32_t FragileFailureWindowSeconds = 20;
    std::int32_t FragileFailureThreshold = 3;
    std::int32_t FragileCooldownSeconds = 6;

    std::int32_t MaxRetries = 2;
    TimeSpan OperationTimeout = TimeSpan::from_seconds(10);
    TimeSpan PerFileTimeout = TimeSpan::zero();
    TimeSpan InitialRetryDelay = TimeSpan::from_milliseconds(250);
    TimeSpan MaxRetryDelay = TimeSpan::from_seconds(8);

    bool ResumeFromJournal = true;
    bool VerifyAfterCopy = true;
    VerificationMode VerificationModeValue = VerificationMode::Full;
    VerificationHashAlgorithm VerificationHashAlgorithmValue = VerificationHashAlgorithm::Sha256;
    std::int32_t SampleVerificationChunkBytes = 128 * 1024;
    std::int32_t SampleVerificationChunkCount = 3;
    bool SalvageUnreadableBlocks = false;
    SalvageFillPattern SalvageFillPatternValue = SalvageFillPattern::Zero;
    bool ContinueOnFileError = false;
    bool PreserveTimestamps = true;

    std::string WorkerProcessPriorityClass = "Normal";
    WorkerTelemetryProfile WorkerTelemetryProfileValue = WorkerTelemetryProfile::Normal;
    std::int32_t WorkerProgressEmitIntervalMs = 75;
    std::int32_t WorkerMaxLogsPerSecond = 100;

    std::int32_t RescueFastScanChunkBytes = 0;
    std::int32_t RescueTrimChunkBytes = 0;
    std::int32_t RescueScrapeChunkBytes = 0;
    std::int32_t RescueRetryChunkBytes = 0;
    std::int32_t RescueSplitMinimumBytes = 0;
    std::int32_t RescueFastScanRetries = 0;
    std::int32_t RescueTrimRetries = 1;
    std::int32_t RescueScrapeRetries = 2;

    void to_json(json::Writer& w) const {
        w.begin_object();
        w.key("OperationMode"); w.value(to_string(OperationMode));
        w.key("SourceRoot"); w.value(SourceRoot);
        w.key("DestinationRoot"); w.value(DestinationRoot);
        w.key("ExpectedSourceIdentity"); w.value(ExpectedSourceIdentity);
        w.key("ExpectedDestinationIdentity"); w.value(ExpectedDestinationIdentity);
        w.key("UseBadRangeMap"); w.value(UseBadRangeMap);
        w.key("SkipKnownBadRanges"); w.value(SkipKnownBadRanges);
        w.key("UpdateBadRangeMapFromRun"); w.value(UpdateBadRangeMapFromRun);
        w.key("BadRangeMapMaxAgeDays"); w.value(BadRangeMapMaxAgeDays);
        w.key("UseExperimentalRawDiskScan"); w.value(UseExperimentalRawDiskScan);
        w.key("ScanPerformanceProfile"); w.value(to_string(ScanPerformanceProfileValue));
        w.key("ResumeJournalPathHint"); w.value(ResumeJournalPathHint);
        w.key("AllowJournalRootRemap"); w.value(AllowJournalRootRemap);
        detail::write_string_array(w, "SelectedRelativePaths", SelectedRelativePaths);
        w.key("OverwritePolicy"); w.value(to_string(OverwritePolicyValue));
        w.key("SymlinkHandling"); w.value(to_string(SymlinkHandling));
        w.key("CopyEmptyDirectories"); w.value(CopyEmptyDirectories);
        w.key("AllowRecoveredOverwriteExisting"); w.value(AllowRecoveredOverwriteExisting);
        w.key("AllowDecryptedDestination"); w.value(AllowDecryptedDestination);
        w.key("BufferSizeBytes"); w.value(BufferSizeBytes);
        w.key("UseAdaptiveBufferSizing"); w.value(UseAdaptiveBufferSizing);
        w.key("TransferEnginePolicy"); w.value(to_string(TransferEnginePolicyValue));
        w.key("MaxThroughputBytesPerSecond"); w.value(MaxThroughputBytesPerSecond);
        w.key("ParallelSmallFileWorkers"); w.value(ParallelSmallFileWorkers);
        w.key("SmallFileThresholdBytes"); w.value(SmallFileThresholdBytes);
        w.key("ParallelScanWorkers"); w.value(ParallelScanWorkers);
        w.key("WaitForMediaAvailability"); w.value(WaitForMediaAvailability);
        w.key("WaitForFileLockRelease"); w.value(WaitForFileLockRelease);
        w.key("TreatAccessDeniedAsContention"); w.value(TreatAccessDeniedAsContention);
        w.key("LockContentionProbeInterval"); w.value_literal_string(LockContentionProbeInterval.to_string());
        w.key("SourceMutationPolicy"); w.value(to_string(SourceMutationPolicyValue));
        w.key("FragileMediaMode"); w.value(FragileMediaMode);
        w.key("SkipFileOnFirstReadError"); w.value(SkipFileOnFirstReadError);
        w.key("PersistFragileSkipAcrossResume"); w.value(PersistFragileSkipAcrossResume);
        w.key("FragileFailureWindowSeconds"); w.value(FragileFailureWindowSeconds);
        w.key("FragileFailureThreshold"); w.value(FragileFailureThreshold);
        w.key("FragileCooldownSeconds"); w.value(FragileCooldownSeconds);
        w.key("MaxRetries"); w.value(MaxRetries);
        w.key("OperationTimeout"); w.value_literal_string(OperationTimeout.to_string());
        w.key("PerFileTimeout"); w.value_literal_string(PerFileTimeout.to_string());
        w.key("InitialRetryDelay"); w.value_literal_string(InitialRetryDelay.to_string());
        w.key("MaxRetryDelay"); w.value_literal_string(MaxRetryDelay.to_string());
        w.key("ResumeFromJournal"); w.value(ResumeFromJournal);
        w.key("VerifyAfterCopy"); w.value(VerifyAfterCopy);
        w.key("VerificationMode"); w.value(to_string(VerificationModeValue));
        w.key("VerificationHashAlgorithm"); w.value(to_string(VerificationHashAlgorithmValue));
        w.key("SampleVerificationChunkBytes"); w.value(SampleVerificationChunkBytes);
        w.key("SampleVerificationChunkCount"); w.value(SampleVerificationChunkCount);
        w.key("SalvageUnreadableBlocks"); w.value(SalvageUnreadableBlocks);
        w.key("SalvageFillPattern"); w.value(to_string(SalvageFillPatternValue));
        w.key("ContinueOnFileError"); w.value(ContinueOnFileError);
        w.key("PreserveTimestamps"); w.value(PreserveTimestamps);
        w.key("WorkerProcessPriorityClass"); w.value(WorkerProcessPriorityClass);
        w.key("WorkerTelemetryProfile"); w.value(to_string(WorkerTelemetryProfileValue));
        w.key("WorkerProgressEmitIntervalMs"); w.value(WorkerProgressEmitIntervalMs);
        w.key("WorkerMaxLogsPerSecond"); w.value(WorkerMaxLogsPerSecond);
        w.key("RescueFastScanChunkBytes"); w.value(RescueFastScanChunkBytes);
        w.key("RescueTrimChunkBytes"); w.value(RescueTrimChunkBytes);
        w.key("RescueScrapeChunkBytes"); w.value(RescueScrapeChunkBytes);
        w.key("RescueRetryChunkBytes"); w.value(RescueRetryChunkBytes);
        w.key("RescueSplitMinimumBytes"); w.value(RescueSplitMinimumBytes);
        w.key("RescueFastScanRetries"); w.value(RescueFastScanRetries);
        w.key("RescueTrimRetries"); w.value(RescueTrimRetries);
        w.key("RescueScrapeRetries"); w.value(RescueScrapeRetries);
        w.end_object();
    }

    static CopyJobOptions from_json(const json::Value& value) {
        CopyJobOptions o;
        const auto* obj = value.as_object();
        if (obj == nullptr) return o;
        o.OperationMode = parse_JobOperationMode(obj->find("OperationMode"), o.OperationMode);
        detail::read(obj, "SourceRoot", o.SourceRoot);
        detail::read(obj, "DestinationRoot", o.DestinationRoot);
        detail::read(obj, "ExpectedSourceIdentity", o.ExpectedSourceIdentity);
        detail::read(obj, "ExpectedDestinationIdentity", o.ExpectedDestinationIdentity);
        detail::read(obj, "UseBadRangeMap", o.UseBadRangeMap);
        detail::read(obj, "SkipKnownBadRanges", o.SkipKnownBadRanges);
        detail::read(obj, "UpdateBadRangeMapFromRun", o.UpdateBadRangeMapFromRun);
        detail::read(obj, "BadRangeMapMaxAgeDays", o.BadRangeMapMaxAgeDays);
        detail::read(obj, "UseExperimentalRawDiskScan", o.UseExperimentalRawDiskScan);
        o.ScanPerformanceProfileValue = parse_ScanPerformanceProfile(obj->find("ScanPerformanceProfile"), o.ScanPerformanceProfileValue);
        detail::read(obj, "ResumeJournalPathHint", o.ResumeJournalPathHint);
        detail::read(obj, "AllowJournalRootRemap", o.AllowJournalRootRemap);
        detail::read(obj, "SelectedRelativePaths", o.SelectedRelativePaths);
        o.OverwritePolicyValue = parse_OverwritePolicy(obj->find("OverwritePolicy"), o.OverwritePolicyValue);
        o.SymlinkHandling = parse_SymlinkHandlingMode(obj->find("SymlinkHandling"), o.SymlinkHandling);
        detail::read(obj, "CopyEmptyDirectories", o.CopyEmptyDirectories);
        detail::read(obj, "AllowRecoveredOverwriteExisting", o.AllowRecoveredOverwriteExisting);
        detail::read(obj, "AllowDecryptedDestination", o.AllowDecryptedDestination);
        detail::read(obj, "BufferSizeBytes", o.BufferSizeBytes);
        detail::read(obj, "UseAdaptiveBufferSizing", o.UseAdaptiveBufferSizing);
        o.TransferEnginePolicyValue = parse_TransferEnginePolicy(obj->find("TransferEnginePolicy"), o.TransferEnginePolicyValue);
        detail::read(obj, "MaxThroughputBytesPerSecond", o.MaxThroughputBytesPerSecond);
        detail::read(obj, "ParallelSmallFileWorkers", o.ParallelSmallFileWorkers);
        detail::read(obj, "SmallFileThresholdBytes", o.SmallFileThresholdBytes);
        detail::read(obj, "ParallelScanWorkers", o.ParallelScanWorkers);
        detail::read(obj, "WaitForMediaAvailability", o.WaitForMediaAvailability);
        detail::read(obj, "WaitForFileLockRelease", o.WaitForFileLockRelease);
        detail::read(obj, "TreatAccessDeniedAsContention", o.TreatAccessDeniedAsContention);
        detail::read(obj, "LockContentionProbeInterval", o.LockContentionProbeInterval);
        o.SourceMutationPolicyValue = parse_SourceMutationPolicy(obj->find("SourceMutationPolicy"), o.SourceMutationPolicyValue);
        detail::read(obj, "FragileMediaMode", o.FragileMediaMode);
        detail::read(obj, "SkipFileOnFirstReadError", o.SkipFileOnFirstReadError);
        detail::read(obj, "PersistFragileSkipAcrossResume", o.PersistFragileSkipAcrossResume);
        detail::read(obj, "FragileFailureWindowSeconds", o.FragileFailureWindowSeconds);
        detail::read(obj, "FragileFailureThreshold", o.FragileFailureThreshold);
        detail::read(obj, "FragileCooldownSeconds", o.FragileCooldownSeconds);
        detail::read(obj, "MaxRetries", o.MaxRetries);
        detail::read(obj, "OperationTimeout", o.OperationTimeout);
        detail::read(obj, "PerFileTimeout", o.PerFileTimeout);
        detail::read(obj, "InitialRetryDelay", o.InitialRetryDelay);
        detail::read(obj, "MaxRetryDelay", o.MaxRetryDelay);
        detail::read(obj, "ResumeFromJournal", o.ResumeFromJournal);
        detail::read(obj, "VerifyAfterCopy", o.VerifyAfterCopy);
        o.VerificationModeValue = parse_VerificationMode(obj->find("VerificationMode"), o.VerificationModeValue);
        o.VerificationHashAlgorithmValue = parse_VerificationHashAlgorithm(obj->find("VerificationHashAlgorithm"), o.VerificationHashAlgorithmValue);
        detail::read(obj, "SampleVerificationChunkBytes", o.SampleVerificationChunkBytes);
        detail::read(obj, "SampleVerificationChunkCount", o.SampleVerificationChunkCount);
        detail::read(obj, "SalvageUnreadableBlocks", o.SalvageUnreadableBlocks);
        o.SalvageFillPatternValue = parse_SalvageFillPattern(obj->find("SalvageFillPattern"), o.SalvageFillPatternValue);
        detail::read(obj, "ContinueOnFileError", o.ContinueOnFileError);
        detail::read(obj, "PreserveTimestamps", o.PreserveTimestamps);
        detail::read(obj, "WorkerProcessPriorityClass", o.WorkerProcessPriorityClass);
        o.WorkerTelemetryProfileValue = parse_WorkerTelemetryProfile(obj->find("WorkerTelemetryProfile"), o.WorkerTelemetryProfileValue);
        detail::read(obj, "WorkerProgressEmitIntervalMs", o.WorkerProgressEmitIntervalMs);
        detail::read(obj, "WorkerMaxLogsPerSecond", o.WorkerMaxLogsPerSecond);
        detail::read(obj, "RescueFastScanChunkBytes", o.RescueFastScanChunkBytes);
        detail::read(obj, "RescueTrimChunkBytes", o.RescueTrimChunkBytes);
        detail::read(obj, "RescueScrapeChunkBytes", o.RescueScrapeChunkBytes);
        detail::read(obj, "RescueRetryChunkBytes", o.RescueRetryChunkBytes);
        detail::read(obj, "RescueSplitMinimumBytes", o.RescueSplitMinimumBytes);
        detail::read(obj, "RescueFastScanRetries", o.RescueFastScanRetries);
        detail::read(obj, "RescueTrimRetries", o.RescueTrimRetries);
        detail::read(obj, "RescueScrapeRetries", o.RescueScrapeRetries);
        if (!o.UseBadRangeMap) o.SkipKnownBadRanges = false;
        return o;
    }

    bool has_unattended_media_binding() const noexcept {
        if (ExpectedSourceIdentity.empty()) return false;
        return OperationMode == JobOperationMode::ScanOnly || !ExpectedDestinationIdentity.empty();
    }

    std::string unattended_policy_issue() const {
        if (!has_unattended_media_binding()) return "source/destination media identity is not bound";
        if (WaitForMediaAvailability || WaitForFileLockRelease) {
            return "an indefinite media or file-lock wait is enabled";
        }
        if (MaxRetries > MaximumNormalRetries) return "the retry count can over-stress unstable media";
        if (SymlinkHandling == SymlinkHandlingMode::Follow) {
            return "symbolic-link following can expand the source outside its saved scope";
        }
        if (WorkerProcessPriorityClass == "High") return "high worker priority requires supervision";
        if (OperationMode == JobOperationMode::ScanOnly) {
            if (UseExperimentalRawDiskScan) return "direct NTFS extent reads require an attended run";
            if (ParallelScanWorkers > 8) return "more than eight scan workers can over-stress media";
            return std::string();
        }
        if (!VerifyAfterCopy || VerificationModeValue != VerificationMode::Full) {
            return "full destination verification is not enabled";
        }
        if (ContinueOnFileError) return "continue-on-error permits an incomplete result";
        if (SalvageUnreadableBlocks && AllowRecoveredOverwriteExisting) {
            return "recovered data may replace an existing file";
        }
        if (AllowDecryptedDestination) return "encrypted source data may be published as plaintext";
        return std::string();
    }

    bool has_safe_unattended_policy() const {
        return unattended_policy_issue().empty();
    }
};

// ---------------------------------------------------------------------------
// CopyJobResult
// ---------------------------------------------------------------------------

struct CopyJobResult {
    bool Succeeded = false;
    bool Cancelled = false;
    std::int32_t TotalFiles = 0;
    std::int32_t CompletedFiles = 0;
    std::int32_t FailedFiles = 0;
    std::int32_t RecoveredFiles = 0;
    std::int32_t SkippedFiles = 0;
    std::int64_t TotalBytes = 0;
    std::int64_t CopiedBytes = 0;
    std::int64_t WorkBytesCompleted = 0;
    std::int64_t BytesRead = 0;
    std::int64_t BytesWritten = 0;
    std::int64_t BytesVerified = 0;
    std::int64_t BytesSkipped = 0;
    std::int64_t BytesReused = 0;
    TransferEnginePolicy TransferEnginePolicyValue = TransferEnginePolicy::Auto;
    std::int64_t ElapsedMilliseconds = 0;
    double AverageBytesPerSecond = 0.0;
    std::int32_t NativeFastPathFiles = 0;
    std::int32_t ParallelNativeFastPathFiles = 0;
    std::int32_t ManagedCopyFiles = 0;
    std::int32_t NativeFallbackFiles = 0;
    std::string JournalPath;
    std::string ErrorMessage;
    // Successful runs can still have reduced assurance or metadata fidelity.
    // Keep those notices separate from ErrorMessage so callers do not mistake
    // an explicitly selected policy (for example, no verification) for an I/O
    // failure.
    std::string IntegrityNotice;
    std::string MetadataNotice;

    // Distinguishes a run that produced a partial/recovered result from a
    // hard startup or validation failure. This is derived from the existing
    // wire fields, so older journals and IPC payloads remain compatible.
    bool is_incomplete() const noexcept {
        if (Succeeded || Cancelled) return false;
        return CompletedFiles > 0 || FailedFiles > 0 || RecoveredFiles > 0 || SkippedFiles > 0 ||
               ErrorMessage.rfind("Source enumeration was incomplete", 0) == 0;
    }

    void to_json(json::Writer& w) const {
        w.begin_object();
        w.key("Succeeded"); w.value(Succeeded);
        w.key("Cancelled"); w.value(Cancelled);
        w.key("TotalFiles"); w.value(TotalFiles);
        w.key("CompletedFiles"); w.value(CompletedFiles);
        w.key("FailedFiles"); w.value(FailedFiles);
        w.key("RecoveredFiles"); w.value(RecoveredFiles);
        w.key("SkippedFiles"); w.value(SkippedFiles);
        w.key("TotalBytes"); w.value(TotalBytes);
        w.key("CopiedBytes"); w.value(CopiedBytes);
        w.key("WorkBytesCompleted"); w.value(WorkBytesCompleted);
        w.key("BytesRead"); w.value(BytesRead);
        w.key("BytesWritten"); w.value(BytesWritten);
        w.key("BytesVerified"); w.value(BytesVerified);
        w.key("BytesSkipped"); w.value(BytesSkipped);
        w.key("BytesReused"); w.value(BytesReused);
        w.key("TransferEnginePolicy"); w.value(to_string(TransferEnginePolicyValue));
        w.key("ElapsedMilliseconds"); w.value(ElapsedMilliseconds);
        w.key("AverageBytesPerSecond"); w.value(AverageBytesPerSecond);
        w.key("NativeFastPathFiles"); w.value(NativeFastPathFiles);
        w.key("ParallelNativeFastPathFiles"); w.value(ParallelNativeFastPathFiles);
        w.key("ManagedCopyFiles"); w.value(ManagedCopyFiles);
        w.key("NativeFallbackFiles"); w.value(NativeFallbackFiles);
        w.key("JournalPath"); w.value(JournalPath);
        w.key("ErrorMessage"); w.value(ErrorMessage);
        w.key("IntegrityNotice"); w.value(IntegrityNotice);
        w.key("MetadataNotice"); w.value(MetadataNotice);
        w.end_object();
    }

    static CopyJobResult from_json(const json::Value& value) {
        CopyJobResult r;
        const auto* obj = value.as_object();
        if (obj == nullptr) return r;
        detail::read(obj, "Succeeded", r.Succeeded);
        detail::read(obj, "Cancelled", r.Cancelled);
        detail::read(obj, "TotalFiles", r.TotalFiles);
        detail::read(obj, "CompletedFiles", r.CompletedFiles);
        detail::read(obj, "FailedFiles", r.FailedFiles);
        detail::read(obj, "RecoveredFiles", r.RecoveredFiles);
        detail::read(obj, "SkippedFiles", r.SkippedFiles);
        detail::read(obj, "TotalBytes", r.TotalBytes);
        detail::read(obj, "CopiedBytes", r.CopiedBytes);
        detail::read(obj, "WorkBytesCompleted", r.WorkBytesCompleted);
        detail::read(obj, "BytesRead", r.BytesRead);
        detail::read(obj, "BytesWritten", r.BytesWritten);
        detail::read(obj, "BytesVerified", r.BytesVerified);
        detail::read(obj, "BytesSkipped", r.BytesSkipped);
        detail::read(obj, "BytesReused", r.BytesReused);
        // Histories written by older workers only have CopiedBytes. Preserve a
        // useful display while keeping the new counters truthful for new runs.
        if (r.WorkBytesCompleted == 0 && r.CopiedBytes > 0) {
            r.WorkBytesCompleted = r.CopiedBytes;
        }
        r.TransferEnginePolicyValue = parse_TransferEnginePolicy(obj->find("TransferEnginePolicy"), r.TransferEnginePolicyValue);
        detail::read(obj, "ElapsedMilliseconds", r.ElapsedMilliseconds);
        detail::read(obj, "AverageBytesPerSecond", r.AverageBytesPerSecond);
        detail::read(obj, "NativeFastPathFiles", r.NativeFastPathFiles);
        detail::read(obj, "ParallelNativeFastPathFiles", r.ParallelNativeFastPathFiles);
        detail::read(obj, "ManagedCopyFiles", r.ManagedCopyFiles);
        detail::read(obj, "NativeFallbackFiles", r.NativeFallbackFiles);
        detail::read(obj, "JournalPath", r.JournalPath);
        detail::read(obj, "ErrorMessage", r.ErrorMessage);
        detail::read(obj, "IntegrityNotice", r.IntegrityNotice);
        detail::read(obj, "MetadataNotice", r.MetadataNotice);
        return r;
    }
};

// ---------------------------------------------------------------------------
// CopyProgressSnapshot — includes the computed read-only properties that
// System.Text.Json serializes from the .NET getters (OverallProgress,
// CurrentFileProgress, LastChunkBufferUtilization).
// ---------------------------------------------------------------------------

struct CopyProgressSnapshot {
    std::string SourceMediaIdentity;
    std::string DestinationMediaIdentity;
    std::string CurrentFile;
    std::int64_t CurrentFileBytesCopied = 0;
    std::int64_t CurrentFileBytesTotal = 0;
    std::int64_t TotalBytesCopied = 0;
    std::int64_t WorkBytesCompleted = 0;
    std::int64_t BytesRead = 0;
    std::int64_t BytesWritten = 0;
    std::int64_t BytesVerified = 0;
    std::int64_t BytesSkipped = 0;
    std::int64_t BytesReused = 0;
    std::int64_t TotalBytes = 0;
    std::int32_t LastChunkBytesTransferred = 0;
    std::int32_t BufferSizeBytes = 0;
    std::int32_t CompletedFiles = 0;
    std::int32_t FailedFiles = 0;
    std::int32_t RecoveredFiles = 0;
    std::int32_t SkippedFiles = 0;
    std::int32_t TotalFiles = 0;
    std::string RescuePass;
    std::int32_t RescueBadRegionCount = 0;
    std::int64_t RescueRemainingBytes = 0;
    std::int32_t ActiveFileCount = 0;
    std::int32_t ScanWorkerCount = 0;
    std::vector<std::string> ActiveFiles;

    double overall_progress() const {
        if (TotalBytes <= 0) {
            if (TotalFiles == 0) return 1.0;
            double done = static_cast<double>(CompletedFiles + FailedFiles + SkippedFiles);
            return done / static_cast<double>(std::max(1, TotalFiles));
        }
        const std::int64_t completed = WorkBytesCompleted > 0
                                           ? WorkBytesCompleted
                                           : TotalBytesCopied;
        double ratio = static_cast<double>(completed) / static_cast<double>(TotalBytes);
        return std::clamp(ratio, 0.0, 1.0);
    }

    double current_file_progress() const {
        if (CurrentFileBytesTotal <= 0) return 0.0;
        double ratio = static_cast<double>(CurrentFileBytesCopied) / static_cast<double>(CurrentFileBytesTotal);
        return std::clamp(ratio, 0.0, 1.0);
    }

    double last_chunk_buffer_utilization() const {
        if (BufferSizeBytes <= 0 || LastChunkBytesTransferred <= 0) return 0.0;
        double ratio = static_cast<double>(LastChunkBytesTransferred) / static_cast<double>(BufferSizeBytes);
        return std::clamp(ratio, 0.0, 1.0);
    }

    void to_json(json::Writer& w) const {
        w.begin_object();
        w.key("SourceMediaIdentity"); w.value(SourceMediaIdentity);
        w.key("DestinationMediaIdentity"); w.value(DestinationMediaIdentity);
        w.key("CurrentFile"); w.value(CurrentFile);
        w.key("CurrentFileBytesCopied"); w.value(CurrentFileBytesCopied);
        w.key("CurrentFileBytesTotal"); w.value(CurrentFileBytesTotal);
        w.key("TotalBytesCopied"); w.value(TotalBytesCopied);
        w.key("WorkBytesCompleted"); w.value(WorkBytesCompleted);
        w.key("BytesRead"); w.value(BytesRead);
        w.key("BytesWritten"); w.value(BytesWritten);
        w.key("BytesVerified"); w.value(BytesVerified);
        w.key("BytesSkipped"); w.value(BytesSkipped);
        w.key("BytesReused"); w.value(BytesReused);
        w.key("TotalBytes"); w.value(TotalBytes);
        w.key("LastChunkBytesTransferred"); w.value(LastChunkBytesTransferred);
        w.key("BufferSizeBytes"); w.value(BufferSizeBytes);
        w.key("CompletedFiles"); w.value(CompletedFiles);
        w.key("FailedFiles"); w.value(FailedFiles);
        w.key("RecoveredFiles"); w.value(RecoveredFiles);
        w.key("SkippedFiles"); w.value(SkippedFiles);
        w.key("TotalFiles"); w.value(TotalFiles);
        w.key("RescuePass"); w.value(RescuePass);
        w.key("RescueBadRegionCount"); w.value(RescueBadRegionCount);
        w.key("RescueRemainingBytes"); w.value(RescueRemainingBytes);
        w.key("ActiveFileCount"); w.value(ActiveFileCount);
        w.key("ScanWorkerCount"); w.value(ScanWorkerCount);
        detail::write_string_array(w, "ActiveFiles", ActiveFiles);
        w.key("OverallProgress"); w.value(overall_progress());
        w.key("CurrentFileProgress"); w.value(current_file_progress());
        w.key("LastChunkBufferUtilization"); w.value(last_chunk_buffer_utilization());
        w.end_object();
    }

    static CopyProgressSnapshot from_json(const json::Value& value) {
        CopyProgressSnapshot s;
        const auto* obj = value.as_object();
        if (obj == nullptr) return s;
        detail::read(obj, "SourceMediaIdentity", s.SourceMediaIdentity);
        detail::read(obj, "DestinationMediaIdentity", s.DestinationMediaIdentity);
        detail::read(obj, "CurrentFile", s.CurrentFile);
        detail::read(obj, "CurrentFileBytesCopied", s.CurrentFileBytesCopied);
        detail::read(obj, "CurrentFileBytesTotal", s.CurrentFileBytesTotal);
        detail::read(obj, "TotalBytesCopied", s.TotalBytesCopied);
        detail::read(obj, "WorkBytesCompleted", s.WorkBytesCompleted);
        detail::read(obj, "BytesRead", s.BytesRead);
        detail::read(obj, "BytesWritten", s.BytesWritten);
        detail::read(obj, "BytesVerified", s.BytesVerified);
        detail::read(obj, "BytesSkipped", s.BytesSkipped);
        detail::read(obj, "BytesReused", s.BytesReused);
        if (s.WorkBytesCompleted == 0 && s.TotalBytesCopied > 0) {
            s.WorkBytesCompleted = s.TotalBytesCopied;
        }
        detail::read(obj, "TotalBytes", s.TotalBytes);
        detail::read(obj, "LastChunkBytesTransferred", s.LastChunkBytesTransferred);
        detail::read(obj, "BufferSizeBytes", s.BufferSizeBytes);
        detail::read(obj, "CompletedFiles", s.CompletedFiles);
        detail::read(obj, "FailedFiles", s.FailedFiles);
        detail::read(obj, "RecoveredFiles", s.RecoveredFiles);
        detail::read(obj, "SkippedFiles", s.SkippedFiles);
        detail::read(obj, "TotalFiles", s.TotalFiles);
        detail::read(obj, "RescuePass", s.RescuePass);
        detail::read(obj, "RescueBadRegionCount", s.RescueBadRegionCount);
        detail::read(obj, "RescueRemainingBytes", s.RescueRemainingBytes);
        detail::read(obj, "ActiveFileCount", s.ActiveFileCount);
        detail::read(obj, "ScanWorkerCount", s.ScanWorkerCount);
        detail::read(obj, "ActiveFiles", s.ActiveFiles);
        return s;
    }
};

} // namespace xact::models
