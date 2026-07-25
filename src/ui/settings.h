// -----------------------------------------------------------------------------
// File: cpp\src\ui\settings.h
// Purpose: Native settings access over the .NET UI's settings.json. Loads the
//          file as a JSON DOM, exposes typed getters/setters for the keys the
//          native UI uses, and saves atomically while PRESERVING every other
//          key untouched — so the file stays fully compatible with the .NET
//          build in both directions.
// -----------------------------------------------------------------------------

#pragma once

#include <algorithm>
#include <optional>
#include <string>

#include "../core/json.h"
#include "../core/models.h"
#include "../storage/stores.h"

namespace xact::ui {

class AppSettings {
public:
    static std::wstring default_path() {
        return storage::fsutil::local_app_data() + L"\\XactCopy\\settings.json";
    }

    explicit AppSettings(std::wstring path = default_path()) : path_(std::move(path)) { load(); }

    void load() {
        root_ = json::Object{};
        auto bytes = storage::fsutil::read_all_bytes(path_);
        if (!bytes.has_value() || bytes->empty()) return;
        try {
            std::string_view text(reinterpret_cast<const char*>(bytes->data()), bytes->size());
            // Tolerate a UTF-8 BOM (some editors add one); the JSON parser would
            // otherwise choke and the whole file would silently reset to defaults.
            if (text.size() >= 3 && static_cast<unsigned char>(text[0]) == 0xEF &&
                static_cast<unsigned char>(text[1]) == 0xBB &&
                static_cast<unsigned char>(text[2]) == 0xBF) {
                text.remove_prefix(3);
            }
            json::Value parsed = json::parse(text);
            if (const auto* obj = parsed.as_object()) root_ = *obj;
        } catch (const std::exception&) {
            root_ = json::Object{};
        }
    }

    void save() {
        json::Writer w(/*indented*/ true);
        write_object(w, root_);
        std::string text = w.take();

        std::wstring directory = storage::fsutil::get_directory_name(path_);
        if (!directory.empty()) storage::fsutil::create_directories(directory);
        storage::fsutil::write_atomic_bytes(path_, text);
    }

    // ---- Typed access ------------------------------------------------------

    std::string get_string(std::string_view key, const std::string& fallback) const {
        const json::Value* value = root_.find(key);
        if (value == nullptr || !value->is_string()) return fallback;
        std::string text = value->as_string();
        std::size_t begin = 0, end = text.size();
        while (begin < end && (text[begin] == ' ' || text[begin] == '\t')) ++begin;
        while (end > begin && (text[end - 1] == ' ' || text[end - 1] == '\t')) --end;
        text = text.substr(begin, end - begin);
        return text.empty() ? fallback : text;
    }

    std::int32_t get_int(std::string_view key, std::int32_t minimum, std::int32_t maximum,
                         std::int32_t fallback) const {
        const json::Value* value = root_.find(key);
        if (value == nullptr || !value->is_number()) return fallback;
        return std::clamp(value->as_int32(), minimum, maximum);
    }

    bool get_bool(std::string_view key, bool fallback) const {
        const json::Value* value = root_.find(key);
        if (value == nullptr || !value->is_bool()) return fallback;
        return value->as_bool();
    }

    void set_string(const std::string& key, const std::string& value) {
        set_value(key, json::Value(value));
    }
    void set_int(const std::string& key, std::int32_t value) {
        set_value(key, json::Value(static_cast<std::int64_t>(value)));
    }
    void set_bool(const std::string& key, bool value) { set_value(key, json::Value(value)); }

    std::string theme() const { return get_string("Theme", "dark"); }

    // ---- Job option defaults (mirrors the MainForm settings->options map) --

    models::CopyJobOptions build_default_options() const {
        using namespace models;
        CopyJobOptions options;

        options.OverwritePolicyValue = parse_kebab<OverwritePolicy>(
            get_string("DefaultOverwritePolicy", "overwrite"),
            {{"skip-existing", OverwritePolicy::SkipExisting},
             {"overwrite-if-newer", OverwritePolicy::OverwriteIfSourceNewer},
             {"ask", OverwritePolicy::Ask}},
            OverwritePolicy::Overwrite);
        options.TransferEnginePolicyValue = parse_kebab<TransferEnginePolicy>(
            get_string("DefaultTransferEnginePolicy", "auto"),
            {{"managed", TransferEnginePolicy::ManagedRescue},
             {"native", TransferEnginePolicy::NativeFast}},
            TransferEnginePolicy::Auto);
        options.SymlinkHandling = parse_kebab<SymlinkHandlingMode>(
            get_string("DefaultSymlinkHandling", "skip"),
            {{"follow", SymlinkHandlingMode::Follow}}, SymlinkHandlingMode::Skip);
        options.VerificationModeValue = parse_kebab<VerificationMode>(
            get_string("DefaultVerificationMode", "full"),
            {{"sampled", VerificationMode::Sampled}, {"full", VerificationMode::Full}},
            VerificationMode::None);
        options.VerificationHashAlgorithmValue = parse_kebab<VerificationHashAlgorithm>(
            get_string("DefaultVerificationHashAlgorithm", "sha256"),
            {{"sha512", VerificationHashAlgorithm::Sha512}}, VerificationHashAlgorithm::Sha256);
        options.SalvageFillPatternValue = parse_kebab<SalvageFillPattern>(
            get_string("DefaultSalvageFillPattern", "zero"),
            {{"ones", SalvageFillPattern::Ones}, {"random", SalvageFillPattern::Random}},
            SalvageFillPattern::Zero);
        options.SourceMutationPolicyValue = parse_kebab<SourceMutationPolicy>(
            get_string("DefaultSourceMutationPolicy", "fail-file"),
            {{"skip-file", SourceMutationPolicy::SkipFile},
             {"wait-for-reappearance", SourceMutationPolicy::WaitForReappearance}},
            SourceMutationPolicy::FailFile);
        options.ScanPerformanceProfileValue = parse_kebab<ScanPerformanceProfile>(
            get_string("DefaultScanPerformanceProfile", "auto"),
            {{"fast", ScanPerformanceProfile::Fast}, {"precise", ScanPerformanceProfile::Precise}},
            ScanPerformanceProfile::Auto);
        options.WorkerTelemetryProfileValue = parse_kebab<WorkerTelemetryProfile>(
            get_string("WorkerTelemetryProfile", "normal"),
            {{"verbose", WorkerTelemetryProfile::Verbose},
             {"debug", WorkerTelemetryProfile::Debug}},
            WorkerTelemetryProfile::Normal);

        options.ResumeFromJournal = get_bool("DefaultResumeFromJournal", true);
        options.SalvageUnreadableBlocks = get_bool("DefaultSalvageUnreadableBlocks", true);
        options.ContinueOnFileError = get_bool("DefaultContinueOnFileError", true);
        options.VerifyAfterCopy = get_bool("DefaultVerifyAfterCopy", false);
        options.UseBadRangeMap = get_bool("DefaultUseBadRangeMap", true);
        options.SkipKnownBadRanges = get_bool("DefaultSkipKnownBadRanges", true);
        options.UpdateBadRangeMapFromRun = get_bool("DefaultUpdateBadRangeMapFromRun", true);
        options.BadRangeMapMaxAgeDays = get_int("DefaultBadRangeMapMaxAgeDays", 0, 3650, 30);
        options.UseExperimentalRawDiskScan = get_bool("DefaultUseExperimentalRawDiskScan", false);
        options.UseAdaptiveBufferSizing = get_bool("DefaultUseAdaptiveBuffer", false);
        options.WaitForMediaAvailability = get_bool("DefaultWaitForMediaAvailability", false);
        options.WaitForFileLockRelease = get_bool("DefaultWaitForFileLockRelease", false);
        options.TreatAccessDeniedAsContention = get_bool("DefaultTreatAccessDeniedAsContention", false);
        options.LockContentionProbeInterval = time::TimeSpan::from_milliseconds(
            get_int("DefaultLockContentionProbeIntervalMs", 100, 10000, 500));
        options.FragileMediaMode = get_bool("DefaultFragileMediaMode", false);
        options.SkipFileOnFirstReadError = get_bool("DefaultSkipFileOnFirstReadError", true);
        options.PersistFragileSkipAcrossResume =
            get_bool("DefaultPersistFragileSkipsAcrossResume", true);
        options.FragileFailureWindowSeconds =
            get_int("DefaultFragileFailureWindowSeconds", 1, 3600, 20);
        options.FragileFailureThreshold = get_int("DefaultFragileFailureThreshold", 1, 1000, 3);
        options.FragileCooldownSeconds = get_int("DefaultFragileCooldownSeconds", 0, 600, 6);
        options.BufferSizeBytes = get_int("DefaultBufferSizeMb", 1, 256, 4) * 1024 * 1024;
        options.MaxRetries = get_int("DefaultMaxRetries", 0, 1000, 12);
        options.OperationTimeout =
            time::TimeSpan::from_seconds(get_int("DefaultOperationTimeoutSeconds", 1, 3600, 10));
        options.PerFileTimeout =
            time::TimeSpan::from_seconds(get_int("DefaultPerFileTimeoutSeconds", 0, 86400, 0));
        options.MaxThroughputBytesPerSecond =
            static_cast<std::int64_t>(get_int("DefaultMaxThroughputMbPerSecond", 0, 4096, 0)) *
            1024 * 1024;
        std::int32_t parallel_small = get_int("DefaultParallelSmallFileWorkers", 0, 64, 0);
        options.ParallelSmallFileWorkers = parallel_small <= 0 ? 1 : parallel_small;
        options.ParallelScanWorkers = get_int("DefaultParallelScanWorkers", 0, 64, 0);
        options.SmallFileThresholdBytes =
            get_int("DefaultSmallFileThresholdKb", 4, 1048576, 256) * 1024;
        options.WorkerProcessPriorityClass =
            get_string("DefaultWorkerProcessPriorityClass", "Normal");
        options.PreserveTimestamps = get_bool("DefaultPreserveTimestamps", true);
        options.CopyEmptyDirectories = get_bool("DefaultCopyEmptyDirectories", true);
        options.WorkerProgressEmitIntervalMs = get_int("WorkerProgressIntervalMs", 20, 1000, 75);
        options.WorkerMaxLogsPerSecond = get_int("WorkerMaxLogsPerSecond", 0, 5000, 100);
        options.RescueFastScanChunkBytes = get_int("DefaultRescueFastScanChunkKb", 0, 262144, 0) * 1024;
        options.RescueTrimChunkBytes = get_int("DefaultRescueTrimChunkKb", 0, 262144, 0) * 1024;
        options.RescueScrapeChunkBytes = get_int("DefaultRescueScrapeChunkKb", 0, 262144, 0) * 1024;
        options.RescueRetryChunkBytes = get_int("DefaultRescueRetryChunkKb", 0, 262144, 0) * 1024;
        options.RescueSplitMinimumBytes = get_int("DefaultRescueSplitMinimumKb", 0, 65536, 0) * 1024;
        options.RescueFastScanRetries = get_int("DefaultRescueFastScanRetries", 0, 1000, 0);
        options.RescueTrimRetries = get_int("DefaultRescueTrimRetries", 0, 1000, 1);
        options.RescueScrapeRetries = get_int("DefaultRescueScrapeRetries", 0, 1000, 2);
        options.SampleVerificationChunkBytes =
            get_int("DefaultSampleVerificationChunkKb", 32, 4096, 128) * 1024;
        options.SampleVerificationChunkCount =
            get_int("DefaultSampleVerificationChunkCount", 1, 64, 3);
        return options;
    }

    bool prompt_resume_after_crash() const { return get_bool("PromptResumeAfterCrash", true); }
    bool auto_resume_after_crash() const { return get_bool("AutoResumeAfterCrash", false); }
    bool keep_resume_prompt_until_resolved() const {
        return get_bool("KeepResumePromptUntilResolved", true);
    }
    bool enable_recovery_autostart() const { return get_bool("EnableRecoveryAutostart", true); }
    std::int32_t recovery_touch_interval_seconds() const {
        return get_int("RecoveryTouchIntervalSeconds", 1, 60, 2);
    }

private:
    std::wstring path_;
    json::Object root_;

    void set_value(const std::string& key, json::Value value) {
        for (auto& member : root_.members) {
            if (models::detail::equals_ignore_case(member.first, key)) {
                member.second = std::move(value);
                return;
            }
        }
        root_.add(key, std::move(value));
    }

    template <typename TEnum>
    static TEnum parse_kebab(const std::string& text,
                             std::initializer_list<std::pair<const char*, TEnum>> table,
                             TEnum fallback) {
        for (const auto& [name, value] : table) {
            if (models::detail::equals_ignore_case(text, name)) return value;
        }
        return fallback;
    }

    static void write_value(json::Writer& w, const json::Value& value) {
        switch (value.kind()) {
            case json::Kind::Null: w.value(nullptr); break;
            case json::Kind::Boolean: w.value(value.as_bool()); break;
            case json::Kind::Int64: w.value(value.as_int64()); break;
            case json::Kind::Double: w.value(value.as_double()); break;
            case json::Kind::String: w.value(value.as_string()); break;
            case json::Kind::ArrayKind: {
                w.begin_array();
                for (const auto& item : *value.as_array()) write_value(w, item);
                w.end_array();
                break;
            }
            case json::Kind::ObjectKind: write_object(w, *value.as_object()); break;
        }
    }

    static void write_object(json::Writer& w, const json::Object& object) {
        w.begin_object();
        for (const auto& [key, value] : object.members) {
            w.key(key);
            write_value(w, value);
        }
        w.end_object();
    }
};

} // namespace xact::ui
