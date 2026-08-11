// -----------------------------------------------------------------------------
// File: src\ui\settings.h
// Purpose: Native settings access over the .NET UI's settings.json. Loads the
//          file as a JSON DOM, exposes typed getters/setters for the keys the
//          native UI uses, and saves atomically while PRESERVING every other
//          key untouched — so the file stays fully compatible with the .NET
//          build in both directions.
// -----------------------------------------------------------------------------

#pragma once

#include <algorithm>
#include <optional>
#include <stdexcept>
#include <string>

#include "../core/json.h"
#include "../core/models.h"
#include "../storage/stores.h"

namespace xact::ui {

inline int verification_combo_index(const models::CopyJobOptions& options) {
    if (!options.VerifyAfterCopy) return 0;
    return options.VerificationModeValue == models::VerificationMode::Sampled ? 1 : 2;
}

class AppSettings {
public:
    static std::wstring default_path() {
        return storage::fsutil::local_app_data() + L"\\XactCopy\\settings.json";
    }

    explicit AppSettings(std::wstring path = default_path()) : path_(std::move(path)) { load(); }

    void load() {
        root_ = json::Object{};
        load_warning_.clear();
        auto bytes = storage::fsutil::read_all_bytes(path_, 16ULL * 1024 * 1024);
        if (!bytes.has_value()) {
            const DWORD read_error = GetLastError();
            if (GetFileAttributesW(path_.c_str()) != INVALID_FILE_ATTRIBUTES) {
                load_warning_ = "Settings file exists but could not be read; defaults were loaded "
                                "(Win32 " + std::to_string(read_error) + ").";
            }
            return;
        }
        try {
            if (bytes->empty()) throw std::runtime_error("settings file is empty");
            std::string_view text(reinterpret_cast<const char*>(bytes->data()), bytes->size());
            // Tolerate a UTF-8 BOM (some editors add one); the JSON parser would
            // otherwise choke and the whole file would silently reset to defaults.
            if (text.size() >= 3 && static_cast<unsigned char>(text[0]) == 0xEF &&
                static_cast<unsigned char>(text[1]) == 0xBB &&
                static_cast<unsigned char>(text[2]) == 0xBF) {
                text.remove_prefix(3);
            }
            json::Value parsed = json::parse(text);
            const auto* obj = parsed.as_object();
            if (obj == nullptr) throw std::runtime_error("settings root must be a JSON object");
            root_ = *obj;
        } catch (const std::exception& ex) {
            root_ = json::Object{};
            const std::wstring quarantine =
                path_ + L".corrupt-" + std::to_wstring(GetTickCount64());
            bool preserved = MoveFileExW(path_.c_str(), quarantine.c_str(),
                                         MOVEFILE_WRITE_THROUGH) != FALSE;
            if (!preserved) {
                preserved = CopyFileW(path_.c_str(), quarantine.c_str(), TRUE) != FALSE;
            }
            load_warning_ = "Settings JSON was invalid and defaults were loaded (" +
                            std::string(ex.what()) + ").";
            if (preserved) {
                load_warning_ += " The original file was preserved as " +
                                 storage::fsutil::wide_to_utf8(quarantine) + ".";
            }
        }
    }

    bool save() {
        last_save_error_.clear();
        json::Writer w(/*indented*/ true);
        write_object(w, root_);
        std::string text = w.take();

        std::wstring directory = storage::fsutil::get_directory_name(path_);
        if (!directory.empty()) storage::fsutil::create_directories(directory);
        if (!storage::fsutil::write_atomic_bytes(path_, text)) {
            last_save_error_ = "Unable to save settings atomically to " +
                               storage::fsutil::wide_to_utf8(path_) +
                               " (Win32 " + std::to_string(GetLastError()) + ").";
            return false;
        }
        return true;
    }

    // Apply a group of related settings as one in-memory/durable transaction.
    // A later unrelated save (for example window placement) must not persist
    // values from an earlier failed or cancelled settings operation.
    template <typename Update>
    bool update_and_save(Update&& update) {
        json::Object previous = root_;
        try {
            update(*this);
        } catch (...) {
            root_ = std::move(previous);
            throw;
        }
        if (save()) return true;
        root_ = std::move(previous);
        return false;
    }

    const std::string& load_warning() const noexcept { return load_warning_; }
    const std::string& last_save_error() const noexcept { return last_save_error_; }
    const std::wstring& path() const noexcept { return path_; }

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
        const std::int64_t parsed = value->as_int64(fallback);
        return static_cast<std::int32_t>(
            std::clamp<std::int64_t>(parsed, minimum, maximum));
    }

    bool get_bool(std::string_view key, bool fallback) const {
        const json::Value* value = root_.find(key);
        if (value == nullptr || !value->is_bool()) return fallback;
        return value->as_bool();
    }

    bool contains(std::string_view key) const noexcept { return root_.find(key) != nullptr; }

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
            {{"follow-internal", SymlinkHandlingMode::FollowInternal},
             {"follow", SymlinkHandlingMode::Follow}}, SymlinkHandlingMode::Skip);
        options.AllowRecoveredOverwriteExisting =
            get_bool("DefaultAllowRecoveredOverwriteExisting", false);
        options.AllowDecryptedDestination =
            get_bool("DefaultAllowDecryptedDestination", false);
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
        // Missing settings use an integrity-first profile. Existing keys are
        // read verbatim, so this changes new-install fallbacks without
        // silently rewriting a user's saved policy.
        options.SalvageUnreadableBlocks = get_bool("DefaultSalvageUnreadableBlocks", false);
        options.ContinueOnFileError = get_bool("DefaultContinueOnFileError", false);
        options.VerifyAfterCopy = get_bool("DefaultVerifyAfterCopy", true);
        options.UseBadRangeMap = get_bool("DefaultUseBadRangeMap", false);
        // Skipping a known range only has meaning when the map is enabled. Keep
        // the in-memory defaults coherent even if an older settings file has
        // the two independent keys set to a contradictory combination.
        options.SkipKnownBadRanges =
            options.UseBadRangeMap && get_bool("DefaultSkipKnownBadRanges", false);
        options.UpdateBadRangeMapFromRun = get_bool("DefaultUpdateBadRangeMapFromRun", false);
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
        options.MaxRetries = get_int("DefaultMaxRetries", 0, 32, 2);
        options.OperationTimeout =
            time::TimeSpan::from_seconds(get_int("DefaultOperationTimeoutSeconds", 1, 3600, 10));
        options.PerFileTimeout =
            time::TimeSpan::from_seconds(get_int("DefaultPerFileTimeoutSeconds", 0, 86400, 0));
        options.MaxThroughputBytesPerSecond =
            static_cast<std::int64_t>(get_int("DefaultMaxThroughputMbPerSecond", 0, 4096, 0)) *
            1024 * 1024;
        std::int32_t parallel_small = get_int("DefaultParallelSmallFileWorkers", 0, 64, 0);
        // Zero is retained as the UI's documented "auto" value; the engine
        // resolves it against the processor count after it knows the file set.
        options.ParallelSmallFileWorkers = parallel_small;
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
        options.RescueFastScanRetries = get_int("DefaultRescueFastScanRetries", 0, 32, 0);
        options.RescueTrimRetries = get_int("DefaultRescueTrimRetries", 0, 32, 1);
        options.RescueScrapeRetries = get_int("DefaultRescueScrapeRetries", 0, 32, 2);
        options.SampleVerificationChunkBytes =
            get_int("DefaultSampleVerificationChunkKb", 32, 4096, 128) * 1024;
        options.SampleVerificationChunkCount =
            get_int("DefaultSampleVerificationChunkCount", 1, 64, 3);
        return options;
    }

    // Persist only the options represented by the main-window run controls.
    // Paths, operation mode, journal identity, and recovery state deliberately
    // remain per-run/per-job instead of becoming global defaults.
    bool save_run_defaults(const models::CopyJobOptions& options) {
        using namespace models;
        json::Object previous = root_;

        auto overwrite_name = [](OverwritePolicy value) {
            switch (value) {
                case OverwritePolicy::SkipExisting: return "skip-existing";
                case OverwritePolicy::OverwriteIfSourceNewer: return "overwrite-if-newer";
                case OverwritePolicy::Ask: return "ask";
                case OverwritePolicy::Overwrite:
                default: return "overwrite";
            }
        };
        auto engine_name = [](TransferEnginePolicy value) {
            switch (value) {
                case TransferEnginePolicy::ManagedRescue: return "managed";
                case TransferEnginePolicy::NativeFast: return "native";
                case TransferEnginePolicy::Auto:
                default: return "auto";
            }
        };

        if (options.OperationMode == JobOperationMode::Copy) {
            set_string("DefaultOverwritePolicy", overwrite_name(options.OverwritePolicyValue));
            set_string("DefaultTransferEnginePolicy",
                       engine_name(options.TransferEnginePolicyValue));
            set_bool("DefaultSalvageUnreadableBlocks", options.SalvageUnreadableBlocks);

            const bool verify = options.VerifyAfterCopy &&
                                options.VerificationModeValue != VerificationMode::None;
            set_bool("DefaultVerifyAfterCopy", verify);
            // Disabling verification preserves the selected mode so turning it
            // back on restores the user's last comparison policy.
            if (verify) {
                set_string("DefaultVerificationMode",
                           options.VerificationModeValue == VerificationMode::Sampled
                               ? "sampled"
                               : "full");
            }
        } else {
            const char* scan_profile =
                options.ScanPerformanceProfileValue == ScanPerformanceProfile::Fast
                    ? "fast"
                    : options.ScanPerformanceProfileValue == ScanPerformanceProfile::Precise
                          ? "precise"
                          : "auto";
            set_string("DefaultScanPerformanceProfile", scan_profile);
            set_bool("DefaultUseExperimentalRawDiskScan", options.UseExperimentalRawDiskScan);
        }
        set_bool("DefaultResumeFromJournal", options.ResumeFromJournal);
        set_bool("DefaultContinueOnFileError", options.ContinueOnFileError);
        set_bool("DefaultUseBadRangeMap", options.UseBadRangeMap);
        set_bool("DefaultSkipKnownBadRanges",
                 options.UseBadRangeMap && options.SkipKnownBadRanges);
        set_bool("DefaultUpdateBadRangeMapFromRun", options.UpdateBadRangeMapFromRun);
        set_bool("DefaultUseAdaptiveBuffer", options.UseAdaptiveBufferSizing);
        set_bool("DefaultWaitForMediaAvailability", options.WaitForMediaAvailability);
        set_bool("DefaultFragileMediaMode", options.FragileMediaMode);

        const std::int32_t buffer_mb = std::clamp(
            options.BufferSizeBytes / (1024 * 1024), static_cast<std::int32_t>(1),
            static_cast<std::int32_t>(256));
        set_int("DefaultBufferSizeMb", buffer_mb);
        set_int("DefaultMaxRetries", std::clamp(options.MaxRetries, 0, 32));
        const std::int64_t timeout_seconds = options.OperationTimeout.ticks / 10000000LL;
        set_int("DefaultOperationTimeoutSeconds",
                static_cast<std::int32_t>(std::clamp<std::int64_t>(
                    timeout_seconds, 1, 3600)));
        if (save()) return true;
        root_ = std::move(previous);
        return false;
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
    std::string load_warning_;
    std::string last_save_error_;

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
