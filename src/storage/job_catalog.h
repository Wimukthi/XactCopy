// Job catalog models + JobCatalogStore (port of XactCopy.Storage
// Infrastructure\JobCatalogStore.vb and the ManagedJob* models).
//
// Saved jobs can run unattended, so the catalog uses a signed snapshot envelope
// with rotations and a mirror. Legacy plain System.Text.Json catalogs remain
// readable and are upgraded on the next successful save.
#pragma once

#include "../core/models.h"
#include "stores.h"

#include <algorithm>
#include <optional>
#include <string>
#include <vector>

namespace xact::storage {

enum class ManagedJobRunStatus : std::int32_t {
    Queued = 0,
    Running = 1,
    Paused = 2,
    Completed = 3,
    Failed = 4,
    Cancelled = 5,
    Interrupted = 6,
    Incomplete = 7,
};

inline constexpr models::detail::EnumEntry<ManagedJobRunStatus> ManagedJobRunStatus_Table[] = {
    {ManagedJobRunStatus::Queued, "Queued"},
    {ManagedJobRunStatus::Running, "Running"},
    {ManagedJobRunStatus::Paused, "Paused"},
    {ManagedJobRunStatus::Completed, "Completed"},
    {ManagedJobRunStatus::Failed, "Failed"},
    {ManagedJobRunStatus::Cancelled, "Cancelled"},
    {ManagedJobRunStatus::Interrupted, "Interrupted"},
    {ManagedJobRunStatus::Incomplete, "Incomplete"},
};

inline std::string_view to_string(ManagedJobRunStatus v) {
    return models::detail::enum_name(ManagedJobRunStatus_Table, v);
}

inline ManagedJobRunStatus parse_ManagedJobRunStatus(const json::Value* v, ManagedJobRunStatus fallback) {
    return models::detail::enum_parse(ManagedJobRunStatus_Table, v, fallback);
}

namespace catalog_detail {

inline void write_optional_time(json::Writer& w, std::string_view name,
                                const std::optional<DateTimeOffset>& value) {
    w.key(name);
    if (value.has_value()) {
        w.value_literal_string(value->to_string());
    } else {
        w.value(nullptr);
    }
}

inline void read_optional_time(const json::Object* obj, std::string_view name,
                               std::optional<DateTimeOffset>& target) {
    const auto* v = obj->find(name);
    if (v == nullptr || v->is_null()) return;
    if (v->is_string()) {
        if (auto parsed = DateTimeOffset::parse(v->as_string())) target = *parsed;
    }
}

inline bool is_blank(const std::string& text) {
    return std::all_of(text.begin(), text.end(),
                       [](unsigned char c) { return c == ' ' || c == '\t' || c == '\r' || c == '\n'; });
}

inline std::string trim_copy(const std::string& text) {
    std::size_t begin = 0;
    std::size_t end = text.size();
    auto is_space = [](unsigned char c) { return c == ' ' || c == '\t' || c == '\r' || c == '\n'; };
    while (begin < end && is_space(static_cast<unsigned char>(text[begin]))) ++begin;
    while (end > begin && is_space(static_cast<unsigned char>(text[end - 1]))) --end;
    return text.substr(begin, end - begin);
}

} // namespace catalog_detail

struct ManagedJob {
    std::string JobId;
    std::string Name;
    models::CopyJobOptions Options;
    DateTimeOffset CreatedUtc = DateTimeOffset::now_utc();
    DateTimeOffset UpdatedUtc = DateTimeOffset::now_utc();

    void to_json(json::Writer& w) const {
        w.begin_object();
        w.key("JobId"); w.value(JobId);
        w.key("Name"); w.value(Name);
        w.key("Options"); Options.to_json(w);
        w.key("CreatedUtc"); w.value_literal_string(CreatedUtc.to_string());
        w.key("UpdatedUtc"); w.value_literal_string(UpdatedUtc.to_string());
        w.end_object();
    }

    static ManagedJob from_json(const json::Value& value) {
        ManagedJob job;
        const auto* obj = value.as_object();
        if (obj == nullptr) return job;
        models::detail::read(obj, "JobId", job.JobId);
        models::detail::read(obj, "Name", job.Name);
        if (const auto* options = obj->find("Options"); options != nullptr && !options->is_null()) {
            job.Options = models::CopyJobOptions::from_json(*options);
        }
        models::detail::read(obj, "CreatedUtc", job.CreatedUtc);
        models::detail::read(obj, "UpdatedUtc", job.UpdatedUtc);
        return job;
    }
};

struct ManagedJobQueueEntry {
    std::string QueueEntryId;
    std::string JobId;
    DateTimeOffset EnqueuedUtc = DateTimeOffset::now_utc();
    DateTimeOffset LastUpdatedUtc = DateTimeOffset::now_utc();
    std::string Trigger = "queued";
    std::string EnqueuedBy;
    std::int32_t AttemptCount = 0;
    std::optional<DateTimeOffset> LastAttemptUtc;
    std::string LastErrorMessage;

    void to_json(json::Writer& w) const {
        w.begin_object();
        w.key("QueueEntryId"); w.value(QueueEntryId);
        w.key("JobId"); w.value(JobId);
        w.key("EnqueuedUtc"); w.value_literal_string(EnqueuedUtc.to_string());
        w.key("LastUpdatedUtc"); w.value_literal_string(LastUpdatedUtc.to_string());
        w.key("Trigger"); w.value(Trigger);
        w.key("EnqueuedBy"); w.value(EnqueuedBy);
        w.key("AttemptCount"); w.value(AttemptCount);
        catalog_detail::write_optional_time(w, "LastAttemptUtc", LastAttemptUtc);
        w.key("LastErrorMessage"); w.value(LastErrorMessage);
        w.end_object();
    }

    static ManagedJobQueueEntry from_json(const json::Value& value) {
        ManagedJobQueueEntry entry;
        const auto* obj = value.as_object();
        if (obj == nullptr) return entry;
        models::detail::read(obj, "QueueEntryId", entry.QueueEntryId);
        models::detail::read(obj, "JobId", entry.JobId);
        models::detail::read(obj, "EnqueuedUtc", entry.EnqueuedUtc);
        models::detail::read(obj, "LastUpdatedUtc", entry.LastUpdatedUtc);
        models::detail::read(obj, "Trigger", entry.Trigger);
        models::detail::read(obj, "EnqueuedBy", entry.EnqueuedBy);
        models::detail::read(obj, "AttemptCount", entry.AttemptCount);
        catalog_detail::read_optional_time(obj, "LastAttemptUtc", entry.LastAttemptUtc);
        models::detail::read(obj, "LastErrorMessage", entry.LastErrorMessage);
        return entry;
    }
};

struct ManagedJobRun {
    std::string RunId;
    std::string JobId;
    std::string DisplayName;
    std::string SourceRoot;
    std::string DestinationRoot;
    std::string Trigger = "manual";
    std::string QueueEntryId;
    std::int32_t QueueAttempt = 0;
    ManagedJobRunStatus Status = ManagedJobRunStatus::Queued;
    DateTimeOffset StartedUtc = DateTimeOffset::now_utc();
    DateTimeOffset LastUpdatedUtc = DateTimeOffset::now_utc();
    std::optional<DateTimeOffset> FinishedUtc;
    std::string JournalPath;
    std::string Summary;
    std::string ErrorMessage;
    std::optional<models::CopyJobResult> Result;

    void to_json(json::Writer& w) const {
        w.begin_object();
        w.key("RunId"); w.value(RunId);
        w.key("JobId"); w.value(JobId);
        w.key("DisplayName"); w.value(DisplayName);
        w.key("SourceRoot"); w.value(SourceRoot);
        w.key("DestinationRoot"); w.value(DestinationRoot);
        w.key("Trigger"); w.value(Trigger);
        w.key("QueueEntryId"); w.value(QueueEntryId);
        w.key("QueueAttempt"); w.value(QueueAttempt);
        w.key("Status"); w.value(to_string(Status));
        w.key("StartedUtc"); w.value_literal_string(StartedUtc.to_string());
        w.key("LastUpdatedUtc"); w.value_literal_string(LastUpdatedUtc.to_string());
        catalog_detail::write_optional_time(w, "FinishedUtc", FinishedUtc);
        w.key("JournalPath"); w.value(JournalPath);
        w.key("Summary"); w.value(Summary);
        w.key("ErrorMessage"); w.value(ErrorMessage);
        w.key("Result");
        if (Result.has_value()) {
            Result->to_json(w);
        } else {
            w.value(nullptr);
        }
        w.end_object();
    }

    static ManagedJobRun from_json(const json::Value& value) {
        ManagedJobRun run;
        const auto* obj = value.as_object();
        if (obj == nullptr) return run;
        models::detail::read(obj, "RunId", run.RunId);
        models::detail::read(obj, "JobId", run.JobId);
        models::detail::read(obj, "DisplayName", run.DisplayName);
        models::detail::read(obj, "SourceRoot", run.SourceRoot);
        models::detail::read(obj, "DestinationRoot", run.DestinationRoot);
        models::detail::read(obj, "Trigger", run.Trigger);
        models::detail::read(obj, "QueueEntryId", run.QueueEntryId);
        models::detail::read(obj, "QueueAttempt", run.QueueAttempt);
        run.Status = parse_ManagedJobRunStatus(obj->find("Status"), run.Status);
        models::detail::read(obj, "StartedUtc", run.StartedUtc);
        models::detail::read(obj, "LastUpdatedUtc", run.LastUpdatedUtc);
        catalog_detail::read_optional_time(obj, "FinishedUtc", run.FinishedUtc);
        models::detail::read(obj, "JournalPath", run.JournalPath);
        models::detail::read(obj, "Summary", run.Summary);
        models::detail::read(obj, "ErrorMessage", run.ErrorMessage);
        if (const auto* result = obj->find("Result"); result != nullptr && !result->is_null()) {
            run.Result = models::CopyJobResult::from_json(*result);
        }
        return run;
    }
};

struct JobCatalog {
    std::int32_t SchemaVersion = 2;
    std::vector<ManagedJob> Jobs;
    std::vector<ManagedJobRun> Runs;
    std::vector<ManagedJobQueueEntry> QueueEntries;
    std::vector<std::string> QueuedJobIds; // legacy (schema v1)

    void to_json(json::Writer& w) const {
        w.begin_object();
        w.key("SchemaVersion"); w.value(SchemaVersion);
        w.key("Jobs");
        w.begin_array();
        for (const auto& job : Jobs) job.to_json(w);
        w.end_array();
        w.key("Runs");
        w.begin_array();
        for (const auto& run : Runs) run.to_json(w);
        w.end_array();
        w.key("QueueEntries");
        w.begin_array();
        for (const auto& entry : QueueEntries) entry.to_json(w);
        w.end_array();
        models::detail::write_string_array(w, "QueuedJobIds", QueuedJobIds);
        w.end_object();
    }

    static JobCatalog from_json(const json::Value& value) {
        JobCatalog catalog;
        const auto* obj = value.as_object();
        if (obj == nullptr) return catalog;
        models::detail::read(obj, "SchemaVersion", catalog.SchemaVersion);
        if (const auto* jobs = obj->find("Jobs"); jobs != nullptr && jobs->as_array() != nullptr) {
            for (const auto& item : *jobs->as_array()) {
                if (!item.is_null()) catalog.Jobs.push_back(ManagedJob::from_json(item));
            }
        }
        if (const auto* runs = obj->find("Runs"); runs != nullptr && runs->as_array() != nullptr) {
            for (const auto& item : *runs->as_array()) {
                if (!item.is_null()) catalog.Runs.push_back(ManagedJobRun::from_json(item));
            }
        }
        if (const auto* entries = obj->find("QueueEntries"); entries != nullptr && entries->as_array() != nullptr) {
            for (const auto& item : *entries->as_array()) {
                if (!item.is_null()) catalog.QueueEntries.push_back(ManagedJobQueueEntry::from_json(item));
            }
        }
        models::detail::read(obj, "QueuedJobIds", catalog.QueuedJobIds);
        return catalog;
    }
};

class JobCatalogStore {
public:
    static constexpr std::int32_t CurrentSchemaVersion = 2;
    static constexpr std::int32_t CurrentEnvelopeSchemaVersion = 1;
    static constexpr int BackupGenerationCount = 2;

    explicit JobCatalogStore(std::wstring catalog_path = std::wstring()) {
        if (catalog_path.empty()) {
            catalog_path_ = fsutil::local_app_data() + L"\\XactCopy\\jobs\\catalog.json";
            mirror_path_ =
                detail::build_mirror_path(catalog_path_, L"jobs-mirror", L"catalog");
        } else {
            catalog_path_ = std::move(catalog_path);
            // Tests and callers with an explicit catalog keep all artifacts in
            // their selected directory; the default production store uses an
            // independent LocalAppData mirror directory.
            mirror_path_ = catalog_path_ + L".mirror";
        }
    }

    const std::wstring& catalog_path() const noexcept { return catalog_path_; }
    const std::wstring& mirror_path() const noexcept { return mirror_path_; }

    JobCatalog load() const {
        const std::vector<std::wstring> candidates = build_candidates();
        const detail::SecurityContext context = create_security_context();
        for (const auto& candidate : candidates) {
            if (auto catalog = try_load_envelope(candidate, context)) return *catalog;
        }
        // Migration path for catalogs written before signed envelopes. A
        // signed backup/mirror always wins, so a damaged primary cannot be
        // downgraded to an unauthenticated payload while a trusted copy exists.
        for (const auto& candidate : candidates) {
            if (auto catalog = try_load_legacy(candidate)) return *catalog;
        }
        return JobCatalog();
    }

    bool save(JobCatalog& catalog) const {
        try {
            normalize(catalog);
            const std::string payload = serialize_catalog(catalog);
            const std::string payload_sha = crypto::sha256_hex(payload);
            const std::int64_t saved_ticks = DateTimeOffset::now_utc().utc_ticks();
            const detail::SecurityContext context = create_security_context();
            const std::string signature =
                compute_envelope_signature(payload_sha, saved_ticks, context);

            json::Writer writer(/*indented*/ true);
            writer.begin_object();
            writer.key("EnvelopeSchemaVersion");
            writer.value(CurrentEnvelopeSchemaVersion);
            writer.key("SavedUtcTicks");
            writer.value(saved_ticks);
            writer.key("PayloadSha256");
            writer.value(payload_sha);
            writer.key("Signature");
            writer.value(signature);
            writer.key("Payload");
            catalog.to_json(writer);
            writer.end_object();
            const std::string envelope = writer.take();

            const std::wstring primary_directory = fsutil::get_directory_name(catalog_path_);
            const std::wstring mirror_directory = fsutil::get_directory_name(mirror_path_);
            if (!primary_directory.empty()) fsutil::create_directories(primary_directory);
            if (!mirror_directory.empty()) fsutil::create_directories(mirror_directory);
            detail::rotate_backups(catalog_path_, BackupGenerationCount);
            if (!fsutil::write_atomic_bytes(catalog_path_, envelope)) return false;
            detail::rotate_backups(mirror_path_, BackupGenerationCount);
            // The primary is authoritative. A failed mirror refresh must not
            // turn a successfully committed catalog update into a false save
            // failure; the previous mirror rotation remains a load candidate.
            (void)fsutil::write_atomic_bytes(mirror_path_, envelope);
            return true;
        } catch (const std::exception&) {
            return false;
        }
    }

    // Mirrors JobCatalogStore.NormalizeCatalog (defaults, legacy queue
    // migration, duplicate queue-entry removal, schema upgrade).
    static void normalize(JobCatalog& catalog) {
        catalog.SchemaVersion = std::max(1, catalog.SchemaVersion);

        for (auto& job : catalog.Jobs) {
            if (catalog_detail::is_blank(job.JobId)) job.JobId = detail::new_guid_n();
        }

        for (auto& run : catalog.Runs) {
            if (catalog_detail::is_blank(run.RunId)) run.RunId = detail::new_guid_n();
            if (run.QueueAttempt < 0) run.QueueAttempt = 0;
        }

        for (auto& entry : catalog.QueueEntries) {
            if (catalog_detail::is_blank(entry.QueueEntryId)) entry.QueueEntryId = detail::new_guid_n();
            if (entry.EnqueuedUtc == DateTimeOffset::min_value()) entry.EnqueuedUtc = DateTimeOffset::now_utc();
            if (entry.LastUpdatedUtc == DateTimeOffset::min_value()) entry.LastUpdatedUtc = entry.EnqueuedUtc;
            if (catalog_detail::is_blank(entry.Trigger)) entry.Trigger = "queued";
            if (entry.AttemptCount < 0) entry.AttemptCount = 0;
        }

        migrate_legacy_queue(catalog);
        remove_duplicate_queue_entries(catalog);
        catalog.SchemaVersion = CurrentSchemaVersion;
    }

private:
    static std::string serialize_catalog(const JobCatalog& catalog) {
        json::Writer writer(/*indented*/ true);
        catalog.to_json(writer);
        return writer.take();
    }

    detail::SecurityContext create_security_context() const {
        detail::SecurityContext context;
        context.path_fingerprint = detail::fingerprint_for_path(catalog_path_);
        context.hmac_key = detail::catalog_key_cache().get_or_create();
        return context;
    }

    static std::string compute_envelope_signature(
        const std::string& payload_sha, std::int64_t saved_ticks,
        const detail::SecurityContext& context) {
        const std::string authenticated =
            payload_sha + "|" + std::to_string(saved_ticks) + "|" + context.path_fingerprint;
        return crypto::to_base64(crypto::hmac_sha256(context.hmac_key, authenticated));
    }

    std::vector<std::wstring> build_candidates() const {
        std::vector<std::wstring> candidates;
        detail::add_unique_path(candidates, catalog_path_);
        for (int generation = 1; generation <= BackupGenerationCount; ++generation) {
            detail::add_unique_path(candidates, detail::backup_path(catalog_path_, generation));
        }
        detail::add_unique_path(candidates, mirror_path_);
        for (int generation = 1; generation <= BackupGenerationCount; ++generation) {
            detail::add_unique_path(candidates, detail::backup_path(mirror_path_, generation));
        }
        return candidates;
    }

    static std::optional<json::Value> read_json(const std::wstring& path) {
        auto bytes = fsutil::read_all_bytes(path, 128ULL * 1024 * 1024);
        if (!bytes.has_value() || bytes->empty()) return std::nullopt;
        std::string text(reinterpret_cast<const char*>(bytes->data()), bytes->size());
        if (text.size() >= 3 && static_cast<unsigned char>(text[0]) == 0xEF &&
            static_cast<unsigned char>(text[1]) == 0xBB &&
            static_cast<unsigned char>(text[2]) == 0xBF) {
            text.erase(0, 3);
        }
        try {
            return json::parse(text);
        } catch (const std::exception&) {
            return std::nullopt;
        }
    }

    static std::optional<JobCatalog> try_load_envelope(
        const std::wstring& path, const detail::SecurityContext& context) {
        auto parsed = read_json(path);
        const json::Object* object = parsed.has_value() ? parsed->as_object() : nullptr;
        if (object == nullptr) return std::nullopt;

        std::int32_t envelope_schema = 0;
        std::int64_t saved_ticks = 0;
        std::string payload_sha;
        std::string signature;
        models::detail::read(object, "EnvelopeSchemaVersion", envelope_schema);
        models::detail::read(object, "SavedUtcTicks", saved_ticks);
        models::detail::read(object, "PayloadSha256", payload_sha);
        models::detail::read(object, "Signature", signature);
        const json::Value* payload_value = object->find("Payload");
        if (envelope_schema != CurrentEnvelopeSchemaVersion || saved_ticks <= 0 ||
            payload_sha.empty() || signature.empty() || payload_value == nullptr ||
            !payload_value->is_object()) {
            return std::nullopt;
        }

        JobCatalog catalog = JobCatalog::from_json(*payload_value);
        const std::string actual_sha = crypto::sha256_hex(serialize_catalog(catalog));
        if (!crypto::secure_equals(actual_sha, payload_sha)) return std::nullopt;
        const std::string expected_signature =
            compute_envelope_signature(actual_sha, saved_ticks, context);
        if (!crypto::secure_equals(expected_signature, signature)) return std::nullopt;
        normalize(catalog);
        return catalog;
    }

    static std::optional<JobCatalog> try_load_legacy(const std::wstring& path) {
        auto parsed = read_json(path);
        const json::Object* object = parsed.has_value() ? parsed->as_object() : nullptr;
        if (object == nullptr || object->find("EnvelopeSchemaVersion") != nullptr ||
            (object->find("Jobs") == nullptr && object->find("Runs") == nullptr &&
             object->find("QueueEntries") == nullptr && object->find("QueuedJobIds") == nullptr)) {
            return std::nullopt;
        }
        JobCatalog catalog = JobCatalog::from_json(*parsed);
        normalize(catalog);
        return catalog;
    }

    static void migrate_legacy_queue(JobCatalog& catalog) {
        if (!catalog.QueueEntries.empty()) return;
        if (catalog.QueuedJobIds.empty()) return;

        DateTimeOffset now_utc = DateTimeOffset::now_utc();
        for (const auto& queued_job_id : catalog.QueuedJobIds) {
            if (catalog_detail::is_blank(queued_job_id)) continue;
            ManagedJobQueueEntry entry;
            entry.QueueEntryId = detail::new_guid_n();
            entry.JobId = catalog_detail::trim_copy(queued_job_id);
            entry.EnqueuedUtc = now_utc;
            entry.LastUpdatedUtc = now_utc;
            entry.Trigger = "legacy-migration";
            entry.EnqueuedBy = "migration";
            catalog.QueueEntries.push_back(std::move(entry));
        }
    }

    static void remove_duplicate_queue_entries(JobCatalog& catalog) {
        if (catalog.QueueEntries.empty()) return;

        std::vector<ManagedJobQueueEntry> deduped;
        deduped.reserve(catalog.QueueEntries.size());
        std::vector<std::string> seen_ids;
        for (auto& entry : catalog.QueueEntries) {
            if (catalog_detail::is_blank(entry.JobId)) continue;
            bool duplicate = false;
            for (const auto& seen : seen_ids) {
                if (models::detail::equals_ignore_case(seen, entry.QueueEntryId)) {
                    duplicate = true;
                    break;
                }
            }
            if (duplicate) continue;
            seen_ids.push_back(entry.QueueEntryId);
            deduped.push_back(std::move(entry));
        }
        catalog.QueueEntries = std::move(deduped);
    }

    std::wstring catalog_path_;
    std::wstring mirror_path_;
};

} // namespace xact::storage
