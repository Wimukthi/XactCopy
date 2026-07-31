// -----------------------------------------------------------------------------
// File: tests\test_storage.cpp
// Purpose: Storage-layer tests: crypto vectors, indented-writer byte parity
//          against the goldens, journal compression and maintenance, native
//          store round-trips and tamper fallback, and job catalog round-trip.
// Usage:   xactcopy_storage_tests.exe [goldenDir]
// -----------------------------------------------------------------------------

#include <chrono>
#include <cstdio>
#include <fstream>
#include <sstream>
#include <string>
#include <vector>

#include "../src/storage/job_catalog.h"
#include "../src/storage/stores.h"

namespace {

using namespace xact;
using storage::fsutil::utf8_to_wide;

int g_failures = 0;
int g_checks = 0;

void check(bool condition, const std::string& label) {
    ++g_checks;
    if (condition) {
        std::printf("  ok: %s\n", label.c_str());
    } else {
        ++g_failures;
        std::printf("  FAIL: %s\n", label.c_str());
    }
}

void check_equal(const std::string& actual, const std::string& expected, const std::string& label) {
    ++g_checks;
    if (actual != expected) {
        ++g_failures;
        std::printf("  FAIL: %s\n    expected: %s\n    actual:   %s\n",
                    label.c_str(), expected.c_str(), actual.c_str());
    } else {
        std::printf("  ok: %s\n", label.c_str());
    }
}

std::string read_file(const std::string& path) {
    std::ifstream stream(path, std::ios::binary);
    if (!stream) return std::string();
    std::ostringstream buffer;
    buffer << stream.rdbuf();
    return buffer.str();
}

// ---------------------------------------------------------------------------
// Canonical journal and map content, shared by the native round-trip and
// tamper-fallback tests below.
// ---------------------------------------------------------------------------

storage::JobJournal build_journal_v1() {
    storage::JobJournal journal;
    journal.JobId = "cross-journal-01";
    journal.SourceRoot = "D:\\CrossSrc";
    journal.DestinationRoot = "E:\\CrossDest";
    journal.CreatedUtc = *time::DateTimeOffset::parse("2026-07-23T00:00:00+00:00");

    storage::JournalFileEntry alpha;
    alpha.RelativePath = "alpha.txt";
    alpha.SourceLength = 1234;
    alpha.SourceLastWriteUtcTicks = 638000000000000000LL;
    alpha.BytesCopied = 1234;
    alpha.State = storage::FileCopyState::Completed;
    alpha.RecoveredRanges.push_back({0, 128});
    alpha.RescueRanges.push_back({512, 64, storage::RescueRangeState::Recovered});
    alpha.LastRescuePass = "Scrape";
    journal.Files.set("alpha.txt", std::move(alpha));

    storage::JournalFileEntry beta;
    beta.RelativePath = "sub\\\xD0\xB1\xD0\xB5\xD1\x82\xD0\xB0.bin"; // sub\бета.bin
    beta.SourceLength = 999999;
    beta.BytesCopied = 4096;
    beta.State = storage::FileCopyState::InProgress;
    beta.LastError = "read timeout";
    beta.DoNotRetry = true;
    std::string beta_key = beta.RelativePath;
    journal.Files.set(std::move(beta_key), std::move(beta));
    return journal;
}

storage::JobJournal build_journal_v2() {
    storage::JobJournal journal = build_journal_v1();
    storage::JournalFileEntry gamma;
    gamma.RelativePath = "gamma.dat";
    gamma.SourceLength = 42;
    gamma.State = storage::FileCopyState::Pending;
    journal.Files.set("gamma.dat", std::move(gamma));
    return journal;
}

storage::BadRangeMap build_map_v1() {
    storage::BadRangeMap map;
    map.SourceRoot = "D:\\CrossSrc";
    map.SourceIdentity = "SER-XYZ-123";

    storage::BadRangeMapFileEntry alpha;
    alpha.RelativePath = "alpha.txt";
    alpha.SourceLength = 1234;
    alpha.LastWriteUtcTicks = 638000000000000000LL;
    alpha.FileFingerprint = "fp-alpha";
    alpha.BadRanges.push_back({300, 25}); // intentionally unsorted
    alpha.BadRanges.push_back({100, 50});
    alpha.LastScanUtc = *time::DateTimeOffset::parse("2026-07-23T06:00:00+00:00");
    alpha.LastError = "CRC error";
    map.Files.set("alpha.txt", std::move(alpha));
    return map;
}

storage::BadRangeMap build_map_v2() {
    storage::BadRangeMap map = build_map_v1();
    storage::BadRangeMapFileEntry delta;
    delta.RelativePath = "delta.iso";
    delta.SourceLength = 777;
    delta.LastWriteUtcTicks = 638111111111111111LL;
    delta.FileFingerprint = "fp-delta";
    delta.BadRanges.push_back({0, 4096});
    delta.LastScanUtc = *time::DateTimeOffset::parse("2026-07-23T07:00:00+00:00");
    map.Files.set("delta.iso", std::move(delta));
    return map;
}

void verify_journal(const std::optional<storage::JobJournal>& journal, const std::string& origin) {
    check(journal.has_value(), origin + ": journal loads");
    if (!journal.has_value()) return;
    check_equal(journal->JobId, "cross-journal-01", origin + ": journal JobId");
    check_equal(journal->SourceRoot, "D:\\CrossSrc", origin + ": journal SourceRoot");
    check(journal->Files.size() == 3, origin + ": journal file count (" + std::to_string(journal->Files.size()) + ")");

    const auto* alpha = journal->Files.find("alpha.txt");
    check(alpha != nullptr, origin + ": alpha present");
    if (alpha != nullptr) {
        check(alpha->State == storage::FileCopyState::Completed, origin + ": alpha state");
        check(alpha->RecoveredRanges.size() == 1 && alpha->RecoveredRanges[0].Length == 128,
              origin + ": alpha recovered range");
        check(alpha->RescueRanges.size() == 1 &&
              alpha->RescueRanges[0].State == storage::RescueRangeState::Recovered,
              origin + ": alpha rescue range state");
        check_equal(alpha->LastRescuePass, "Scrape", origin + ": alpha rescue pass");
    }

    const auto* beta = journal->Files.find("sub\\\xD0\xB1\xD0\xB5\xD1\x82\xD0\xB0.bin");
    check(beta != nullptr, origin + ": beta present");
    if (beta != nullptr) {
        check(beta->DoNotRetry, origin + ": beta DoNotRetry");
        check(beta->State == storage::FileCopyState::InProgress, origin + ": beta state");
        check_equal(beta->LastError, "read timeout", origin + ": beta error text");
    }

    const auto* gamma = journal->Files.find("gamma.dat");
    check(gamma != nullptr, origin + ": gamma present");
    if (gamma != nullptr) {
        check(gamma->State == storage::FileCopyState::Pending, origin + ": gamma state");
    }
}

// After primary tampering the map store falls back to .bak1 (the previous
// save), because — unlike the journal — the map trust model has no ledger
// sequencing. That is exactly the .NET candidate order, so tamper checks
// expect the v1 content.
void verify_map_v1(const std::optional<storage::BadRangeMap>& map, const std::string& origin) {
    check(map.has_value(), origin + ": map loads");
    if (!map.has_value()) return;
    check_equal(map->SourceIdentity, "SER-XYZ-123", origin + ": map identity");
    check(map->Files.size() == 1, origin + ": map v1 file count (" + std::to_string(map->Files.size()) + ")");
    const auto* alpha = map->Files.find("alpha.txt");
    check(alpha != nullptr, origin + ": map alpha present");
    if (alpha != nullptr) {
        check(alpha->BadRanges.size() == 2 &&
              alpha->BadRanges[0].Offset == 100 && alpha->BadRanges[1].Offset == 300,
              origin + ": map alpha ranges sorted");
    }
    check(map->Files.find("delta.iso") == nullptr, origin + ": map is pre-delta (bak1) content");
}

void verify_map(const std::optional<storage::BadRangeMap>& map, const std::string& origin) {
    check(map.has_value(), origin + ": map loads");
    if (!map.has_value()) return;
    check_equal(map->SourceIdentity, "SER-XYZ-123", origin + ": map identity");
    check(map->Files.size() == 2, origin + ": map file count (" + std::to_string(map->Files.size()) + ")");

    const auto* alpha = map->Files.find("alpha.txt");
    check(alpha != nullptr, origin + ": map alpha present");
    if (alpha != nullptr) {
        check(alpha->BadRanges.size() == 2 &&
              alpha->BadRanges[0].Offset == 100 && alpha->BadRanges[0].Length == 50 &&
              alpha->BadRanges[1].Offset == 300 && alpha->BadRanges[1].Length == 25,
              origin + ": map alpha ranges sorted");
        check_equal(alpha->LastError, "CRC error", origin + ": map alpha error");
    }

    const auto* delta = map->Files.find("delta.iso");
    check(delta != nullptr, origin + ": map delta present");
    if (delta != nullptr) {
        check(delta->BadRanges.size() == 1 && delta->BadRanges[0].Length == 4096,
              origin + ": map delta range");
    }
}

// ---------------------------------------------------------------------------
// Unit tests
// ---------------------------------------------------------------------------

void test_crypto_vectors() {
    check_equal(crypto::sha256_hex(std::string_view("abc")),
                "ba7816bf8f01cfea414140de5dae2223b00361a396177a9cb410ff61f20015ad",
                "sha256(abc)");

    std::vector<unsigned char> jefe = {'J', 'e', 'f', 'e'};
    check_equal(crypto::to_hex_lower(crypto::hmac_sha256(jefe, "what do ya want for nothing?")),
                "5bdcc146bf60754e6a042426089575c75a003f089d2739839dec58b964ec3843",
                "hmac-sha256 RFC4231 case 2");

    check_equal(crypto::to_base64({}), "", "base64 empty");
    check_equal(crypto::to_base64({'f'}), "Zg==", "base64 f");
    check_equal(crypto::to_base64({'f', 'o'}), "Zm8=", "base64 fo");
    check_equal(crypto::to_base64({'f', 'o', 'o'}), "Zm9v", "base64 foo");

    check(crypto::secure_equals("same", "same"), "secure_equals same");
    check(!crypto::secure_equals("same", "sane"), "secure_equals differ");
    check(!crypto::secure_equals("same", "longer"), "secure_equals length");

    check(storage::detail::checksum32(nullptr, 0) == 0x811C9DC5u, "fnv1a empty");
    const unsigned char hello[] = {'h', 'e', 'l', 'l', 'o'};
    check(storage::detail::checksum32(hello, 5) == 0x4F9F2CABu, "fnv1a hello");

    std::vector<unsigned char> secret = {1, 2, 3, 4, 5, 6, 7, 8};
    auto protected_bytes = crypto::dpapi_protect(secret);
    check(protected_bytes.size() > secret.size(), "dpapi output larger than input");
    check(crypto::dpapi_unprotect(protected_bytes) == secret, "dpapi round-trip");
}

void test_indented_writer_basics() {
    json::Writer w(/*indented*/ true);
    w.begin_object();
    w.key("A"); w.value(static_cast<std::int64_t>(1));
    w.key("B");
    w.begin_object();
    w.end_object();
    w.key("C");
    w.begin_array();
    w.value(static_cast<std::int64_t>(1));
    w.value(static_cast<std::int64_t>(2));
    w.end_array();
    w.end_object();
    check_equal(w.take(),
                "{\r\n  \"A\": 1,\r\n  \"B\": {},\r\n  \"C\": [\r\n    1,\r\n    2\r\n  ]\r\n}",
                "indented writer layout");
}

void test_ordered_file_map_scale() {
    // Regression for the O(n^2) journal build that hung whole-drive scans: the
    // map (used by JobJournal.Files / BadRangeMap.Files) must be O(1) find/set.
    const int count = 100000;
    storage::OrderedFileMap<storage::JournalFileEntry> map;

    auto t0 = std::chrono::steady_clock::now();
    for (int i = 0; i < count; ++i) {
        storage::JournalFileEntry entry;
        entry.RelativePath = "dir" + std::to_string(i % 500) + "\\file" + std::to_string(i) + ".bin";
        entry.SourceLength = i;
        map.set(entry.RelativePath, entry);
    }
    double build_seconds = std::chrono::duration<double>(std::chrono::steady_clock::now() - t0).count();
    std::printf("  built %d-entry map in %.3fs\n", count, build_seconds);

    check(map.size() == static_cast<std::size_t>(count), "all entries inserted, none lost");
    // O(1) build finishes in well under a second; the old O(n^2) build was
    // billions of compares (minutes) at this scale.
    check(build_seconds < 3.0, "100k-entry map builds fast (O(1), not O(n^2))");

    // Lookups are correct, order-preserving, and case-insensitive (ASCII).
    auto t1 = std::chrono::steady_clock::now();
    bool all_found = true;
    for (int i = 0; i < count; ++i) {
        std::string key = "DIR" + std::to_string(i % 500) + "\\FILE" + std::to_string(i) + ".BIN";
        const storage::JournalFileEntry* e = map.find(key);
        if (e == nullptr || e->SourceLength != i) { all_found = false; break; }
    }
    double find_seconds = std::chrono::duration<double>(std::chrono::steady_clock::now() - t1).count();
    check(all_found, "case-insensitive find returns the right entry for every key");
    check(find_seconds < 2.0, "100k lookups fast");

    // Upsert keeps position + count; the folded key must not create a duplicate.
    storage::JournalFileEntry replacement;
    replacement.RelativePath = "dir0\\file0.bin";
    replacement.SourceLength = 999;
    map.set("DIR0\\FILE0.BIN", replacement);
    check(map.size() == static_cast<std::size_t>(count), "case-variant upsert does not duplicate");
    check(map.find("dir0\\file0.bin") != nullptr && map.find("dir0\\file0.bin")->SourceLength == 999,
          "upsert replaced value under the existing key");
    check(map.entries.front().first == "dir0\\file0.bin", "upsert preserved insertion order");
}

// The object golden_journal_payload.json is expected to deserialize into,
// whichever writer produced the bytes.
storage::JobJournal build_golden_journal() {
    storage::JobJournal journal;
    journal.JobId = "golden-journal-01";
    journal.SourceRoot = "D:\\GoldenSrc";
    journal.DestinationRoot = "E:\\GoldenDest";
    journal.CreatedUtc = *time::DateTimeOffset::parse("2026-07-23T00:00:00+00:00");
    journal.UpdatedUtc = *time::DateTimeOffset::parse("2026-07-23T01:30:00.5+00:00");

    storage::JournalFileEntry alpha;
    alpha.RelativePath = "alpha.txt";
    alpha.SourceLength = 1234;
    alpha.SourceLastWriteUtcTicks = 638000000000000000LL;
    alpha.BytesCopied = 1234;
    alpha.State = storage::FileCopyState::CompletedWithRecovery;
    alpha.RecoveredRanges.push_back({0, 128});
    alpha.RescueRanges.push_back({512, 64, storage::RescueRangeState::Recovered});
    alpha.LastRescuePass = "Scrape";
    journal.Files.set("alpha.txt", std::move(alpha));

    storage::JournalFileEntry empty;
    empty.RelativePath = "empty.dat";
    empty.State = storage::FileCopyState::Pending;
    journal.Files.set("empty.dat", std::move(empty));
    return journal;
}

bool entries_equivalent(const storage::JournalFileEntry& a, const storage::JournalFileEntry& b) {
    if (a.RelativePath != b.RelativePath || a.SourceLength != b.SourceLength ||
        a.SourceLastWriteUtcTicks != b.SourceLastWriteUtcTicks || a.BytesCopied != b.BytesCopied ||
        a.State != b.State || a.LastError != b.LastError || a.DoNotRetry != b.DoNotRetry ||
        a.LastRescuePass != b.LastRescuePass ||
        a.RecoveredRanges.size() != b.RecoveredRanges.size() ||
        a.RescueRanges.size() != b.RescueRanges.size()) {
        return false;
    }
    for (std::size_t i = 0; i < a.RecoveredRanges.size(); ++i) {
        if (a.RecoveredRanges[i].Offset != b.RecoveredRanges[i].Offset ||
            a.RecoveredRanges[i].Length != b.RecoveredRanges[i].Length) {
            return false;
        }
    }
    for (std::size_t i = 0; i < a.RescueRanges.size(); ++i) {
        if (a.RescueRanges[i].Offset != b.RescueRanges[i].Offset ||
            a.RescueRanges[i].Length != b.RescueRanges[i].Length ||
            a.RescueRanges[i].State != b.RescueRanges[i].State) {
            return false;
        }
    }
    return true;
}

bool journals_equivalent(const storage::JobJournal& a, const storage::JobJournal& b) {
    if (a.JobId != b.JobId || a.SourceRoot != b.SourceRoot ||
        a.DestinationRoot != b.DestinationRoot || !(a.CreatedUtc == b.CreatedUtc) ||
        !(a.UpdatedUtc == b.UpdatedUtc) || a.Files.size() != b.Files.size()) {
        return false;
    }
    for (std::size_t i = 0; i < a.Files.entries.size(); ++i) {
        if (a.Files.entries[i].first != b.Files.entries[i].first) return false;
        if (!entries_equivalent(a.Files.entries[i].second, b.Files.entries[i].second)) return false;
    }
    return true;
}

void test_storage_goldens(const std::string& golden_dir) {
    std::string map_golden = read_file(golden_dir + "/golden_map_payload.json");
    if (map_golden.empty()) {
        std::printf("NOTE: storage goldens not found in %s; pass the goldens directory.\n",
                    golden_dir.c_str());
        return;
    }

    // Reconstruct the exact .NET golden objects and byte-compare our indented writer.
    storage::BadRangeMap map;
    map.SchemaVersion = 1;
    map.SourceRoot = "D:\\GoldenSrc";
    map.SourceIdentity = "SER-0042 & <id>";
    map.UpdatedUtc = *time::DateTimeOffset::parse("2026-07-23T12:00:00.123+00:00");
    {
        storage::BadRangeMapFileEntry alpha;
        alpha.RelativePath = "alpha.txt";
        alpha.SourceLength = 1234;
        alpha.LastWriteUtcTicks = 638000000000000000LL;
        alpha.FileFingerprint = "fp-alpha";
        alpha.BadRanges.push_back({100, 50});
        alpha.BadRanges.push_back({300, 25});
        alpha.LastScanUtc = *time::DateTimeOffset::parse("2026-07-23T12:00:00+00:00");
        alpha.LastError = "CRC error";
        map.Files.set("alpha.txt", std::move(alpha));

        storage::BadRangeMapFileEntry beta;
        beta.RelativePath = "sub\\\xD0\xB1\xD0\xB5\xD1\x82\xD0\xB0.bin";
        beta.LastScanUtc = *time::DateTimeOffset::parse("2026-07-23T12:00:00+05:30");
        std::string beta_key = beta.RelativePath;
        map.Files.set(std::move(beta_key), std::move(beta));
    }
    check_equal(storage::BadRangeMapStore::serialize_map(map), map_golden,
                "golden map payload byte-for-byte");

    // Round-trip: parse the golden and re-serialize.
    storage::BadRangeMap reparsed = storage::BadRangeMap::from_json(json::parse(map_golden));
    check_equal(storage::BadRangeMapStore::serialize_map(reparsed), map_golden,
                "golden map reparse byte-for-byte");

    // The journal writer omits default-valued members, so its output is
    // deliberately no longer byte-identical to the .NET writer's. What still
    // has to hold is interop, and that is what these three checks cover:
    //   1. the .NET-shaped golden still loads into an identical object,
    //   2. the slim output is byte-stable (regression guard on the writer),
    //   3. re-reading the slim output reproduces the same object, which is the
    //      property that keeps a .NET reader (defaults for absent members)
    //      equivalent to ours.
    std::string journal_golden = read_file(golden_dir + "/golden_journal_payload.json");
    if (!journal_golden.empty()) {
        storage::JobJournal from_dotnet = storage::JobJournal::from_json(json::parse(journal_golden));
        check(journals_equivalent(from_dotnet, build_golden_journal()),
              "golden journal: .NET-shaped payload still loads to the expected object");

        json::Writer w(/*indented*/ true);
        from_dotnet.to_json(w);
        std::string slim = w.take();

        std::string slim_golden = read_file(golden_dir + "/golden_journal_payload_slim.json");
        if (!slim_golden.empty()) {
            check_equal(slim, slim_golden, "golden journal slim payload byte-for-byte");
        }

        storage::JobJournal round_tripped = storage::JobJournal::from_json(json::parse(slim));
        storage::JobJournalStore::normalize_journal(round_tripped);
        storage::JobJournal expected = from_dotnet;
        storage::JobJournalStore::normalize_journal(expected);
        check(journals_equivalent(round_tripped, expected),
              "golden journal: slim payload round-trips to the same object");
        check(slim.size() < journal_golden.size(), "golden journal: slim payload is smaller");
    }

    std::string empty_golden = read_file(golden_dir + "/golden_journal_empty.json");
    if (!empty_golden.empty()) {
        storage::JobJournal empty_journal = storage::JobJournal::from_json(json::parse(empty_golden));
        json::Writer w(/*indented*/ true);
        empty_journal.to_json(w);
        check_equal(w.take(), empty_golden, "golden empty-files journal byte-for-byte");
    }
}

// Journal snapshots are compressed above a size threshold. The container must
// round-trip, small journals must stay plain JSON, and — the one that actually
// matters for upgrades — a plain-JSON journal already on disk must still load.
void test_journal_compression(const std::wstring& work_dir) {
    std::printf("--- journal compression: container + legacy plain-JSON reads ---\n");
    storage::fsutil::create_directories(work_dir);

    // Small journal: stays readable JSON, no container header.
    {
        std::wstring path = work_dir + L"\\job-small.json";
        storage::JobJournalStore store;
        storage::JobJournal journal;
        journal.JobId = "small";
        journal.SourceRoot = "C:\\src";
        journal.DestinationRoot = "C:\\dst";
        store.save(path, journal);

        auto bytes = storage::fsutil::read_all_bytes(path);
        check(bytes.has_value() && !bytes->empty(), "compression: small journal written");
        check(bytes.has_value() && (*bytes)[0] == '{',
              "compression: small journal stays plain JSON");
        auto loaded = store.load(path);
        check(loaded.has_value() && loaded->JobId == "small", "compression: small journal loads");
        storage::JobJournalStore::remove_journal_set(path);
    }

    // Large journal: compressed container, and it must survive a full round-trip
    // through the store including the ledger's trust check.
    {
        std::wstring path = work_dir + L"\\job-large.json";
        storage::JobJournalStore store;
        storage::JobJournal journal;
        journal.JobId = "large";
        journal.SourceRoot = "C:\\src";
        journal.DestinationRoot = "C:\\dst";
        for (int i = 0; i < 4000; ++i) {
            storage::JournalFileEntry entry;
            entry.RelativePath = "folder\\subfolder\\file-" + std::to_string(i) + ".bin";
            entry.SourceLength = 1024 * (i + 1);
            entry.BytesCopied = entry.SourceLength;
            entry.State = storage::FileCopyState::Completed;
            entry.LastRescuePass = "FastHealthScan";
            std::string key = entry.RelativePath;
            journal.Files.set(std::move(key), std::move(entry));
        }
        store.save(path, journal);

        auto bytes = storage::fsutil::read_all_bytes(path);
        check(bytes.has_value() && bytes->size() >= 4, "compression: large journal written");
        bool is_container = bytes.has_value() && bytes->size() >= 16 && (*bytes)[0] == 'X' &&
                            (*bytes)[1] == 'C' && (*bytes)[2] == 'J' && (*bytes)[3] == 'Z';
        check(is_container, "compression: large journal is a compressed container");

        json::Writer w(/*indented*/ true);
        journal.to_json(w);
        std::size_t plain_size = w.take().size();
        check(bytes.has_value() && bytes->size() < plain_size / 2,
              "compression: container is less than half the plain size (" +
                  std::to_string(bytes.has_value() ? bytes->size() : 0) + " vs " +
                  std::to_string(plain_size) + ")");

        auto loaded = store.load(path);
        check(loaded.has_value() && loaded->JobId == "large",
              "compression: compressed journal loads through the ledger");
        check(loaded.has_value() && loaded->Files.size() == 4000,
              "compression: every entry survives the round-trip");
        if (loaded.has_value()) {
            const auto* entry = loaded->Files.find("folder\\subfolder\\file-1234.bin");
            check(entry != nullptr && entry->SourceLength == 1024 * 1235 &&
                      entry->State == storage::FileCopyState::Completed &&
                      entry->RelativePath == "folder\\subfolder\\file-1234.bin",
                  "compression: entry contents and key-derived path are intact");
        }
        storage::JobJournalStore::remove_journal_set(path);
    }

    // Upgrade path: a plain-JSON journal sitting on disk from an older build
    // (or the legacy .NET build) still loads, with no ledger present at all.
    {
        std::wstring path = work_dir + L"\\job-legacy.json";
        std::string legacy =
            "{\r\n  \"JobId\": \"legacy\",\r\n  \"SourceRoot\": \"D:\\\\Old\",\r\n"
            "  \"DestinationRoot\": \"E:\\\\Old\",\r\n"
            "  \"CreatedUtc\": \"2026-01-01T00:00:00+00:00\",\r\n"
            "  \"UpdatedUtc\": \"2026-01-01T00:00:00+00:00\",\r\n  \"Files\": {\r\n"
            "    \"legacy.bin\": {\r\n      \"RelativePath\": \"legacy.bin\",\r\n"
            "      \"SourceLength\": 4096,\r\n      \"BytesCopied\": 4096,\r\n"
            "      \"State\": \"Completed\",\r\n      \"LastError\": \"\",\r\n"
            "      \"DoNotRetry\": false,\r\n      \"RecoveredRanges\": [],\r\n"
            "      \"RescueRanges\": [],\r\n      \"LastRescuePass\": \"\"\r\n    }\r\n  }\r\n}";
        storage::fsutil::write_file_raw(path, reinterpret_cast<const unsigned char*>(legacy.data()),
                                        legacy.size(), false, false);

        storage::JobJournalStore store;
        auto loaded = store.load(path);
        check(loaded.has_value() && loaded->JobId == "legacy",
              "compression: legacy plain-JSON journal still loads");
        if (loaded.has_value()) {
            const auto* entry = loaded->Files.find("legacy.bin");
            check(entry != nullptr && entry->BytesCopied == 4096 &&
                      entry->State == storage::FileCopyState::Completed,
                  "compression: legacy entry read correctly");
        }
        storage::JobJournalStore::remove_journal_set(path);
    }
}

// Journal maintenance: compaction must drop only the rotating backups, and
// pruning must never touch a journal that is still resumable.
void test_journal_maintenance(const std::wstring& work_dir) {
    std::printf("--- journal maintenance: compaction + retention ---\n");
    storage::fsutil::create_directories(work_dir);
    // Named like a real journal so the store's strict job-*.json sweep sees it.
    std::wstring journal_path = work_dir + L"\\job-maint.json";

    storage::JobJournalStore store;
    storage::JobJournal journal;
    journal.JobId = "maintenance";
    journal.SourceRoot = "C:\\src";
    journal.DestinationRoot = "C:\\dst";

    // Three saves produce the full backup rotation (bak1..bak3).
    for (int i = 0; i < 4; ++i) {
        journal.UpdatedUtc = time::DateTimeOffset::now_utc();
        store.save(journal_path, journal);
    }

    int backups_before = 0;
    for (int generation = 1; generation <= storage::JobJournalStore::BackupGenerationCount;
         ++generation) {
        if (storage::fsutil::file_exists(journal_path + L".bak" + std::to_wstring(generation))) {
            ++backups_before;
        }
    }
    check(backups_before > 0, "maintenance: rotation produced backup generations");

    std::int64_t reclaimed = storage::JobJournalStore::compact_completed(journal_path);
    check(reclaimed > 0, "maintenance: compaction reclaimed bytes");

    int backups_after = 0;
    for (int generation = 1; generation <= storage::JobJournalStore::BackupGenerationCount;
         ++generation) {
        if (storage::fsutil::file_exists(journal_path + L".bak" + std::to_wstring(generation))) {
            ++backups_after;
        }
    }
    check(backups_after == 0, "maintenance: compaction removed every backup generation");

    // The snapshot and its trust chain must survive, and still load.
    check(storage::fsutil::file_exists(journal_path), "maintenance: snapshot kept after compaction");
    check(storage::fsutil::file_exists(journal_path + L".ledger"),
          "maintenance: ledger kept after compaction");
    auto reloaded = store.load(journal_path);
    check(reloaded.has_value() && reloaded->JobId == "maintenance",
          "maintenance: journal still loads after compaction");

    // Retention must skip a protected (still resumable) journal.
    storage::JobJournalStore::PruneOptions options;
    options.retention_days = 1;
    options.keep_minimum = 0;
    options.protected_paths.push_back(journal_path);
    options.journals_root = work_dir;
    storage::JobJournalStore::PruneResult guarded = storage::JobJournalStore::prune(options);
    check(storage::fsutil::file_exists(journal_path),
          "maintenance: protected journal survives pruning");
    (void)guarded;

    // keep_minimum alone must also hold a journal back, regardless of age.
    storage::JobJournalStore::PruneOptions keep_all;
    keep_all.retention_days = 1;
    keep_all.keep_minimum = 1000;
    keep_all.journals_root = work_dir;
    storage::JobJournalStore::prune(keep_all);
    check(storage::fsutil::file_exists(journal_path),
          "maintenance: keep_minimum holds journals back");

    // Retention disabled must be a no-op.
    storage::JobJournalStore::PruneOptions disabled;
    disabled.retention_days = 0;
    disabled.journals_root = work_dir;
    storage::JobJournalStore::PruneResult none = storage::JobJournalStore::prune(disabled);
    check(none.journals_removed == 0, "maintenance: retention_days=0 disables pruning");

    // Tear down: save() also writes into the real mirror directory, so remove
    // this journal's whole artifact set rather than leaving strays behind.
    storage::JobJournalStore::remove_journal_set(journal_path);
    check(!storage::fsutil::file_exists(journal_path),
          "maintenance: remove_journal_set clears the journal");
}

void test_native_store_roundtrip(const std::wstring& work_dir) {
    storage::fsutil::create_directories(work_dir);
    std::wstring journal_path = work_dir + L"\\rt-journal.json";
    std::wstring map_path = work_dir + L"\\rt-map.json";

    storage::JobJournalStore journal_store;
    journal_store.save(journal_path, build_journal_v1());
    journal_store.save(journal_path, build_journal_v2());
    verify_journal(journal_store.load(journal_path), "native-roundtrip");

    auto records = journal_store.read_ledger(journal_path);
    check(records.size() == 2, "native ledger has 2 chained records");
    if (records.size() == 2) {
        check(records[1].Sequence == records[0].Sequence + 1, "native ledger sequence increments");
        check(records[1].PreviousRecordHash == records[0].RecordHash, "native ledger hash chain");
    }

    storage::BadRangeMapStore map_store;
    map_store.save(map_path, build_map_v1());
    map_store.save(map_path, build_map_v2());
    verify_map(map_store.load(map_path), "native-roundtrip");

    // Tamper: overwrite primaries with parseable-but-wrong content; trusted
    // paths must fall back to backups/mirrors.
    storage::JobJournal tampered_journal = build_journal_v2();
    tampered_journal.JobId = "TAMPERED";
    json::Writer jw(true);
    tampered_journal.to_json(jw);
    storage::fsutil::write_atomic_bytes(journal_path, jw.take());

    auto recovered_journal = journal_store.load(journal_path);
    check(recovered_journal.has_value() && recovered_journal->JobId != "TAMPERED",
          "native tamper: journal fell back to trusted snapshot");
    verify_journal(recovered_journal, "native-tamper");

    storage::BadRangeMap tampered_map = build_map_v2();
    tampered_map.SourceIdentity = "TAMPERED";
    storage::fsutil::write_atomic_bytes(map_path, storage::BadRangeMapStore::serialize_map(tampered_map));

    auto recovered_map = map_store.load(map_path);
    check(recovered_map.has_value() && recovered_map->SourceIdentity != "TAMPERED",
          "native tamper: map fell back to signed snapshot");
    verify_map_v1(recovered_map, "native-tamper");
}

// --- Job catalog round-trip ------------------------------------------------
//
// This fixture was originally the shared canonical content for the .NET
// StorageProbe cross-check. That tooling is retired, but the content is still
// the most thorough catalog exercise we have — non-ASCII names and paths,
// nullable members that must stay null, enum members, and an optional result
// present on one run and absent on another — so it now drives a native
// save/load round-trip instead.

storage::JobCatalog build_canonical_catalog() {
    auto fixed_time = [](const char* text) {
        return *time::DateTimeOffset::parse(text);
    };

    storage::JobCatalog catalog;

    storage::ManagedJob job;
    job.JobId = "cat-job-0001";
    job.Name = "Кросс Job <&>";
    job.Options.SourceRoot = "D:\\CrossSrc";
    job.Options.DestinationRoot = "E:\\CrossDest";
    job.Options.MaxRetries = 9;
    job.Options.VerificationModeValue = models::VerificationMode::Full;
    job.Options.OverwritePolicyValue = models::OverwritePolicy::SkipExisting;
    job.CreatedUtc = fixed_time("2026-07-23T01:02:03+00:00");
    job.UpdatedUtc = fixed_time("2026-07-23T01:02:03+00:00");
    catalog.Jobs.push_back(job);

    storage::ManagedJobQueueEntry attempted;
    attempted.QueueEntryId = "cat-queue-0001";
    attempted.JobId = job.JobId;
    attempted.EnqueuedUtc = fixed_time("2026-07-23T01:30:00+00:00");
    attempted.LastUpdatedUtc = fixed_time("2026-07-23T02:00:00+00:00");
    attempted.Trigger = "manual";
    attempted.EnqueuedBy = "probe";
    attempted.AttemptCount = 2;
    attempted.LastAttemptUtc = fixed_time("2026-07-23T02:00:00+00:00");
    attempted.LastErrorMessage = "prior failure";
    catalog.QueueEntries.push_back(attempted);

    storage::ManagedJobQueueEntry fresh;
    fresh.QueueEntryId = "cat-queue-0002";
    fresh.JobId = job.JobId;
    fresh.EnqueuedUtc = fixed_time("2026-07-23T01:45:00+00:00");
    fresh.LastUpdatedUtc = fixed_time("2026-07-23T01:45:00+00:00");
    fresh.Trigger = "queued";
    catalog.QueueEntries.push_back(fresh);

    storage::ManagedJobRun completed;
    completed.RunId = "cat-run-0001";
    completed.JobId = job.JobId;
    completed.DisplayName = job.Name;
    completed.SourceRoot = job.Options.SourceRoot;
    completed.DestinationRoot = job.Options.DestinationRoot;
    completed.Trigger = "queued-manual";
    completed.QueueEntryId = attempted.QueueEntryId;
    completed.QueueAttempt = 2;
    completed.Status = storage::ManagedJobRunStatus::Completed;
    completed.StartedUtc = fixed_time("2026-07-23T02:10:00+00:00");
    completed.LastUpdatedUtc = fixed_time("2026-07-23T02:20:00+00:00");
    completed.FinishedUtc = fixed_time("2026-07-23T02:20:00+00:00");
    completed.JournalPath = "C:\\журнал\\j.json";
    completed.Summary = "Completed: 3/3 files at 1 MB/s.";
    models::CopyJobResult result;
    result.Succeeded = true;
    result.TotalFiles = 3;
    result.CompletedFiles = 3;
    result.TransferEnginePolicyValue = models::TransferEnginePolicy::NativeFast;
    result.AverageBytesPerSecond = 1048576.0;
    result.JournalPath = completed.JournalPath;
    completed.Result = result;
    catalog.Runs.push_back(completed);

    storage::ManagedJobRun interrupted;
    interrupted.RunId = "cat-run-0002";
    interrupted.JobId = job.JobId;
    interrupted.DisplayName = "Manual Copy";
    interrupted.SourceRoot = job.Options.SourceRoot;
    interrupted.DestinationRoot = job.Options.DestinationRoot;
    interrupted.Trigger = "manual";
    interrupted.Status = storage::ManagedJobRunStatus::Interrupted;
    interrupted.StartedUtc = fixed_time("2026-07-23T03:00:00+00:00");
    interrupted.LastUpdatedUtc = fixed_time("2026-07-23T03:05:00+00:00");
    interrupted.ErrorMessage = "Application closed unexpectedly.";
    interrupted.Summary = "Application closed unexpectedly.";
    catalog.Runs.push_back(interrupted);

    return catalog;
}

void verify_canonical_catalog(const storage::JobCatalog& catalog, const std::string& tag) {
    check(catalog.SchemaVersion == 2, tag + ": catalog schema 2");
    check(catalog.Jobs.size() == 1, tag + ": one job");
    if (catalog.Jobs.size() == 1) {
        const auto& job = catalog.Jobs[0];
        check(job.JobId == "cat-job-0001", tag + ": job id");
        check(job.Name == "Кросс Job <&>", tag + ": job unicode name");
        check(job.Options.MaxRetries == 9, tag + ": job options retries");
        check(job.Options.VerificationModeValue == models::VerificationMode::Full,
              tag + ": job options verification enum");
        check(job.Options.OverwritePolicyValue == models::OverwritePolicy::SkipExisting,
              tag + ": job options overwrite enum");
        check(job.CreatedUtc == *time::DateTimeOffset::parse("2026-07-23T01:02:03+00:00"),
              tag + ": job created utc");
    }

    check(catalog.QueueEntries.size() == 2, tag + ": two queue entries");
    if (catalog.QueueEntries.size() == 2) {
        const auto& attempted = catalog.QueueEntries[0];
        check(attempted.AttemptCount == 2 && attempted.LastAttemptUtc.has_value() &&
                  *attempted.LastAttemptUtc == *time::DateTimeOffset::parse("2026-07-23T02:00:00+00:00"),
              tag + ": attempted entry retains LastAttemptUtc");
        check(attempted.LastErrorMessage == "prior failure", tag + ": attempted entry error message");
        check(!catalog.QueueEntries[1].LastAttemptUtc.has_value(), tag + ": fresh entry LastAttemptUtc null");
    }

    check(catalog.Runs.size() == 2, tag + ": two runs");
    if (catalog.Runs.size() == 2) {
        const auto& completed = catalog.Runs[0];
        check(completed.Status == storage::ManagedJobRunStatus::Completed, tag + ": run 1 completed");
        check(completed.FinishedUtc.has_value(), tag + ": run 1 finished utc set");
        check(completed.JournalPath == "C:\\журнал\\j.json", tag + ": run 1 unicode journal path");
        check(completed.Result.has_value() && completed.Result->TotalFiles == 3 &&
                  completed.Result->TransferEnginePolicyValue == models::TransferEnginePolicy::NativeFast &&
                  completed.Result->AverageBytesPerSecond == 1048576.0,
              tag + ": run 1 result round-trip");

        const auto& interrupted = catalog.Runs[1];
        check(interrupted.Status == storage::ManagedJobRunStatus::Interrupted, tag + ": run 2 interrupted");
        check(!interrupted.Result.has_value(), tag + ": run 2 null result");
        check(!interrupted.FinishedUtc.has_value(), tag + ": run 2 null finished utc");
        check(interrupted.ErrorMessage == "Application closed unexpectedly.", tag + ": run 2 error message");
    }
}

void test_catalog_roundtrip(const std::wstring& work_dir) {
    std::printf("--- job catalog: canonical content round-trip ---\n");
    storage::fsutil::create_directories(work_dir);
    std::wstring path = work_dir + L"\\rt-catalog.json";

    storage::JobCatalogStore store(path);
    // save() normalizes in place, so it takes a mutable reference.
    storage::JobCatalog catalog = build_canonical_catalog();
    check(store.save(catalog), "catalog round-trip: save");
    verify_canonical_catalog(store.load(), "catalog round-trip");

    storage::fsutil::delete_file(path);
}

} // namespace

int main(int argc, char** argv) {
    test_crypto_vectors();
    test_indented_writer_basics();
    test_ordered_file_map_scale();

    std::string golden_dir = argc > 1 ? argv[1] : "tests/golden";
    test_storage_goldens(golden_dir);

    wchar_t temp_root[MAX_PATH];
    GetTempPathW(MAX_PATH, temp_root);
    std::wstring work_dir = std::wstring(temp_root) + L"xactcopy-storage-test-" +
                            storage::fsutil::random_temp_suffix().substr(0, 12);
    test_native_store_roundtrip(work_dir);
    test_journal_compression(work_dir);
    test_journal_maintenance(work_dir);
    test_catalog_roundtrip(work_dir);

    if (g_failures == 0) {
        std::printf("STORAGE PASS: %d checks\n", g_checks);
        return 0;
    }
    std::printf("STORAGE FAILED: %d of %d checks\n", g_failures, g_checks);
    return 1;
}
