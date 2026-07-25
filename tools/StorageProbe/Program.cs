// -----------------------------------------------------------------------------
// File: cpp\tools\StorageProbe\Program.cs
// Purpose: Cross-compatibility driver for the storage port. Uses the REAL .NET
//          JobJournalStore/BadRangeMapStore to write canonical artifacts for
//          the C++ side to read, and to read/verify artifacts written by the
//          C++ stores — including the tamper-fallback trust path.
// Modes:   write <journalPath> <mapPath>
//          verify <journalPath> <mapPath>
//          verify-tamper <journalPath> <mapPath>   (corrupts primaries first)
// -----------------------------------------------------------------------------

using System.Text;
using XactCopy.Infrastructure;
using XactCopy.Models;

if (args.Length < 3)
{
    Console.Error.WriteLine("usage: StorageProbe <write|verify|verify-tamper> <journalPath> <mapPath>");
    return 2;
}

var mode = args[0].ToLowerInvariant();
var journalPath = Path.GetFullPath(args[1]);
var mapPath = Path.GetFullPath(args[2]);

int failures = 0;
int checks = 0;

void Check(bool condition, string label)
{
    checks++;
    if (condition)
    {
        Console.WriteLine($"  ok: {label}");
    }
    else
    {
        failures++;
        Console.WriteLine($"  FAIL: {label}");
    }
}

// --- Canonical content shared with cpp/tests/test_storage.cpp ---------------

static JobJournal BuildJournalV1()
{
    var journal = new JobJournal
    {
        JobId = "cross-journal-01",
        SourceRoot = "D:\\CrossSrc",
        DestinationRoot = "E:\\CrossDest",
        CreatedUtc = new DateTimeOffset(2026, 7, 23, 0, 0, 0, TimeSpan.Zero)
    };
    journal.Files["alpha.txt"] = new JournalFileEntry
    {
        RelativePath = "alpha.txt",
        SourceLength = 1234,
        SourceLastWriteUtcTicks = 638000000000000000,
        BytesCopied = 1234,
        State = FileCopyState.Completed,
        RecoveredRanges = { new ByteRange { Offset = 0, Length = 128 } },
        RescueRanges = { new RescueRange { Offset = 512, Length = 64, State = RescueRangeState.Recovered } },
        LastRescuePass = "Scrape"
    };
    journal.Files["sub\\бета.bin"] = new JournalFileEntry
    {
        RelativePath = "sub\\бета.bin",
        SourceLength = 999999,
        BytesCopied = 4096,
        State = FileCopyState.InProgress,
        LastError = "read timeout",
        DoNotRetry = true
    };
    return journal;
}

static JobJournal BuildJournalV2()
{
    var journal = BuildJournalV1();
    journal.Files["gamma.dat"] = new JournalFileEntry
    {
        RelativePath = "gamma.dat",
        SourceLength = 42,
        State = FileCopyState.Pending
    };
    return journal;
}

static BadRangeMap BuildMapV1()
{
    var map = new BadRangeMap
    {
        SourceRoot = "D:\\CrossSrc",
        SourceIdentity = "SER-XYZ-123"
    };
    map.Files["alpha.txt"] = new BadRangeMapFileEntry
    {
        RelativePath = "alpha.txt",
        SourceLength = 1234,
        LastWriteUtcTicks = 638000000000000000,
        FileFingerprint = "fp-alpha",
        // Intentionally unsorted; the store must sort by offset.
        BadRanges = { new ByteRange { Offset = 300, Length = 25 }, new ByteRange { Offset = 100, Length = 50 } },
        LastScanUtc = new DateTimeOffset(2026, 7, 23, 6, 0, 0, TimeSpan.Zero),
        LastError = "CRC error"
    };
    return map;
}

static BadRangeMap BuildMapV2()
{
    var map = BuildMapV1();
    map.Files["delta.iso"] = new BadRangeMapFileEntry
    {
        RelativePath = "delta.iso",
        SourceLength = 777,
        LastWriteUtcTicks = 638111111111111111,
        FileFingerprint = "fp-delta",
        BadRanges = { new ByteRange { Offset = 0, Length = 4096 } },
        LastScanUtc = new DateTimeOffset(2026, 7, 23, 7, 0, 0, TimeSpan.Zero)
    };
    return map;
}

void VerifyJournal(JobJournal? journal, string origin)
{
    Check(journal is not null, $"{origin}: journal loads");
    if (journal is null) return;
    Check(journal.JobId == "cross-journal-01", $"{origin}: journal JobId ({journal.JobId})");
    Check(journal.SourceRoot == "D:\\CrossSrc", $"{origin}: journal SourceRoot");
    Check(journal.Files.Count == 3, $"{origin}: journal file count ({journal.Files.Count})");

    Check(journal.Files.TryGetValue("alpha.txt", out var alpha), $"{origin}: alpha present");
    if (alpha is not null)
    {
        Check(alpha.State == FileCopyState.Completed, $"{origin}: alpha state");
        Check(alpha.RecoveredRanges.Count == 1 && alpha.RecoveredRanges[0].Length == 128,
            $"{origin}: alpha recovered range");
        Check(alpha.RescueRanges.Count == 1 && alpha.RescueRanges[0].State == RescueRangeState.Recovered,
            $"{origin}: alpha rescue range state");
        Check(alpha.LastRescuePass == "Scrape", $"{origin}: alpha rescue pass");
    }

    Check(journal.Files.TryGetValue("sub\\бета.bin", out var beta), $"{origin}: беta present");
    if (beta is not null)
    {
        Check(beta.DoNotRetry, $"{origin}: беta DoNotRetry");
        Check(beta.State == FileCopyState.InProgress, $"{origin}: беta state");
        Check(beta.LastError == "read timeout", $"{origin}: беta error text");
    }

    Check(journal.Files.TryGetValue("gamma.dat", out var gamma), $"{origin}: gamma present");
    if (gamma is not null)
    {
        Check(gamma.State == FileCopyState.Pending, $"{origin}: gamma state");
    }
}

// After primary tampering the map store falls back to .bak1 (the previous
// save): unlike the journal, the map trust model has no ledger sequencing.
void VerifyMapV1(BadRangeMap? map, string origin)
{
    Check(map is not null, $"{origin}: map loads");
    if (map is null) return;
    Check(map.SourceIdentity == "SER-XYZ-123", $"{origin}: map identity ({map.SourceIdentity})");
    Check(map.Files.Count == 1, $"{origin}: map v1 file count ({map.Files.Count})");
    Check(map.Files.TryGetValue("alpha.txt", out var alphaV1), $"{origin}: map alpha present");
    if (alphaV1 is not null)
    {
        Check(alphaV1.BadRanges.Count == 2 &&
              alphaV1.BadRanges[0].Offset == 100 && alphaV1.BadRanges[1].Offset == 300,
            $"{origin}: map alpha ranges sorted");
    }
    Check(!map.Files.ContainsKey("delta.iso"), $"{origin}: map is pre-delta (bak1) content");
}

void VerifyMap(BadRangeMap? map, string origin)
{
    Check(map is not null, $"{origin}: map loads");
    if (map is null) return;
    Check(map.SourceIdentity == "SER-XYZ-123", $"{origin}: map identity ({map.SourceIdentity})");
    Check(map.Files.Count == 2, $"{origin}: map file count ({map.Files.Count})");

    Check(map.Files.TryGetValue("alpha.txt", out var alpha), $"{origin}: map alpha present");
    if (alpha is not null)
    {
        Check(alpha.BadRanges.Count == 2 &&
              alpha.BadRanges[0].Offset == 100 && alpha.BadRanges[0].Length == 50 &&
              alpha.BadRanges[1].Offset == 300 && alpha.BadRanges[1].Length == 25,
            $"{origin}: map alpha ranges sorted");
        Check(alpha.LastError == "CRC error", $"{origin}: map alpha error");
    }

    Check(map.Files.TryGetValue("delta.iso", out var delta), $"{origin}: map delta present");
    if (delta is not null)
    {
        Check(delta.BadRanges.Count == 1 && delta.BadRanges[0].Length == 4096,
            $"{origin}: map delta range");
    }
}

// --- Canonical job catalog shared with cpp/tests/test_storage.cpp -----------

static JobCatalog BuildCrossCatalog()
{
    var fixedCreated = DateTimeOffset.Parse("2026-07-23T01:02:03+00:00");
    var catalog = new JobCatalog();

    var job = new ManagedJob
    {
        JobId = "cat-job-0001",
        Name = "Кросс Job <&>",
        CreatedUtc = fixedCreated,
        UpdatedUtc = fixedCreated
    };
    job.Options.SourceRoot = "D:\\CrossSrc";
    job.Options.DestinationRoot = "E:\\CrossDest";
    job.Options.MaxRetries = 9;
    job.Options.VerificationMode = VerificationMode.Full;
    job.Options.OverwritePolicy = OverwritePolicy.SkipExisting;
    catalog.Jobs.Add(job);

    catalog.QueueEntries.Add(new ManagedJobQueueEntry
    {
        QueueEntryId = "cat-queue-0001",
        JobId = job.JobId,
        EnqueuedUtc = DateTimeOffset.Parse("2026-07-23T01:30:00+00:00"),
        LastUpdatedUtc = DateTimeOffset.Parse("2026-07-23T02:00:00+00:00"),
        Trigger = "manual",
        EnqueuedBy = "probe",
        AttemptCount = 2,
        LastAttemptUtc = DateTimeOffset.Parse("2026-07-23T02:00:00+00:00"),
        LastErrorMessage = "prior failure"
    });
    catalog.QueueEntries.Add(new ManagedJobQueueEntry
    {
        QueueEntryId = "cat-queue-0002",
        JobId = job.JobId,
        EnqueuedUtc = DateTimeOffset.Parse("2026-07-23T01:45:00+00:00"),
        LastUpdatedUtc = DateTimeOffset.Parse("2026-07-23T01:45:00+00:00"),
        Trigger = "queued"
    });

    catalog.Runs.Add(new ManagedJobRun
    {
        RunId = "cat-run-0001",
        JobId = job.JobId,
        DisplayName = job.Name,
        SourceRoot = job.Options.SourceRoot,
        DestinationRoot = job.Options.DestinationRoot,
        Trigger = "queued-manual",
        QueueEntryId = "cat-queue-0001",
        QueueAttempt = 2,
        Status = ManagedJobRunStatus.Completed,
        StartedUtc = DateTimeOffset.Parse("2026-07-23T02:10:00+00:00"),
        LastUpdatedUtc = DateTimeOffset.Parse("2026-07-23T02:20:00+00:00"),
        FinishedUtc = DateTimeOffset.Parse("2026-07-23T02:20:00+00:00"),
        JournalPath = "C:\\журнал\\j.json",
        Summary = "Completed: 3/3 files at 1 MB/s.",
        Result = new CopyJobResult
        {
            Succeeded = true,
            TotalFiles = 3,
            CompletedFiles = 3,
            TransferEnginePolicy = TransferEnginePolicy.NativeFast,
            AverageBytesPerSecond = 1048576.0,
            JournalPath = "C:\\журнал\\j.json"
        }
    });
    catalog.Runs.Add(new ManagedJobRun
    {
        RunId = "cat-run-0002",
        JobId = job.JobId,
        DisplayName = "Manual Copy",
        SourceRoot = job.Options.SourceRoot,
        DestinationRoot = job.Options.DestinationRoot,
        Trigger = "manual",
        Status = ManagedJobRunStatus.Interrupted,
        StartedUtc = DateTimeOffset.Parse("2026-07-23T03:00:00+00:00"),
        LastUpdatedUtc = DateTimeOffset.Parse("2026-07-23T03:05:00+00:00"),
        ErrorMessage = "Application closed unexpectedly.",
        Summary = "Application closed unexpectedly."
    });

    return catalog;
}

void VerifyCrossCatalog(JobCatalog catalog, string tag)
{
    Check(catalog.SchemaVersion == 2, $"{tag}: catalog schema 2");
    Check(catalog.Jobs.Count == 1, $"{tag}: one job");
    if (catalog.Jobs.Count == 1)
    {
        var job = catalog.Jobs[0];
        Check(job.JobId == "cat-job-0001", $"{tag}: job id");
        Check(job.Name == "Кросс Job <&>", $"{tag}: job unicode name");
        Check(job.Options.MaxRetries == 9, $"{tag}: job options retries");
        Check(job.Options.VerificationMode == VerificationMode.Full, $"{tag}: job options verification enum");
        Check(job.Options.OverwritePolicy == OverwritePolicy.SkipExisting, $"{tag}: job options overwrite enum");
        Check(job.CreatedUtc == DateTimeOffset.Parse("2026-07-23T01:02:03+00:00"), $"{tag}: job created utc");
    }

    Check(catalog.QueueEntries.Count == 2, $"{tag}: two queue entries");
    if (catalog.QueueEntries.Count == 2)
    {
        var attempted = catalog.QueueEntries[0];
        Check(attempted.AttemptCount == 2 &&
              attempted.LastAttemptUtc == DateTimeOffset.Parse("2026-07-23T02:00:00+00:00"),
            $"{tag}: attempted entry retains LastAttemptUtc");
        Check(attempted.LastErrorMessage == "prior failure", $"{tag}: attempted entry error message");
        Check(catalog.QueueEntries[1].LastAttemptUtc is null, $"{tag}: fresh entry LastAttemptUtc null");
    }

    Check(catalog.Runs.Count == 2, $"{tag}: two runs");
    if (catalog.Runs.Count == 2)
    {
        var completed = catalog.Runs[0];
        Check(completed.Status == ManagedJobRunStatus.Completed, $"{tag}: run 1 completed");
        Check(completed.FinishedUtc is not null, $"{tag}: run 1 finished utc set");
        Check(completed.JournalPath == "C:\\журнал\\j.json", $"{tag}: run 1 unicode journal path");
        Check(completed.Result is not null && completed.Result.TotalFiles == 3 &&
              completed.Result.TransferEnginePolicy == TransferEnginePolicy.NativeFast &&
              completed.Result.AverageBytesPerSecond == 1048576.0,
            $"{tag}: run 1 result round-trip");

        var interrupted = catalog.Runs[1];
        Check(interrupted.Status == ManagedJobRunStatus.Interrupted, $"{tag}: run 2 interrupted");
        Check(interrupted.Result is null, $"{tag}: run 2 null result");
        Check(interrupted.FinishedUtc is null, $"{tag}: run 2 null finished utc");
        Check(interrupted.ErrorMessage == "Application closed unexpectedly.", $"{tag}: run 2 error message");
    }
}

var journalStore = new JobJournalStore();
var mapStore = new BadRangeMapStore();

switch (mode)
{
    case "write":
    {
        Directory.CreateDirectory(Path.GetDirectoryName(journalPath)!);
        Directory.CreateDirectory(Path.GetDirectoryName(mapPath)!);
        await journalStore.SaveAsync(journalPath, BuildJournalV1(), CancellationToken.None);
        await journalStore.SaveAsync(journalPath, BuildJournalV2(), CancellationToken.None);
        await mapStore.SaveAsync(mapPath, BuildMapV1(), CancellationToken.None);
        await mapStore.SaveAsync(mapPath, BuildMapV2(), CancellationToken.None);
        Console.WriteLine("STORAGE WRITE OK (journal x2, map x2)");
        return 0;
    }

    case "verify":
    {
        VerifyJournal(await journalStore.LoadAsync(journalPath, CancellationToken.None), "dotnet-load");
        VerifyMap(await mapStore.LoadAsync(mapPath, CancellationToken.None), "dotnet-load");
        break;
    }

    case "verify-tamper":
    {
        // Tampered primaries: parseable content with wrong values. The trusted
        // ledger/envelope paths must reject these and fall back to backups.
        var tamperedJournal = BuildJournalV2();
        tamperedJournal.JobId = "TAMPERED";
        await File.WriteAllTextAsync(journalPath,
            System.Text.Json.JsonSerializer.Serialize(tamperedJournal),
            new UTF8Encoding(false));

        var tamperedMap = BuildMapV2();
        tamperedMap.SourceIdentity = "TAMPERED";
        await File.WriteAllTextAsync(mapPath,
            System.Text.Json.JsonSerializer.Serialize(tamperedMap),
            new UTF8Encoding(false));

        var journal = await journalStore.LoadAsync(journalPath, CancellationToken.None);
        Check(journal is not null && journal.JobId != "TAMPERED",
            $"tamper: journal fell back to trusted snapshot (JobId {journal?.JobId})");
        VerifyJournal(journal, "tamper-load");

        var map = await mapStore.LoadAsync(mapPath, CancellationToken.None);
        Check(map is not null && map.SourceIdentity != "TAMPERED",
            $"tamper: map fell back to signed snapshot (identity {map?.SourceIdentity})");
        VerifyMapV1(map, "tamper-load");
        break;
    }

    case "write-catalog":
    {
        Directory.CreateDirectory(Path.GetDirectoryName(journalPath)!);
        var catalogStore = new JobCatalogStore(journalPath);
        catalogStore.Save(BuildCrossCatalog());
        Console.WriteLine("STORAGE CATALOG WRITE OK");
        return 0;
    }

    case "verify-catalog":
    {
        var catalogStore = new JobCatalogStore(journalPath);
        VerifyCrossCatalog(catalogStore.Load(), "dotnet-load");
        break;
    }

    default:
        Console.Error.WriteLine($"unknown mode: {mode}");
        return 2;
}

Console.WriteLine(failures == 0
    ? $"STORAGE PROBE PASS: {checks} checks"
    : $"STORAGE PROBE FAILED: {failures} of {checks} checks");
return failures == 0 ? 0 : 1;
