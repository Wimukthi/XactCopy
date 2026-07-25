// -----------------------------------------------------------------------------
// File: cpp\tools\GoldenGen\Program.cs
// Purpose: Emits golden JSON artifacts using the real XactCopy.Core serializer
//          (System.Text.Json) so the native C++ port can byte-compare its own
//          output. Values are deterministic; timestamps use fixed exotic
//          fractions to pin the trimming/formatting rules.
// -----------------------------------------------------------------------------

using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using XactCopy.Ipc;
using XactCopy.Ipc.Messages;
using XactCopy.Models;

var outputDir = args.Length > 0 ? args[0] : Path.Combine("..", "..", "tests", "golden");
Directory.CreateDirectory(outputDir);

// Identical to IpcSerializer.BuildOptions().
var serializerOptions = new JsonSerializerOptions { PropertyNameCaseInsensitive = true };
serializerOptions.Converters.Add(new JsonStringEnumConverter());

void WriteGolden(string fileName, string content)
{
    var path = Path.Combine(outputDir, fileName);
    File.WriteAllText(path, content, new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));
    Console.WriteLine($"wrote {path} ({content.Length} chars)");
}

// ---------------------------------------------------------------------------
// 1. String escaping table: every ASCII char plus multilingual/emoji samples.
// ---------------------------------------------------------------------------
var escapeSamples = new List<string>();
for (int c = 0; c < 128; c++)
{
    escapeSamples.Add(((char)c).ToString());
}
escapeSamples.Add("plain ascii text");
escapeSamples.Add("quote\" backslash\\ <tag> & 'x' +1 `t");
escapeSamples.Add("newline\ntab\tcr\r");
escapeSamples.Add("café € 😀 ñ 日本語 сложно");
escapeSamples.Add("D:\\Source\\folder\\file (copy) [1] {x} ~!@#$%^&*()_+-=");
escapeSamples.Add(" ");
WriteGolden("string_escapes.json", JsonSerializer.Serialize(escapeSamples, serializerOptions));

// ---------------------------------------------------------------------------
// 2. Default CopyJobOptions (all defaults, exactly as the UI constructs one).
// ---------------------------------------------------------------------------
WriteGolden("options_default.json", JsonSerializer.Serialize(new CopyJobOptions(), serializerOptions));

// ---------------------------------------------------------------------------
// 3. Envelopes with fixed timestamps (constructed like IpcSerializer does).
// ---------------------------------------------------------------------------
string SerializeEnvelope<T>(string messageType, T payload, string correlationId, DateTimeOffset sentUtc)
{
    var envelope = new IpcEnvelope<T>
    {
        ProtocolVersion = ProtocolConstants.Version,
        MessageType = messageType,
        CorrelationId = correlationId,
        SentUtc = sentUtc,
        Payload = payload
    };
    return JsonSerializer.Serialize(envelope, serializerOptions);
}

var logEvent = new WorkerLogEvent
{
    JobId = "8f2c1f6f2ab84a1f9d3e5b7c9a1d2e3f",
    Message = "Copy failed for <D:\\data\\file+name's.bin> & retried \"twice\" — done.",
    TimestampUtc = new DateTimeOffset(2026, 7, 23, 1, 2, 3, TimeSpan.Zero).AddTicks(40005)
};
WriteGolden("envelope_log.json",
    SerializeEnvelope(IpcMessageTypes.WorkerLogEvent, logEvent, "corr-9",
        new DateTimeOffset(2026, 7, 23, 1, 2, 3, TimeSpan.Zero)));

var heartbeat = new WorkerHeartbeatEvent
{
    WorkerProcessId = 4242,
    IsJobRunning = true,
    IsJobPaused = false,
    ActiveJobId = "job-active-1",
    LastProgressUtc = DateTimeOffset.MinValue, // sentinel round-trip case
    TimestampUtc = new DateTimeOffset(2026, 7, 23, 4, 5, 6, TimeSpan.Zero).AddTicks(1234567)
};
WriteGolden("envelope_heartbeat.json",
    SerializeEnvelope(IpcMessageTypes.WorkerHeartbeatEvent, heartbeat, "",
        new DateTimeOffset(2026, 7, 23, 4, 5, 6, TimeSpan.Zero).AddTicks(5000000)));

var startCommand = new StartJobCommand
{
    JobId = "0123456789abcdef0123456789abcdef",
    Options = new CopyJobOptions
    {
        OperationMode = JobOperationMode.ScanOnly,
        SourceRoot = "D:\\",
        DestinationRoot = "E:\\Backup dir\\данные",
        ExpectedSourceIdentity = "VOLSER:1234-ABCD",
        UseBadRangeMap = true,
        BadRangeMapMaxAgeDays = 45,
        ScanPerformanceProfile = ScanPerformanceProfile.Fast,
        SelectedRelativePaths = { "a.txt", "sub\\file+plus.bin", "日本語.dat" },
        OverwritePolicy = OverwritePolicy.OverwriteIfSourceNewer,
        SymlinkHandling = SymlinkHandlingMode.Follow,
        BufferSizeBytes = 1 * 1024 * 1024,
        UseAdaptiveBufferSizing = true,
        TransferEnginePolicy = TransferEnginePolicy.ManagedRescue,
        MaxThroughputBytesPerSecond = 50_000_000,
        ParallelSmallFileWorkers = 4,
        LockContentionProbeInterval = TimeSpan.FromMilliseconds(750),
        SourceMutationPolicy = SourceMutationPolicy.WaitForReappearance,
        FragileMediaMode = true,
        MaxRetries = 3,
        OperationTimeout = TimeSpan.FromSeconds(90),
        PerFileTimeout = TimeSpan.FromMinutes(2.5),
        InitialRetryDelay = TimeSpan.FromMilliseconds(125),
        MaxRetryDelay = new TimeSpan(1, 2, 3, 4, 567),
        VerifyAfterCopy = true,
        VerificationMode = VerificationMode.Sampled,
        VerificationHashAlgorithm = VerificationHashAlgorithm.Sha512,
        SalvageFillPattern = SalvageFillPattern.Random,
        WorkerProcessPriorityClass = "BelowNormal",
        WorkerTelemetryProfile = WorkerTelemetryProfile.Verbose,
        RescueFastScanChunkBytes = 65536,
        RescueScrapeRetries = 5
    }
};
WriteGolden("envelope_start.json",
    SerializeEnvelope(IpcMessageTypes.StartJobCommand, startCommand, "",
        new DateTimeOffset(2026, 7, 23, 9, 0, 0, TimeSpan.Zero).AddTicks(1)));

var progressEvent = new WorkerProgressEvent
{
    JobId = "job-progress",
    Snapshot = new CopyProgressSnapshot
    {
        CurrentFile = "D:\\big\\payload.iso",
        CurrentFileBytesCopied = 1_073_741_824,
        CurrentFileBytesTotal = 4_294_967_296,
        TotalBytesCopied = 5_000_000_000,
        TotalBytes = 10_000_000_000,
        LastChunkBytesTransferred = 1048576,
        BufferSizeBytes = 4194304,
        CompletedFiles = 12,
        FailedFiles = 1,
        RecoveredFiles = 2,
        SkippedFiles = 3,
        TotalFiles = 40,
        RescuePass = "Trim",
        RescueBadRegionCount = 7,
        RescueRemainingBytes = 123456,
        ActiveFileCount = 2,
        ScanWorkerCount = 4,
        ActiveFiles = { "a.bin", "b.bin" }
    },
    TimestampUtc = new DateTimeOffset(2026, 7, 23, 10, 20, 30, TimeSpan.Zero).AddTicks(9999999)
};
WriteGolden("envelope_progress.json",
    SerializeEnvelope(IpcMessageTypes.WorkerProgressEvent, progressEvent, "",
        new DateTimeOffset(2026, 7, 23, 10, 20, 30, TimeSpan.Zero)));

var resultEvent = new WorkerJobResultEvent
{
    JobId = "job-result",
    Result = new CopyJobResult
    {
        Succeeded = true,
        Cancelled = false,
        TotalFiles = 100,
        CompletedFiles = 97,
        FailedFiles = 1,
        RecoveredFiles = 2,
        SkippedFiles = 0,
        TotalBytes = 123_456_789_012,
        CopiedBytes = 123_000_000_000,
        TransferEnginePolicy = TransferEnginePolicy.NativeFast,
        ElapsedMilliseconds = 654321,
        AverageBytesPerSecond = 188034187.25,
        NativeFastPathFiles = 90,
        ParallelNativeFastPathFiles = 45,
        ManagedCopyFiles = 7,
        NativeFallbackFiles = 3,
        JournalPath = "C:\\Users\\user\\AppData\\Local\\XactCopy\\journals\\j1.json",
        ErrorMessage = ""
    },
    TimestampUtc = new DateTimeOffset(2026, 7, 23, 11, 0, 0, TimeSpan.FromHours(5.5))
};
WriteGolden("envelope_result.json",
    SerializeEnvelope(IpcMessageTypes.WorkerJobResultEvent, resultEvent, "",
        new DateTimeOffset(2026, 7, 23, 11, 0, 0, TimeSpan.FromHours(-7))));

// ---------------------------------------------------------------------------
// 4. Storage payload goldens (WriteIndented=true — the .NET storage stores'
//    serializer). Pins newline style, indent width, empty-container layout.
// ---------------------------------------------------------------------------
var storageSerializer = new JsonSerializerOptions { WriteIndented = true };
storageSerializer.Converters.Add(new JsonStringEnumConverter());

var goldenMap = new BadRangeMap
{
    SchemaVersion = 1,
    SourceRoot = "D:\\GoldenSrc",
    SourceIdentity = "SER-0042 & <id>",
    UpdatedUtc = new DateTimeOffset(2026, 7, 23, 12, 0, 0, TimeSpan.Zero).AddTicks(1230000)
};
goldenMap.Files["alpha.txt"] = new BadRangeMapFileEntry
{
    RelativePath = "alpha.txt",
    SourceLength = 1234,
    LastWriteUtcTicks = 638000000000000000,
    FileFingerprint = "fp-alpha",
    BadRanges = { new ByteRange { Offset = 100, Length = 50 }, new ByteRange { Offset = 300, Length = 25 } },
    LastScanUtc = new DateTimeOffset(2026, 7, 23, 12, 0, 0, TimeSpan.Zero),
    LastError = "CRC error"
};
goldenMap.Files["sub\\бета.bin"] = new BadRangeMapFileEntry
{
    RelativePath = "sub\\бета.bin",
    SourceLength = 0,
    LastWriteUtcTicks = 0,
    FileFingerprint = "",
    LastScanUtc = new DateTimeOffset(2026, 7, 23, 12, 0, 0, TimeSpan.FromHours(5.5)),
    LastError = ""
};
WriteGolden("golden_map_payload.json", JsonSerializer.Serialize(goldenMap, storageSerializer));

var goldenJournal = new JobJournal
{
    JobId = "golden-journal-01",
    SourceRoot = "D:\\GoldenSrc",
    DestinationRoot = "E:\\GoldenDest",
    CreatedUtc = new DateTimeOffset(2026, 7, 23, 0, 0, 0, TimeSpan.Zero),
    UpdatedUtc = new DateTimeOffset(2026, 7, 23, 1, 30, 0, TimeSpan.Zero).AddTicks(5000000)
};
goldenJournal.Files["alpha.txt"] = new JournalFileEntry
{
    RelativePath = "alpha.txt",
    SourceLength = 1234,
    SourceLastWriteUtcTicks = 638000000000000000,
    BytesCopied = 1234,
    State = FileCopyState.CompletedWithRecovery,
    LastError = "",
    DoNotRetry = false,
    RecoveredRanges = { new ByteRange { Offset = 0, Length = 128 } },
    RescueRanges = { new RescueRange { Offset = 512, Length = 64, State = RescueRangeState.Recovered } },
    LastRescuePass = "Scrape"
};
goldenJournal.Files["empty.dat"] = new JournalFileEntry
{
    RelativePath = "empty.dat",
    State = FileCopyState.Pending
};
WriteGolden("golden_journal_payload.json", JsonSerializer.Serialize(goldenJournal, storageSerializer));

var emptyJournal = new JobJournal
{
    JobId = "empty-journal",
    SourceRoot = "D:\\X",
    DestinationRoot = "E:\\Y",
    CreatedUtc = new DateTimeOffset(2026, 1, 1, 0, 0, 0, TimeSpan.Zero),
    UpdatedUtc = new DateTimeOffset(2026, 1, 1, 0, 0, 0, TimeSpan.Zero)
};
WriteGolden("golden_journal_empty.json", JsonSerializer.Serialize(emptyJournal, storageSerializer));

Console.WriteLine("golden generation complete");
