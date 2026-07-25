// -----------------------------------------------------------------------------
// File: cpp\tools\InteropProbe\Program.cs
// Purpose: End-to-end tests for the native XactCopyExecutive driven through
//          the real .NET IPC stack (JsonMessagePipe + IpcSerializer), exactly
//          as WorkerSupervisor does. Covers the protocol conversation, clean
//          copies (native + managed engines), verification, and fault-injected
//          rescue/salvage with byte-exact destination validation.
// -----------------------------------------------------------------------------

using System.Diagnostics;
using System.IO.Pipes;
using System.Security.Cryptography;
using XactCopy.Ipc;
using XactCopy.Ipc.Messages;
using XactCopy.Models;

if (args.Length < 1)
{
    Console.Error.WriteLine("usage: InteropProbe <path-to-XactCopyExecutive.exe>");
    return 2;
}

var workerPath = Path.GetFullPath(args[0]);
if (!File.Exists(workerPath))
{
    Console.Error.WriteLine($"worker not found: {workerPath}");
    return 2;
}

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

var workRoot = Path.Combine(Path.GetTempPath(), $"xactcopy-interop-{Guid.NewGuid():N}");
Directory.CreateDirectory(workRoot);

byte[] DeterministicBytes(int length, int seed)
{
    var data = new byte[length];
    var random = new Random(seed);
    random.NextBytes(data);
    return data;
}

void BuildSourceTree(string sourceRoot)
{
    Directory.CreateDirectory(sourceRoot);
    Directory.CreateDirectory(Path.Combine(sourceRoot, "sub"));
    Directory.CreateDirectory(Path.Combine(sourceRoot, "empty-dir"));
    File.WriteAllBytes(Path.Combine(sourceRoot, "a.txt"), DeterministicBytes(1024, 1));
    File.WriteAllBytes(Path.Combine(sourceRoot, "sub", "бета.bin"), DeterministicBytes(300 * 1024, 2));
    File.WriteAllBytes(Path.Combine(sourceRoot, "sub", "large.dat"), DeterministicBytes(2 * 1024 * 1024, 3));
    File.WriteAllBytes(Path.Combine(sourceRoot, "zero.bin"), Array.Empty<byte>());
}

bool TreesEqual(string sourceRoot, string destinationRoot, out string mismatch)
{
    mismatch = string.Empty;
    foreach (var sourceFile in Directory.EnumerateFiles(sourceRoot, "*", SearchOption.AllDirectories))
    {
        var relative = Path.GetRelativePath(sourceRoot, sourceFile);
        var destinationFile = Path.Combine(destinationRoot, relative);
        if (!File.Exists(destinationFile))
        {
            mismatch = $"missing {relative}";
            return false;
        }
        var sourceBytes = File.ReadAllBytes(sourceFile);
        var destinationBytes = File.ReadAllBytes(destinationFile);
        if (!sourceBytes.AsSpan().SequenceEqual(destinationBytes))
        {
            mismatch = $"content mismatch {relative}";
            return false;
        }
    }
    return true;
}

// --- Worker session helper --------------------------------------------------

async Task<WorkerSession> StartWorkerAsync(Dictionary<string, string>? environment = null)
{
    var pipeName = $"xactcopy-{Guid.NewGuid():N}";
    var startInfo = new ProcessStartInfo(workerPath, $"--pipe \"{pipeName}\"")
    {
        UseShellExecute = false,
        CreateNoWindow = true,
        WindowStyle = ProcessWindowStyle.Hidden
    };
    if (environment is not null)
    {
        foreach (var (key, value) in environment)
        {
            startInfo.Environment[key] = value;
        }
    }

    var process = Process.Start(startInfo)!;
    var pipe = new NamedPipeClientStream(".", pipeName, PipeDirection.InOut, PipeOptions.Asynchronous);
    await pipe.ConnectAsync((int)TimeSpan.FromSeconds(15).TotalMilliseconds);
    return new WorkerSession(process, pipe);
}

CopyJobOptions DefaultOptions(string sourceRoot, string destinationRoot) => new()
{
    SourceRoot = sourceRoot,
    DestinationRoot = destinationRoot,
    ResumeFromJournal = true,
    PreserveTimestamps = true
};

// --- Scenario A: protocol + clean copy via native fast path ------------------
{
    Console.WriteLine("--- scenario A: protocol + native fast-path copy ---");
    using var session = await StartWorkerAsync();

    var connectLog = await session.ExpectAsync(IpcMessageTypes.WorkerLogEvent);
    Check(IpcSerializer.DeserializeEnvelope<WorkerLogEvent>(connectLog).Payload!.Message ==
          "Worker connected to supervisor.", "connection log message");

    var hb = IpcSerializer.DeserializeEnvelope<WorkerHeartbeatEvent>(
        await session.ExpectAsync(IpcMessageTypes.WorkerHeartbeatEvent));
    Check(hb.Payload!.WorkerProcessId == session.Process.Id, "heartbeat PID matches");
    Check(!hb.Payload!.IsJobRunning, "heartbeat reports no job running");

    await session.SendAsync(IpcMessageTypes.PingCommand, new PingCommand());
    await session.ExpectAsync(IpcMessageTypes.WorkerHeartbeatEvent);
    Check(true, "ping answered with heartbeat");

    await session.SendAsync(IpcMessageTypes.PauseJobCommand, new PauseJobCommand { JobId = "x", Reason = "t" });
    await session.ExpectAsync(IpcMessageTypes.WorkerLogEvent,
        j => IpcSerializer.DeserializeEnvelope<WorkerLogEvent>(j).Payload!.Message.Contains("ignored"));
    Check(true, "pause without job ignored");

    var sourceRoot = Path.Combine(workRoot, "a-src");
    var destinationRoot = Path.Combine(workRoot, "a-dst");
    BuildSourceTree(sourceRoot);

    var jobId = Guid.NewGuid().ToString("N");
    await session.SendAsync(IpcMessageTypes.StartJobCommand,
        new StartJobCommand { JobId = jobId, Options = DefaultOptions(sourceRoot, destinationRoot) });

    await session.ExpectAsync(IpcMessageTypes.WorkerLogEvent,
        j => IpcSerializer.DeserializeEnvelope<WorkerLogEvent>(j).Payload!.Message == "Job accepted by worker.");
    Check(true, "job accepted");

    var result = IpcSerializer.DeserializeEnvelope<WorkerJobResultEvent>(
        await session.ExpectAsync(IpcMessageTypes.WorkerJobResultEvent)).Payload!;
    Check(result.Result.Succeeded, $"A: job succeeded ({result.Result.ErrorMessage})");
    Check(result.JobId == jobId, "A: result JobId");
    Check(result.Result.TotalFiles == 4, $"A: total files ({result.Result.TotalFiles})");
    Check(result.Result.CompletedFiles == 4, $"A: completed files ({result.Result.CompletedFiles})");
    Check(result.Result.FailedFiles == 0, "A: no failures");
    Check(result.Result.NativeFastPathFiles > 0, $"A: native fast path used ({result.Result.NativeFastPathFiles})");
    Check(TreesEqual(sourceRoot, destinationRoot, out var mismatchA), $"A: destination matches source {mismatchA}");
    Check(Directory.Exists(Path.Combine(destinationRoot, "empty-dir")), "A: empty directory created");
    Check(session.ProgressEvents > 0, $"A: progress events observed ({session.ProgressEvents})");

    var sourceInfo = new FileInfo(Path.Combine(sourceRoot, "sub", "large.dat"));
    var destinationInfo = new FileInfo(Path.Combine(destinationRoot, "sub", "large.dat"));
    Check(Math.Abs((sourceInfo.LastWriteTimeUtc - destinationInfo.LastWriteTimeUtc).TotalSeconds) < 2,
        "A: timestamps preserved");

    await session.ShutdownAsync();
    Check(session.Process.WaitForExit(10000) && session.Process.ExitCode == 0, "A: clean worker exit");
}

// --- Scenario B: managed engine + full verification --------------------------
{
    Console.WriteLine("--- scenario B: managed rescue engine + full verification ---");
    using var session = await StartWorkerAsync();
    await session.DrainStartupAsync();

    var sourceRoot = Path.Combine(workRoot, "b-src");
    var destinationRoot = Path.Combine(workRoot, "b-dst");
    BuildSourceTree(sourceRoot);

    var options = DefaultOptions(sourceRoot, destinationRoot);
    options.TransferEnginePolicy = TransferEnginePolicy.ManagedRescue;
    options.VerificationMode = VerificationMode.Full;
    options.VerificationHashAlgorithm = VerificationHashAlgorithm.Sha256;

    var jobId = Guid.NewGuid().ToString("N");
    await session.SendAsync(IpcMessageTypes.StartJobCommand,
        new StartJobCommand { JobId = jobId, Options = options });

    var result = IpcSerializer.DeserializeEnvelope<WorkerJobResultEvent>(
        await session.ExpectAsync(IpcMessageTypes.WorkerJobResultEvent)).Payload!;
    Check(result.Result.Succeeded, $"B: job succeeded ({result.Result.ErrorMessage})");
    Check(result.Result.CompletedFiles == 4, $"B: completed files ({result.Result.CompletedFiles})");
    Check(result.Result.ManagedCopyFiles > 0, $"B: managed engine used ({result.Result.ManagedCopyFiles})");
    Check(result.Result.NativeFastPathFiles == 0, "B: native path suppressed by policy");
    Check(TreesEqual(sourceRoot, destinationRoot, out var mismatchB), $"B: destination matches source {mismatchB}");
    Check(session.Logs.Any(l => l.Contains("Verifying (full hash)")), "B: verification ran");

    await session.ShutdownAsync();
    Check(session.Process.WaitForExit(10000) && session.Process.ExitCode == 0, "B: clean worker exit");
}

// --- Scenario C: injected read faults -> rescue passes + salvage fill --------
{
    Console.WriteLine("--- scenario C: fault-injected rescue + salvage ---");
    var environment = new Dictionary<string, string>
    {
        // All reads overlapping [65536, 65536+4096) fail with an I/O error.
        ["XACTCOPY_DEV_FAULT_RULES"] = "read,65536,4096,always,io",
        ["XACTCOPY_DEV_FAULT_SEED"] = "42"
    };
    using var session = await StartWorkerAsync(environment);
    await session.DrainStartupAsync();

    var sourceRoot = Path.Combine(workRoot, "c-src");
    var destinationRoot = Path.Combine(workRoot, "c-dst");
    Directory.CreateDirectory(sourceRoot);
    var payload = DeterministicBytes(1024 * 1024, 7);
    File.WriteAllBytes(Path.Combine(sourceRoot, "big.bin"), payload);

    var options = DefaultOptions(sourceRoot, destinationRoot);
    options.SalvageUnreadableBlocks = true;
    options.SalvageFillPattern = SalvageFillPattern.Zero;
    options.MaxRetries = 1;
    options.InitialRetryDelay = TimeSpan.FromMilliseconds(5);
    options.MaxRetryDelay = TimeSpan.FromMilliseconds(20);

    var jobId = Guid.NewGuid().ToString("N");
    await session.SendAsync(IpcMessageTypes.StartJobCommand,
        new StartJobCommand { JobId = jobId, Options = options });

    var result = IpcSerializer.DeserializeEnvelope<WorkerJobResultEvent>(
        await session.ExpectAsync(IpcMessageTypes.WorkerJobResultEvent, timeoutSeconds: 120)).Payload!;
    Check(result.Result.Succeeded, $"C: job succeeded with salvage ({result.Result.ErrorMessage})");
    Check(result.Result.RecoveredFiles == 1, $"C: recovered files ({result.Result.RecoveredFiles})");
    Check(session.Logs.Any(l => l.Contains("[DevFault] Enabled")), "C: fault injector active");
    Check(session.Logs.Any(l => l.Contains("Salvaging remaining unreadable regions")), "C: salvage pass ran");

    var copied = File.ReadAllBytes(Path.Combine(destinationRoot, "big.bin"));
    Check(copied.Length == payload.Length, "C: destination length");

    // Expected: identical everywhere except the fault-aligned block(s)
    // [65536, 69632), which salvage zero-filled.
    var expected = (byte[])payload.Clone();
    Array.Clear(expected, 65536, 4096);
    bool matches = copied.AsSpan().SequenceEqual(expected);
    Check(matches, "C: destination equals source outside salvaged block, zeros inside");
    if (!matches)
    {
        int firstDiff = -1;
        for (int i = 0; i < copied.Length; i++)
        {
            if (copied[i] != expected[i]) { firstDiff = i; break; }
        }
        Console.WriteLine($"    first diff at offset {firstDiff}");
    }

    // Journal must carry the recovered range and CompletedWithRecovery state.
    var journalStore = new XactCopy.Infrastructure.JobJournalStore();
    var journal = await journalStore.LoadAsync(result.Result.JournalPath, CancellationToken.None);
    Check(journal is not null, "C: journal loads via .NET store");
    if (journal is not null)
    {
        Check(journal.Files.TryGetValue("big.bin", out var entry), "C: journal entry present");
        if (entry is not null)
        {
            Check(entry.State == FileCopyState.CompletedWithRecovery,
                $"C: journal state CompletedWithRecovery ({entry.State})");
            Check(entry.RecoveredRanges.Count == 1 &&
                  entry.RecoveredRanges[0].Offset == 65536 &&
                  entry.RecoveredRanges[0].Length == 4096,
                $"C: journal recovered range ({string.Join(",", entry.RecoveredRanges.Select(r => $"{r.Offset}+{r.Length}"))})");
        }
    }

    await session.ShutdownAsync();
    Check(session.Process.WaitForExit(10000) && session.Process.ExitCode == 0, "C: clean worker exit");
}

// --- Scenario D: faults without salvage -> file fails ------------------------
{
    Console.WriteLine("--- scenario D: fault without salvage fails file ---");
    var environment = new Dictionary<string, string>
    {
        ["XACTCOPY_DEV_FAULT_RULES"] = "read,0,0,always,io"
    };
    using var session = await StartWorkerAsync(environment);
    await session.DrainStartupAsync();

    var sourceRoot = Path.Combine(workRoot, "d-src");
    var destinationRoot = Path.Combine(workRoot, "d-dst");
    Directory.CreateDirectory(sourceRoot);
    File.WriteAllBytes(Path.Combine(sourceRoot, "doomed.bin"), DeterministicBytes(64 * 1024, 9));

    var options = DefaultOptions(sourceRoot, destinationRoot);
    options.SalvageUnreadableBlocks = false;
    options.MaxRetries = 1;
    options.InitialRetryDelay = TimeSpan.FromMilliseconds(5);
    options.MaxRetryDelay = TimeSpan.FromMilliseconds(20);

    var jobId = Guid.NewGuid().ToString("N");
    await session.SendAsync(IpcMessageTypes.StartJobCommand,
        new StartJobCommand { JobId = jobId, Options = options });

    var result = IpcSerializer.DeserializeEnvelope<WorkerJobResultEvent>(
        await session.ExpectAsync(IpcMessageTypes.WorkerJobResultEvent, timeoutSeconds: 120)).Payload!;
    Check(!result.Result.Succeeded, "D: job reports failure");
    Check(result.Result.FailedFiles == 1, $"D: failed files ({result.Result.FailedFiles})");
    Check(result.Result.CompletedFiles == 0, "D: no completions");

    await session.ShutdownAsync();
    Check(session.Process.WaitForExit(10000) && session.Process.ExitCode == 0, "D: clean worker exit");
}

// --- Scenario E: scan-only (precise) on a healthy tree ------------------------
{
    Console.WriteLine("--- scenario E: precise scan-only on healthy tree ---");
    using var session = await StartWorkerAsync();
    await session.DrainStartupAsync();

    var sourceRoot = Path.Combine(workRoot, "e-src");
    BuildSourceTree(sourceRoot);

    var options = DefaultOptions(sourceRoot, string.Empty);
    options.OperationMode = JobOperationMode.ScanOnly;
    options.ScanPerformanceProfile = ScanPerformanceProfile.Precise;
    options.ResumeFromJournal = false;

    var jobId = Guid.NewGuid().ToString("N");
    await session.SendAsync(IpcMessageTypes.StartJobCommand,
        new StartJobCommand { JobId = jobId, Options = options });

    var result = IpcSerializer.DeserializeEnvelope<WorkerJobResultEvent>(
        await session.ExpectAsync(IpcMessageTypes.WorkerJobResultEvent, timeoutSeconds: 120)).Payload!;
    Check(result.Result.Succeeded, $"E: scan succeeded ({result.Result.ErrorMessage})");
    Check(result.Result.TotalFiles == 4, $"E: total files ({result.Result.TotalFiles})");
    Check(result.Result.CompletedFiles == 4, $"E: completed files ({result.Result.CompletedFiles})");
    Check(result.Result.RecoveredFiles == 0, "E: no bad ranges detected");
    Check(session.Logs.Any(l => l.Contains("Operation mode: Scan only")), "E: scan mode log");

    var journalStore = new XactCopy.Infrastructure.JobJournalStore();
    var journal = await journalStore.LoadAsync(result.Result.JournalPath, CancellationToken.None);
    Check(journal is not null && journal.Files.Count == 4 &&
          journal.Files.Values.All(f => f.State == FileCopyState.Completed),
        "E: journal entries all Completed");

    await session.ShutdownAsync();
    Check(session.Process.WaitForExit(10000) && session.Process.ExitCode == 0, "E: clean worker exit");
}

// --- Scenario F: fast scan with faults -> bad-range map written ---------------
var scanSourceRoot = Path.Combine(workRoot, "f-src");
{
    Console.WriteLine("--- scenario F: fast scan detects faults, writes bad-range map ---");
    var environment = new Dictionary<string, string>
    {
        ["XACTCOPY_DEV_FAULT_RULES"] = "read,65536,4096,always,io"
    };
    using var session = await StartWorkerAsync(environment);
    await session.DrainStartupAsync();

    Directory.CreateDirectory(scanSourceRoot);
    File.WriteAllBytes(Path.Combine(scanSourceRoot, "big.bin"), DeterministicBytes(1024 * 1024, 7));

    var options = DefaultOptions(scanSourceRoot, string.Empty);
    options.OperationMode = JobOperationMode.ScanOnly;
    options.ScanPerformanceProfile = ScanPerformanceProfile.Fast;
    options.ResumeFromJournal = false;
    options.UpdateBadRangeMapFromRun = true;
    options.MaxRetries = 1;
    options.InitialRetryDelay = TimeSpan.FromMilliseconds(5);
    options.MaxRetryDelay = TimeSpan.FromMilliseconds(20);

    var jobId = Guid.NewGuid().ToString("N");
    await session.SendAsync(IpcMessageTypes.StartJobCommand,
        new StartJobCommand { JobId = jobId, Options = options });

    var result = IpcSerializer.DeserializeEnvelope<WorkerJobResultEvent>(
        await session.ExpectAsync(IpcMessageTypes.WorkerJobResultEvent, timeoutSeconds: 120)).Payload!;
    Check(result.Result.Succeeded, $"F: scan succeeded ({result.Result.ErrorMessage})");
    Check(result.Result.CompletedFiles == 1, $"F: completed files ({result.Result.CompletedFiles})");
    Check(result.Result.RecoveredFiles == 1, $"F: bad-range detections ({result.Result.RecoveredFiles})");
    Check(session.Logs.Any(l => l.Contains("Fast scan engine:")), "F: fast scan engine ran");
    Check(session.Logs.Any(l => l.Contains("Fast scan fallback queued: big.bin")), "F: fallback queued");
    Check(session.Logs.Any(l => l.Contains("Scan detected unreadable ranges on big.bin")),
        "F: detection log");

    // The map must load through the real .NET store with the exact bad block.
    var mapPath = XactCopy.Infrastructure.BadRangeMapStore.GetDefaultMapPath(scanSourceRoot);
    var mapStore = new XactCopy.Infrastructure.BadRangeMapStore();
    var map = await mapStore.LoadAsync(mapPath, CancellationToken.None);
    Check(map is not null, "F: bad-range map loads via .NET store");
    if (map is not null)
    {
        Check(map.Files.TryGetValue("big.bin", out var mapEntry), "F: map entry present");
        if (mapEntry is not null)
        {
            Check(mapEntry.BadRanges.Count == 1 &&
                  mapEntry.BadRanges[0].Offset == 65536 && mapEntry.BadRanges[0].Length == 4096,
                $"F: map bad range ({string.Join(",", mapEntry.BadRanges.Select(r => $"{r.Offset}+{r.Length}"))})");
            Check(mapEntry.SourceLength == 1024 * 1024, "F: map entry source length");
        }
    }

    await session.ShutdownAsync();
    Check(session.Process.WaitForExit(10000) && session.Process.ExitCode == 0, "F: clean worker exit");
}

// --- Scenario G: copy with map hints skips known-bad reads --------------------
{
    Console.WriteLine("--- scenario G: copy with bad-range map hints + salvage ---");
    var environment = new Dictionary<string, string>
    {
        // Same fault region; with hints applied the engine never reads it.
        ["XACTCOPY_DEV_FAULT_RULES"] = "read,65536,4096,always,io"
    };
    using var session = await StartWorkerAsync(environment);
    await session.DrainStartupAsync();

    var destinationRoot = Path.Combine(workRoot, "g-dst");
    var options = DefaultOptions(scanSourceRoot, destinationRoot);
    options.UseBadRangeMap = true;
    options.SkipKnownBadRanges = true;
    options.UpdateBadRangeMapFromRun = true;
    options.SalvageUnreadableBlocks = true;
    options.SalvageFillPattern = SalvageFillPattern.Zero;
    options.MaxRetries = 1;
    options.InitialRetryDelay = TimeSpan.FromMilliseconds(5);
    options.MaxRetryDelay = TimeSpan.FromMilliseconds(20);

    var jobId = Guid.NewGuid().ToString("N");
    await session.SendAsync(IpcMessageTypes.StartJobCommand,
        new StartJobCommand { JobId = jobId, Options = options });

    var result = IpcSerializer.DeserializeEnvelope<WorkerJobResultEvent>(
        await session.ExpectAsync(IpcMessageTypes.WorkerJobResultEvent, timeoutSeconds: 120)).Payload!;
    Check(result.Result.Succeeded, $"G: copy succeeded ({result.Result.ErrorMessage})");
    Check(result.Result.RecoveredFiles == 1, $"G: recovered files ({result.Result.RecoveredFiles})");
    Check(session.Logs.Any(l => l.Contains("Bad-range map loaded:")), "G: map loaded");
    Check(session.Logs.Any(l => l.Contains("Applied bad-range map hints to big.bin")),
        "G: hints applied");
    Check(!session.Logs.Any(l => l.Contains("Read retry")),
        "G: no read retries (known-bad block never read)");
    Check(session.Logs.Any(l => l.Contains("Salvaging remaining unreadable regions")),
        "G: salvage ran");

    var payload = DeterministicBytes(1024 * 1024, 7);
    var expected = (byte[])payload.Clone();
    Array.Clear(expected, 65536, 4096);
    var copied = File.ReadAllBytes(Path.Combine(destinationRoot, "big.bin"));
    Check(copied.AsSpan().SequenceEqual(expected),
        "G: destination equals source with zeros in known-bad block");

    await session.ShutdownAsync();
    Check(session.Process.WaitForExit(10000) && session.Process.ExitCode == 0, "G: clean worker exit");
}

// --- Scenario H: parallel small-file phase ------------------------------------
{
    Console.WriteLine("--- scenario H: parallel small-file phase ---");
    using var session = await StartWorkerAsync();
    await session.DrainStartupAsync();

    var sourceRoot = Path.Combine(workRoot, "h-src");
    var destinationRoot = Path.Combine(workRoot, "h-dst");
    Directory.CreateDirectory(sourceRoot);
    for (int i = 0; i < 12; i++)
    {
        File.WriteAllBytes(Path.Combine(sourceRoot, $"small-{i:D2}.bin"), DeterministicBytes(8192, 100 + i));
    }

    var options = DefaultOptions(sourceRoot, destinationRoot);
    options.ParallelSmallFileWorkers = 4;

    var jobId = Guid.NewGuid().ToString("N");
    await session.SendAsync(IpcMessageTypes.StartJobCommand,
        new StartJobCommand { JobId = jobId, Options = options });

    var result = IpcSerializer.DeserializeEnvelope<WorkerJobResultEvent>(
        await session.ExpectAsync(IpcMessageTypes.WorkerJobResultEvent, timeoutSeconds: 120)).Payload!;
    Check(result.Result.Succeeded, $"H: job succeeded ({result.Result.ErrorMessage})");
    Check(result.Result.CompletedFiles == 12, $"H: completed files ({result.Result.CompletedFiles})");
    Check(result.Result.ParallelNativeFastPathFiles > 0,
        $"H: parallel native path used ({result.Result.ParallelNativeFastPathFiles})");
    Check(session.Logs.Any(l => l.Contains("Parallel small file phase:")), "H: phase log");
    Check(TreesEqual(sourceRoot, destinationRoot, out var mismatchH), $"H: destination matches source {mismatchH}");

    await session.ShutdownAsync();
    Check(session.Process.WaitForExit(10000) && session.Process.ExitCode == 0, "H: clean worker exit");
}

// --- Scenario I: adaptive buffer sizing through the managed engine ------------
{
    Console.WriteLine("--- scenario I: adaptive buffers + managed engine ---");
    using var session = await StartWorkerAsync();
    await session.DrainStartupAsync();

    var sourceRoot = Path.Combine(workRoot, "i-src");
    var destinationRoot = Path.Combine(workRoot, "i-dst");
    BuildSourceTree(sourceRoot);

    var options = DefaultOptions(sourceRoot, destinationRoot);
    options.TransferEnginePolicy = TransferEnginePolicy.ManagedRescue;
    options.UseAdaptiveBufferSizing = true;

    var jobId = Guid.NewGuid().ToString("N");
    await session.SendAsync(IpcMessageTypes.StartJobCommand,
        new StartJobCommand { JobId = jobId, Options = options });

    var result = IpcSerializer.DeserializeEnvelope<WorkerJobResultEvent>(
        await session.ExpectAsync(IpcMessageTypes.WorkerJobResultEvent, timeoutSeconds: 120)).Payload!;
    Check(result.Result.Succeeded, $"I: job succeeded ({result.Result.ErrorMessage})");
    Check(result.Result.CompletedFiles == 4, $"I: completed files ({result.Result.CompletedFiles})");
    Check(TreesEqual(sourceRoot, destinationRoot, out var mismatchI), $"I: destination matches source {mismatchI}");

    await session.ShutdownAsync();
    Check(session.Process.WaitForExit(10000) && session.Process.ExitCode == 0, "I: clean worker exit");
}

// Remove probe artifacts, including the bad-range map + mirror for f-src.
try
{
    var mapPath = XactCopy.Infrastructure.BadRangeMapStore.GetDefaultMapPath(scanSourceRoot);
    var mapDirectory = Path.GetDirectoryName(mapPath)!;
    var baseName = Path.GetFileNameWithoutExtension(mapPath);
    foreach (var file in Directory.EnumerateFiles(mapDirectory, baseName + "*"))
    {
        File.Delete(file);
    }
    var mirrorDirectory = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "XactCopy", "badmaps-mirror");
    if (Directory.Exists(mirrorDirectory))
    {
        foreach (var file in Directory.EnumerateFiles(mirrorDirectory, baseName + "*"))
        {
            File.Delete(file);
        }
    }
}
catch { }
try { Directory.Delete(workRoot, recursive: true); } catch { }

Console.WriteLine(failures == 0
    ? $"INTEROP PASS: {checks} checks"
    : $"INTEROP FAILED: {failures} of {checks} checks");
return failures == 0 ? 0 : 1;

// --- Session plumbing --------------------------------------------------------

sealed class WorkerSession : IDisposable
{
    public Process Process { get; }
    public List<string> Logs { get; } = new();
    public int ProgressEvents { get; private set; }

    private readonly NamedPipeClientStream _pipe;
    private readonly Queue<(string Type, string Json)> _pending = new();

    public WorkerSession(Process process, NamedPipeClientStream pipe)
    {
        Process = process;
        _pipe = pipe;
    }

    public async Task<(string Type, string Json)> ReadEventAsync(CancellationToken token)
    {
        if (_pending.Count > 0) return _pending.Dequeue();
        var raw = await JsonMessagePipe.ReadMessageAsync(_pipe, token)
            ?? throw new EndOfStreamException("pipe closed unexpectedly");
        var messageType = string.Empty;
        if (!IpcSerializer.TryReadMessageType(raw, ref messageType))
            throw new InvalidDataException($"malformed message: {raw[..Math.Min(200, raw.Length)]}");

        if (messageType == IpcMessageTypes.WorkerLogEvent)
        {
            Logs.Add(IpcSerializer.DeserializeEnvelope<WorkerLogEvent>(raw).Payload!.Message);
        }
        else if (messageType == IpcMessageTypes.WorkerProgressEvent)
        {
            ProgressEvents++;
        }
        return (messageType, raw);
    }

    public async Task<string> ExpectAsync(string expectedType, Func<string, bool>? predicate = null,
                                          int timeoutSeconds = 30)
    {
        using var cts = new CancellationTokenSource(TimeSpan.FromSeconds(timeoutSeconds));
        while (true)
        {
            var (type, json) = await ReadEventAsync(cts.Token);
            if (type == expectedType && (predicate is null || predicate(json)))
                return json;
        }
    }

    public async Task DrainStartupAsync()
    {
        await ExpectAsync(IpcMessageTypes.WorkerLogEvent,
            j => IpcSerializer.DeserializeEnvelope<WorkerLogEvent>(j).Payload!.Message ==
                 "Worker connected to supervisor.");
    }

    public async Task SendAsync<T>(string messageType, T payload)
    {
        var json = IpcSerializer.SerializeEnvelope(messageType, payload);
        await JsonMessagePipe.WriteMessageAsync(_pipe, json, CancellationToken.None);
    }

    public async Task ShutdownAsync()
    {
        await SendAsync(IpcMessageTypes.ShutdownCommand, new ShutdownCommand { Reason = "probe done" });
        try
        {
            await ExpectAsync(IpcMessageTypes.WorkerLogEvent,
                j => IpcSerializer.DeserializeEnvelope<WorkerLogEvent>(j).Payload!.Message ==
                     "Shutdown command received.", timeoutSeconds: 10);
        }
        catch
        {
        }
    }

    public void Dispose()
    {
        _pipe.Dispose();
        try { if (!Process.HasExited) Process.Kill(entireProcessTree: true); } catch { }
        Process.Dispose();
    }
}
