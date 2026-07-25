Imports System.Diagnostics
Imports System.Globalization
Imports System.IO
Imports System.Text
Imports System.Text.Json
Imports System.Threading
Imports XactCopy.Models
Imports XactCopy.Services

Module Program
    Private Const DefaultSmallFileCount As Integer = 500
    Private Const DefaultSmallFileSizeBytes As Integer = 8 * 1024
    Private Const DefaultLargeFileCount As Integer = 3
    Private Const DefaultLargeFileSizeBytes As Integer = 16 * 1024 * 1024
    Private Const DefaultMixedSmallFileCount As Integer = 250
    Private Const DefaultMixedLargeFileCount As Integer = 2

    Private ReadOnly JsonOptions As New JsonSerializerOptions() With {
        .WriteIndented = True
    }

    Function Main(args As String()) As Integer
        Try
            Return MainAsync(args).GetAwaiter().GetResult()
        Catch ex As Exception
            Console.Error.WriteLine(ex.ToString())
            Return 1
        End Try
    End Function

    Private Async Function MainAsync(args As String()) As Task(Of Integer)
        Dim config = BenchmarkConfig.Parse(args)
        If config.ShowHelp Then
            Console.WriteLine(BenchmarkConfig.GetUsage())
            Return 0
        End If

        Directory.CreateDirectory(config.WorkRoot)
        Directory.CreateDirectory(config.OutputDirectory)

        Dim summary As New BenchmarkSummary() With {
            .StartedUtc = DateTimeOffset.UtcNow,
            .WorkRoot = config.WorkRoot,
            .OutputDirectory = config.OutputDirectory,
            .Iterations = config.Iterations,
            .SmallFileWorkers = config.SmallFileWorkers,
            .ScanWorkers = config.ScanWorkers,
            .SmallFileThresholdBytes = config.SmallFileThresholdBytes,
            .Modes = config.Modes.ToList()
        }

        Console.WriteLine($"XactCopy benchmark root: {config.WorkRoot}")
        Console.WriteLine($"Report output: {config.OutputDirectory}")

        For Each scenario In config.Scenarios
            Dim sourceRoot = Path.Combine(config.WorkRoot, "sources", scenario)
            Dim manifest = Await EnsureScenarioAsync(config, scenario, sourceRoot, CancellationToken.None).ConfigureAwait(False)
            summary.Scenarios.Add(manifest)

            For Each mode In config.Modes
                For iteration = 1 To config.Iterations
                    For Each policy In config.Policies
                        Dim run = Await ExecuteRunAsync(
                            config,
                            mode,
                            scenario,
                            policy,
                            iteration,
                            sourceRoot,
                            CancellationToken.None).ConfigureAwait(False)
                        summary.Runs.Add(run)

                        Dim statusText = If(run.Succeeded, "ok", "failed")
                        Console.WriteLine(
                            $"{scenario} {mode} {policy} iter {iteration}: {statusText}, {FormatBytesPerSecond(run.BytesPerSecond)}, {run.ElapsedMilliseconds} ms")
                    Next
                Next
            Next
        Next

        summary.CompletedUtc = DateTimeOffset.UtcNow
        WriteReports(summary, config.OutputDirectory)

        If Not config.KeepWorkRoot Then
            Try
                Directory.Delete(config.WorkRoot, recursive:=True)
            Catch ex As Exception
                Console.Error.WriteLine($"Work root cleanup skipped: {ex.Message}")
            End Try
        End If

        Return If(summary.Runs.Any(Function(run) Not run.Succeeded), 2, 0)
    End Function

    Private Async Function EnsureScenarioAsync(
        config As BenchmarkConfig,
        scenario As String,
        sourceRoot As String,
        cancellationToken As CancellationToken) As Task(Of ScenarioManifest)

        If Directory.Exists(sourceRoot) Then
            Directory.Delete(sourceRoot, recursive:=True)
        End If

        Directory.CreateDirectory(sourceRoot)

        Dim manifest As New ScenarioManifest() With {
            .Name = scenario,
            .SourceRoot = sourceRoot
        }

        Select Case scenario.ToLowerInvariant()
            Case "small"
                Await CreateFileSetAsync(sourceRoot, "small", config.SmallFileCount, config.SmallFileSizeBytes, manifest, cancellationToken).ConfigureAwait(False)

            Case "large"
                Await CreateFileSetAsync(sourceRoot, "large", config.LargeFileCount, config.LargeFileSizeBytes, manifest, cancellationToken).ConfigureAwait(False)

            Case "mixed"
                Await CreateFileSetAsync(sourceRoot, "mixed-small", config.MixedSmallFileCount, config.SmallFileSizeBytes, manifest, cancellationToken).ConfigureAwait(False)
                Await CreateFileSetAsync(sourceRoot, "mixed-large", config.MixedLargeFileCount, config.LargeFileSizeBytes, manifest, cancellationToken).ConfigureAwait(False)

            Case Else
                Throw New InvalidOperationException($"Unsupported scenario: {scenario}")
        End Select

        Return manifest
    End Function

    Private Async Function CreateFileSetAsync(
        root As String,
        prefix As String,
        fileCount As Integer,
        fileSizeBytes As Integer,
        manifest As ScenarioManifest,
        cancellationToken As CancellationToken) As Task

        For index = 0 To fileCount - 1
            cancellationToken.ThrowIfCancellationRequested()

            Dim relativeDirectory = If(index Mod 10 = 0, Path.Combine(prefix, $"group-{index \ 10:D3}"), prefix)
            Dim targetDirectory = Path.Combine(root, relativeDirectory)
            Directory.CreateDirectory(targetDirectory)

            Dim relativePath = Path.Combine(relativeDirectory, $"{prefix}-{index:D5}.bin")
            Dim fullPath = Path.Combine(root, relativePath)
            Await WritePatternFileAsync(fullPath, fileSizeBytes, seed:=index + prefix.Length, cancellationToken).ConfigureAwait(False)

            manifest.FileCount += 1
            manifest.TotalBytes += fileSizeBytes
        Next
    End Function

    Private Async Function WritePatternFileAsync(
        pathValue As String,
        length As Integer,
        seed As Integer,
        cancellationToken As CancellationToken) As Task

        Dim buffer = New Byte(Math.Min(1024 * 1024, Math.Max(4096, length)) - 1) {}
        Using stream As New FileStream(pathValue, FileMode.Create, FileAccess.Write, FileShare.Read, buffer.Length, FileOptions.SequentialScan)
            Dim remaining = length
            Dim offset = 0
            While remaining > 0
                cancellationToken.ThrowIfCancellationRequested()
                Dim count = Math.Min(buffer.Length, remaining)
                FillPattern(buffer, count, seed, offset)
                Await stream.WriteAsync(buffer.AsMemory(0, count), cancellationToken).ConfigureAwait(False)
                remaining -= count
                offset += count
            End While
        End Using
    End Function

    Private Sub FillPattern(buffer As Byte(), count As Integer, seed As Integer, offset As Integer)
        For index = 0 To count - 1
            buffer(index) = CByte((seed + offset + index) Mod 251)
        Next
    End Sub

    Private Async Function ExecuteRunAsync(
        config As BenchmarkConfig,
        mode As String,
        scenario As String,
        policy As TransferEnginePolicy,
        iteration As Integer,
        sourceRoot As String,
        cancellationToken As CancellationToken) As Task(Of BenchmarkRunResult)

        Dim scanMode = String.Equals(mode, "scan", StringComparison.OrdinalIgnoreCase)
        Dim destinationRoot = String.Empty
        If Not scanMode Then
            destinationRoot = Path.Combine(config.WorkRoot, "destinations", scenario, policy.ToString(), $"iter-{iteration:D3}")
            If Directory.Exists(destinationRoot) Then
                Directory.Delete(destinationRoot, recursive:=True)
            End If
            Directory.CreateDirectory(destinationRoot)
        End If

        Dim options As New CopyJobOptions() With {
            .OperationMode = If(scanMode, JobOperationMode.ScanOnly, JobOperationMode.Copy),
            .SourceRoot = sourceRoot,
            .DestinationRoot = destinationRoot,
            .ResumeFromJournal = False,
            .ContinueOnFileError = False,
            .SalvageUnreadableBlocks = False,
            .UseAdaptiveBufferSizing = config.UseAdaptiveBufferSizing,
            .UseBadRangeMap = False,
            .UpdateBadRangeMapFromRun = False,
            .SkipKnownBadRanges = False,
            .ScanPerformanceProfile = config.ScanPerformanceProfile,
            .TransferEnginePolicy = policy,
            .ParallelSmallFileWorkers = config.SmallFileWorkers,
            .ParallelScanWorkers = config.ScanWorkers,
            .SmallFileThresholdBytes = config.SmallFileThresholdBytes,
            .BufferSizeBytes = config.BufferSizeBytes,
            .MaxRetries = 1,
            .OperationTimeout = TimeSpan.FromSeconds(30),
            .PreserveTimestamps = False,
            .WorkerTelemetryProfile = WorkerTelemetryProfile.Normal
        }

        Dim progressCount = 0
        Dim logCount = 0
        Dim observedPasses As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
        Dim lastSnapshot As CopyProgressSnapshot = Nothing

        Dim service As New ResilientCopyService(options)
        AddHandler service.ProgressChanged,
            Sub(sender, snapshot)
                progressCount += 1
                lastSnapshot = snapshot
                If snapshot IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(snapshot.RescuePass) Then
                    observedPasses.Add(snapshot.RescuePass)
                End If
            End Sub
        AddHandler service.LogMessage,
            Sub(sender, message)
                logCount += 1
            End Sub

        Dim watch = Stopwatch.StartNew()
        Dim result = Await service.RunAsync(cancellationToken).ConfigureAwait(False)
        watch.Stop()

        Dim bytesPerSecond = If(watch.Elapsed.TotalSeconds > 0, CDbl(result.CopiedBytes) / watch.Elapsed.TotalSeconds, 0.0R)
        Return New BenchmarkRunResult() With {
            .Mode = If(scanMode, "scan", "copy"),
            .Scenario = scenario,
            .Policy = policy.ToString(),
            .Iteration = iteration,
            .Succeeded = result.Succeeded,
            .Cancelled = result.Cancelled,
            .ErrorMessage = result.ErrorMessage,
            .TotalFiles = result.TotalFiles,
            .CompletedFiles = result.CompletedFiles,
            .SkippedFiles = result.SkippedFiles,
            .FailedFiles = result.FailedFiles,
            .RecoveredFiles = result.RecoveredFiles,
            .TotalBytes = result.TotalBytes,
            .CopiedBytes = result.CopiedBytes,
            .ElapsedMilliseconds = CLng(watch.Elapsed.TotalMilliseconds),
            .BytesPerSecond = bytesPerSecond,
            .ProgressEvents = progressCount,
            .LogEvents = logCount,
            .LastProgressBytes = If(lastSnapshot Is Nothing, 0L, lastSnapshot.TotalBytesCopied),
            .NativeFastPathFiles = result.NativeFastPathFiles,
            .ParallelNativeFastPathFiles = result.ParallelNativeFastPathFiles,
            .ManagedCopyFiles = result.ManagedCopyFiles,
            .NativeFallbackFiles = result.NativeFallbackFiles,
            .ObservedPasses = observedPasses.OrderBy(Function(value) value, StringComparer.OrdinalIgnoreCase).ToList()
        }
    End Function

    Private Sub WriteReports(summary As BenchmarkSummary, outputDirectory As String)
        Dim jsonPath = Path.Combine(outputDirectory, "benchmark-results.json")
        File.WriteAllText(jsonPath, JsonSerializer.Serialize(summary, JsonOptions), Encoding.UTF8)

        Dim markdownPath = Path.Combine(outputDirectory, "benchmark-results.md")
        File.WriteAllText(markdownPath, BuildMarkdown(summary), Encoding.UTF8)

        Console.WriteLine($"Wrote {jsonPath}")
        Console.WriteLine($"Wrote {markdownPath}")
    End Sub

    Private Function BuildMarkdown(summary As BenchmarkSummary) As String
        Dim builder As New StringBuilder()
        builder.AppendLine("# XactCopy Benchmark Results")
        builder.AppendLine()
        builder.AppendLine($"Started: {summary.StartedUtc:O}")
        builder.AppendLine($"Completed: {summary.CompletedUtc:O}")
        builder.AppendLine($"Work root: `{summary.WorkRoot}`")
        builder.AppendLine($"Iterations: {summary.Iterations}")
        builder.AppendLine($"Modes: {String.Join(", ", summary.Modes)}")
        builder.AppendLine($"Small-file workers: {summary.SmallFileWorkers}")
        builder.AppendLine($"Scan workers: {summary.ScanWorkers}")
        builder.AppendLine($"Small-file threshold: {FormatBytes(summary.SmallFileThresholdBytes)}")
        builder.AppendLine()

        builder.AppendLine("## Scenarios")
        builder.AppendLine()
        builder.AppendLine("| Scenario | Files | Bytes |")
        builder.AppendLine("|---|---:|---:|")
        For Each scenario In summary.Scenarios
            builder.AppendLine($"| {scenario.Name} | {scenario.FileCount} | {FormatBytes(scenario.TotalBytes)} |")
        Next
        builder.AppendLine()

        builder.AppendLine("## Runs")
        builder.AppendLine()
        builder.AppendLine("| Scenario | Mode | Policy | Iteration | Status | Files | Bytes | Time | Throughput | Engine Mix | Passes |")
        builder.AppendLine("|---|---|---|---:|---|---:|---:|---:|---:|---|---|")
        For Each run In summary.Runs
            Dim statusText = If(run.Succeeded, "Succeeded", If(run.Cancelled, "Cancelled", "Failed"))
            Dim passes = If(run.ObservedPasses Is Nothing OrElse run.ObservedPasses.Count = 0, "", String.Join(", ", run.ObservedPasses))
            Dim engineMix = $"native {run.NativeFastPathFiles}, parallel {run.ParallelNativeFastPathFiles}, managed {run.ManagedCopyFiles}, fallback {run.NativeFallbackFiles}"
            builder.AppendLine(
                $"| {run.Scenario} | {run.Mode} | {run.Policy} | {run.Iteration} | {statusText} | {run.CompletedFiles}/{run.TotalFiles} | {FormatBytes(run.CopiedBytes)} | {run.ElapsedMilliseconds} ms | {FormatBytesPerSecond(run.BytesPerSecond)} | {engineMix} | {passes} |")
        Next

        builder.AppendLine()
        builder.AppendLine("## Mode/Policy Averages")
        builder.AppendLine()
        builder.AppendLine("| Scenario | Mode | Policy | Runs | Avg Throughput | Best Throughput | Avg Time |")
        builder.AppendLine("|---|---|---|---:|---:|---:|---:|")
        For Each group In summary.Runs.Where(Function(run) run.Succeeded).GroupBy(Function(run) $"{run.Scenario}|{run.Mode}|{run.Policy}")
            Dim first = group.First()
            builder.AppendLine(
                $"| {first.Scenario} | {first.Mode} | {first.Policy} | {group.Count()} | {FormatBytesPerSecond(group.Average(Function(run) run.BytesPerSecond))} | {FormatBytesPerSecond(group.Max(Function(run) run.BytesPerSecond))} | {group.Average(Function(run) run.ElapsedMilliseconds):0} ms |")
        Next

        Return builder.ToString()
    End Function

    Private Function FormatBytes(value As Long) As String
        If value < 1024 Then
            Return $"{value} B"
        End If

        Dim units = {"KB", "MB", "GB", "TB"}
        Dim size = CDbl(value)
        Dim unitIndex = -1
        Do
            size /= 1024.0R
            unitIndex += 1
        Loop While size >= 1024.0R AndAlso unitIndex < units.Length - 1

        Return $"{size:0.##} {units(unitIndex)}"
    End Function

    Private Function FormatBytesPerSecond(value As Double) As String
        Return $"{FormatBytes(CLng(Math.Max(0.0R, value)))}/s"
    End Function

    Private NotInheritable Class BenchmarkConfig
        Public Property ShowHelp As Boolean
        Public Property WorkRoot As String = Path.Combine(Path.GetTempPath(), $"xactcopy-bench-{DateTimeOffset.UtcNow:yyyyMMdd-HHmmss}")
        Public Property OutputDirectory As String = Path.GetFullPath(Path.Combine("artifacts", "benchmarks", DateTimeOffset.UtcNow.ToString("yyyyMMdd-HHmmss", CultureInfo.InvariantCulture)))
        Public Property Modes As New List(Of String) From {"copy"}
        Public Property Scenarios As New List(Of String) From {"small", "mixed", "large"}
        Public Property Policies As New List(Of TransferEnginePolicy) From {
            TransferEnginePolicy.Auto,
            TransferEnginePolicy.ManagedRescue,
            TransferEnginePolicy.NativeFast
        }
        Public Property Iterations As Integer = 1
        Public Property SmallFileCount As Integer = DefaultSmallFileCount
        Public Property SmallFileSizeBytes As Integer = DefaultSmallFileSizeBytes
        Public Property LargeFileCount As Integer = DefaultLargeFileCount
        Public Property LargeFileSizeBytes As Integer = DefaultLargeFileSizeBytes
        Public Property MixedSmallFileCount As Integer = DefaultMixedSmallFileCount
        Public Property MixedLargeFileCount As Integer = DefaultMixedLargeFileCount
        Public Property SmallFileWorkers As Integer = Math.Max(1, Math.Min(8, Environment.ProcessorCount))
        Public Property ScanWorkers As Integer = Math.Max(1, Math.Min(8, Environment.ProcessorCount))
        Public Property SmallFileThresholdBytes As Integer = 256 * 1024
        Public Property BufferSizeBytes As Integer = 4 * 1024 * 1024
        Public Property UseAdaptiveBufferSizing As Boolean = False
        Public Property ScanPerformanceProfile As ScanPerformanceProfile = ScanPerformanceProfile.Auto
        Public Property KeepWorkRoot As Boolean

        Public Shared Function Parse(args As String()) As BenchmarkConfig
            Dim config As New BenchmarkConfig()
            Dim index = 0
            While index < args.Length
                Dim arg = args(index)
                Select Case arg.ToLowerInvariant()
                    Case "-h", "--help", "/?"
                        config.ShowHelp = True

                    Case "--work-root"
                        config.WorkRoot = RequireValue(args, index)
                        index += 1

                    Case "--output"
                        config.OutputDirectory = Path.GetFullPath(RequireValue(args, index))
                        index += 1

                    Case "--scenario", "--scenarios"
                        config.Scenarios = SplitList(RequireValue(args, index)).
                            Select(Function(value) value.ToLowerInvariant()).
                            Where(Function(value) value.Length > 0).
                            ToList()
                        index += 1

                    Case "--mode", "--modes"
                        config.Modes = SplitList(RequireValue(args, index)).
                            Select(AddressOf ParseMode).
                            ToList()
                        index += 1

                    Case "--policies", "--policy"
                        config.Policies = SplitList(RequireValue(args, index)).
                            Select(AddressOf ParsePolicy).
                            ToList()
                        index += 1

                    Case "--iterations"
                        config.Iterations = ParsePositiveInteger(RequireValue(args, index), "iterations")
                        index += 1

                    Case "--small-files"
                        config.SmallFileCount = ParsePositiveInteger(RequireValue(args, index), "small-files")
                        index += 1

                    Case "--small-size-kb"
                        config.SmallFileSizeBytes = ParsePositiveInteger(RequireValue(args, index), "small-size-kb") * 1024
                        index += 1

                    Case "--large-files"
                        config.LargeFileCount = ParsePositiveInteger(RequireValue(args, index), "large-files")
                        index += 1

                    Case "--large-size-mb"
                        config.LargeFileSizeBytes = ParsePositiveInteger(RequireValue(args, index), "large-size-mb") * 1024 * 1024
                        index += 1

                    Case "--workers"
                        config.SmallFileWorkers = Math.Max(1, Math.Min(64, ParsePositiveInteger(RequireValue(args, index), "workers")))
                        index += 1

                    Case "--scan-workers"
                        config.ScanWorkers = Math.Max(1, Math.Min(64, ParsePositiveInteger(RequireValue(args, index), "scan-workers")))
                        index += 1

                    Case "--scan-profile"
                        config.ScanPerformanceProfile = ParseScanPerformanceProfile(RequireValue(args, index))
                        index += 1

                    Case "--threshold-kb"
                        config.SmallFileThresholdBytes = Math.Max(4096, ParsePositiveInteger(RequireValue(args, index), "threshold-kb") * 1024)
                        index += 1

                    Case "--buffer-mb"
                        config.BufferSizeBytes = Math.Max(4096, ParsePositiveInteger(RequireValue(args, index), "buffer-mb") * 1024 * 1024)
                        index += 1

                    Case "--adaptive-buffer"
                        config.UseAdaptiveBufferSizing = True

                    Case "--keep"
                        config.KeepWorkRoot = True

                    Case Else
                        Throw New ArgumentException($"Unknown argument: {arg}")
                End Select

                index += 1
            End While

            If config.Scenarios.Count = 0 Then
                Throw New ArgumentException("At least one scenario is required.")
            End If

            If config.Modes.Count = 0 Then
                Throw New ArgumentException("At least one mode is required.")
            End If

            If config.Policies.Count = 0 Then
                Throw New ArgumentException("At least one policy is required.")
            End If

            Return config
        End Function

        Public Shared Function GetUsage() As String
            Return String.Join(
                Environment.NewLine,
                "XactCopy.Benchmarks",
                "",
                "Usage:",
                "  dotnet run --project tools/XactCopy.Benchmarks -- [options]",
                "",
                "Options:",
                "  --mode copy,scan                  Operation modes. Default: copy.",
                "  --scenario small,mixed,large      Scenario list. Default: all.",
                "  --policies auto,managed,native    Policy list. Default: all.",
                "  --iterations N                    Iterations per scenario/policy. Default: 1.",
                "  --workers N                       Small-file workers. Default: min(Environment.ProcessorCount, 8).",
                "  --scan-workers N                  Fast scan workers. Default: min(Environment.ProcessorCount, 8).",
                "  --scan-profile auto|fast|precise  Scan-only engine profile. Default: auto.",
                "  --threshold-kb N                  Small-file threshold. Default: 256.",
                "  --small-files N                   Small-file count. Default: 500.",
                "  --small-size-kb N                 Small-file size. Default: 8.",
                "  --large-files N                   Large-file count. Default: 3.",
                "  --large-size-mb N                 Large-file size. Default: 16.",
                "  --buffer-mb N                     Copy buffer size. Default: 4.",
                "  --adaptive-buffer                 Enable adaptive buffer sizing.",
                "  --work-root PATH                  Benchmark scratch root. Default: temp.",
                "  --output PATH                     Report output directory. Default: artifacts/benchmarks/<timestamp>.",
                "  --keep                            Keep generated source/destination trees.")
        End Function

        Private Shared Function RequireValue(args As String(), index As Integer) As String
            If index + 1 >= args.Length Then
                Throw New ArgumentException($"Missing value for {args(index)}.")
            End If

            Return args(index + 1)
        End Function

        Private Shared Function SplitList(value As String) As IEnumerable(Of String)
            Return If(value, String.Empty).
                Split(","c, StringSplitOptions.RemoveEmptyEntries Or StringSplitOptions.TrimEntries)
        End Function

        Private Shared Function ParsePositiveInteger(value As String, name As String) As Integer
            Dim parsed As Integer
            If Not Integer.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, parsed) OrElse parsed <= 0 Then
                Throw New ArgumentException($"{name} must be a positive integer.")
            End If

            Return parsed
        End Function

        Private Shared Function ParsePolicy(value As String) As TransferEnginePolicy
            Select Case If(value, String.Empty).Trim().ToLowerInvariant()
                Case "auto"
                    Return TransferEnginePolicy.Auto
                Case "managed", "managedrescue", "managed-rescue"
                    Return TransferEnginePolicy.ManagedRescue
                Case "native", "nativefast", "native-fast"
                    Return TransferEnginePolicy.NativeFast
                Case Else
                    Throw New ArgumentException($"Unsupported transfer policy: {value}")
            End Select
        End Function

        Private Shared Function ParseScanPerformanceProfile(value As String) As ScanPerformanceProfile
            Select Case If(value, String.Empty).Trim().ToLowerInvariant()
                Case "auto"
                    Return ScanPerformanceProfile.Auto
                Case "fast"
                    Return ScanPerformanceProfile.Fast
                Case "precise"
                    Return ScanPerformanceProfile.Precise
                Case Else
                    Throw New ArgumentException($"Unsupported scan profile: {value}")
            End Select
        End Function

        Private Shared Function ParseMode(value As String) As String
            Select Case If(value, String.Empty).Trim().ToLowerInvariant()
                Case "copy"
                    Return "copy"
                Case "scan", "scanonly", "scan-only"
                    Return "scan"
                Case Else
                    Throw New ArgumentException($"Unsupported benchmark mode: {value}")
            End Select
        End Function
    End Class

    Private NotInheritable Class BenchmarkSummary
        Public Property StartedUtc As DateTimeOffset
        Public Property CompletedUtc As DateTimeOffset
        Public Property WorkRoot As String = String.Empty
        Public Property OutputDirectory As String = String.Empty
        Public Property Iterations As Integer
        Public Property SmallFileWorkers As Integer
        Public Property ScanWorkers As Integer
        Public Property SmallFileThresholdBytes As Integer
        Public Property Modes As New List(Of String)()
        Public Property Scenarios As New List(Of ScenarioManifest)()
        Public Property Runs As New List(Of BenchmarkRunResult)()
    End Class

    Private NotInheritable Class ScenarioManifest
        Public Property Name As String = String.Empty
        Public Property SourceRoot As String = String.Empty
        Public Property FileCount As Integer
        Public Property TotalBytes As Long
    End Class

    Private NotInheritable Class BenchmarkRunResult
        Public Property Mode As String = String.Empty
        Public Property Scenario As String = String.Empty
        Public Property Policy As String = String.Empty
        Public Property Iteration As Integer
        Public Property Succeeded As Boolean
        Public Property Cancelled As Boolean
        Public Property ErrorMessage As String = String.Empty
        Public Property TotalFiles As Integer
        Public Property CompletedFiles As Integer
        Public Property SkippedFiles As Integer
        Public Property FailedFiles As Integer
        Public Property RecoveredFiles As Integer
        Public Property TotalBytes As Long
        Public Property CopiedBytes As Long
        Public Property ElapsedMilliseconds As Long
        Public Property BytesPerSecond As Double
        Public Property ProgressEvents As Integer
        Public Property LogEvents As Integer
        Public Property LastProgressBytes As Long
        Public Property NativeFastPathFiles As Integer
        Public Property ParallelNativeFastPathFiles As Integer
        Public Property ManagedCopyFiles As Integer
        Public Property NativeFallbackFiles As Integer
        Public Property ObservedPasses As New List(Of String)()
    End Class
End Module
