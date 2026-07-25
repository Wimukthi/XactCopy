' -----------------------------------------------------------------------------
' File: tests\XactCopy.Tests\ResilientCopyServiceIntegrationTests.vb
' Purpose: Source file for XactCopy runtime behavior.
' -----------------------------------------------------------------------------

Imports System.IO
Imports System.Collections.Generic
Imports System.Linq
Imports System.Threading
Imports System.Threading.Tasks
Imports XactCopy.Infrastructure
Imports XactCopy.Models
Imports XactCopy.Services
Imports Xunit

''' <summary>
''' Class ResilientCopyServiceIntegrationTests.
''' </summary>
Public Class ResilientCopyServiceIntegrationTests
    Private Const DevFaultRulesEnvVar As String = "XACTCOPY_DEV_FAULT_RULES"
    Private Shared ReadOnly DevFaultEnvGate As New SemaphoreSlim(1, 1)

    ''' <summary>
    ''' Computes CopyProgressSnapshot_CurrentFileProgress_HandlesLargeFiles.
    ''' </summary>
    <Fact>
    Public Sub CopyProgressSnapshot_CurrentFileProgress_HandlesLargeFiles()
        Dim fileLength = 40L * 1024L * 1024L * 1024L
        Dim copied = 30L * 1024L * 1024L * 1024L

        Dim snapshot As New CopyProgressSnapshot() With {
            .CurrentFile = "large.bin",
            .CurrentFileBytesCopied = copied,
            .CurrentFileBytesTotal = fileLength
        }

        Assert.InRange(snapshot.CurrentFileProgress, 0.7499R, 0.7501R)
    End Sub

    ''' <summary>
    ''' Computes RescueRangeProgress_TracksBeyondInt32Range.
    ''' </summary>
    <Fact>
    Public Sub RescueRangeProgress_TracksBeyondInt32Range()
        Dim fileLength = 40L * 1024L * 1024L * 1024L
        Dim targetCopied = 30L * 1024L * 1024L * 1024L
        Dim chunkLength = 1024 * 1024 * 1024

        Dim ranges As New List(Of RescueRange) From {
            New RescueRange() With {
                .Offset = 0L,
                .Length = fileLength,
                .State = RescueRangeState.Pending
            }
        }

        Dim setState = GetType(ResilientCopyService).GetMethod(
            "SetRescueRangeState",
            Reflection.BindingFlags.NonPublic Or Reflection.BindingFlags.Static)
        Assert.NotNull(setState)

        Dim offset = 0L
        While offset < targetCopied
            setState.Invoke(
                Nothing,
                New Object() {
                    ranges,
                    offset,
                    CLng(chunkLength),
                    RescueRangeState.Good
                })
            offset += chunkLength
        End While

        Dim displayProgress = GetType(ResilientCopyService).GetMethod(
            "GetDisplayProgressBytes",
            Reflection.BindingFlags.NonPublic Or Reflection.BindingFlags.Static)
        Assert.NotNull(displayProgress)

        Dim entry As New JournalFileEntry() With {
            .RelativePath = "large.bin",
            .SourceLength = fileLength,
            .RescueRanges = ranges
        }

        Dim displayBytes = CLng(displayProgress.Invoke(Nothing, New Object() {entry, fileLength, 0L}))

        Assert.Equal(targetCopied, displayBytes)
        Assert.Equal(fileLength, ranges.Sum(Function(range) range.Length))
        Assert.Contains(ranges, Function(range) range.State = RescueRangeState.Good AndAlso range.Length = targetCopied)
        Assert.Contains(ranges, Function(range) range.State = RescueRangeState.Pending AndAlso range.Length = fileLength - targetCopied)
    End Sub

    ''' <summary>
    ''' Computes RescueRangeProgress_AcceptsSingleRangeBeyondInt32Range.
    ''' </summary>
    <Fact>
    Public Sub RescueRangeProgress_AcceptsSingleRangeBeyondInt32Range()
        Dim fileLength = 5L * 1024L * 1024L * 1024L
        Dim knownBadBytes = 3L * 1024L * 1024L * 1024L

        Dim ranges As New List(Of RescueRange) From {
            New RescueRange() With {
                .Offset = 0L,
                .Length = fileLength,
                .State = RescueRangeState.Pending
            }
        }

        Dim setState = GetType(ResilientCopyService).GetMethod(
            "SetRescueRangeState",
            Reflection.BindingFlags.NonPublic Or Reflection.BindingFlags.Static)
        Assert.NotNull(setState)
        setState.Invoke(Nothing, New Object() {ranges, 0L, knownBadBytes, RescueRangeState.KnownBad})

        Dim displayProgress = GetType(ResilientCopyService).GetMethod(
            "GetDisplayProgressBytes",
            Reflection.BindingFlags.NonPublic Or Reflection.BindingFlags.Static)
        Assert.NotNull(displayProgress)

        Dim entry As New JournalFileEntry() With {
            .RelativePath = "mapped-large.bin",
            .SourceLength = fileLength,
            .RescueRanges = ranges
        }

        Dim displayBytes = CLng(displayProgress.Invoke(Nothing, New Object() {entry, fileLength, 0L}))

        Assert.Equal(knownBadBytes, displayBytes)
        Assert.Equal(fileLength, ranges.Sum(Function(range) range.Length))
        Assert.Contains(ranges, Function(range) range.State = RescueRangeState.KnownBad AndAlso range.Length = knownBadBytes)
        Assert.Contains(ranges, Function(range) range.State = RescueRangeState.Pending AndAlso range.Length = fileLength - knownBadBytes)
    End Sub

    ''' <summary>
    ''' Computes RunAsync_CopiesFileAndEmitsRescueTelemetry.
    ''' </summary>
    <Fact>
    Public Async Function RunAsync_CopiesFileAndEmitsRescueTelemetry() As Task
        Dim tempRoot = Path.Combine(Path.GetTempPath(), $"xactcopy-tests-{Guid.NewGuid():N}")
        Dim sourceRoot = Path.Combine(tempRoot, "src")
        Dim destinationRoot = Path.Combine(tempRoot, "dst")
        Directory.CreateDirectory(sourceRoot)
        Directory.CreateDirectory(destinationRoot)

        Dim sourcePath = Path.Combine(sourceRoot, "sample.bin")
        Dim content = New Byte(256 * 1024 - 1) {}
        For index = 0 To content.Length - 1
            content(index) = CByte(index Mod 251)
        Next
        File.WriteAllBytes(sourcePath, content)

        Dim options As New CopyJobOptions() With {
            .SourceRoot = sourceRoot,
            .DestinationRoot = destinationRoot,
            .ResumeFromJournal = True,
            .SalvageUnreadableBlocks = False,
            .ContinueOnFileError = False,
            .UseAdaptiveBufferSizing = True,
            .BufferSizeBytes = 4 * 1024 * 1024,
            .MaxRetries = 1,
            .OperationTimeout = TimeSpan.FromSeconds(5)
        }

        Dim lastSnapshot As CopyProgressSnapshot = Nothing
        Dim service As New ResilientCopyService(options)
        AddHandler service.ProgressChanged, Sub(sender, snapshot) lastSnapshot = snapshot

        Try
            Dim result = Await service.RunAsync(CancellationToken.None)

            Assert.True(result.Succeeded)
            Assert.False(result.Cancelled)
            Assert.True(result.ElapsedMilliseconds >= 0)
            Assert.True(result.AverageBytesPerSecond >= 0)
            Assert.Equal(TransferEnginePolicy.Auto, result.TransferEnginePolicy)
            Assert.True(result.NativeFastPathFiles + result.ManagedCopyFiles > 0)
            Assert.True(File.Exists(Path.Combine(destinationRoot, "sample.bin")))
            Assert.Equal(content.Length, New FileInfo(Path.Combine(destinationRoot, "sample.bin")).Length)
            Assert.NotNull(lastSnapshot)
            Assert.True(Not String.IsNullOrWhiteSpace(lastSnapshot.RescuePass))
            Assert.True(lastSnapshot.RescueRemainingBytes >= 0)
        Finally
            CleanupJournal(sourceRoot, destinationRoot)

            If Directory.Exists(tempRoot) Then
                Directory.Delete(tempRoot, recursive:=True)
            End If
        End Try
    End Function

    ''' <summary>
    ''' Computes RunAsync_ParallelSmallFilePhase_CopiesEligibleFiles.
    ''' </summary>
    <Fact>
    Public Async Function RunAsync_ParallelSmallFilePhase_CopiesEligibleFiles() As Task
        Dim tempRoot = Path.Combine(Path.GetTempPath(), $"xactcopy-tests-parallel-small-{Guid.NewGuid():N}")
        Dim sourceRoot = Path.Combine(tempRoot, "src")
        Dim destinationRoot = Path.Combine(tempRoot, "dst")
        Directory.CreateDirectory(sourceRoot)
        Directory.CreateDirectory(destinationRoot)

        Dim expectedFiles As New Dictionary(Of String, Byte())(StringComparer.OrdinalIgnoreCase)
        For index = 0 To 11
            Dim relativePath = $"small-{index:D2}.bin"
            Dim content = New Byte(8192 + index - 1) {}
            For byteIndex = 0 To content.Length - 1
                content(byteIndex) = CByte((index + byteIndex) Mod 251)
            Next

            expectedFiles(relativePath) = content
            File.WriteAllBytes(Path.Combine(sourceRoot, relativePath), content)
        Next

        Dim options As New CopyJobOptions() With {
            .SourceRoot = sourceRoot,
            .DestinationRoot = destinationRoot,
            .ResumeFromJournal = False,
            .SalvageUnreadableBlocks = False,
            .ContinueOnFileError = False,
            .UseAdaptiveBufferSizing = False,
            .TransferEnginePolicy = TransferEnginePolicy.NativeFast,
            .ParallelSmallFileWorkers = 4,
            .SmallFileThresholdBytes = 64 * 1024,
            .BufferSizeBytes = 4096,
            .MaxRetries = 1,
            .OperationTimeout = TimeSpan.FromSeconds(5)
        }

        Dim observedPasses As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
        Dim observedSnapshots As New List(Of CopyProgressSnapshot)()
        Dim service As New ResilientCopyService(options)
        AddHandler service.ProgressChanged,
            Sub(sender, snapshot)
                If snapshot Is Nothing Then
                    Return
                End If

                SyncLock observedSnapshots
                    observedSnapshots.Add(snapshot)
                End SyncLock

                If Not String.IsNullOrWhiteSpace(snapshot.RescuePass) Then
                    SyncLock observedPasses
                        observedPasses.Add(snapshot.RescuePass)
                    End SyncLock
                End If
            End Sub

        Try
            Dim result = Await service.RunAsync(CancellationToken.None)

            Assert.True(result.Succeeded)
            Assert.Equal(expectedFiles.Count, result.CompletedFiles)
            Assert.Equal(0, result.SkippedFiles)
            Assert.Equal(expectedFiles.Sum(Function(item) CLng(item.Value.Length)), result.CopiedBytes)
            Assert.True(result.ParallelNativeFastPathFiles > 0)
            Assert.Contains("ParallelNativeFast", observedPasses)
            Assert.All(
                observedSnapshots,
                Sub(snapshot)
                    Assert.True(snapshot.CurrentFileBytesCopied <= snapshot.CurrentFileBytesTotal)
                    Assert.True(snapshot.TotalBytesCopied <= snapshot.TotalBytes)
                End Sub)

            For Each expected In expectedFiles
                Dim destinationPath = Path.Combine(destinationRoot, expected.Key)
                Assert.True(File.Exists(destinationPath))
                Assert.Equal(expected.Value, File.ReadAllBytes(destinationPath))
            Next
        Finally
            CleanupJournal(sourceRoot, destinationRoot)

            If Directory.Exists(tempRoot) Then
                Directory.Delete(tempRoot, recursive:=True)
            End If
        End Try
    End Function

    ''' <summary>
    ''' Computes RunAsync_AdaptiveBuffer_UsesSmallBuffersWithoutUnderBufferingLargeFiles.
    ''' </summary>
    <Fact>
    Public Async Function RunAsync_AdaptiveBuffer_UsesSmallBuffersWithoutUnderBufferingLargeFiles() As Task
        Dim tempRoot = Path.Combine(Path.GetTempPath(), $"xactcopy-tests-adaptive-buffer-{Guid.NewGuid():N}")
        Dim sourceRoot = Path.Combine(tempRoot, "src")
        Dim destinationRoot = Path.Combine(tempRoot, "dst")
        Directory.CreateDirectory(sourceRoot)
        Directory.CreateDirectory(destinationRoot)

        File.WriteAllBytes(Path.Combine(sourceRoot, "small.bin"), New Byte(8 * 1024 - 1) {})
        File.WriteAllBytes(Path.Combine(sourceRoot, "large.bin"), New Byte(2 * 1024 * 1024 - 1) {})

        Dim options As New CopyJobOptions() With {
            .SourceRoot = sourceRoot,
            .DestinationRoot = destinationRoot,
            .ResumeFromJournal = False,
            .SalvageUnreadableBlocks = False,
            .ContinueOnFileError = False,
            .UseAdaptiveBufferSizing = True,
            .TransferEnginePolicy = TransferEnginePolicy.ManagedRescue,
            .ParallelSmallFileWorkers = 1,
            .SmallFileThresholdBytes = 4 * 1024,
            .BufferSizeBytes = 512 * 1024,
            .MaxRetries = 1,
            .OperationTimeout = TimeSpan.FromSeconds(5)
        }

        Dim observedBufferSizes As New Dictionary(Of String, Integer)(StringComparer.OrdinalIgnoreCase)
        Dim service As New ResilientCopyService(options)
        AddHandler service.ProgressChanged,
            Sub(sender, snapshot)
                If snapshot Is Nothing OrElse String.IsNullOrWhiteSpace(snapshot.CurrentFile) OrElse snapshot.BufferSizeBytes <= 0 Then
                    Return
                End If

                SyncLock observedBufferSizes
                    If Not observedBufferSizes.ContainsKey(snapshot.CurrentFile) Then
                        observedBufferSizes(snapshot.CurrentFile) = snapshot.BufferSizeBytes
                    End If
                End SyncLock
            End Sub

        Try
            Dim result = Await service.RunAsync(CancellationToken.None)

            Assert.True(result.Succeeded)
            Assert.True(observedBufferSizes.ContainsKey("small.bin"))
            Assert.True(observedBufferSizes.ContainsKey("large.bin"))
            Assert.True(observedBufferSizes("small.bin") < options.BufferSizeBytes)
            Assert.True(observedBufferSizes("large.bin") > options.BufferSizeBytes)
        Finally
            CleanupJournal(sourceRoot, destinationRoot)

            If Directory.Exists(tempRoot) Then
                Directory.Delete(tempRoot, recursive:=True)
            End If
        End Try
    End Function

    ''' <summary>
    ''' Computes RunAsync_AdaptiveBuffer_GrowsBufferForLargeManagedFiles.
    ''' </summary>
    <Fact>
    Public Async Function RunAsync_AdaptiveBuffer_GrowsBufferForLargeManagedFiles() As Task
        Dim tempRoot = Path.Combine(Path.GetTempPath(), $"xactcopy-tests-adaptive-buffer-growth-{Guid.NewGuid():N}")
        Dim sourceRoot = Path.Combine(tempRoot, "src")
        Dim destinationRoot = Path.Combine(tempRoot, "dst")
        Directory.CreateDirectory(sourceRoot)
        Directory.CreateDirectory(destinationRoot)

        Dim sourcePath = Path.Combine(sourceRoot, "large.bin")
        Using stream = File.Create(sourcePath)
            stream.SetLength(24L * 1024L * 1024L)
        End Using

        Dim options As New CopyJobOptions() With {
            .SourceRoot = sourceRoot,
            .DestinationRoot = destinationRoot,
            .ResumeFromJournal = False,
            .SalvageUnreadableBlocks = False,
            .ContinueOnFileError = False,
            .UseAdaptiveBufferSizing = True,
            .TransferEnginePolicy = TransferEnginePolicy.ManagedRescue,
            .ParallelSmallFileWorkers = 1,
            .SmallFileThresholdBytes = 4 * 1024,
            .BufferSizeBytes = 512 * 1024,
            .MaxRetries = 1,
            .OperationTimeout = TimeSpan.FromSeconds(5)
        }

        Dim observedBufferSizes As New List(Of Integer)()
        Dim observedLogLines As New List(Of String)()
        Dim service As New ResilientCopyService(options)
        AddHandler service.ProgressChanged,
            Sub(sender, snapshot)
                If snapshot Is Nothing OrElse
                    Not String.Equals(snapshot.CurrentFile, "large.bin", StringComparison.OrdinalIgnoreCase) OrElse
                    snapshot.BufferSizeBytes <= 0 Then

                    Return
                End If

                SyncLock observedBufferSizes
                    observedBufferSizes.Add(snapshot.BufferSizeBytes)
                End SyncLock
            End Sub
        AddHandler service.LogMessage,
            Sub(sender, message)
                If String.IsNullOrWhiteSpace(message) Then
                    Return
                End If

                SyncLock observedLogLines
                    observedLogLines.Add(message)
                End SyncLock
            End Sub

        Try
            Dim result = Await service.RunAsync(CancellationToken.None)

            Assert.True(result.Succeeded)
            Assert.True(observedBufferSizes.Count >= 4)
            Assert.True(observedBufferSizes.Max() > observedBufferSizes.Min())
            Assert.True(observedBufferSizes.Max() > options.BufferSizeBytes)
            Assert.Contains(observedLogLines, Function(line) line.Contains("Adaptive buffer summary:", StringComparison.OrdinalIgnoreCase))
            Assert.DoesNotContain(
                observedLogLines,
                Function(line) line.Contains("Adaptive buffer during", StringComparison.OrdinalIgnoreCase) AndAlso
                    line.Contains("throughput stayed stable", StringComparison.OrdinalIgnoreCase))
            Assert.DoesNotContain(
                observedLogLines,
                Function(line) line.Contains("throughput dropped below the best observed rate", StringComparison.OrdinalIgnoreCase))
            Assert.Equal(New FileInfo(sourcePath).Length, New FileInfo(Path.Combine(destinationRoot, "large.bin")).Length)
        Finally
            CleanupJournal(sourceRoot, destinationRoot)

            If Directory.Exists(tempRoot) Then
                Directory.Delete(tempRoot, recursive:=True)
            End If
        End Try
    End Function

    ''' <summary>
    ''' Computes RunAsync_ScanOnlyMode_CreatesBadRangeMapWithoutDestination.
    ''' </summary>
    <Fact>
    Public Async Function RunAsync_ScanOnlyMode_CreatesBadRangeMapWithoutDestination() As Task
        Dim tempRoot = Path.Combine(Path.GetTempPath(), $"xactcopy-tests-scan-{Guid.NewGuid():N}")
        Dim sourceRoot = Path.Combine(tempRoot, "src")
        Directory.CreateDirectory(sourceRoot)

        Dim sourcePath = Path.Combine(sourceRoot, "scan.bin")
        File.WriteAllBytes(sourcePath, New Byte(64 * 1024 - 1) {})

        Dim options As New CopyJobOptions() With {
            .OperationMode = JobOperationMode.ScanOnly,
            .SourceRoot = sourceRoot,
            .DestinationRoot = String.Empty,
            .UseBadRangeMap = True,
            .SkipKnownBadRanges = True,
            .UpdateBadRangeMapFromRun = True,
            .ScanPerformanceProfile = ScanPerformanceProfile.Precise,
            .ResumeFromJournal = False,
            .SalvageUnreadableBlocks = False,
            .ContinueOnFileError = False,
            .UseAdaptiveBufferSizing = False,
            .BufferSizeBytes = 4096,
            .MaxRetries = 1,
            .OperationTimeout = TimeSpan.FromSeconds(5)
        }

        Dim service As New ResilientCopyService(options)
        Dim mapPath = BadRangeMapStore.GetDefaultMapPath(sourceRoot)

        Try
            Dim result = Await service.RunAsync(CancellationToken.None)

            Assert.True(result.Succeeded)
            Assert.True(File.Exists(mapPath))

            Dim mapStore As New BadRangeMapStore()
            Dim loadedMap = Await mapStore.LoadAsync(mapPath, CancellationToken.None)
            Assert.NotNull(loadedMap)
            Assert.NotNull(loadedMap.Files)
        Finally
            CleanupJournal(sourceRoot, sourceRoot)
            DeleteBadMapArtifacts(mapPath)
            If Directory.Exists(tempRoot) Then
                Directory.Delete(tempRoot, recursive:=True)
            End If
        End Try
    End Function

    ''' <summary>
    ''' Computes RunAsync_ScanOnlyMode_UsesSmallFileFastPass.
    ''' </summary>
    <Fact>
    Public Async Function RunAsync_ScanOnlyMode_UsesSmallFileFastPass() As Task
        Dim tempRoot = Path.Combine(Path.GetTempPath(), $"xactcopy-tests-scan-fast-{Guid.NewGuid():N}")
        Dim sourceRoot = Path.Combine(tempRoot, "src")
        Directory.CreateDirectory(sourceRoot)

        File.WriteAllBytes(Path.Combine(sourceRoot, "a.bin"), New Byte(8 * 1024 - 1) {})
        File.WriteAllBytes(Path.Combine(sourceRoot, "b.bin"), New Byte(12 * 1024 - 1) {})

        Dim options As New CopyJobOptions() With {
            .OperationMode = JobOperationMode.ScanOnly,
            .SourceRoot = sourceRoot,
            .DestinationRoot = String.Empty,
            .UseBadRangeMap = True,
            .SkipKnownBadRanges = True,
            .UpdateBadRangeMapFromRun = True,
            .ScanPerformanceProfile = ScanPerformanceProfile.Precise,
            .ResumeFromJournal = False,
            .SalvageUnreadableBlocks = False,
            .ContinueOnFileError = False,
            .UseAdaptiveBufferSizing = False,
            .BufferSizeBytes = 4 * 1024,
            .SmallFileThresholdBytes = 256 * 1024,
            .MaxRetries = 1,
            .OperationTimeout = TimeSpan.FromSeconds(5)
        }

        Dim observedPasses As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
        Dim service As New ResilientCopyService(options)
        AddHandler service.ProgressChanged,
            Sub(sender, snapshot)
                If snapshot IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(snapshot.RescuePass) Then
                    SyncLock observedPasses
                        observedPasses.Add(snapshot.RescuePass)
                    End SyncLock
                End If
            End Sub

        Dim mapPath = BadRangeMapStore.GetDefaultMapPath(sourceRoot)

        Try
            Dim result = Await service.RunAsync(CancellationToken.None)

            Assert.True(result.Succeeded)
            Assert.Contains("ScanSmallFast", observedPasses)
        Finally
            CleanupJournal(sourceRoot, sourceRoot)
            DeleteBadMapArtifacts(mapPath)
            If Directory.Exists(tempRoot) Then
                Directory.Delete(tempRoot, recursive:=True)
            End If
        End Try
    End Function

    ''' <summary>
    ''' Computes RunAsync_ScanOnlyMode_FastProfileScansInParallel.
    ''' </summary>
    <Fact>
    Public Async Function RunAsync_ScanOnlyMode_FastProfileScansInParallel() As Task
        Dim tempRoot = Path.Combine(Path.GetTempPath(), $"xactcopy-tests-fast-scan-{Guid.NewGuid():N}")
        Dim sourceRoot = Path.Combine(tempRoot, "src")
        Directory.CreateDirectory(sourceRoot)

        Dim totalBytes = 0L
        For index = 0 To 15
            Dim content = New Byte(256 * 1024 - 1) {}
            For byteIndex = 0 To content.Length - 1
                content(byteIndex) = CByte((index + byteIndex) Mod 251)
            Next

            totalBytes += content.Length
            File.WriteAllBytes(Path.Combine(sourceRoot, $"fast-{index:D2}.bin"), content)
        Next

        Dim options As New CopyJobOptions() With {
            .OperationMode = JobOperationMode.ScanOnly,
            .SourceRoot = sourceRoot,
            .DestinationRoot = String.Empty,
            .UseBadRangeMap = False,
            .SkipKnownBadRanges = False,
            .UpdateBadRangeMapFromRun = False,
            .ScanPerformanceProfile = ScanPerformanceProfile.Fast,
            .ParallelScanWorkers = 4,
            .ResumeFromJournal = False,
            .SalvageUnreadableBlocks = False,
            .ContinueOnFileError = False,
            .UseAdaptiveBufferSizing = True,
            .BufferSizeBytes = 64 * 1024,
            .SmallFileThresholdBytes = 256 * 1024,
            .MaxRetries = 1,
            .OperationTimeout = TimeSpan.FromSeconds(5)
        }

        Dim observedFastPass = False
        Dim observedParallelWorkers = False
        Dim observedSnapshots As New List(Of CopyProgressSnapshot)()
        Dim service As New ResilientCopyService(options)
        AddHandler service.ProgressChanged,
            Sub(sender, snapshot)
                If snapshot Is Nothing Then
                    Return
                End If

                SyncLock observedSnapshots
                    observedSnapshots.Add(snapshot)
                End SyncLock

                If String.Equals(snapshot.RescuePass, "FastHealthScan", StringComparison.OrdinalIgnoreCase) Then
                    observedFastPass = True
                End If

                If snapshot.ScanWorkerCount > 1 Then
                    observedParallelWorkers = True
                End If
            End Sub

        Try
            Dim result = Await service.RunAsync(CancellationToken.None)

            Assert.True(result.Succeeded)
            Assert.Equal(16, result.CompletedFiles)
            Assert.Equal(0, result.FailedFiles)
            Assert.Equal(totalBytes, result.CopiedBytes)
            Assert.True(observedFastPass)
            Assert.True(observedParallelWorkers)
            Assert.NotEmpty(observedSnapshots)
            Assert.All(
                observedSnapshots,
                Sub(snapshot)
                    Assert.True(snapshot.CurrentFileBytesCopied <= snapshot.CurrentFileBytesTotal)
                    Assert.True(snapshot.TotalBytesCopied <= snapshot.TotalBytes)
                End Sub)
        Finally
            CleanupJournal(sourceRoot, sourceRoot)
            Dim mapPath = BadRangeMapStore.GetDefaultMapPath(sourceRoot)
            DeleteBadMapArtifacts(mapPath)
            If Directory.Exists(tempRoot) Then
                Directory.Delete(tempRoot, recursive:=True)
            End If
        End Try
    End Function

    ''' <summary>
    ''' Computes RunAsync_FastScanResume_DoesNotRereadCompletedFiles.
    ''' </summary>
    <Fact>
    Public Async Function RunAsync_FastScanResume_DoesNotRereadCompletedFiles() As Task
        Dim tempRoot = Path.Combine(Path.GetTempPath(), $"xactcopy-tests-fast-scan-resume-{Guid.NewGuid():N}")
        Dim sourceRoot = Path.Combine(tempRoot, "src")
        Directory.CreateDirectory(sourceRoot)

        Dim totalBytes = 0L
        For index = 0 To 7
            Dim content = New Byte(256 * 1024 - 1) {}
            For byteIndex = 0 To content.Length - 1
                content(byteIndex) = CByte((index + byteIndex) Mod 251)
            Next

            totalBytes += content.Length
            File.WriteAllBytes(Path.Combine(sourceRoot, $"resume-fast-{index:D2}.bin"), content)
        Next

        Dim options As New CopyJobOptions() With {
            .OperationMode = JobOperationMode.ScanOnly,
            .SourceRoot = sourceRoot,
            .DestinationRoot = String.Empty,
            .UseBadRangeMap = False,
            .SkipKnownBadRanges = False,
            .UpdateBadRangeMapFromRun = False,
            .ScanPerformanceProfile = ScanPerformanceProfile.Fast,
            .ParallelScanWorkers = 4,
            .ResumeFromJournal = True,
            .SalvageUnreadableBlocks = False,
            .ContinueOnFileError = False,
            .UseAdaptiveBufferSizing = True,
            .BufferSizeBytes = 64 * 1024,
            .SmallFileThresholdBytes = 256 * 1024,
            .MaxRetries = 1,
            .OperationTimeout = TimeSpan.FromSeconds(5)
        }

        Try
            Dim firstService As New ResilientCopyService(options)
            Dim firstResult = Await firstService.RunAsync(CancellationToken.None)
            Assert.True(firstResult.Succeeded)
            Assert.Equal(8, firstResult.CompletedFiles)
            Assert.Equal(totalBytes, firstResult.CopiedBytes)

            Dim rereadChunkSnapshots = 0
            Dim resumeSnapshots = 0
            Dim secondService As New ResilientCopyService(options)
            AddHandler secondService.ProgressChanged,
                Sub(sender, snapshot)
                    If snapshot Is Nothing Then
                        Return
                    End If

                    If String.Equals(snapshot.RescuePass, "FastHealthScan", StringComparison.OrdinalIgnoreCase) AndAlso
                        snapshot.LastChunkBytesTransferred > 0 Then
                        Interlocked.Increment(rereadChunkSnapshots)
                    End If

                    If String.Equals(snapshot.RescuePass, "FastHealthScanResume", StringComparison.OrdinalIgnoreCase) Then
                        Interlocked.Increment(resumeSnapshots)
                    End If
                End Sub

            Dim secondResult = Await secondService.RunAsync(CancellationToken.None)
            Assert.True(secondResult.Succeeded)
            Assert.Equal(8, secondResult.CompletedFiles)
            Assert.Equal(totalBytes, secondResult.CopiedBytes)
            Assert.Equal(0, rereadChunkSnapshots)
            Assert.True(resumeSnapshots > 0)
        Finally
            CleanupJournal(sourceRoot, sourceRoot)
            Dim mapPath = BadRangeMapStore.GetDefaultMapPath(sourceRoot)
            DeleteBadMapArtifacts(mapPath)
            If Directory.Exists(tempRoot) Then
                Directory.Delete(tempRoot, recursive:=True)
            End If
        End Try
    End Function

    ''' <summary>
    ''' Computes RunAsync_ScanOnlyMode_EmitsAccurateProgressBounds.
    ''' </summary>
    <Fact>
    Public Async Function RunAsync_ScanOnlyMode_EmitsAccurateProgressBounds() As Task
        Dim tempRoot = Path.Combine(Path.GetTempPath(), $"xactcopy-tests-scan-progress-{Guid.NewGuid():N}")
        Dim sourceRoot = Path.Combine(tempRoot, "src")
        Directory.CreateDirectory(sourceRoot)

        File.WriteAllBytes(Path.Combine(sourceRoot, "first.bin"), New Byte(8 * 1024 - 1) {})
        File.WriteAllBytes(Path.Combine(sourceRoot, "second.bin"), New Byte(12 * 1024 - 1) {})

        Dim options As New CopyJobOptions() With {
            .OperationMode = JobOperationMode.ScanOnly,
            .SourceRoot = sourceRoot,
            .DestinationRoot = String.Empty,
            .UseBadRangeMap = False,
            .SkipKnownBadRanges = False,
            .UpdateBadRangeMapFromRun = False,
            .ResumeFromJournal = False,
            .SalvageUnreadableBlocks = False,
            .ContinueOnFileError = False,
            .UseAdaptiveBufferSizing = True,
            .BufferSizeBytes = 4 * 1024,
            .SmallFileThresholdBytes = 256 * 1024,
            .MaxRetries = 1,
            .OperationTimeout = TimeSpan.FromSeconds(5)
        }

        Dim observedSnapshots As New List(Of CopyProgressSnapshot)()
        Dim service As New ResilientCopyService(options)
        AddHandler service.ProgressChanged,
            Sub(sender, snapshot)
                If snapshot Is Nothing Then
                    Return
                End If

                SyncLock observedSnapshots
                    observedSnapshots.Add(snapshot)
                End SyncLock
            End Sub

        Dim mapPath = BadRangeMapStore.GetDefaultMapPath(sourceRoot)

        Try
            Dim result = Await service.RunAsync(CancellationToken.None)

            Assert.True(result.Succeeded)
            Assert.Equal(2, result.CompletedFiles)
            Assert.Equal(result.TotalBytes, result.CopiedBytes)
            Assert.NotEmpty(observedSnapshots)
            Assert.All(
                observedSnapshots,
                Sub(snapshot)
                    Assert.True(snapshot.CurrentFileBytesCopied >= 0)
                    Assert.True(snapshot.CurrentFileBytesTotal >= 0)
                    Assert.True(snapshot.CurrentFileBytesCopied <= snapshot.CurrentFileBytesTotal)
                    Assert.True(snapshot.TotalBytesCopied >= 0)
                    Assert.True(snapshot.TotalBytesCopied <= snapshot.TotalBytes)
                    Assert.True(snapshot.LastChunkBytesTransferred >= 0)
                    Assert.True(snapshot.BufferSizeBytes >= 0)
                End Sub)

            Dim lastSnapshot = observedSnapshots.Last()
            Assert.Equal(result.TotalBytes, lastSnapshot.TotalBytes)
            Assert.Equal(result.CopiedBytes, lastSnapshot.TotalBytesCopied)
            Assert.Equal(lastSnapshot.CurrentFileBytesTotal, lastSnapshot.CurrentFileBytesCopied)
        Finally
            CleanupJournal(sourceRoot, sourceRoot)
            DeleteBadMapArtifacts(mapPath)
            If Directory.Exists(tempRoot) Then
                Directory.Delete(tempRoot, recursive:=True)
            End If
        End Try
    End Function

    ''' <summary>
    ''' Computes RunAsync_UseBadRangeMap_SkipsMappedUnreadableRangesAndSalvages.
    ''' </summary>
    <Fact>
    Public Async Function RunAsync_UseBadRangeMap_SkipsMappedUnreadableRangesAndSalvages() As Task
        Dim tempRoot = Path.Combine(Path.GetTempPath(), $"xactcopy-tests-map-{Guid.NewGuid():N}")
        Dim sourceRoot = Path.Combine(tempRoot, "src")
        Dim destinationRoot = Path.Combine(tempRoot, "dst")
        Directory.CreateDirectory(sourceRoot)
        Directory.CreateDirectory(destinationRoot)

        Dim sourcePath = Path.Combine(sourceRoot, "mapped.bin")
        Dim content = New Byte(128 * 1024 - 1) {}
        For index = 0 To content.Length - 1
            content(index) = CByte((index Mod 251) + 1)
        Next
        File.WriteAllBytes(sourcePath, content)

        Dim sourceInfo As New FileInfo(sourcePath)
        Dim relativePath = "mapped.bin"
        Dim mapPath = BadRangeMapStore.GetDefaultMapPath(sourceRoot)
        Dim mapStore As New BadRangeMapStore()
        Dim map As New BadRangeMap() With {
            .SourceRoot = sourceRoot,
            .UpdatedUtc = DateTimeOffset.UtcNow
        }
        map.Files(relativePath) = New BadRangeMapFileEntry() With {
            .RelativePath = relativePath,
            .SourceLength = sourceInfo.Length,
            .LastWriteUtcTicks = sourceInfo.LastWriteTimeUtc.Ticks,
            .FileFingerprint = $"{sourceInfo.Length:X16}:{sourceInfo.LastWriteTimeUtc.Ticks:X16}",
            .BadRanges = New List(Of ByteRange) From {
                New ByteRange() With {.Offset = 0, .Length = CInt(sourceInfo.Length)}
            },
            .LastScanUtc = DateTimeOffset.UtcNow
        }
        Await mapStore.SaveAsync(mapPath, map, CancellationToken.None)

        Dim options As New CopyJobOptions() With {
            .OperationMode = JobOperationMode.Copy,
            .SourceRoot = sourceRoot,
            .DestinationRoot = destinationRoot,
            .UseBadRangeMap = True,
            .SkipKnownBadRanges = True,
            .UpdateBadRangeMapFromRun = False,
            .ResumeFromJournal = False,
            .SalvageUnreadableBlocks = True,
            .SalvageFillPattern = SalvageFillPattern.Zero,
            .ContinueOnFileError = False,
            .UseAdaptiveBufferSizing = False,
            .BufferSizeBytes = 4096,
            .SmallFileThresholdBytes = 4096,
            .MaxRetries = 1,
            .OperationTimeout = TimeSpan.FromSeconds(5)
        }

        Dim service As New ResilientCopyService(options)

        Try
            Dim result = Await service.RunAsync(CancellationToken.None)

            Dim destinationPath = Path.Combine(destinationRoot, relativePath)
            Assert.True(result.Succeeded)
            Assert.True(result.RecoveredFiles > 0)
            Assert.True(File.Exists(destinationPath))

            Dim destinationBytes = File.ReadAllBytes(destinationPath)
            Assert.Equal(content.Length, destinationBytes.Length)
            Assert.False(content.SequenceEqual(destinationBytes))
        Finally
            CleanupJournal(sourceRoot, destinationRoot)
            DeleteBadMapArtifacts(mapPath)
            If Directory.Exists(tempRoot) Then
                Directory.Delete(tempRoot, recursive:=True)
            End If
        End Try
    End Function

#If DEBUG Then
    ''' <summary>
    ''' Computes RunAsync_DebugFaultReadRule_RecoversWithSalvage.
    ''' </summary>
    <Fact>
    Public Async Function RunAsync_DebugFaultReadRule_RecoversWithSalvage() As Task
        Await DevFaultEnvGate.WaitAsync()
        Dim tempRoot = Path.Combine(Path.GetTempPath(), $"xactcopy-tests-fault-read-{Guid.NewGuid():N}")
        Dim sourceRoot = Path.Combine(tempRoot, "src")
        Dim destinationRoot = Path.Combine(tempRoot, "dst")

        Try
            Environment.SetEnvironmentVariable(DevFaultRulesEnvVar, "read,4096,4096,always,io")

            Directory.CreateDirectory(sourceRoot)
            Directory.CreateDirectory(destinationRoot)

            Dim sourcePath = Path.Combine(sourceRoot, "fault-sample.bin")
            Dim content = New Byte(16384 - 1) {}
            For index = 0 To content.Length - 1
                content(index) = CByte(index Mod 251)
            Next
            File.WriteAllBytes(sourcePath, content)

            Dim options As New CopyJobOptions() With {
                .SourceRoot = sourceRoot,
                .DestinationRoot = destinationRoot,
                .ResumeFromJournal = True,
                .SalvageUnreadableBlocks = True,
                .ContinueOnFileError = False,
                .UseAdaptiveBufferSizing = False,
                .BufferSizeBytes = 4096,
                .SmallFileThresholdBytes = 256 * 1024,
                .MaxRetries = 1,
                .OperationTimeout = TimeSpan.FromSeconds(5)
            }

            Dim service As New ResilientCopyService(options)
            Dim result = Await service.RunAsync(CancellationToken.None)

            Assert.True(result.Succeeded)
            Assert.True(result.RecoveredFiles > 0)
            Assert.True(File.Exists(Path.Combine(destinationRoot, "fault-sample.bin")))
            Assert.Equal(content.Length, New FileInfo(Path.Combine(destinationRoot, "fault-sample.bin")).Length)
        Finally
            Environment.SetEnvironmentVariable(DevFaultRulesEnvVar, Nothing)
            CleanupJournal(sourceRoot, destinationRoot)
            If Directory.Exists(tempRoot) Then
                Directory.Delete(tempRoot, recursive:=True)
            End If
            DevFaultEnvGate.Release()
        End Try
    End Function

    ''' <summary>
    ''' Computes RunAsync_DebugFaultWriteRule_FailsWhenContinueDisabled.
    ''' </summary>
    <Fact>
    Public Async Function RunAsync_DebugFaultWriteRule_FailsWhenContinueDisabled() As Task
        Await DevFaultEnvGate.WaitAsync()
        Dim tempRoot = Path.Combine(Path.GetTempPath(), $"xactcopy-tests-fault-write-{Guid.NewGuid():N}")
        Dim sourceRoot = Path.Combine(tempRoot, "src")
        Dim destinationRoot = Path.Combine(tempRoot, "dst")

        Try
            Environment.SetEnvironmentVariable(DevFaultRulesEnvVar, "write,0,0,always,io")

            Directory.CreateDirectory(sourceRoot)
            Directory.CreateDirectory(destinationRoot)

            Dim sourcePath = Path.Combine(sourceRoot, "fault-sample.bin")
            File.WriteAllBytes(sourcePath, New Byte(8192 - 1) {})

            Dim options As New CopyJobOptions() With {
                .SourceRoot = sourceRoot,
                .DestinationRoot = destinationRoot,
                .ResumeFromJournal = True,
                .SalvageUnreadableBlocks = False,
                .ContinueOnFileError = False,
                .UseAdaptiveBufferSizing = False,
                .BufferSizeBytes = 4096,
                .MaxRetries = 1,
                .OperationTimeout = TimeSpan.FromSeconds(5)
            }

            Dim service As New ResilientCopyService(options)
            Dim result = Await service.RunAsync(CancellationToken.None)

            Assert.False(result.Succeeded)
            Assert.True(result.FailedFiles > 0)
            Assert.False(String.IsNullOrWhiteSpace(result.ErrorMessage))
        Finally
            Environment.SetEnvironmentVariable(DevFaultRulesEnvVar, Nothing)
            CleanupJournal(sourceRoot, destinationRoot)
            If Directory.Exists(tempRoot) Then
                Directory.Delete(tempRoot, recursive:=True)
            End If
            DevFaultEnvGate.Release()
        End Try
    End Function
#End If

    Private Shared Sub CleanupJournal(sourceRoot As String, destinationRoot As String)
        Try
            Dim jobId = JobJournalStore.BuildJobId(sourceRoot, destinationRoot)
            Dim journalPath = JobJournalStore.GetDefaultJournalPath(jobId)
            If File.Exists(journalPath) Then
                File.Delete(journalPath)
            End If
        Catch
        End Try
    End Sub

    Private Shared Sub DeleteBadMapArtifacts(mapPath As String)
        Try
            If String.IsNullOrWhiteSpace(mapPath) Then
                Return
            End If

            For Each candidatePath In New String() {
                mapPath,
                $"{mapPath}.bak1",
                $"{mapPath}.bak2"}
                If File.Exists(candidatePath) Then
                    File.Delete(candidatePath)
                End If
            Next
        Catch
        End Try
    End Sub
End Class
