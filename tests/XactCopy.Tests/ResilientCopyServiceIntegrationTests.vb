Imports System.IO
Imports System.Linq
Imports System.Threading
Imports System.Threading.Tasks
Imports XactCopy.Infrastructure
Imports XactCopy.Models
Imports XactCopy.Services
Imports Xunit

Public Class ResilientCopyServiceIntegrationTests
    Private Const DevFaultRulesEnvVar As String = "XACTCOPY_DEV_FAULT_RULES"
    Private Shared ReadOnly DevFaultEnvGate As New SemaphoreSlim(1, 1)

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
