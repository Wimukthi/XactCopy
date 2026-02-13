Imports System.IO
Imports System.Threading
Imports System.Threading.Tasks
Imports XactCopy.Infrastructure
Imports XactCopy.Models
Imports XactCopy.Services
Imports Xunit

Public Class ResilientCopyServiceIntegrationTests
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
            Try
                Dim jobId = JobJournalStore.BuildJobId(sourceRoot, destinationRoot)
                Dim journalPath = JobJournalStore.GetDefaultJournalPath(jobId)
                If File.Exists(journalPath) Then
                    File.Delete(journalPath)
                End If
            Catch
            End Try

            If Directory.Exists(tempRoot) Then
                Directory.Delete(tempRoot, recursive:=True)
            End If
        End Try
    End Function
End Class
