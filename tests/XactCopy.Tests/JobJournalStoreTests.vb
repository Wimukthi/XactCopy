Imports System.Collections.Generic
Imports System.IO
Imports System.Text.Json
Imports System.Threading
Imports System.Threading.Tasks
Imports XactCopy.Infrastructure
Imports XactCopy.Models
Imports Xunit

Public Class JobJournalStoreTests
    <Fact>
    Public Async Function SaveAndLoadAsync_RoundTripsJournalData() As Task
        Dim store As New JobJournalStore()
        Dim tempDirectory = Path.Combine(Path.GetTempPath(), $"xactcopy-tests-{Guid.NewGuid():N}")
        Directory.CreateDirectory(tempDirectory)

        Try
            Dim journalPath = Path.Combine(tempDirectory, "journal.json")
            Dim journal As New JobJournal() With {
                .JobId = "job-1",
                .SourceRoot = "C:\source",
                .DestinationRoot = "D:\dest"
            }

            journal.Files("a.txt") = New JournalFileEntry() With {
                .RelativePath = "a.txt",
                .SourceLength = 100,
                .BytesCopied = 64,
                .State = FileCopyState.InProgress,
                .LastRescuePass = "TrimSweep",
                .RescueRanges = New List(Of RescueRange) From {
                    New RescueRange() With {.Offset = 0, .Length = 64, .State = RescueRangeState.Good},
                    New RescueRange() With {.Offset = 64, .Length = 36, .State = RescueRangeState.Bad}
                }
            }

            Await store.SaveAsync(journalPath, journal, CancellationToken.None)
            Dim loaded = Await store.LoadAsync(journalPath, CancellationToken.None)

            Assert.NotNull(loaded)
            Assert.Equal("job-1", loaded.JobId)
            Assert.True(loaded.Files.ContainsKey("a.txt"))
            Assert.Equal(64, loaded.Files("a.txt").BytesCopied)
            Assert.Equal(FileCopyState.InProgress, loaded.Files("a.txt").State)
            Assert.Equal("TrimSweep", loaded.Files("a.txt").LastRescuePass)
            Assert.Equal(2, loaded.Files("a.txt").RescueRanges.Count)
            Assert.Equal(RescueRangeState.Bad, loaded.Files("a.txt").RescueRanges(1).State)
        Finally
            If Directory.Exists(tempDirectory) Then
                Directory.Delete(tempDirectory, recursive:=True)
            End If
        End Try
    End Function

    <Fact>
    Public Async Function LoadAsync_LoadsLegacyJsonSnapshotWithoutLedger() As Task
        Dim store As New JobJournalStore()
        Dim tempDirectory = Path.Combine(Path.GetTempPath(), $"xactcopy-tests-{Guid.NewGuid():N}")
        Directory.CreateDirectory(tempDirectory)

        Try
            Dim journalPath = Path.Combine(tempDirectory, "journal.json")
            Dim journal As New JobJournal() With {
                .JobId = "legacy-job",
                .SourceRoot = "C:\legacy-source",
                .DestinationRoot = "D:\legacy-destination"
            }
            journal.Files("legacy.bin") = New JournalFileEntry() With {
                .RelativePath = "legacy.bin",
                .SourceLength = 200,
                .BytesCopied = 10,
                .State = FileCopyState.InProgress
            }

            Dim legacyPayload = JsonSerializer.Serialize(journal)
            File.WriteAllText(journalPath, legacyPayload)

            Dim loaded = Await store.LoadAsync(journalPath, CancellationToken.None)

            Assert.NotNull(loaded)
            Assert.Equal("legacy-job", loaded.JobId)
            Assert.True(loaded.Files.ContainsKey("legacy.bin"))
            Assert.Equal(10, loaded.Files("legacy.bin").BytesCopied)
        Finally
            If Directory.Exists(tempDirectory) Then
                Directory.Delete(tempDirectory, recursive:=True)
            End If
        End Try
    End Function

    <Fact>
    Public Async Function LoadAsync_UsesBackupWhenPrimarySnapshotIsTampered() As Task
        Dim store As New JobJournalStore()
        Dim tempDirectory = Path.Combine(Path.GetTempPath(), $"xactcopy-tests-{Guid.NewGuid():N}")
        Directory.CreateDirectory(tempDirectory)

        Try
            Dim journalPath = Path.Combine(tempDirectory, "journal.json")
            Dim journal As New JobJournal() With {
                .JobId = "tamper-job",
                .SourceRoot = "C:\source",
                .DestinationRoot = "D:\dest"
            }
            journal.Files("file.dat") = New JournalFileEntry() With {
                .RelativePath = "file.dat",
                .SourceLength = 1024,
                .BytesCopied = 64,
                .State = FileCopyState.InProgress
            }

            Await store.SaveAsync(journalPath, journal, CancellationToken.None)

            journal.Files("file.dat").BytesCopied = 128
            Await store.SaveAsync(journalPath, journal, CancellationToken.None)

            ' Tamper the primary snapshot so it no longer matches any signed ledger hash.
            journal.Files("file.dat").BytesCopied = 900
            File.WriteAllText(journalPath, JsonSerializer.Serialize(journal))

            Dim loaded = Await store.LoadAsync(journalPath, CancellationToken.None)

            Assert.NotNull(loaded)
            Assert.True(loaded.Files.ContainsKey("file.dat"))
            Assert.Equal(128, loaded.Files("file.dat").BytesCopied)
        Finally
            If Directory.Exists(tempDirectory) Then
                Directory.Delete(tempDirectory, recursive:=True)
            End If
        End Try
    End Function

    <Fact>
    Public Async Function LoadAsync_UsesMirrorWhenPrimaryFilesAreMissing() As Task
        Dim store As New JobJournalStore()
        Dim tempDirectory = Path.Combine(Path.GetTempPath(), $"xactcopy-tests-{Guid.NewGuid():N}")
        Directory.CreateDirectory(tempDirectory)

        Try
            Dim journalPath = Path.Combine(tempDirectory, "journal.json")
            Dim journal As New JobJournal() With {
                .JobId = "mirror-job",
                .SourceRoot = "C:\source",
                .DestinationRoot = "D:\dest"
            }
            journal.Files("mirror.bin") = New JournalFileEntry() With {
                .RelativePath = "mirror.bin",
                .SourceLength = 400,
                .BytesCopied = 333,
                .State = FileCopyState.InProgress
            }

            Await store.SaveAsync(journalPath, journal, CancellationToken.None)

            DeleteIfExists(journalPath)
            DeleteIfExists($"{journalPath}.bak1")
            DeleteIfExists($"{journalPath}.bak2")
            DeleteIfExists($"{journalPath}.bak3")
            DeleteIfExists($"{journalPath}.ledger")
            DeleteIfExists($"{journalPath}.anchor")

            Dim loaded = Await store.LoadAsync(journalPath, CancellationToken.None)

            Assert.NotNull(loaded)
            Assert.Equal("mirror-job", loaded.JobId)
            Assert.True(loaded.Files.ContainsKey("mirror.bin"))
            Assert.Equal(333, loaded.Files("mirror.bin").BytesCopied)
        Finally
            If Directory.Exists(tempDirectory) Then
                Directory.Delete(tempDirectory, recursive:=True)
            End If
        End Try
    End Function

    <Fact>
    Public Async Function LoadAsync_UsesMirrorLedgerWhenPrimaryLedgerIsCorrupted() As Task
        Dim store As New JobJournalStore()
        Dim tempDirectory = Path.Combine(Path.GetTempPath(), $"xactcopy-tests-{Guid.NewGuid():N}")
        Directory.CreateDirectory(tempDirectory)

        Try
            Dim journalPath = Path.Combine(tempDirectory, "journal.json")
            Dim journal As New JobJournal() With {
                .JobId = "mirror-ledger-job",
                .SourceRoot = "C:\source",
                .DestinationRoot = "D:\dest"
            }
            journal.Files("hash.bin") = New JournalFileEntry() With {
                .RelativePath = "hash.bin",
                .SourceLength = 800,
                .BytesCopied = 456,
                .State = FileCopyState.InProgress
            }

            Await store.SaveAsync(journalPath, journal, CancellationToken.None)

            ' Corrupt primary ledger and snapshot.
            File.WriteAllBytes($"{journalPath}.ledger", {CByte(1), CByte(2), CByte(3), CByte(4)})
            journal.Files("hash.bin").BytesCopied = 999
            File.WriteAllText(journalPath, JsonSerializer.Serialize(journal))

            Dim loaded = Await store.LoadAsync(journalPath, CancellationToken.None)

            Assert.NotNull(loaded)
            Assert.Equal("mirror-ledger-job", loaded.JobId)
            Assert.Equal(456, loaded.Files("hash.bin").BytesCopied)
        Finally
            If Directory.Exists(tempDirectory) Then
                Directory.Delete(tempDirectory, recursive:=True)
            End If
        End Try
    End Function

    <Fact>
    Public Sub BuildJobId_IsDeterministicAndCaseInsensitive()
        Dim first = JobJournalStore.BuildJobId("c:\src", "d:\dst")
        Dim second = JobJournalStore.BuildJobId("C:\SRC", "D:\DST")

        Assert.Equal(first, second)
        Assert.Equal(20, first.Length)
    End Sub

    Private Shared Sub DeleteIfExists(path As String)
        If File.Exists(path) Then
            File.Delete(path)
        End If
    End Sub
End Class
