Imports System.IO
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
                .State = FileCopyState.InProgress
            }

            Await store.SaveAsync(journalPath, journal, CancellationToken.None)
            Dim loaded = Await store.LoadAsync(journalPath, CancellationToken.None)

            Assert.NotNull(loaded)
            Assert.Equal("job-1", loaded.JobId)
            Assert.True(loaded.Files.ContainsKey("a.txt"))
            Assert.Equal(64, loaded.Files("a.txt").BytesCopied)
            Assert.Equal(FileCopyState.InProgress, loaded.Files("a.txt").State)
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
End Class
