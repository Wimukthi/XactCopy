Imports System.IO
Imports XactCopy.Infrastructure
Imports XactCopy.Models
Imports Xunit

Public Class RecoveryAndJobStoreTests
    <Fact>
    Public Sub JobCatalogStore_SaveAndLoad_RoundTripsJobsRunsAndQueue()
        Dim tempDirectory = Path.Combine(Path.GetTempPath(), "xactcopy-tests-" & Guid.NewGuid().ToString("N"))
        Directory.CreateDirectory(tempDirectory)

        Try
            Dim catalogPath = Path.Combine(tempDirectory, "catalog.json")
            Dim store As New JobCatalogStore(catalogPath)

            Dim job As New ManagedJob() With {
                .JobId = "job-alpha",
                .Name = "Alpha",
                .Options = New CopyJobOptions() With {
                    .SourceRoot = "C:\source",
                    .DestinationRoot = "D:\dest"
                }
            }

            Dim run As New ManagedJobRun() With {
                .RunId = "run-alpha",
                .JobId = "job-alpha",
                .DisplayName = "Alpha",
                .SourceRoot = "C:\source",
                .DestinationRoot = "D:\dest",
                .Status = ManagedJobRunStatus.Completed,
                .Summary = "Completed"
            }

            Dim catalog As New JobCatalog()
            catalog.Jobs.Add(job)
            catalog.Runs.Add(run)
            catalog.QueuedJobIds.Add("job-alpha")

            store.Save(catalog)
            Dim loaded = store.Load()

            Assert.Single(loaded.Jobs)
            Assert.Single(loaded.Runs)
            Assert.Single(loaded.QueuedJobIds)
            Assert.Equal("job-alpha", loaded.Jobs(0).JobId)
            Assert.Equal(ManagedJobRunStatus.Completed, loaded.Runs(0).Status)
            Assert.Equal("job-alpha", loaded.QueuedJobIds(0))
        Finally
            If Directory.Exists(tempDirectory) Then
                Directory.Delete(tempDirectory, recursive:=True)
            End If
        End Try
    End Sub

    <Fact>
    Public Sub RecoveryStateStore_SaveAndLoad_RoundTripsActiveRun()
        Dim tempDirectory = Path.Combine(Path.GetTempPath(), "xactcopy-tests-" & Guid.NewGuid().ToString("N"))
        Directory.CreateDirectory(tempDirectory)

        Try
            Dim statePath = Path.Combine(tempDirectory, "recovery-state.json")
            Dim store As New RecoveryStateStore(statePath)

            Dim state As New RecoveryState() With {
                .CleanShutdown = False,
                .PendingResumePrompt = True,
                .LastInterruptionReason = "Test interruption",
                .ActiveRun = New RecoveryActiveRun() With {
                    .RunId = "run-1",
                    .JobId = "job-1",
                    .JobName = "Recover Me",
                    .Trigger = "manual",
                    .Options = New CopyJobOptions() With {
                        .SourceRoot = "C:\src",
                        .DestinationRoot = "D:\dst",
                        .ResumeFromJournal = True
                    },
                    .JournalPath = "C:\journal.json"
                }
            }

            store.Save(state)
            Dim loaded = store.Load()

            Assert.False(loaded.CleanShutdown)
            Assert.True(loaded.PendingResumePrompt)
            Assert.NotNull(loaded.ActiveRun)
            Assert.Equal("run-1", loaded.ActiveRun.RunId)
            Assert.Equal("C:\src", loaded.ActiveRun.Options.SourceRoot)
            Assert.Equal("C:\journal.json", loaded.ActiveRun.JournalPath)
        Finally
            If Directory.Exists(tempDirectory) Then
                Directory.Delete(tempDirectory, recursive:=True)
            End If
        End Try
    End Sub
End Class
