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
                    .DestinationRoot = "D:\dest",
                    .ExpectedSourceIdentity = "vol:C::11111111",
                    .ExpectedDestinationIdentity = "vol:D::22222222",
                    .ResumeJournalPathHint = "C:\journals\job-alpha.json",
                    .AllowJournalRootRemap = True,
                    .WaitForFileLockRelease = True,
                    .TreatAccessDeniedAsContention = True,
                    .LockContentionProbeInterval = TimeSpan.FromMilliseconds(700),
                    .SourceMutationPolicy = SourceMutationPolicy.WaitForReappearance
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
            Assert.Single(loaded.QueueEntries)
            Assert.Equal("job-alpha", loaded.Jobs(0).JobId)
            Assert.Equal("vol:C::11111111", loaded.Jobs(0).Options.ExpectedSourceIdentity)
            Assert.Equal("vol:D::22222222", loaded.Jobs(0).Options.ExpectedDestinationIdentity)
            Assert.Equal("C:\journals\job-alpha.json", loaded.Jobs(0).Options.ResumeJournalPathHint)
            Assert.True(loaded.Jobs(0).Options.AllowJournalRootRemap)
            Assert.True(loaded.Jobs(0).Options.WaitForFileLockRelease)
            Assert.True(loaded.Jobs(0).Options.TreatAccessDeniedAsContention)
            Assert.Equal(TimeSpan.FromMilliseconds(700), loaded.Jobs(0).Options.LockContentionProbeInterval)
            Assert.Equal(SourceMutationPolicy.WaitForReappearance, loaded.Jobs(0).Options.SourceMutationPolicy)
            Assert.Equal(ManagedJobRunStatus.Completed, loaded.Runs(0).Status)
            Assert.Equal("job-alpha", loaded.QueueEntries(0).JobId)
            Assert.Equal(2, loaded.SchemaVersion)
        Finally
            If Directory.Exists(tempDirectory) Then
                Directory.Delete(tempDirectory, recursive:=True)
            End If
        End Try
    End Sub

    <Fact>
    Public Sub JobCatalogStore_Load_MigratesLegacyQueuedJobIdsToQueueEntries()
        Dim tempDirectory = Path.Combine(Path.GetTempPath(), "xactcopy-tests-" & Guid.NewGuid().ToString("N"))
        Directory.CreateDirectory(tempDirectory)

        Try
            Dim catalogPath = Path.Combine(tempDirectory, "catalog.json")
            Dim legacyJson =
"{
  ""Jobs"": [
    {
      ""JobId"": ""job-legacy"",
      ""Name"": ""Legacy Job"",
      ""Options"": {
        ""SourceRoot"": ""C:\\legacy-src"",
        ""DestinationRoot"": ""D:\\legacy-dst""
      }
    }
  ],
  ""Runs"": [],
  ""QueuedJobIds"": [ ""job-legacy"" ]
}"

            File.WriteAllText(catalogPath, legacyJson)
            Dim store As New JobCatalogStore(catalogPath)
            Dim loaded = store.Load()

            Assert.Single(loaded.QueueEntries)
            Assert.Equal("job-legacy", loaded.QueueEntries(0).JobId)
            Assert.Equal("legacy-migration", loaded.QueueEntries(0).Trigger)
            Assert.Equal(2, loaded.SchemaVersion)
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
                        .ExpectedSourceIdentity = "unc:\\SOURCE\share",
                        .ExpectedDestinationIdentity = "vol:D::DEADBEEF",
                        .ResumeJournalPathHint = "C:\journals\run-1.json",
                        .AllowJournalRootRemap = True,
                        .WaitForFileLockRelease = True,
                        .TreatAccessDeniedAsContention = True,
                        .LockContentionProbeInterval = TimeSpan.FromMilliseconds(900),
                        .SourceMutationPolicy = SourceMutationPolicy.SkipFile,
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
            Assert.Equal("unc:\\SOURCE\share", loaded.ActiveRun.Options.ExpectedSourceIdentity)
            Assert.Equal("vol:D::DEADBEEF", loaded.ActiveRun.Options.ExpectedDestinationIdentity)
            Assert.Equal("C:\journals\run-1.json", loaded.ActiveRun.Options.ResumeJournalPathHint)
            Assert.True(loaded.ActiveRun.Options.AllowJournalRootRemap)
            Assert.True(loaded.ActiveRun.Options.WaitForFileLockRelease)
            Assert.True(loaded.ActiveRun.Options.TreatAccessDeniedAsContention)
            Assert.Equal(TimeSpan.FromMilliseconds(900), loaded.ActiveRun.Options.LockContentionProbeInterval)
            Assert.Equal(SourceMutationPolicy.SkipFile, loaded.ActiveRun.Options.SourceMutationPolicy)
            Assert.Equal("C:\journal.json", loaded.ActiveRun.JournalPath)
        Finally
            If Directory.Exists(tempDirectory) Then
                Directory.Delete(tempDirectory, recursive:=True)
            End If
        End Try
    End Sub
End Class
