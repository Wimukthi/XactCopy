Imports System.Linq
Imports System.Collections.Generic
Imports XactCopy.Infrastructure
Imports XactCopy.Models

Namespace Services
    Friend Class JobManagerService
        Private Const MaximumRunHistory As Integer = 500

        Private ReadOnly _syncRoot As New Object()
        Private ReadOnly _catalogStore As JobCatalogStore
        Private _catalog As JobCatalog

        Public Sub New(Optional catalogStore As JobCatalogStore = Nothing)
            _catalogStore = If(catalogStore, New JobCatalogStore())
            _catalog = _catalogStore.Load()
        End Sub

        Public Function GetJobs() As IReadOnlyList(Of ManagedJob)
            SyncLock _syncRoot
                Dim jobs = _catalog.Jobs.Where(Function(job) job IsNot Nothing).Select(AddressOf CloneJob).ToList()
                jobs.Sort(Function(left, right) StringComparer.OrdinalIgnoreCase.Compare(left.Name, right.Name))
                Return jobs
            End SyncLock
        End Function

        Public Function GetRecentRuns(Optional take As Integer = 200) As IReadOnlyList(Of ManagedJobRun)
            SyncLock _syncRoot
                Dim bounded = Math.Max(1, take)
                Return _catalog.Runs.
                    Where(Function(run) run IsNot Nothing).
                    Take(bounded).
                    Select(AddressOf CloneRun).
                    ToList()
            End SyncLock
        End Function

        Public Function GetQueuedJobs() As IReadOnlyList(Of ManagedJob)
            SyncLock _syncRoot
                Dim queued As New List(Of ManagedJob)()
                For Each jobId In _catalog.QueuedJobIds
                    Dim job = FindJobByIdLocked(jobId)
                    If job IsNot Nothing Then
                        queued.Add(CloneJob(job))
                    End If
                Next

                Return queued
            End SyncLock
        End Function

        Public Function GetJobById(jobId As String) As ManagedJob
            If String.IsNullOrWhiteSpace(jobId) Then
                Return Nothing
            End If

            SyncLock _syncRoot
                Dim job = FindJobByIdLocked(jobId)
                Return CloneJob(job)
            End SyncLock
        End Function

        Public Function SaveJob(name As String, options As CopyJobOptions, Optional existingJobId As String = Nothing) As ManagedJob
            If String.IsNullOrWhiteSpace(name) Then
                Throw New ArgumentException("Job name is required.", NameOf(name))
            End If

            If options Is Nothing Then
                Throw New ArgumentNullException(NameOf(options))
            End If

            SyncLock _syncRoot
                Dim nowUtc = DateTimeOffset.UtcNow
                Dim trimmedName = name.Trim()

                Dim job As ManagedJob = Nothing
                If Not String.IsNullOrWhiteSpace(existingJobId) Then
                    job = FindJobByIdLocked(existingJobId)
                End If

                If job Is Nothing Then
                    job = New ManagedJob() With {
                        .JobId = Guid.NewGuid().ToString("N"),
                        .CreatedUtc = nowUtc
                    }
                    _catalog.Jobs.Add(job)
                End If

                job.Name = trimmedName
                job.Options = CloneOptions(options)
                job.UpdatedUtc = nowUtc

                SaveCatalogLocked()
                Return CloneJob(job)
            End SyncLock
        End Function

        Public Function DeleteJob(jobId As String) As Boolean
            If String.IsNullOrWhiteSpace(jobId) Then
                Return False
            End If

            SyncLock _syncRoot
                Dim removed = _catalog.Jobs.RemoveAll(Function(job) job IsNot Nothing AndAlso String.Equals(job.JobId, jobId, StringComparison.OrdinalIgnoreCase)) > 0
                If Not removed Then
                    Return False
                End If

                _catalog.QueuedJobIds.RemoveAll(Function(queuedId) String.Equals(queuedId, jobId, StringComparison.OrdinalIgnoreCase))
                SaveCatalogLocked()
                Return True
            End SyncLock
        End Function

        Public Function QueueJob(jobId As String) As Boolean
            If String.IsNullOrWhiteSpace(jobId) Then
                Return False
            End If

            SyncLock _syncRoot
                If FindJobByIdLocked(jobId) Is Nothing Then
                    Return False
                End If

                If _catalog.QueuedJobIds.Any(Function(existing) String.Equals(existing, jobId, StringComparison.OrdinalIgnoreCase)) Then
                    Return False
                End If

                _catalog.QueuedJobIds.Add(jobId)
                SaveCatalogLocked()
                Return True
            End SyncLock
        End Function

        Public Function RemoveQueuedJob(jobId As String) As Boolean
            If String.IsNullOrWhiteSpace(jobId) Then
                Return False
            End If

            SyncLock _syncRoot
                Dim removed = _catalog.QueuedJobIds.RemoveAll(Function(existing) String.Equals(existing, jobId, StringComparison.OrdinalIgnoreCase)) > 0
                If removed Then
                    SaveCatalogLocked()
                End If

                Return removed
            End SyncLock
        End Function

        Public Function TryDequeueNextJob(ByRef job As ManagedJob) As Boolean
            SyncLock _syncRoot
                Dim mutated = False
                While _catalog.QueuedJobIds.Count > 0
                    Dim nextJobId = _catalog.QueuedJobIds(0)
                    _catalog.QueuedJobIds.RemoveAt(0)
                    mutated = True

                    Dim nextJob = FindJobByIdLocked(nextJobId)
                    If nextJob Is Nothing Then
                        Continue While
                    End If

                    SaveCatalogLocked()
                    job = CloneJob(nextJob)
                    Return True
                End While

                If mutated Then
                    SaveCatalogLocked()
                End If

                job = Nothing
                Return False
            End SyncLock
        End Function

        Public Function CreateRunForJob(jobId As String, trigger As String) As ManagedJobRun
            If String.IsNullOrWhiteSpace(jobId) Then
                Throw New ArgumentException("Job ID is required.", NameOf(jobId))
            End If

            SyncLock _syncRoot
                Dim job = FindJobByIdLocked(jobId)
                If job Is Nothing Then
                    Throw New InvalidOperationException("Selected job no longer exists.")
                End If

                Dim run = CreateRunLocked(
                    sourceJobId:=job.JobId,
                    displayName:=If(String.IsNullOrWhiteSpace(job.Name), "Saved Job", job.Name),
                    options:=job.Options,
                    trigger:=trigger)

                SaveCatalogLocked()
                Return CloneRun(run)
            End SyncLock
        End Function

        Public Function CreateAdHocRun(options As CopyJobOptions, displayName As String, trigger As String) As ManagedJobRun
            If options Is Nothing Then
                Throw New ArgumentNullException(NameOf(options))
            End If

            SyncLock _syncRoot
                Dim resolvedName = displayName
                If String.IsNullOrWhiteSpace(resolvedName) Then
                    resolvedName = "Manual Copy"
                End If

                Dim run = CreateRunLocked(
                    sourceJobId:=String.Empty,
                    displayName:=resolvedName,
                    options:=options,
                    trigger:=trigger)

                SaveCatalogLocked()
                Return CloneRun(run)
            End SyncLock
        End Function

        Public Sub MarkRunRunning(runId As String, journalPath As String)
            If String.IsNullOrWhiteSpace(runId) Then
                Return
            End If

            SyncLock _syncRoot
                Dim run = FindRunByIdLocked(runId)
                If run Is Nothing Then
                    Return
                End If

                run.Status = ManagedJobRunStatus.Running
                run.JournalPath = If(journalPath, String.Empty)
                run.LastUpdatedUtc = DateTimeOffset.UtcNow
                SaveCatalogLocked()
            End SyncLock
        End Sub

        Public Sub MarkRunPaused(runId As String)
            UpdateRunStatus(runId, ManagedJobRunStatus.Paused)
        End Sub

        Public Sub MarkRunResumed(runId As String)
            UpdateRunStatus(runId, ManagedJobRunStatus.Running)
        End Sub

        Public Sub MarkRunCompleted(runId As String, result As CopyJobResult)
            If String.IsNullOrWhiteSpace(runId) Then
                Return
            End If

            SyncLock _syncRoot
                Dim run = FindRunByIdLocked(runId)
                If run Is Nothing Then
                    Return
                End If

                Dim nowUtc = DateTimeOffset.UtcNow
                run.Result = CloneResult(result)
                run.JournalPath = If(result?.JournalPath, run.JournalPath)
                run.ErrorMessage = If(result?.ErrorMessage, String.Empty)
                run.LastUpdatedUtc = nowUtc
                run.FinishedUtc = nowUtc

                If result Is Nothing Then
                    run.Status = ManagedJobRunStatus.Failed
                    run.Summary = "Run finished with unknown result."
                ElseIf result.Cancelled Then
                    run.Status = ManagedJobRunStatus.Cancelled
                    run.Summary = "Cancelled by user."
                ElseIf result.Succeeded Then
                    run.Status = ManagedJobRunStatus.Completed
                    run.Summary = $"Completed: {result.CompletedFiles}/{result.TotalFiles} files."
                Else
                    run.Status = ManagedJobRunStatus.Failed
                    run.Summary = $"Completed with failures: failed {result.FailedFiles}, recovered {result.RecoveredFiles}."
                End If

                SaveCatalogLocked()
            End SyncLock
        End Sub

        Public Function MarkRunInterrupted(runId As String, reason As String) As Boolean
            If String.IsNullOrWhiteSpace(runId) Then
                Return False
            End If

            SyncLock _syncRoot
                Dim run = FindRunByIdLocked(runId)
                If run Is Nothing Then
                    Return False
                End If

                Dim changed = False
                Select Case run.Status
                    Case ManagedJobRunStatus.Running, ManagedJobRunStatus.Paused, ManagedJobRunStatus.Queued
                        run.Status = ManagedJobRunStatus.Interrupted
                        run.LastUpdatedUtc = DateTimeOffset.UtcNow
                        run.FinishedUtc = DateTimeOffset.UtcNow
                        run.ErrorMessage = If(String.IsNullOrWhiteSpace(reason), "Run interrupted.", reason)
                        run.Summary = run.ErrorMessage
                        changed = True
                End Select

                If changed Then
                    SaveCatalogLocked()
                End If

                Return changed
            End SyncLock
        End Function

        Public Function MarkAnyRunningRunsInterrupted(reason As String) As Integer
            SyncLock _syncRoot
                Dim count = 0
                Dim message = If(String.IsNullOrWhiteSpace(reason), "Run interrupted.", reason)
                For Each run In _catalog.Runs
                    If run Is Nothing Then
                        Continue For
                    End If

                    If run.Status = ManagedJobRunStatus.Running OrElse run.Status = ManagedJobRunStatus.Paused Then
                        run.Status = ManagedJobRunStatus.Interrupted
                        run.LastUpdatedUtc = DateTimeOffset.UtcNow
                        run.FinishedUtc = DateTimeOffset.UtcNow
                        run.ErrorMessage = message
                        run.Summary = message
                        count += 1
                    End If
                Next

                If count > 0 Then
                    SaveCatalogLocked()
                End If

                Return count
            End SyncLock
        End Function

        Private Sub UpdateRunStatus(runId As String, status As ManagedJobRunStatus)
            If String.IsNullOrWhiteSpace(runId) Then
                Return
            End If

            SyncLock _syncRoot
                Dim run = FindRunByIdLocked(runId)
                If run Is Nothing Then
                    Return
                End If

                run.Status = status
                run.LastUpdatedUtc = DateTimeOffset.UtcNow
                SaveCatalogLocked()
            End SyncLock
        End Sub

        Private Function CreateRunLocked(sourceJobId As String, displayName As String, options As CopyJobOptions, trigger As String) As ManagedJobRun
            Dim safeOptions = CloneOptions(options)
            Dim nowUtc = DateTimeOffset.UtcNow
            Dim run As New ManagedJobRun() With {
                .RunId = Guid.NewGuid().ToString("N"),
                .JobId = If(sourceJobId, String.Empty),
                .DisplayName = If(displayName, String.Empty),
                .SourceRoot = safeOptions.SourceRoot,
                .DestinationRoot = safeOptions.DestinationRoot,
                .Trigger = If(trigger, String.Empty),
                .Status = ManagedJobRunStatus.Queued,
                .StartedUtc = nowUtc,
                .LastUpdatedUtc = nowUtc,
                .JournalPath = String.Empty,
                .Summary = "Queued"
            }

            _catalog.Runs.Insert(0, run)
            TrimRunHistoryLocked()
            Return run
        End Function

        Private Sub TrimRunHistoryLocked()
            If _catalog.Runs.Count <= MaximumRunHistory Then
                Return
            End If

            _catalog.Runs = _catalog.Runs.Take(MaximumRunHistory).ToList()
        End Sub

        Private Function FindJobByIdLocked(jobId As String) As ManagedJob
            Return _catalog.Jobs.FirstOrDefault(
                Function(job) job IsNot Nothing AndAlso String.Equals(job.JobId, jobId, StringComparison.OrdinalIgnoreCase))
        End Function

        Private Function FindRunByIdLocked(runId As String) As ManagedJobRun
            Return _catalog.Runs.FirstOrDefault(
                Function(run) run IsNot Nothing AndAlso String.Equals(run.RunId, runId, StringComparison.OrdinalIgnoreCase))
        End Function

        Private Sub SaveCatalogLocked()
            _catalogStore.Save(_catalog)
        End Sub

        Private Shared Function CloneJob(job As ManagedJob) As ManagedJob
            If job Is Nothing Then
                Return Nothing
            End If

            Return New ManagedJob() With {
                .JobId = job.JobId,
                .Name = job.Name,
                .Options = CloneOptions(job.Options),
                .CreatedUtc = job.CreatedUtc,
                .UpdatedUtc = job.UpdatedUtc
            }
        End Function

        Private Shared Function CloneRun(run As ManagedJobRun) As ManagedJobRun
            If run Is Nothing Then
                Return Nothing
            End If

            Return New ManagedJobRun() With {
                .RunId = run.RunId,
                .JobId = run.JobId,
                .DisplayName = run.DisplayName,
                .SourceRoot = run.SourceRoot,
                .DestinationRoot = run.DestinationRoot,
                .Trigger = run.Trigger,
                .Status = run.Status,
                .StartedUtc = run.StartedUtc,
                .LastUpdatedUtc = run.LastUpdatedUtc,
                .FinishedUtc = run.FinishedUtc,
                .JournalPath = run.JournalPath,
                .Summary = run.Summary,
                .ErrorMessage = run.ErrorMessage,
                .Result = CloneResult(run.Result)
            }
        End Function

        Private Shared Function CloneResult(result As CopyJobResult) As CopyJobResult
            If result Is Nothing Then
                Return Nothing
            End If

            Return New CopyJobResult() With {
                .Succeeded = result.Succeeded,
                .Cancelled = result.Cancelled,
                .TotalFiles = result.TotalFiles,
                .CompletedFiles = result.CompletedFiles,
                .FailedFiles = result.FailedFiles,
                .RecoveredFiles = result.RecoveredFiles,
                .SkippedFiles = result.SkippedFiles,
                .TotalBytes = result.TotalBytes,
                .CopiedBytes = result.CopiedBytes,
                .JournalPath = result.JournalPath,
                .ErrorMessage = result.ErrorMessage
            }
        End Function

        Private Shared Function CloneOptions(options As CopyJobOptions) As CopyJobOptions
            If options Is Nothing Then
                Return New CopyJobOptions()
            End If

            Return New CopyJobOptions() With {
                .SourceRoot = options.SourceRoot,
                .DestinationRoot = options.DestinationRoot,
                .SelectedRelativePaths = New List(Of String)(If(options.SelectedRelativePaths, New List(Of String)())),
                .OverwritePolicy = options.OverwritePolicy,
                .SymlinkHandling = options.SymlinkHandling,
                .CopyEmptyDirectories = options.CopyEmptyDirectories,
                .BufferSizeBytes = options.BufferSizeBytes,
                .UseAdaptiveBufferSizing = options.UseAdaptiveBufferSizing,
                .MaxThroughputBytesPerSecond = options.MaxThroughputBytesPerSecond,
                .ParallelSmallFileWorkers = options.ParallelSmallFileWorkers,
                .SmallFileThresholdBytes = options.SmallFileThresholdBytes,
                .WaitForMediaAvailability = options.WaitForMediaAvailability,
                .MaxRetries = options.MaxRetries,
                .OperationTimeout = options.OperationTimeout,
                .PerFileTimeout = options.PerFileTimeout,
                .InitialRetryDelay = options.InitialRetryDelay,
                .MaxRetryDelay = options.MaxRetryDelay,
                .ResumeFromJournal = options.ResumeFromJournal,
                .VerifyAfterCopy = options.VerifyAfterCopy,
                .VerificationMode = options.VerificationMode,
                .VerificationHashAlgorithm = options.VerificationHashAlgorithm,
                .SampleVerificationChunkBytes = options.SampleVerificationChunkBytes,
                .SampleVerificationChunkCount = options.SampleVerificationChunkCount,
                .SalvageUnreadableBlocks = options.SalvageUnreadableBlocks,
                .SalvageFillPattern = options.SalvageFillPattern,
                .ContinueOnFileError = options.ContinueOnFileError,
                .PreserveTimestamps = options.PreserveTimestamps,
                .WorkerProcessPriorityClass = options.WorkerProcessPriorityClass
            }
        End Function
    End Class
End Namespace
