Imports System.Collections.Generic
Imports System.Linq
Imports XactCopy.Infrastructure
Imports XactCopy.Models

Namespace Services
    Friend Enum QueueMoveDirection
        Up = 0
        Down = 1
        Top = 2
        Bottom = 3
    End Enum

    Friend NotInheritable Class JobQueueEntryView
        Public Property QueueEntryId As String = String.Empty
        Public Property JobId As String = String.Empty
        Public Property Position As Integer
        Public Property JobName As String = String.Empty
        Public Property SourceRoot As String = String.Empty
        Public Property DestinationRoot As String = String.Empty
        Public Property EnqueuedUtc As DateTimeOffset = DateTimeOffset.MinValue
        Public Property LastUpdatedUtc As DateTimeOffset = DateTimeOffset.MinValue
        Public Property Trigger As String = String.Empty
        Public Property EnqueuedBy As String = String.Empty
        Public Property AttemptCount As Integer
        Public Property LastAttemptUtc As DateTimeOffset?
        Public Property LastErrorMessage As String = String.Empty
    End Class

    Friend NotInheritable Class QueuedJobWorkItem
        Public Property QueueEntryId As String = String.Empty
        Public Property Trigger As String = "queued"
        Public Property EnqueuedBy As String = String.Empty
        Public Property EnqueuedUtc As DateTimeOffset = DateTimeOffset.MinValue
        Public Property Attempt As Integer
        Public Property Job As ManagedJob
    End Class

    Friend Class JobManagerService
        Private Const MaximumRunHistory As Integer = 1000

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
                Dim sanitizedQueue = BuildSanitizedQueueLocked()

                For Each queueEntry In sanitizedQueue
                    Dim job = FindJobByIdLocked(queueEntry.JobId)
                    If job IsNot Nothing Then
                        queued.Add(CloneJob(job))
                    End If
                Next

                Return queued
            End SyncLock
        End Function

        Public Function GetQueueEntries() As IReadOnlyList(Of JobQueueEntryView)
            SyncLock _syncRoot
                Dim queueEntries = BuildSanitizedQueueLocked()
                Dim results As New List(Of JobQueueEntryView)()
                Dim position = 1

                For Each queueEntry In queueEntries
                    Dim job = FindJobByIdLocked(queueEntry.JobId)
                    If job Is Nothing Then
                        Continue For
                    End If

                    results.Add(New JobQueueEntryView() With {
                        .QueueEntryId = queueEntry.QueueEntryId,
                        .JobId = job.JobId,
                        .Position = position,
                        .JobName = job.Name,
                        .SourceRoot = job.Options.SourceRoot,
                        .DestinationRoot = job.Options.DestinationRoot,
                        .EnqueuedUtc = queueEntry.EnqueuedUtc,
                        .LastUpdatedUtc = queueEntry.LastUpdatedUtc,
                        .Trigger = queueEntry.Trigger,
                        .EnqueuedBy = queueEntry.EnqueuedBy,
                        .AttemptCount = queueEntry.AttemptCount,
                        .LastAttemptUtc = queueEntry.LastAttemptUtc,
                        .LastErrorMessage = queueEntry.LastErrorMessage
                    })
                    position += 1
                Next

                Return results
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
                Dim templateOptions = CloneOptions(options)
                templateOptions.ExpectedSourceIdentity = String.Empty
                templateOptions.ExpectedDestinationIdentity = String.Empty
                templateOptions.ResumeJournalPathHint = String.Empty
                templateOptions.AllowJournalRootRemap = False
                job.Options = templateOptions
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
                Dim removed = _catalog.Jobs.RemoveAll(
                    Function(job) job IsNot Nothing AndAlso String.Equals(job.JobId, jobId, StringComparison.OrdinalIgnoreCase)) > 0
                If Not removed Then
                    Return False
                End If

                _catalog.QueueEntries.RemoveAll(
                    Function(queueEntry) queueEntry IsNot Nothing AndAlso
                        String.Equals(queueEntry.JobId, jobId, StringComparison.OrdinalIgnoreCase))

                _catalog.QueuedJobIds.RemoveAll(
                    Function(queuedId) String.Equals(queuedId, jobId, StringComparison.OrdinalIgnoreCase))

                SaveCatalogLocked()
                Return True
            End SyncLock
        End Function

        Public Function QueueJob(jobId As String, Optional allowDuplicate As Boolean = False, Optional trigger As String = "manual", Optional enqueuedBy As String = "") As Boolean
            If String.IsNullOrWhiteSpace(jobId) Then
                Return False
            End If

            SyncLock _syncRoot
                Dim job = FindJobByIdLocked(jobId)
                If job Is Nothing Then
                    Return False
                End If

                BuildSanitizedQueueLocked()

                If Not allowDuplicate Then
                    Dim alreadyQueued = _catalog.QueueEntries.Any(
                        Function(existing) existing IsNot Nothing AndAlso
                            String.Equals(existing.JobId, jobId, StringComparison.OrdinalIgnoreCase))
                    If alreadyQueued Then
                        Return False
                    End If
                End If

                Dim nowUtc = DateTimeOffset.UtcNow
                Dim entry As New ManagedJobQueueEntry() With {
                    .QueueEntryId = Guid.NewGuid().ToString("N"),
                    .JobId = job.JobId,
                    .EnqueuedUtc = nowUtc,
                    .LastUpdatedUtc = nowUtc,
                    .Trigger = If(String.IsNullOrWhiteSpace(trigger), "manual", trigger.Trim()),
                    .EnqueuedBy = If(enqueuedBy, String.Empty),
                    .AttemptCount = 0,
                    .LastAttemptUtc = Nothing,
                    .LastErrorMessage = String.Empty
                }

                _catalog.QueueEntries.Add(entry)
                SaveCatalogLocked()
                Return True
            End SyncLock
        End Function

        Public Function RemoveQueuedJob(queueEntryIdOrJobId As String) As Boolean
            If String.IsNullOrWhiteSpace(queueEntryIdOrJobId) Then
                Return False
            End If

            SyncLock _syncRoot
                Dim key = queueEntryIdOrJobId.Trim()

                Dim removed = _catalog.QueueEntries.RemoveAll(
                    Function(existing) existing IsNot Nothing AndAlso
                        String.Equals(existing.QueueEntryId, key, StringComparison.OrdinalIgnoreCase)) > 0

                If Not removed Then
                    removed = _catalog.QueueEntries.RemoveAll(
                        Function(existing) existing IsNot Nothing AndAlso
                            String.Equals(existing.JobId, key, StringComparison.OrdinalIgnoreCase)) > 0
                End If

                If removed Then
                    SaveCatalogLocked()
                End If

                Return removed
            End SyncLock
        End Function

        Public Function MoveQueueEntry(queueEntryId As String, direction As QueueMoveDirection) As Boolean
            If String.IsNullOrWhiteSpace(queueEntryId) Then
                Return False
            End If

            SyncLock _syncRoot
                BuildSanitizedQueueLocked()
                If _catalog.QueueEntries.Count <= 1 Then
                    Return False
                End If

                Dim currentIndex = _catalog.QueueEntries.FindIndex(
                    Function(queueEntryCandidate) queueEntryCandidate IsNot Nothing AndAlso
                        String.Equals(queueEntryCandidate.QueueEntryId, queueEntryId, StringComparison.OrdinalIgnoreCase))
                If currentIndex < 0 Then
                    Return False
                End If

                Dim targetIndex = currentIndex
                Select Case direction
                    Case QueueMoveDirection.Up
                        targetIndex = Math.Max(0, currentIndex - 1)
                    Case QueueMoveDirection.Down
                        targetIndex = Math.Min(_catalog.QueueEntries.Count - 1, currentIndex + 1)
                    Case QueueMoveDirection.Top
                        targetIndex = 0
                    Case QueueMoveDirection.Bottom
                        targetIndex = _catalog.QueueEntries.Count - 1
                End Select

                If targetIndex = currentIndex Then
                    Return False
                End If

                Dim entry = _catalog.QueueEntries(currentIndex)
                _catalog.QueueEntries.RemoveAt(currentIndex)
                _catalog.QueueEntries.Insert(targetIndex, entry)
                entry.LastUpdatedUtc = DateTimeOffset.UtcNow
                SaveCatalogLocked()
                Return True
            End SyncLock
        End Function

        Public Function ClearQueue() As Integer
            SyncLock _syncRoot
                Dim removedCount = _catalog.QueueEntries.Count
                If removedCount = 0 Then
                    Return 0
                End If

                _catalog.QueueEntries.Clear()
                SaveCatalogLocked()
                Return removedCount
            End SyncLock
        End Function

        Public Function TryDequeueNextJob(ByRef job As ManagedJob) As Boolean
            Dim workItem As QueuedJobWorkItem = Nothing
            Dim dequeued = TryDequeueNextJob(workItem)
            job = If(workItem IsNot Nothing, CloneJob(workItem.Job), Nothing)
            Return dequeued
        End Function

        Public Function TryDequeueNextJob(ByRef workItem As QueuedJobWorkItem) As Boolean
            Return TryDequeueInternal(queueEntryId:=Nothing, workItem:=workItem)
        End Function

        Public Function TryDequeueQueuedEntry(queueEntryId As String, ByRef workItem As QueuedJobWorkItem) As Boolean
            If String.IsNullOrWhiteSpace(queueEntryId) Then
                workItem = Nothing
                Return False
            End If

            Return TryDequeueInternal(queueEntryId.Trim(), workItem)
        End Function

        Public Function CreateRunForJob(jobId As String, trigger As String, Optional queueEntryId As String = Nothing, Optional queueAttempt As Integer = 0) As ManagedJobRun
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
                    trigger:=trigger,
                    queueEntryId:=queueEntryId,
                    queueAttempt:=queueAttempt)

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
                    trigger:=trigger,
                    queueEntryId:=String.Empty,
                    queueAttempt:=0)

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

        Public Function DeleteRun(runId As String) As Boolean
            If String.IsNullOrWhiteSpace(runId) Then
                Return False
            End If

            SyncLock _syncRoot
                Dim removed = _catalog.Runs.RemoveAll(
                    Function(run) run IsNot Nothing AndAlso String.Equals(run.RunId, runId, StringComparison.OrdinalIgnoreCase)) > 0
                If removed Then
                    SaveCatalogLocked()
                End If

                Return removed
            End SyncLock
        End Function

        Public Function ClearRunHistory(Optional keepLatest As Integer = 0) As Integer
            SyncLock _syncRoot
                Dim safeKeep = Math.Max(0, keepLatest)
                If _catalog.Runs.Count <= safeKeep Then
                    Return 0
                End If

                Dim beforeCount = _catalog.Runs.Count
                _catalog.Runs = _catalog.Runs.Take(safeKeep).ToList()
                Dim removed = beforeCount - _catalog.Runs.Count
                SaveCatalogLocked()
                Return removed
            End SyncLock
        End Function

        Private Function TryDequeueInternal(queueEntryId As String, ByRef workItem As QueuedJobWorkItem) As Boolean
            SyncLock _syncRoot
                Dim nowUtc = DateTimeOffset.UtcNow
                Dim mutated = False
                Dim entryIndex = 0
                Dim specificEntryRequested = Not String.IsNullOrWhiteSpace(queueEntryId)

                While entryIndex < _catalog.QueueEntries.Count
                    Dim entry = _catalog.QueueEntries(entryIndex)
                    If entry Is Nothing OrElse String.IsNullOrWhiteSpace(entry.JobId) Then
                        _catalog.QueueEntries.RemoveAt(entryIndex)
                        mutated = True
                        Continue While
                    End If

                    If specificEntryRequested AndAlso
                        Not String.Equals(entry.QueueEntryId, queueEntryId, StringComparison.OrdinalIgnoreCase) Then
                        entryIndex += 1
                        Continue While
                    End If

                    Dim job = FindJobByIdLocked(entry.JobId)
                    If job Is Nothing Then
                        _catalog.QueueEntries.RemoveAt(entryIndex)
                        mutated = True
                        Continue While
                    End If

                    _catalog.QueueEntries.RemoveAt(entryIndex)
                    mutated = True

                    entry.AttemptCount = Math.Max(0, entry.AttemptCount) + 1
                    entry.LastAttemptUtc = nowUtc
                    entry.LastUpdatedUtc = nowUtc

                    SaveCatalogLocked()

                    workItem = New QueuedJobWorkItem() With {
                        .QueueEntryId = entry.QueueEntryId,
                        .Trigger = If(String.IsNullOrWhiteSpace(entry.Trigger), "queued", entry.Trigger),
                        .EnqueuedBy = If(entry.EnqueuedBy, String.Empty),
                        .EnqueuedUtc = entry.EnqueuedUtc,
                        .Attempt = entry.AttemptCount,
                        .Job = CloneJob(job)
                    }
                    Return True
                End While

                If mutated Then
                    SaveCatalogLocked()
                End If

                workItem = Nothing
                Return False
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

        Private Function CreateRunLocked(sourceJobId As String, displayName As String, options As CopyJobOptions, trigger As String, queueEntryId As String, queueAttempt As Integer) As ManagedJobRun
            Dim safeOptions = CloneOptions(options)
            Dim nowUtc = DateTimeOffset.UtcNow
            Dim run As New ManagedJobRun() With {
                .RunId = Guid.NewGuid().ToString("N"),
                .JobId = If(sourceJobId, String.Empty),
                .DisplayName = If(displayName, String.Empty),
                .SourceRoot = safeOptions.SourceRoot,
                .DestinationRoot = safeOptions.DestinationRoot,
                .Trigger = If(trigger, String.Empty),
                .QueueEntryId = If(queueEntryId, String.Empty),
                .QueueAttempt = Math.Max(0, queueAttempt),
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

        Private Function BuildSanitizedQueueLocked() As List(Of ManagedJobQueueEntry)
            Dim sanitized As New List(Of ManagedJobQueueEntry)()
            Dim seenEntryIds As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

            For Each queueEntry In _catalog.QueueEntries
                If queueEntry Is Nothing Then
                    Continue For
                End If

                Dim entryId = If(queueEntry.QueueEntryId, String.Empty).Trim()
                If entryId.Length = 0 Then
                    queueEntry.QueueEntryId = Guid.NewGuid().ToString("N")
                    entryId = queueEntry.QueueEntryId
                End If

                If Not seenEntryIds.Add(entryId) Then
                    Continue For
                End If

                If String.IsNullOrWhiteSpace(queueEntry.JobId) Then
                    Continue For
                End If

                If FindJobByIdLocked(queueEntry.JobId) Is Nothing Then
                    Continue For
                End If

                sanitized.Add(queueEntry)
            Next

            If _catalog.QueueEntries.Count <> sanitized.Count Then
                _catalog.QueueEntries = sanitized
                SaveCatalogLocked()
            End If

            Return sanitized
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
                .QueueEntryId = run.QueueEntryId,
                .QueueAttempt = run.QueueAttempt,
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
                .ExpectedSourceIdentity = options.ExpectedSourceIdentity,
                .ExpectedDestinationIdentity = options.ExpectedDestinationIdentity,
                .ResumeJournalPathHint = options.ResumeJournalPathHint,
                .AllowJournalRootRemap = options.AllowJournalRootRemap,
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
                .WaitForFileLockRelease = options.WaitForFileLockRelease,
                .TreatAccessDeniedAsContention = options.TreatAccessDeniedAsContention,
                .LockContentionProbeInterval = options.LockContentionProbeInterval,
                .SourceMutationPolicy = options.SourceMutationPolicy,
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
                .WorkerProcessPriorityClass = options.WorkerProcessPriorityClass,
                .WorkerTelemetryProfile = options.WorkerTelemetryProfile,
                .WorkerProgressEmitIntervalMs = options.WorkerProgressEmitIntervalMs,
                .WorkerMaxLogsPerSecond = options.WorkerMaxLogsPerSecond,
                .RescueFastScanChunkBytes = options.RescueFastScanChunkBytes,
                .RescueTrimChunkBytes = options.RescueTrimChunkBytes,
                .RescueScrapeChunkBytes = options.RescueScrapeChunkBytes,
                .RescueRetryChunkBytes = options.RescueRetryChunkBytes,
                .RescueSplitMinimumBytes = options.RescueSplitMinimumBytes,
                .RescueFastScanRetries = options.RescueFastScanRetries,
                .RescueTrimRetries = options.RescueTrimRetries,
                .RescueScrapeRetries = options.RescueScrapeRetries
            }
        End Function
    End Class
End Namespace
