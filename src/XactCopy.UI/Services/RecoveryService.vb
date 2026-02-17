Imports System.Collections.Generic
Imports XactCopy.Configuration
Imports XactCopy.Infrastructure
Imports XactCopy.Models

Namespace Services
    Friend NotInheritable Class RecoveryStartupInfo
        Public Property InterruptedRun As RecoveryActiveRun
        Public Property InterruptionReason As String = String.Empty
        Public Property ShouldPrompt As Boolean
        Public Property ShouldAutoResume As Boolean
        Public Property WasRecoveryAutostart As Boolean

        Public ReadOnly Property HasInterruptedRun As Boolean
            Get
                Return InterruptedRun IsNot Nothing
            End Get
        End Property
    End Class

    Friend Class RecoveryService
        Private ReadOnly _syncRoot As New Object()
        Private ReadOnly _stateStore As RecoveryStateStore
        Private ReadOnly _startupService As WindowsStartupService

        Private _state As RecoveryState
        Private _lastTouchUtc As DateTimeOffset = DateTimeOffset.MinValue
        Private _touchInterval As TimeSpan = TimeSpan.FromSeconds(2)

        Public Sub New(Optional stateStore As RecoveryStateStore = Nothing, Optional startupService As WindowsStartupService = Nothing)
            _stateStore = If(stateStore, New RecoveryStateStore())
            _startupService = If(startupService, New WindowsStartupService())
            _state = _stateStore.Load()
        End Sub

        Public Function InitializeSession(settings As AppSettings, launchOptions As LaunchOptions) As RecoveryStartupInfo
            If settings Is Nothing Then
                settings = New AppSettings()
            End If

            If launchOptions Is Nothing Then
                launchOptions = New LaunchOptions()
            End If

            SyncLock _syncRoot
                _state = _stateStore.Load()

                Dim startupInfo As New RecoveryStartupInfo()
                Dim hasPendingInterruptedRun = _state.ActiveRun IsNot Nothing AndAlso _state.PendingResumePrompt
                Dim previousRunInterrupted = ((Not _state.CleanShutdown) AndAlso _state.ActiveRun IsNot Nothing) OrElse hasPendingInterruptedRun

                If previousRunInterrupted Then
                    startupInfo.InterruptedRun = CloneActiveRun(_state.ActiveRun)
                    startupInfo.InterruptionReason = If(
                        String.IsNullOrWhiteSpace(_state.LastInterruptionReason),
                        "The previous copy session ended unexpectedly.",
                        _state.LastInterruptionReason)

                    _state.PendingResumePrompt = True
                    _state.LastInterruptionReason = startupInfo.InterruptionReason
                Else
                    startupInfo.InterruptedRun = Nothing
                    startupInfo.InterruptionReason = String.Empty
                End If

                Dim forcePrompt = launchOptions.IsRecoveryAutostart OrElse launchOptions.ForceResumePrompt
                Dim promptPending = startupInfo.HasInterruptedRun AndAlso _state.PendingResumePrompt

                startupInfo.ShouldAutoResume = promptPending AndAlso settings.AutoResumeAfterCrash
                startupInfo.ShouldPrompt = (Not startupInfo.ShouldAutoResume) AndAlso promptPending AndAlso (forcePrompt OrElse settings.PromptResumeAfterCrash)
                startupInfo.WasRecoveryAutostart = launchOptions.IsRecoveryAutostart

                _state.ProcessSessionId = Guid.NewGuid().ToString("N")
                _state.LastStartUtc = DateTimeOffset.UtcNow
                _state.CleanShutdown = False
                SaveStateNoThrow()

                Return startupInfo
            End SyncLock
        End Function

        Public Function GetPendingInterruptedRun() As RecoveryActiveRun
            SyncLock _syncRoot
                If _state Is Nothing OrElse _state.ActiveRun Is Nothing Then
                    Return Nothing
                End If

                Return CloneActiveRun(_state.ActiveRun)
            End SyncLock
        End Function

        Public Sub MarkResumePromptDeferred(settings As AppSettings, suppressForThisRun As Boolean)
            If settings Is Nothing Then
                settings = New AppSettings()
            End If

            SyncLock _syncRoot
                If _state Is Nothing OrElse _state.ActiveRun Is Nothing Then
                    Return
                End If

                If suppressForThisRun Then
                    _state.PendingResumePrompt = False
                Else
                    _state.PendingResumePrompt = settings.KeepResumePromptUntilResolved
                End If

                SaveStateNoThrow()
            End SyncLock
        End Sub

        Public Sub MarkJobStarted(run As ManagedJobRun, options As CopyJobOptions, journalPath As String, settings As AppSettings)
            If run Is Nothing Then
                Throw New ArgumentNullException(NameOf(run))
            End If

            If options Is Nothing Then
                Throw New ArgumentNullException(NameOf(options))
            End If

            If settings Is Nothing Then
                settings = New AppSettings()
            End If

            SyncLock _syncRoot
                Dim nowUtc = DateTimeOffset.UtcNow
                _state.ActiveRun = New RecoveryActiveRun() With {
                    .RunId = run.RunId,
                    .JobId = run.JobId,
                    .JobName = run.DisplayName,
                    .Trigger = run.Trigger,
                    .Options = CloneOptions(options),
                    .StartedUtc = run.StartedUtc,
                    .LastHeartbeatUtc = nowUtc,
                    .LastProgressUtc = nowUtc,
                    .JournalPath = If(journalPath, String.Empty)
                }

                _state.PendingResumePrompt = False
                _state.LastInterruptionReason = String.Empty
                _state.CleanShutdown = False
                _state.LastStartUtc = nowUtc
                _lastTouchUtc = nowUtc
                _touchInterval = TimeSpan.FromSeconds(Math.Max(1, settings.RecoveryTouchIntervalSeconds))

                SaveStateNoThrow()
                If settings.EnableRecoveryAutostart Then
                    _startupService.RegisterRecoveryRunOnce()
                Else
                    _startupService.ClearRecoveryRunOnce()
                End If
            End SyncLock
        End Sub

        Public Sub TouchActiveRun(runId As String)
            If String.IsNullOrWhiteSpace(runId) Then
                Return
            End If

            SyncLock _syncRoot
                If _state Is Nothing OrElse _state.ActiveRun Is Nothing Then
                    Return
                End If

                If Not String.Equals(_state.ActiveRun.RunId, runId, StringComparison.OrdinalIgnoreCase) Then
                    Return
                End If

                Dim nowUtc = DateTimeOffset.UtcNow
                If _lastTouchUtc <> DateTimeOffset.MinValue AndAlso (nowUtc - _lastTouchUtc) < _touchInterval Then
                    Return
                End If

                _state.ActiveRun.LastHeartbeatUtc = nowUtc
                _state.ActiveRun.LastProgressUtc = nowUtc
                _lastTouchUtc = nowUtc
                SaveStateNoThrow()
            End SyncLock
        End Sub

        Public Sub MarkActiveRunPaused(runId As String)
            If String.IsNullOrWhiteSpace(runId) Then
                Return
            End If

            SyncLock _syncRoot
                If _state Is Nothing OrElse _state.ActiveRun Is Nothing Then
                    Return
                End If

                If Not String.Equals(_state.ActiveRun.RunId, runId, StringComparison.OrdinalIgnoreCase) Then
                    Return
                End If

                _state.ActiveRun.LastHeartbeatUtc = DateTimeOffset.UtcNow
                SaveStateNoThrow()
            End SyncLock
        End Sub

        Public Sub MarkJobEnded(runId As String)
            If String.IsNullOrWhiteSpace(runId) Then
                Return
            End If

            SyncLock _syncRoot
                If _state IsNot Nothing AndAlso _state.ActiveRun IsNot Nothing AndAlso
                    String.Equals(_state.ActiveRun.RunId, runId, StringComparison.OrdinalIgnoreCase) Then

                    _state.ActiveRun = Nothing
                    _state.PendingResumePrompt = False
                    _state.LastInterruptionReason = String.Empty
                    SaveStateNoThrow()
                End If

                _startupService.ClearRecoveryRunOnce()
            End SyncLock
        End Sub

        Public Sub MarkRunInterrupted(reason As String)
            SyncLock _syncRoot
                If _state Is Nothing OrElse _state.ActiveRun Is Nothing Then
                    Return
                End If

                _state.PendingResumePrompt = True
                _state.LastInterruptionReason = If(String.IsNullOrWhiteSpace(reason), "Copy session was interrupted.", reason)
                SaveStateNoThrow()
            End SyncLock
        End Sub

        Public Sub MarkCleanShutdown()
            SyncLock _syncRoot
                If _state Is Nothing Then
                    _state = New RecoveryState()
                End If

                _state.CleanShutdown = True
                _state.LastExitUtc = DateTimeOffset.UtcNow
                SaveStateNoThrow()
                _startupService.ClearRecoveryRunOnce()
            End SyncLock
        End Sub

        Private Sub SaveStateNoThrow()
            Try
                _stateStore.Save(_state)
            Catch
                ' Ignore persistence failures and keep runtime flow alive.
            End Try
        End Sub

        Private Shared Function CloneActiveRun(source As RecoveryActiveRun) As RecoveryActiveRun
            If source Is Nothing Then
                Return Nothing
            End If

            Return New RecoveryActiveRun() With {
                .RunId = source.RunId,
                .JobId = source.JobId,
                .JobName = source.JobName,
                .Trigger = source.Trigger,
                .Options = CloneOptions(source.Options),
                .StartedUtc = source.StartedUtc,
                .LastHeartbeatUtc = source.LastHeartbeatUtc,
                .LastProgressUtc = source.LastProgressUtc,
                .JournalPath = source.JournalPath
            }
        End Function

        Private Shared Function CloneOptions(options As CopyJobOptions) As CopyJobOptions
            If options Is Nothing Then
                Return New CopyJobOptions()
            End If

            Return New CopyJobOptions() With {
                .OperationMode = options.OperationMode,
                .SourceRoot = options.SourceRoot,
                .DestinationRoot = options.DestinationRoot,
                .ExpectedSourceIdentity = options.ExpectedSourceIdentity,
                .ExpectedDestinationIdentity = options.ExpectedDestinationIdentity,
                .UseBadRangeMap = options.UseBadRangeMap,
                .SkipKnownBadRanges = options.SkipKnownBadRanges,
                .UpdateBadRangeMapFromRun = options.UpdateBadRangeMapFromRun,
                .BadRangeMapMaxAgeDays = options.BadRangeMapMaxAgeDays,
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
