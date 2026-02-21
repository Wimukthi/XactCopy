Namespace Models
    ''' <summary>
    ''' Complete run configuration shared across UI, supervisor, worker, and persistence.
    ''' </summary>
    Public Class CopyJobOptions
        ' Run mode and roots.
        ''' <summary>
        ''' Gets or sets OperationMode.
        ''' </summary>
        Public Property OperationMode As JobOperationMode = JobOperationMode.Copy
        ''' <summary>
        ''' Gets or sets SourceRoot.
        ''' </summary>
        Public Property SourceRoot As String = String.Empty
        ''' <summary>
        ''' Gets or sets DestinationRoot.
        ''' </summary>
        Public Property DestinationRoot As String = String.Empty
        ''' <summary>
        ''' Gets or sets ExpectedSourceIdentity.
        ''' </summary>
        Public Property ExpectedSourceIdentity As String = String.Empty
        ''' <summary>
        ''' Gets or sets ExpectedDestinationIdentity.
        ''' </summary>
        Public Property ExpectedDestinationIdentity As String = String.Empty

        ' Bad-range map behavior.
        ''' <summary>
        ''' Gets or sets UseBadRangeMap.
        ''' </summary>
        Public Property UseBadRangeMap As Boolean = False
        ''' <summary>
        ''' Gets or sets SkipKnownBadRanges.
        ''' </summary>
        Public Property SkipKnownBadRanges As Boolean = True
        ''' <summary>
        ''' Gets or sets UpdateBadRangeMapFromRun.
        ''' </summary>
        Public Property UpdateBadRangeMapFromRun As Boolean = True
        ''' <summary>
        ''' Gets or sets BadRangeMapMaxAgeDays.
        ''' </summary>
        Public Property BadRangeMapMaxAgeDays As Integer = 30
        ''' <summary>
        ''' Gets or sets UseExperimentalRawDiskScan.
        ''' </summary>
        Public Property UseExperimentalRawDiskScan As Boolean = False

        ' Resume/remap state used by recovery paths.
        ''' <summary>
        ''' Gets or sets ResumeJournalPathHint.
        ''' </summary>
        Public Property ResumeJournalPathHint As String = String.Empty
        ''' <summary>
        ''' Gets or sets AllowJournalRootRemap.
        ''' </summary>
        Public Property AllowJournalRootRemap As Boolean = False

        ' Optional Explorer-driven subset selection.
        ''' <summary>
        ''' Gets or sets SelectedRelativePaths.
        ''' </summary>
        Public Property SelectedRelativePaths As New List(Of String)()

        ' Copy policy.
        ''' <summary>
        ''' Gets or sets OverwritePolicy.
        ''' </summary>
        Public Property OverwritePolicy As OverwritePolicy = OverwritePolicy.Overwrite
        ''' <summary>
        ''' Gets or sets SymlinkHandling.
        ''' </summary>
        Public Property SymlinkHandling As SymlinkHandlingMode = SymlinkHandlingMode.Skip
        ''' <summary>
        ''' Gets or sets CopyEmptyDirectories.
        ''' </summary>
        Public Property CopyEmptyDirectories As Boolean = True

        ' Throughput and scheduling.
        ''' <summary>
        ''' Gets or sets BufferSizeBytes.
        ''' </summary>
        Public Property BufferSizeBytes As Integer = 4 * 1024 * 1024
        ''' <summary>
        ''' Gets or sets UseAdaptiveBufferSizing.
        ''' </summary>
        Public Property UseAdaptiveBufferSizing As Boolean = False
        ''' <summary>
        ''' Gets or sets MaxThroughputBytesPerSecond.
        ''' </summary>
        Public Property MaxThroughputBytesPerSecond As Long = 0
        ''' <summary>
        ''' Gets or sets ParallelSmallFileWorkers.
        ''' </summary>
        Public Property ParallelSmallFileWorkers As Integer = 1
        ''' <summary>
        ''' Gets or sets SmallFileThresholdBytes.
        ''' </summary>
        Public Property SmallFileThresholdBytes As Integer = 256 * 1024

        ' Availability/contention behavior.
        ''' <summary>
        ''' Gets or sets WaitForMediaAvailability.
        ''' </summary>
        Public Property WaitForMediaAvailability As Boolean = False
        ''' <summary>
        ''' Gets or sets WaitForFileLockRelease.
        ''' </summary>
        Public Property WaitForFileLockRelease As Boolean = False
        ''' <summary>
        ''' Gets or sets TreatAccessDeniedAsContention.
        ''' </summary>
        Public Property TreatAccessDeniedAsContention As Boolean = False
        ''' <summary>
        ''' Gets or sets LockContentionProbeInterval.
        ''' </summary>
        Public Property LockContentionProbeInterval As TimeSpan = TimeSpan.FromMilliseconds(500)
        ''' <summary>
        ''' Gets or sets SourceMutationPolicy.
        ''' </summary>
        Public Property SourceMutationPolicy As SourceMutationPolicy = SourceMutationPolicy.FailFile
        ''' <summary>
        ''' Gets or sets FragileMediaMode.
        ''' </summary>
        Public Property FragileMediaMode As Boolean = False
        ''' <summary>
        ''' Gets or sets SkipFileOnFirstReadError.
        ''' </summary>
        Public Property SkipFileOnFirstReadError As Boolean = True
        ''' <summary>
        ''' Gets or sets PersistFragileSkipAcrossResume.
        ''' </summary>
        Public Property PersistFragileSkipAcrossResume As Boolean = True
        ''' <summary>
        ''' Gets or sets FragileFailureWindowSeconds.
        ''' </summary>
        Public Property FragileFailureWindowSeconds As Integer = 20
        ''' <summary>
        ''' Gets or sets FragileFailureThreshold.
        ''' </summary>
        Public Property FragileFailureThreshold As Integer = 3
        ''' <summary>
        ''' Gets or sets FragileCooldownSeconds.
        ''' </summary>
        Public Property FragileCooldownSeconds As Integer = 6

        ' Retry/timeout profile.
        ''' <summary>
        ''' Gets or sets MaxRetries.
        ''' </summary>
        Public Property MaxRetries As Integer = 12
        ''' <summary>
        ''' Gets or sets OperationTimeout.
        ''' </summary>
        Public Property OperationTimeout As TimeSpan = TimeSpan.FromSeconds(10)
        ''' <summary>
        ''' Gets or sets PerFileTimeout.
        ''' </summary>
        Public Property PerFileTimeout As TimeSpan = TimeSpan.Zero
        ''' <summary>
        ''' Gets or sets InitialRetryDelay.
        ''' </summary>
        Public Property InitialRetryDelay As TimeSpan = TimeSpan.FromMilliseconds(250)
        ''' <summary>
        ''' Gets or sets MaxRetryDelay.
        ''' </summary>
        Public Property MaxRetryDelay As TimeSpan = TimeSpan.FromSeconds(8)

        ' Resume/verification/salvage policy.
        ''' <summary>
        ''' Gets or sets ResumeFromJournal.
        ''' </summary>
        Public Property ResumeFromJournal As Boolean = True
        ''' <summary>
        ''' Gets or sets VerifyAfterCopy.
        ''' </summary>
        Public Property VerifyAfterCopy As Boolean = False
        ''' <summary>
        ''' Gets or sets VerificationMode.
        ''' </summary>
        Public Property VerificationMode As VerificationMode = VerificationMode.None
        ''' <summary>
        ''' Gets or sets VerificationHashAlgorithm.
        ''' </summary>
        Public Property VerificationHashAlgorithm As VerificationHashAlgorithm = VerificationHashAlgorithm.Sha256
        ''' <summary>
        ''' Gets or sets SampleVerificationChunkBytes.
        ''' </summary>
        Public Property SampleVerificationChunkBytes As Integer = 128 * 1024
        ''' <summary>
        ''' Gets or sets SampleVerificationChunkCount.
        ''' </summary>
        Public Property SampleVerificationChunkCount As Integer = 3
        ''' <summary>
        ''' Gets or sets SalvageUnreadableBlocks.
        ''' </summary>
        Public Property SalvageUnreadableBlocks As Boolean = True
        ''' <summary>
        ''' Gets or sets SalvageFillPattern.
        ''' </summary>
        Public Property SalvageFillPattern As SalvageFillPattern = SalvageFillPattern.Zero
        ''' <summary>
        ''' Gets or sets ContinueOnFileError.
        ''' </summary>
        Public Property ContinueOnFileError As Boolean = True
        ''' <summary>
        ''' Gets or sets PreserveTimestamps.
        ''' </summary>
        Public Property PreserveTimestamps As Boolean = True

        ' Worker telemetry and process hints.
        ''' <summary>
        ''' Gets or sets WorkerProcessPriorityClass.
        ''' </summary>
        Public Property WorkerProcessPriorityClass As String = "Normal"
        ''' <summary>
        ''' Gets or sets WorkerTelemetryProfile.
        ''' </summary>
        Public Property WorkerTelemetryProfile As WorkerTelemetryProfile = WorkerTelemetryProfile.Normal
        ''' <summary>
        ''' Gets or sets WorkerProgressEmitIntervalMs.
        ''' </summary>
        Public Property WorkerProgressEmitIntervalMs As Integer = 75
        ''' <summary>
        ''' Gets or sets WorkerMaxLogsPerSecond.
        ''' </summary>
        Public Property WorkerMaxLogsPerSecond As Integer = 100

        ' Rescue tuning (0 keeps built-in defaults).
        ''' <summary>
        ''' Gets or sets RescueFastScanChunkBytes.
        ''' </summary>
        Public Property RescueFastScanChunkBytes As Integer = 0
        ''' <summary>
        ''' Gets or sets RescueTrimChunkBytes.
        ''' </summary>
        Public Property RescueTrimChunkBytes As Integer = 0
        ''' <summary>
        ''' Gets or sets RescueScrapeChunkBytes.
        ''' </summary>
        Public Property RescueScrapeChunkBytes As Integer = 0
        ''' <summary>
        ''' Gets or sets RescueRetryChunkBytes.
        ''' </summary>
        Public Property RescueRetryChunkBytes As Integer = 0
        ''' <summary>
        ''' Gets or sets RescueSplitMinimumBytes.
        ''' </summary>
        Public Property RescueSplitMinimumBytes As Integer = 0
        ''' <summary>
        ''' Gets or sets RescueFastScanRetries.
        ''' </summary>
        Public Property RescueFastScanRetries As Integer = 0
        ''' <summary>
        ''' Gets or sets RescueTrimRetries.
        ''' </summary>
        Public Property RescueTrimRetries As Integer = 1
        ''' <summary>
        ''' Gets or sets RescueScrapeRetries.
        ''' </summary>
        Public Property RescueScrapeRetries As Integer = 2
    End Class
End Namespace
