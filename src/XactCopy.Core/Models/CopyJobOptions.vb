Namespace Models
    ''' <summary>
    ''' Complete run configuration shared across UI, supervisor, worker, and persistence.
    ''' </summary>
    Public Class CopyJobOptions
        ' Run mode and roots.
        Public Property OperationMode As JobOperationMode = JobOperationMode.Copy
        Public Property SourceRoot As String = String.Empty
        Public Property DestinationRoot As String = String.Empty
        Public Property ExpectedSourceIdentity As String = String.Empty
        Public Property ExpectedDestinationIdentity As String = String.Empty

        ' Bad-range map behavior.
        Public Property UseBadRangeMap As Boolean = False
        Public Property SkipKnownBadRanges As Boolean = True
        Public Property UpdateBadRangeMapFromRun As Boolean = True
        Public Property BadRangeMapMaxAgeDays As Integer = 30
        Public Property UseExperimentalRawDiskScan As Boolean = False

        ' Resume/remap state used by recovery paths.
        Public Property ResumeJournalPathHint As String = String.Empty
        Public Property AllowJournalRootRemap As Boolean = False

        ' Optional Explorer-driven subset selection.
        Public Property SelectedRelativePaths As New List(Of String)()

        ' Copy policy.
        Public Property OverwritePolicy As OverwritePolicy = OverwritePolicy.Overwrite
        Public Property SymlinkHandling As SymlinkHandlingMode = SymlinkHandlingMode.Skip
        Public Property CopyEmptyDirectories As Boolean = True

        ' Throughput and scheduling.
        Public Property BufferSizeBytes As Integer = 4 * 1024 * 1024
        Public Property UseAdaptiveBufferSizing As Boolean = False
        Public Property MaxThroughputBytesPerSecond As Long = 0
        Public Property ParallelSmallFileWorkers As Integer = 1
        Public Property SmallFileThresholdBytes As Integer = 256 * 1024

        ' Availability/contention behavior.
        Public Property WaitForMediaAvailability As Boolean = False
        Public Property WaitForFileLockRelease As Boolean = False
        Public Property TreatAccessDeniedAsContention As Boolean = False
        Public Property LockContentionProbeInterval As TimeSpan = TimeSpan.FromMilliseconds(500)
        Public Property SourceMutationPolicy As SourceMutationPolicy = SourceMutationPolicy.FailFile
        Public Property FragileMediaMode As Boolean = False
        Public Property SkipFileOnFirstReadError As Boolean = True
        Public Property PersistFragileSkipAcrossResume As Boolean = True
        Public Property FragileFailureWindowSeconds As Integer = 20
        Public Property FragileFailureThreshold As Integer = 3
        Public Property FragileCooldownSeconds As Integer = 6

        ' Retry/timeout profile.
        Public Property MaxRetries As Integer = 12
        Public Property OperationTimeout As TimeSpan = TimeSpan.FromSeconds(10)
        Public Property PerFileTimeout As TimeSpan = TimeSpan.Zero
        Public Property InitialRetryDelay As TimeSpan = TimeSpan.FromMilliseconds(250)
        Public Property MaxRetryDelay As TimeSpan = TimeSpan.FromSeconds(8)

        ' Resume/verification/salvage policy.
        Public Property ResumeFromJournal As Boolean = True
        Public Property VerifyAfterCopy As Boolean = False
        Public Property VerificationMode As VerificationMode = VerificationMode.None
        Public Property VerificationHashAlgorithm As VerificationHashAlgorithm = VerificationHashAlgorithm.Sha256
        Public Property SampleVerificationChunkBytes As Integer = 128 * 1024
        Public Property SampleVerificationChunkCount As Integer = 3
        Public Property SalvageUnreadableBlocks As Boolean = True
        Public Property SalvageFillPattern As SalvageFillPattern = SalvageFillPattern.Zero
        Public Property ContinueOnFileError As Boolean = True
        Public Property PreserveTimestamps As Boolean = True

        ' Worker telemetry and process hints.
        Public Property WorkerProcessPriorityClass As String = "Normal"
        Public Property WorkerTelemetryProfile As WorkerTelemetryProfile = WorkerTelemetryProfile.Normal
        Public Property WorkerProgressEmitIntervalMs As Integer = 75
        Public Property WorkerMaxLogsPerSecond As Integer = 100

        ' Rescue tuning (0 keeps built-in defaults).
        Public Property RescueFastScanChunkBytes As Integer = 0
        Public Property RescueTrimChunkBytes As Integer = 0
        Public Property RescueScrapeChunkBytes As Integer = 0
        Public Property RescueRetryChunkBytes As Integer = 0
        Public Property RescueSplitMinimumBytes As Integer = 0
        Public Property RescueFastScanRetries As Integer = 0
        Public Property RescueTrimRetries As Integer = 1
        Public Property RescueScrapeRetries As Integer = 2
    End Class
End Namespace
