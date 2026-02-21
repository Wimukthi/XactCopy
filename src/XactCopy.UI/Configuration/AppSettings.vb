Namespace Configuration
    ''' <summary>
    ''' User-scoped persisted preferences for UI behavior, defaults, integration, and resiliency policy.
    ''' </summary>
    Public NotInheritable Class AppSettings
        Public Const DefaultUpdateReleaseUrl As String = "https://api.github.com/repos/Wimukthi/XactCopy/releases/latest"

        ' Appearance.
        ''' <summary>
        ''' Gets or sets Theme.
        ''' </summary>
        Public Property Theme As String = "dark"
        ''' <summary>
        ''' Gets or sets AccentColorMode.
        ''' </summary>
        Public Property AccentColorMode As String = "auto"
        ''' <summary>
        ''' Gets or sets AccentColorHex.
        ''' </summary>
        Public Property AccentColorHex As String = "#5A78C8"
        ''' <summary>
        ''' Gets or sets WindowChromeMode.
        ''' </summary>
        Public Property WindowChromeMode As String = "themed"
        ''' <summary>
        ''' Gets or sets UiDensity.
        ''' </summary>
        Public Property UiDensity As String = "normal"
        ''' <summary>
        ''' Gets or sets UiScalePercent.
        ''' </summary>
        Public Property UiScalePercent As Integer = 100
        ''' <summary>
        ''' Gets or sets LogFontFamily.
        ''' </summary>
        Public Property LogFontFamily As String = "Consolas"
        ''' <summary>
        ''' Gets or sets LogFontSizePoints.
        ''' </summary>
        Public Property LogFontSizePoints As Integer = 9
        ''' <summary>
        ''' Gets or sets UiColorizeLogBySeverity.
        ''' </summary>
        Public Property UiColorizeLogBySeverity As Boolean = True
        ''' <summary>
        ''' Gets or sets GridAlternatingRows.
        ''' </summary>
        Public Property GridAlternatingRows As Boolean = True
        ''' <summary>
        ''' Gets or sets GridRowHeight.
        ''' </summary>
        Public Property GridRowHeight As Integer = 24
        ''' <summary>
        ''' Gets or sets GridHeaderStyle.
        ''' </summary>
        Public Property GridHeaderStyle As String = "default"
        ''' <summary>
        ''' Gets or sets ProgressBarStyle.
        ''' </summary>
        Public Property ProgressBarStyle As String = "standard"
        ''' <summary>
        ''' Gets or sets ProgressBarShowPercentage.
        ''' </summary>
        Public Property ProgressBarShowPercentage As Boolean = False
        ''' <summary>
        ''' Gets or sets ShowBufferStatusRow.
        ''' </summary>
        Public Property ShowBufferStatusRow As Boolean = True
        ''' <summary>
        ''' Gets or sets ShowRescueStatusRow.
        ''' </summary>
        Public Property ShowRescueStatusRow As Boolean = True

        ' Updates.
        ''' <summary>
        ''' Gets or sets CheckUpdatesOnLaunch.
        ''' </summary>
        Public Property CheckUpdatesOnLaunch As Boolean = False
        ''' <summary>
        ''' Gets or sets UpdateReleaseUrl.
        ''' </summary>
        Public Property UpdateReleaseUrl As String = DefaultUpdateReleaseUrl
        ''' <summary>
        ''' Gets or sets UserAgent.
        ''' </summary>
        Public Property UserAgent As String = String.Empty

        ' Recovery/startup.
        ''' <summary>
        ''' Gets or sets EnableRecoveryAutostart.
        ''' </summary>
        Public Property EnableRecoveryAutostart As Boolean = True
        ''' <summary>
        ''' Gets or sets PromptResumeAfterCrash.
        ''' </summary>
        Public Property PromptResumeAfterCrash As Boolean = True
        ''' <summary>
        ''' Gets or sets AutoResumeAfterCrash.
        ''' </summary>
        Public Property AutoResumeAfterCrash As Boolean = False
        ''' <summary>
        ''' Gets or sets KeepResumePromptUntilResolved.
        ''' </summary>
        Public Property KeepResumePromptUntilResolved As Boolean = True
        ''' <summary>
        ''' Gets or sets AutoRunQueuedJobsOnStartup.
        ''' </summary>
        Public Property AutoRunQueuedJobsOnStartup As Boolean = False
        ''' <summary>
        ''' Gets or sets RecoveryTouchIntervalSeconds.
        ''' </summary>
        Public Property RecoveryTouchIntervalSeconds As Integer = 2

        ' Explorer integration.
        ''' <summary>
        ''' Gets or sets EnableExplorerContextMenu.
        ''' </summary>
        Public Property EnableExplorerContextMenu As Boolean = False
        ''' <summary>
        ''' Gets or sets ExplorerSelectionMode.
        ''' </summary>
        Public Property ExplorerSelectionMode As String = "selected-items"

        ' Default copy/scan policy.
        ''' <summary>
        ''' Gets or sets DefaultResumeFromJournal.
        ''' </summary>
        Public Property DefaultResumeFromJournal As Boolean = True
        ''' <summary>
        ''' Gets or sets DefaultSalvageUnreadableBlocks.
        ''' </summary>
        Public Property DefaultSalvageUnreadableBlocks As Boolean = True
        ''' <summary>
        ''' Gets or sets DefaultContinueOnFileError.
        ''' </summary>
        Public Property DefaultContinueOnFileError As Boolean = True
        ''' <summary>
        ''' Gets or sets DefaultVerifyAfterCopy.
        ''' </summary>
        Public Property DefaultVerifyAfterCopy As Boolean = False
        ''' <summary>
        ''' Gets or sets DefaultUseBadRangeMap.
        ''' </summary>
        Public Property DefaultUseBadRangeMap As Boolean = True
        ''' <summary>
        ''' Gets or sets DefaultSkipKnownBadRanges.
        ''' </summary>
        Public Property DefaultSkipKnownBadRanges As Boolean = True
        ''' <summary>
        ''' Gets or sets DefaultUpdateBadRangeMapFromRun.
        ''' </summary>
        Public Property DefaultUpdateBadRangeMapFromRun As Boolean = True
        ''' <summary>
        ''' Gets or sets DefaultBadRangeMapMaxAgeDays.
        ''' </summary>
        Public Property DefaultBadRangeMapMaxAgeDays As Integer = 30
        ''' <summary>
        ''' Gets or sets DefaultUseExperimentalRawDiskScan.
        ''' </summary>
        Public Property DefaultUseExperimentalRawDiskScan As Boolean = False
        ''' <summary>
        ''' Gets or sets DefaultUseAdaptiveBuffer.
        ''' </summary>
        Public Property DefaultUseAdaptiveBuffer As Boolean = False
        ''' <summary>
        ''' Gets or sets DefaultWaitForMediaAvailability.
        ''' </summary>
        Public Property DefaultWaitForMediaAvailability As Boolean = False
        ''' <summary>
        ''' Gets or sets DefaultWaitForFileLockRelease.
        ''' </summary>
        Public Property DefaultWaitForFileLockRelease As Boolean = False
        ''' <summary>
        ''' Gets or sets DefaultTreatAccessDeniedAsContention.
        ''' </summary>
        Public Property DefaultTreatAccessDeniedAsContention As Boolean = False
        ''' <summary>
        ''' Gets or sets DefaultLockContentionProbeIntervalMs.
        ''' </summary>
        Public Property DefaultLockContentionProbeIntervalMs As Integer = 500
        ''' <summary>
        ''' Gets or sets DefaultSourceMutationPolicy.
        ''' </summary>
        Public Property DefaultSourceMutationPolicy As String = "fail-file"
        ''' <summary>
        ''' Gets or sets DefaultFragileMediaMode.
        ''' </summary>
        Public Property DefaultFragileMediaMode As Boolean = False
        ''' <summary>
        ''' Gets or sets DefaultSkipFileOnFirstReadError.
        ''' </summary>
        Public Property DefaultSkipFileOnFirstReadError As Boolean = True
        ''' <summary>
        ''' Gets or sets DefaultPersistFragileSkipsAcrossResume.
        ''' </summary>
        Public Property DefaultPersistFragileSkipsAcrossResume As Boolean = True
        ''' <summary>
        ''' Gets or sets DefaultFragileFailureWindowSeconds.
        ''' </summary>
        Public Property DefaultFragileFailureWindowSeconds As Integer = 20
        ''' <summary>
        ''' Gets or sets DefaultFragileFailureThreshold.
        ''' </summary>
        Public Property DefaultFragileFailureThreshold As Integer = 3
        ''' <summary>
        ''' Gets or sets DefaultFragileCooldownSeconds.
        ''' </summary>
        Public Property DefaultFragileCooldownSeconds As Integer = 6
        ''' <summary>
        ''' Gets or sets DefaultBufferSizeMb.
        ''' </summary>
        Public Property DefaultBufferSizeMb As Integer = 4
        ''' <summary>
        ''' Gets or sets DefaultMaxRetries.
        ''' </summary>
        Public Property DefaultMaxRetries As Integer = 12
        ''' <summary>
        ''' Gets or sets DefaultOperationTimeoutSeconds.
        ''' </summary>
        Public Property DefaultOperationTimeoutSeconds As Integer = 10
        ''' <summary>
        ''' Gets or sets DefaultPerFileTimeoutSeconds.
        ''' </summary>
        Public Property DefaultPerFileTimeoutSeconds As Integer = 0
        ''' <summary>
        ''' Gets or sets DefaultMaxThroughputMbPerSecond.
        ''' </summary>
        Public Property DefaultMaxThroughputMbPerSecond As Integer = 0
        ''' <summary>
        ''' Gets or sets DefaultPreserveTimestamps.
        ''' </summary>
        Public Property DefaultPreserveTimestamps As Boolean = True
        ''' <summary>
        ''' Gets or sets DefaultCopyEmptyDirectories.
        ''' </summary>
        Public Property DefaultCopyEmptyDirectories As Boolean = True
        ''' <summary>
        ''' Gets or sets WorkerTelemetryProfile.
        ''' </summary>
        Public Property WorkerTelemetryProfile As String = "normal"
        ''' <summary>
        ''' Gets or sets WorkerProgressIntervalMs.
        ''' </summary>
        Public Property WorkerProgressIntervalMs As Integer = 75
        ''' <summary>
        ''' Gets or sets WorkerMaxLogsPerSecond.
        ''' </summary>
        Public Property WorkerMaxLogsPerSecond As Integer = 100
        ''' <summary>
        ''' Gets or sets UiShowDiagnostics.
        ''' </summary>
        Public Property UiShowDiagnostics As Boolean = True
        ''' <summary>
        ''' Gets or sets UiDiagnosticsRefreshMs.
        ''' </summary>
        Public Property UiDiagnosticsRefreshMs As Integer = 250
        ''' <summary>
        ''' Gets or sets UiMaxLogLines.
        ''' </summary>
        Public Property UiMaxLogLines As Integer = 50000
        ''' <summary>
        ''' Gets or sets DefaultOverwritePolicy.
        ''' </summary>
        Public Property DefaultOverwritePolicy As String = "overwrite"
        ''' <summary>
        ''' Gets or sets DefaultSymlinkHandling.
        ''' </summary>
        Public Property DefaultSymlinkHandling As String = "skip"
        ''' <summary>
        ''' Gets or sets DefaultRescueFastScanChunkKb.
        ''' </summary>
        Public Property DefaultRescueFastScanChunkKb As Integer = 0
        ''' <summary>
        ''' Gets or sets DefaultRescueTrimChunkKb.
        ''' </summary>
        Public Property DefaultRescueTrimChunkKb As Integer = 0
        ''' <summary>
        ''' Gets or sets DefaultRescueScrapeChunkKb.
        ''' </summary>
        Public Property DefaultRescueScrapeChunkKb As Integer = 0
        ''' <summary>
        ''' Gets or sets DefaultRescueRetryChunkKb.
        ''' </summary>
        Public Property DefaultRescueRetryChunkKb As Integer = 0
        ''' <summary>
        ''' Gets or sets DefaultRescueSplitMinimumKb.
        ''' </summary>
        Public Property DefaultRescueSplitMinimumKb As Integer = 0
        ''' <summary>
        ''' Gets or sets DefaultRescueFastScanRetries.
        ''' </summary>
        Public Property DefaultRescueFastScanRetries As Integer = 0
        ''' <summary>
        ''' Gets or sets DefaultRescueTrimRetries.
        ''' </summary>
        Public Property DefaultRescueTrimRetries As Integer = 1
        ''' <summary>
        ''' Gets or sets DefaultRescueScrapeRetries.
        ''' </summary>
        Public Property DefaultRescueScrapeRetries As Integer = 2
        ''' <summary>
        ''' Gets or sets DefaultVerificationMode.
        ''' </summary>
        Public Property DefaultVerificationMode As String = "full"
        ''' <summary>
        ''' Gets or sets DefaultVerificationHashAlgorithm.
        ''' </summary>
        Public Property DefaultVerificationHashAlgorithm As String = "sha256"
        ''' <summary>
        ''' Gets or sets DefaultSampleVerificationChunkKb.
        ''' </summary>
        Public Property DefaultSampleVerificationChunkKb As Integer = 128
        ''' <summary>
        ''' Gets or sets DefaultSampleVerificationChunkCount.
        ''' </summary>
        Public Property DefaultSampleVerificationChunkCount As Integer = 3
        ''' <summary>
        ''' Gets or sets DefaultSalvageFillPattern.
        ''' </summary>
        Public Property DefaultSalvageFillPattern As String = "zero"
    End Class
End Namespace
