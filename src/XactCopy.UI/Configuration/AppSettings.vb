Namespace Configuration
    Public NotInheritable Class AppSettings
        Public Const DefaultUpdateReleaseUrl As String = "https://api.github.com/repos/Wimukthi/XactCopy/releases/latest"

        Public Property Theme As String = "dark"

        Public Property CheckUpdatesOnLaunch As Boolean = False
        Public Property UpdateReleaseUrl As String = DefaultUpdateReleaseUrl
        Public Property UserAgent As String = String.Empty

        Public Property EnableRecoveryAutostart As Boolean = True
        Public Property PromptResumeAfterCrash As Boolean = True
        Public Property AutoResumeAfterCrash As Boolean = False
        Public Property KeepResumePromptUntilResolved As Boolean = True
        Public Property AutoRunQueuedJobsOnStartup As Boolean = False
        Public Property RecoveryTouchIntervalSeconds As Integer = 2

        Public Property EnableExplorerContextMenu As Boolean = False
        Public Property ExplorerSelectionMode As String = "selected-items"

        Public Property DefaultResumeFromJournal As Boolean = True
        Public Property DefaultSalvageUnreadableBlocks As Boolean = True
        Public Property DefaultContinueOnFileError As Boolean = True
        Public Property DefaultVerifyAfterCopy As Boolean = False
        Public Property DefaultUseAdaptiveBuffer As Boolean = False
        Public Property DefaultWaitForMediaAvailability As Boolean = False
        Public Property DefaultBufferSizeMb As Integer = 4
        Public Property DefaultMaxRetries As Integer = 12
        Public Property DefaultOperationTimeoutSeconds As Integer = 10
        Public Property DefaultPerFileTimeoutSeconds As Integer = 0
        Public Property DefaultMaxThroughputMbPerSecond As Integer = 0
        Public Property DefaultPreserveTimestamps As Boolean = True
        Public Property DefaultCopyEmptyDirectories As Boolean = True
        Public Property DefaultOverwritePolicy As String = "overwrite"
        Public Property DefaultSymlinkHandling As String = "skip"
        Public Property DefaultRescueFastScanChunkKb As Integer = 0
        Public Property DefaultRescueTrimChunkKb As Integer = 0
        Public Property DefaultRescueScrapeChunkKb As Integer = 0
        Public Property DefaultRescueRetryChunkKb As Integer = 0
        Public Property DefaultRescueSplitMinimumKb As Integer = 0
        Public Property DefaultRescueFastScanRetries As Integer = 0
        Public Property DefaultRescueTrimRetries As Integer = 1
        Public Property DefaultRescueScrapeRetries As Integer = 2
        Public Property DefaultVerificationMode As String = "full"
        Public Property DefaultVerificationHashAlgorithm As String = "sha256"
        Public Property DefaultSampleVerificationChunkKb As Integer = 128
        Public Property DefaultSampleVerificationChunkCount As Integer = 3
        Public Property DefaultSalvageFillPattern As String = "zero"
    End Class
End Namespace
