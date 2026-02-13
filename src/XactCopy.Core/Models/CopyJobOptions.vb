Namespace Models
    Public Class CopyJobOptions
        Public Property SourceRoot As String = String.Empty
        Public Property DestinationRoot As String = String.Empty
        Public Property SelectedRelativePaths As New List(Of String)()
        Public Property OverwritePolicy As OverwritePolicy = OverwritePolicy.Overwrite
        Public Property SymlinkHandling As SymlinkHandlingMode = SymlinkHandlingMode.Skip
        Public Property CopyEmptyDirectories As Boolean = True
        Public Property BufferSizeBytes As Integer = 4 * 1024 * 1024
        Public Property UseAdaptiveBufferSizing As Boolean = False
        Public Property MaxThroughputBytesPerSecond As Long = 0
        Public Property ParallelSmallFileWorkers As Integer = 1
        Public Property SmallFileThresholdBytes As Integer = 256 * 1024
        Public Property WaitForMediaAvailability As Boolean = False
        Public Property MaxRetries As Integer = 12
        Public Property OperationTimeout As TimeSpan = TimeSpan.FromSeconds(10)
        Public Property PerFileTimeout As TimeSpan = TimeSpan.Zero
        Public Property InitialRetryDelay As TimeSpan = TimeSpan.FromMilliseconds(250)
        Public Property MaxRetryDelay As TimeSpan = TimeSpan.FromSeconds(8)
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
        Public Property WorkerProcessPriorityClass As String = "Normal"
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
