Namespace Models
    Public Class RecoveryState
        Public Property ProcessSessionId As String = String.Empty
        Public Property LastStartUtc As DateTimeOffset = DateTimeOffset.MinValue
        Public Property LastExitUtc As DateTimeOffset = DateTimeOffset.MinValue
        Public Property CleanShutdown As Boolean = True
        Public Property PendingResumePrompt As Boolean = False
        Public Property LastInterruptionReason As String = String.Empty
        Public Property ActiveRun As RecoveryActiveRun
    End Class
End Namespace
