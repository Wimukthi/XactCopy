Namespace Models
    Public Class RecoveryActiveRun
        Public Property RunId As String = String.Empty
        Public Property JobId As String = String.Empty
        Public Property JobName As String = String.Empty
        Public Property Trigger As String = "manual"
        Public Property Options As New CopyJobOptions()
        Public Property StartedUtc As DateTimeOffset = DateTimeOffset.UtcNow
        Public Property LastHeartbeatUtc As DateTimeOffset = DateTimeOffset.UtcNow
        Public Property LastProgressUtc As DateTimeOffset = DateTimeOffset.UtcNow
        Public Property JournalPath As String = String.Empty
    End Class
End Namespace
