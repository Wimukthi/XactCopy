Namespace Models
    Public Class ManagedJobQueueEntry
        Public Property QueueEntryId As String = Guid.NewGuid().ToString("N")
        Public Property JobId As String = String.Empty
        Public Property EnqueuedUtc As DateTimeOffset = DateTimeOffset.UtcNow
        Public Property LastUpdatedUtc As DateTimeOffset = DateTimeOffset.UtcNow
        Public Property Trigger As String = "queued"
        Public Property EnqueuedBy As String = String.Empty
        Public Property AttemptCount As Integer = 0
        Public Property LastAttemptUtc As DateTimeOffset?
        Public Property LastErrorMessage As String = String.Empty
    End Class
End Namespace
