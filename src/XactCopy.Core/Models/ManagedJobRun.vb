Namespace Models
    Public Class ManagedJobRun
        Public Property RunId As String = Guid.NewGuid().ToString("N")
        Public Property JobId As String = String.Empty
        Public Property DisplayName As String = String.Empty
        Public Property SourceRoot As String = String.Empty
        Public Property DestinationRoot As String = String.Empty
        Public Property Trigger As String = "manual"
        Public Property Status As ManagedJobRunStatus = ManagedJobRunStatus.Queued
        Public Property StartedUtc As DateTimeOffset = DateTimeOffset.UtcNow
        Public Property LastUpdatedUtc As DateTimeOffset = DateTimeOffset.UtcNow
        Public Property FinishedUtc As DateTimeOffset?
        Public Property JournalPath As String = String.Empty
        Public Property Summary As String = String.Empty
        Public Property ErrorMessage As String = String.Empty
        Public Property Result As CopyJobResult
    End Class
End Namespace
