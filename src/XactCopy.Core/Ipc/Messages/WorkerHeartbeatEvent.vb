Namespace Ipc.Messages
    Public Class WorkerHeartbeatEvent
        Public Property WorkerProcessId As Integer
        Public Property IsJobRunning As Boolean
        Public Property IsJobPaused As Boolean
        Public Property ActiveJobId As String = String.Empty
        Public Property LastProgressUtc As DateTimeOffset
        Public Property TimestampUtc As DateTimeOffset = DateTimeOffset.UtcNow
    End Class
End Namespace
