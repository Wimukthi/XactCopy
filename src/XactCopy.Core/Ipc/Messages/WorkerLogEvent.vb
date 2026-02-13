Namespace Ipc.Messages
    Public Class WorkerLogEvent
        Public Property JobId As String = String.Empty
        Public Property Message As String = String.Empty
        Public Property TimestampUtc As DateTimeOffset = DateTimeOffset.UtcNow
    End Class
End Namespace
