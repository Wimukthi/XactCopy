Namespace Ipc.Messages
    Public Class WorkerFatalEvent
        Public Property JobId As String = String.Empty
        Public Property ErrorMessage As String = String.Empty
        Public Property StackTraceText As String = String.Empty
        Public Property TimestampUtc As DateTimeOffset = DateTimeOffset.UtcNow
    End Class
End Namespace
