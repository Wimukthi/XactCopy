Imports XactCopy.Models

Namespace Ipc.Messages
    Public Class WorkerJobResultEvent
        Public Property JobId As String = String.Empty
        Public Property Result As CopyJobResult = New CopyJobResult()
        Public Property TimestampUtc As DateTimeOffset = DateTimeOffset.UtcNow
    End Class
End Namespace
