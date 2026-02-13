Imports XactCopy.Models

Namespace Ipc.Messages
    Public Class WorkerProgressEvent
        Public Property JobId As String = String.Empty
        Public Property Snapshot As CopyProgressSnapshot = New CopyProgressSnapshot()
        Public Property TimestampUtc As DateTimeOffset = DateTimeOffset.UtcNow
    End Class
End Namespace
