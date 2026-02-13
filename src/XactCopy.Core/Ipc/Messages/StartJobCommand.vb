Imports XactCopy.Models

Namespace Ipc.Messages
    Public Class StartJobCommand
        Public Property JobId As String = String.Empty
        Public Property Options As CopyJobOptions = New CopyJobOptions()
    End Class
End Namespace
