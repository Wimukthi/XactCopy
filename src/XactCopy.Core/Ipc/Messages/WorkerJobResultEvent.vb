' -----------------------------------------------------------------------------
' File: src\XactCopy.Core\Ipc\Messages\WorkerJobResultEvent.vb
' Purpose: Source file for XactCopy runtime behavior.
' -----------------------------------------------------------------------------

Imports XactCopy.Models

Namespace Ipc.Messages
    ''' <summary>
    ''' Class WorkerJobResultEvent.
    ''' </summary>
    Public Class WorkerJobResultEvent
        ''' <summary>
        ''' Gets or sets JobId.
        ''' </summary>
        Public Property JobId As String = String.Empty
        ''' <summary>
        ''' Gets or sets Result.
        ''' </summary>
        Public Property Result As CopyJobResult = New CopyJobResult()
        ''' <summary>
        ''' Gets or sets TimestampUtc.
        ''' </summary>
        Public Property TimestampUtc As DateTimeOffset = DateTimeOffset.UtcNow
    End Class
End Namespace
