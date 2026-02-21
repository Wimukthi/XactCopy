' -----------------------------------------------------------------------------
' File: src\XactCopy.Core\Ipc\Messages\WorkerProgressEvent.vb
' Purpose: Source file for XactCopy runtime behavior.
' -----------------------------------------------------------------------------

Imports XactCopy.Models

Namespace Ipc.Messages
    ''' <summary>
    ''' Class WorkerProgressEvent.
    ''' </summary>
    Public Class WorkerProgressEvent
        ''' <summary>
        ''' Gets or sets JobId.
        ''' </summary>
        Public Property JobId As String = String.Empty
        ''' <summary>
        ''' Gets or sets Snapshot.
        ''' </summary>
        Public Property Snapshot As CopyProgressSnapshot = New CopyProgressSnapshot()
        ''' <summary>
        ''' Gets or sets TimestampUtc.
        ''' </summary>
        Public Property TimestampUtc As DateTimeOffset = DateTimeOffset.UtcNow
    End Class
End Namespace
