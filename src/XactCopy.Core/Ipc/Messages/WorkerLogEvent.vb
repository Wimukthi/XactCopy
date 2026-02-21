' -----------------------------------------------------------------------------
' File: src\XactCopy.Core\Ipc\Messages\WorkerLogEvent.vb
' Purpose: Source file for XactCopy runtime behavior.
' -----------------------------------------------------------------------------

Namespace Ipc.Messages
    ''' <summary>
    ''' Class WorkerLogEvent.
    ''' </summary>
    Public Class WorkerLogEvent
        ''' <summary>
        ''' Gets or sets JobId.
        ''' </summary>
        Public Property JobId As String = String.Empty
        ''' <summary>
        ''' Gets or sets Message.
        ''' </summary>
        Public Property Message As String = String.Empty
        ''' <summary>
        ''' Gets or sets TimestampUtc.
        ''' </summary>
        Public Property TimestampUtc As DateTimeOffset = DateTimeOffset.UtcNow
    End Class
End Namespace
