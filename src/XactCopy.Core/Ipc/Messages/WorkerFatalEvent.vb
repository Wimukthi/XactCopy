' -----------------------------------------------------------------------------
' File: src\XactCopy.Core\Ipc\Messages\WorkerFatalEvent.vb
' Purpose: Source file for XactCopy runtime behavior.
' -----------------------------------------------------------------------------

Namespace Ipc.Messages
    ''' <summary>
    ''' Class WorkerFatalEvent.
    ''' </summary>
    Public Class WorkerFatalEvent
        ''' <summary>
        ''' Gets or sets JobId.
        ''' </summary>
        Public Property JobId As String = String.Empty
        ''' <summary>
        ''' Gets or sets ErrorMessage.
        ''' </summary>
        Public Property ErrorMessage As String = String.Empty
        ''' <summary>
        ''' Gets or sets StackTraceText.
        ''' </summary>
        Public Property StackTraceText As String = String.Empty
        ''' <summary>
        ''' Gets or sets TimestampUtc.
        ''' </summary>
        Public Property TimestampUtc As DateTimeOffset = DateTimeOffset.UtcNow
    End Class
End Namespace
