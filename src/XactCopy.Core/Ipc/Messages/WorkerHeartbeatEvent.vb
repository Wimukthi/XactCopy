' -----------------------------------------------------------------------------
' File: src\XactCopy.Core\Ipc\Messages\WorkerHeartbeatEvent.vb
' Purpose: Source file for XactCopy runtime behavior.
' -----------------------------------------------------------------------------

Namespace Ipc.Messages
    ''' <summary>
    ''' Class WorkerHeartbeatEvent.
    ''' </summary>
    Public Class WorkerHeartbeatEvent
        ''' <summary>
        ''' Gets or sets WorkerProcessId.
        ''' </summary>
        Public Property WorkerProcessId As Integer
        ''' <summary>
        ''' Gets or sets IsJobRunning.
        ''' </summary>
        Public Property IsJobRunning As Boolean
        ''' <summary>
        ''' Gets or sets IsJobPaused.
        ''' </summary>
        Public Property IsJobPaused As Boolean
        ''' <summary>
        ''' Gets or sets ActiveJobId.
        ''' </summary>
        Public Property ActiveJobId As String = String.Empty
        ''' <summary>
        ''' Gets or sets LastProgressUtc.
        ''' </summary>
        Public Property LastProgressUtc As DateTimeOffset
        ''' <summary>
        ''' Gets or sets TimestampUtc.
        ''' </summary>
        Public Property TimestampUtc As DateTimeOffset = DateTimeOffset.UtcNow
    End Class
End Namespace
