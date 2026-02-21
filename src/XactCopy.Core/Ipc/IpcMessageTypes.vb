' -----------------------------------------------------------------------------
' File: src\XactCopy.Core\Ipc\IpcMessageTypes.vb
' Purpose: Source file for XactCopy runtime behavior.
' -----------------------------------------------------------------------------

Namespace Ipc
    ''' <summary>
    ''' Class IpcMessageTypes.
    ''' </summary>
    Public NotInheritable Class IpcMessageTypes
        Public Const StartJobCommand As String = "StartJobCommand"
        Public Const CancelJobCommand As String = "CancelJobCommand"
        Public Const PauseJobCommand As String = "PauseJobCommand"
        Public Const ResumeJobCommand As String = "ResumeJobCommand"
        Public Const PingCommand As String = "PingCommand"
        Public Const ShutdownCommand As String = "ShutdownCommand"

        Public Const WorkerHeartbeatEvent As String = "WorkerHeartbeatEvent"
        Public Const WorkerProgressEvent As String = "WorkerProgressEvent"
        Public Const WorkerLogEvent As String = "WorkerLogEvent"
        Public Const WorkerJobResultEvent As String = "WorkerJobResultEvent"
        Public Const WorkerFatalEvent As String = "WorkerFatalEvent"

        Private Sub New()
        End Sub
    End Class
End Namespace
