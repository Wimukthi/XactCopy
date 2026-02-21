' -----------------------------------------------------------------------------
' File: src\XactCopy.Core\Ipc\Messages\PauseJobCommand.vb
' Purpose: Source file for XactCopy runtime behavior.
' -----------------------------------------------------------------------------

Namespace Ipc.Messages
    ''' <summary>
    ''' Class PauseJobCommand.
    ''' </summary>
    Public Class PauseJobCommand
        ''' <summary>
        ''' Gets or sets JobId.
        ''' </summary>
        Public Property JobId As String = String.Empty
        ''' <summary>
        ''' Gets or sets Reason.
        ''' </summary>
        Public Property Reason As String = String.Empty
    End Class
End Namespace
