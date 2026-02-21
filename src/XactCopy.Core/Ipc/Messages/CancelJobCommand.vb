' -----------------------------------------------------------------------------
' File: src\XactCopy.Core\Ipc\Messages\CancelJobCommand.vb
' Purpose: Source file for XactCopy runtime behavior.
' -----------------------------------------------------------------------------

Namespace Ipc.Messages
    ''' <summary>
    ''' Class CancelJobCommand.
    ''' </summary>
    Public Class CancelJobCommand
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
