' -----------------------------------------------------------------------------
' File: src\XactCopy.Core\Ipc\Messages\ResumeJobCommand.vb
' Purpose: Source file for XactCopy runtime behavior.
' -----------------------------------------------------------------------------

Namespace Ipc.Messages
    ''' <summary>
    ''' Class ResumeJobCommand.
    ''' </summary>
    Public Class ResumeJobCommand
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
