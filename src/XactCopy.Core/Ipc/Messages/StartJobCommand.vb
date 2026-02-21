' -----------------------------------------------------------------------------
' File: src\XactCopy.Core\Ipc\Messages\StartJobCommand.vb
' Purpose: Source file for XactCopy runtime behavior.
' -----------------------------------------------------------------------------

Imports XactCopy.Models

Namespace Ipc.Messages
    ''' <summary>
    ''' Class StartJobCommand.
    ''' </summary>
    Public Class StartJobCommand
        ''' <summary>
        ''' Gets or sets JobId.
        ''' </summary>
        Public Property JobId As String = String.Empty
        ''' <summary>
        ''' Gets or sets Options.
        ''' </summary>
        Public Property Options As CopyJobOptions = New CopyJobOptions()
    End Class
End Namespace
