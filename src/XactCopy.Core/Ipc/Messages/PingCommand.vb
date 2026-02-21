' -----------------------------------------------------------------------------
' File: src\XactCopy.Core\Ipc\Messages\PingCommand.vb
' Purpose: Source file for XactCopy runtime behavior.
' -----------------------------------------------------------------------------

Namespace Ipc.Messages
    ''' <summary>
    ''' Class PingCommand.
    ''' </summary>
    Public Class PingCommand
        ''' <summary>
        ''' Gets or sets Utc.
        ''' </summary>
        Public Property Utc As DateTimeOffset = DateTimeOffset.UtcNow
    End Class
End Namespace
