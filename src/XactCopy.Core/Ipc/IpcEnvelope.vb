' -----------------------------------------------------------------------------
' File: src\XactCopy.Core\Ipc\IpcEnvelope.vb
' Purpose: Source file for XactCopy runtime behavior.
' -----------------------------------------------------------------------------

Namespace Ipc
    ''' <summary>
    ''' Class IpcEnvelope.
    ''' </summary>
    Public Class IpcEnvelope(Of TPayload)
        ''' <summary>
        ''' Gets or sets ProtocolVersion.
        ''' </summary>
        Public Property ProtocolVersion As Integer = ProtocolConstants.Version
        ''' <summary>
        ''' Gets or sets MessageType.
        ''' </summary>
        Public Property MessageType As String = String.Empty
        ''' <summary>
        ''' Gets or sets CorrelationId.
        ''' </summary>
        Public Property CorrelationId As String = String.Empty
        ''' <summary>
        ''' Gets or sets SentUtc.
        ''' </summary>
        Public Property SentUtc As DateTimeOffset = DateTimeOffset.UtcNow
        ''' <summary>
        ''' Gets or sets Payload.
        ''' </summary>
        Public Property Payload As TPayload
    End Class
End Namespace
