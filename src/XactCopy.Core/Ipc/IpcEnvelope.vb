Namespace Ipc
    Public Class IpcEnvelope(Of TPayload)
        Public Property ProtocolVersion As Integer = ProtocolConstants.Version
        Public Property MessageType As String = String.Empty
        Public Property CorrelationId As String = String.Empty
        Public Property SentUtc As DateTimeOffset = DateTimeOffset.UtcNow
        Public Property Payload As TPayload
    End Class
End Namespace
