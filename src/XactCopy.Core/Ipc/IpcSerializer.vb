Imports System.Text.Json
Imports System.Text.Json.Serialization

Namespace Ipc
    Public NotInheritable Class IpcSerializer
        Private Shared ReadOnly SerializerOptions As JsonSerializerOptions = BuildOptions()

        Private Sub New()
        End Sub

        Public Shared Function SerializeEnvelope(Of TPayload)(
            messageType As String,
            payload As TPayload,
            Optional correlationId As String = "") As String

            Dim envelope As New IpcEnvelope(Of TPayload)() With {
                .ProtocolVersion = ProtocolConstants.Version,
                .MessageType = messageType,
                .CorrelationId = correlationId,
                .SentUtc = DateTimeOffset.UtcNow,
                .Payload = payload
            }

            Return JsonSerializer.Serialize(envelope, SerializerOptions)
        End Function

        Public Shared Function DeserializeEnvelope(Of TPayload)(json As String) As IpcEnvelope(Of TPayload)
            Dim envelope = JsonSerializer.Deserialize(Of IpcEnvelope(Of TPayload))(json, SerializerOptions)
            If envelope Is Nothing Then
                Throw New InvalidOperationException("IPC envelope deserialization returned null.")
            End If

            If envelope.ProtocolVersion <> ProtocolConstants.Version Then
                Throw New InvalidOperationException(
                    $"Unsupported protocol version {envelope.ProtocolVersion}. Expected {ProtocolConstants.Version}.")
            End If

            Return envelope
        End Function

        Public Shared Function TryReadMessageType(json As String, ByRef messageType As String) As Boolean
            messageType = String.Empty
            Try
                Using document = JsonDocument.Parse(json)
                    Dim root = document.RootElement
                    Dim protocolElement As JsonElement
                    If Not root.TryGetProperty(NameOf(IpcEnvelope(Of Object).ProtocolVersion), protocolElement) Then
                        Return False
                    End If

                    Dim protocolVersion = protocolElement.GetInt32()
                    If protocolVersion <> ProtocolConstants.Version Then
                        Return False
                    End If

                    Dim messageTypeElement As JsonElement
                    If Not root.TryGetProperty(NameOf(IpcEnvelope(Of Object).MessageType), messageTypeElement) Then
                        Return False
                    End If

                    messageType = messageTypeElement.GetString()
                    Return Not String.IsNullOrWhiteSpace(messageType)
                End Using
            Catch
                Return False
            End Try
        End Function

        Private Shared Function BuildOptions() As JsonSerializerOptions
            Dim options As New JsonSerializerOptions() With {
                .PropertyNameCaseInsensitive = True
            }
            options.Converters.Add(New JsonStringEnumConverter())
            Return options
        End Function
    End Class
End Namespace
