' -----------------------------------------------------------------------------
' File: src\XactCopy.Core\Ipc\IpcSerializer.vb
' Purpose: Source file for XactCopy runtime behavior.
' -----------------------------------------------------------------------------

Imports System.Text.Json
Imports System.Text.Json.Serialization

Namespace Ipc
    ''' <summary>
    ''' Class IpcSerializer.
    ''' </summary>
    Public NotInheritable Class IpcSerializer
        Private Shared ReadOnly SerializerOptions As JsonSerializerOptions = BuildOptions()

        Private Sub New()
        End Sub

        ''' <summary>
        ''' Computes SerializeEnvelope.
        ''' </summary>
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

        ''' <summary>
        ''' Computes DeserializeEnvelope.
        ''' </summary>
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

        ''' <summary>
        ''' Computes TryReadMessageType.
        ''' </summary>
        Public Shared Function TryReadMessageType(json As String, ByRef messageType As String) As Boolean
            messageType = String.Empty
            If String.IsNullOrEmpty(json) Then Return False

            Try
                ' Fast path: scan for "ProtocolVersion" and "MessageType" fields directly in the
                ' JSON string without allocating a JsonDocument DOM tree. This method is called for
                ' every incoming IPC message, so avoiding DOM allocation is worthwhile.
                ' Assumptions that make this safe:
                '   - The envelope is serialized by IpcSerializer.SerializeEnvelope using
                '     SerializerOptions (compact, PascalCase, no extra whitespace).
                '   - MessageType values are simple ASCII identifiers with no embedded quotes.

                ' -- Locate and validate the ProtocolVersion field. --------------------------
                Const pvKey As String = """ProtocolVersion"":"
                Dim pvIdx = json.IndexOf(pvKey, StringComparison.Ordinal)
                If pvIdx < 0 Then Return False

                Dim vStart = pvIdx + pvKey.Length
                Dim vEnd = vStart
                While vEnd < json.Length AndAlso Char.IsDigit(json(vEnd))
                    vEnd += 1
                End While
                If vEnd = vStart Then Return False

                Dim protocolVersion As Integer
                If Not Integer.TryParse(json.AsSpan(vStart, vEnd - vStart), protocolVersion) Then Return False
                If protocolVersion <> ProtocolConstants.Version Then Return False

                ' -- Locate and extract the MessageType string value. -----------------------
                Const mtKey As String = """MessageType"":"""
                Dim mtIdx = json.IndexOf(mtKey, StringComparison.Ordinal)
                If mtIdx < 0 Then Return False

                Dim valStart = mtIdx + mtKey.Length
                Dim valEnd = json.IndexOf("""", valStart, StringComparison.Ordinal)
                If valEnd < valStart Then Return False

                messageType = json.Substring(valStart, valEnd - valStart)
                Return Not String.IsNullOrWhiteSpace(messageType)
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
