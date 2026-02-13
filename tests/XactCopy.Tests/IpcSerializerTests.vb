Imports XactCopy.Ipc
Imports XactCopy.Ipc.Messages
Imports XactCopy.Models
Imports Xunit

Public Class IpcSerializerTests
    <Fact>
    Public Sub SerializeAndDeserializeEnvelope_RoundTripsStartJobCommand()
        Dim payload As New StartJobCommand() With {
            .JobId = "job-123",
            .Options = New CopyJobOptions() With {
                .SourceRoot = "C:\src",
                .DestinationRoot = "D:\dst",
                .UseAdaptiveBufferSizing = True,
                .WaitForMediaAvailability = True,
                .MaxRetries = 5
            }
        }

        Dim json = IpcSerializer.SerializeEnvelope(IpcMessageTypes.StartJobCommand, payload, "corr-1")
        Dim envelope = IpcSerializer.DeserializeEnvelope(Of StartJobCommand)(json)

        Assert.Equal(ProtocolConstants.Version, envelope.ProtocolVersion)
        Assert.Equal(IpcMessageTypes.StartJobCommand, envelope.MessageType)
        Assert.Equal("corr-1", envelope.CorrelationId)
        Assert.Equal("job-123", envelope.Payload.JobId)
        Assert.Equal("C:\src", envelope.Payload.Options.SourceRoot)
        Assert.Equal("D:\dst", envelope.Payload.Options.DestinationRoot)
        Assert.True(envelope.Payload.Options.UseAdaptiveBufferSizing)
        Assert.True(envelope.Payload.Options.WaitForMediaAvailability)
        Assert.Equal(5, envelope.Payload.Options.MaxRetries)
    End Sub

    <Fact>
    Public Sub TryReadMessageType_ReturnsFalseForMismatchedProtocolVersion()
        Dim payload As New PingCommand()
        Dim validJson = IpcSerializer.SerializeEnvelope(IpcMessageTypes.PingCommand, payload)
        Dim invalidJson = validJson.Replace(
            $"""ProtocolVersion"":{ProtocolConstants.Version}",
            """ProtocolVersion"":999")

        Dim parsedType As String = String.Empty
        Dim parsed = IpcSerializer.TryReadMessageType(invalidJson, parsedType)

        Assert.False(parsed)
        Assert.True(String.IsNullOrWhiteSpace(parsedType))
    End Sub
End Class
