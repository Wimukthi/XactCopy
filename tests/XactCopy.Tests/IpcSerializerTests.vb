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
                .MaxRetries = 5,
                .RescueTrimChunkBytes = 64 * 1024,
                .RescueScrapeRetries = 4
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
        Assert.Equal(64 * 1024, envelope.Payload.Options.RescueTrimChunkBytes)
        Assert.Equal(4, envelope.Payload.Options.RescueScrapeRetries)
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

    <Fact>
    Public Sub SerializeAndDeserializeEnvelope_RoundTripsWorkerProgressRescueTelemetry()
        Dim payload As New WorkerProgressEvent() With {
            .JobId = "job-456",
            .Snapshot = New CopyProgressSnapshot() With {
                .CurrentFile = "data.bin",
                .CurrentFileBytesCopied = 512,
                .CurrentFileBytesTotal = 2048,
                .TotalBytesCopied = 1024,
                .TotalBytes = 4096,
                .RescuePass = "Scrape",
                .RescueBadRegionCount = 3,
                .RescueRemainingBytes = 8192
            }
        }

        Dim json = IpcSerializer.SerializeEnvelope(IpcMessageTypes.WorkerProgressEvent, payload)
        Dim envelope = IpcSerializer.DeserializeEnvelope(Of WorkerProgressEvent)(json)

        Assert.Equal("job-456", envelope.Payload.JobId)
        Assert.Equal("Scrape", envelope.Payload.Snapshot.RescuePass)
        Assert.Equal(3, envelope.Payload.Snapshot.RescueBadRegionCount)
        Assert.Equal(8192, envelope.Payload.Snapshot.RescueRemainingBytes)
    End Sub
End Class
