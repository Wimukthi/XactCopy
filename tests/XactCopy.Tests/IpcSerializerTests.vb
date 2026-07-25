' -----------------------------------------------------------------------------
' File: tests\XactCopy.Tests\IpcSerializerTests.vb
' Purpose: Source file for XactCopy runtime behavior.
' -----------------------------------------------------------------------------

Imports XactCopy.Ipc
Imports XactCopy.Ipc.Messages
Imports XactCopy.Models
Imports Xunit

''' <summary>
''' Class IpcSerializerTests.
''' </summary>
Public Class IpcSerializerTests
    ''' <summary>
    ''' Executes SerializeAndDeserializeEnvelope_RoundTripsStartJobCommand.
    ''' </summary>
    <Fact>
    Public Sub SerializeAndDeserializeEnvelope_RoundTripsStartJobCommand()
        Dim payload As New StartJobCommand() With {
            .JobId = "job-123",
            .Options = New CopyJobOptions() With {
                .SourceRoot = "C:\src",
                .DestinationRoot = "D:\dst",
                .UseAdaptiveBufferSizing = True,
                .TransferEnginePolicy = TransferEnginePolicy.ManagedRescue,
                .ScanPerformanceProfile = ScanPerformanceProfile.Fast,
                .ParallelScanWorkers = 4,
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
        Assert.Equal(TransferEnginePolicy.ManagedRescue, envelope.Payload.Options.TransferEnginePolicy)
        Assert.Equal(ScanPerformanceProfile.Fast, envelope.Payload.Options.ScanPerformanceProfile)
        Assert.Equal(4, envelope.Payload.Options.ParallelScanWorkers)
        Assert.True(envelope.Payload.Options.WaitForMediaAvailability)
        Assert.Equal(5, envelope.Payload.Options.MaxRetries)
        Assert.Equal(64 * 1024, envelope.Payload.Options.RescueTrimChunkBytes)
        Assert.Equal(4, envelope.Payload.Options.RescueScrapeRetries)
    End Sub

    ''' <summary>
    ''' Executes TryReadMessageType_ReturnsFalseForMismatchedProtocolVersion.
    ''' </summary>
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

    ''' <summary>
    ''' Executes SerializeAndDeserializeEnvelope_RoundTripsWorkerProgressRescueTelemetry.
    ''' </summary>
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
                .RescueRemainingBytes = 8192,
                .ActiveFileCount = 2,
                .ScanWorkerCount = 4,
                .ActiveFiles = New List(Of String) From {"a.bin", "b.bin"}
            }
        }

        Dim json = IpcSerializer.SerializeEnvelope(IpcMessageTypes.WorkerProgressEvent, payload)
        Dim envelope = IpcSerializer.DeserializeEnvelope(Of WorkerProgressEvent)(json)

        Assert.Equal("job-456", envelope.Payload.JobId)
        Assert.Equal("Scrape", envelope.Payload.Snapshot.RescuePass)
        Assert.Equal(3, envelope.Payload.Snapshot.RescueBadRegionCount)
        Assert.Equal(8192, envelope.Payload.Snapshot.RescueRemainingBytes)
        Assert.Equal(2, envelope.Payload.Snapshot.ActiveFileCount)
        Assert.Equal(4, envelope.Payload.Snapshot.ScanWorkerCount)
        Assert.Equal(2, envelope.Payload.Snapshot.ActiveFiles.Count)
    End Sub

    ''' <summary>
    ''' Executes SerializeAndDeserializeEnvelope_RoundTripsWorkerResultEngineTelemetry.
    ''' </summary>
    <Fact>
    Public Sub SerializeAndDeserializeEnvelope_RoundTripsWorkerResultEngineTelemetry()
        Dim payload As New WorkerJobResultEvent() With {
            .JobId = "job-789",
            .Result = New CopyJobResult() With {
                .Succeeded = True,
                .TotalFiles = 5,
                .CompletedFiles = 5,
                .TotalBytes = 40960,
                .CopiedBytes = 40960,
                .TransferEnginePolicy = TransferEnginePolicy.NativeFast,
                .ElapsedMilliseconds = 250,
                .AverageBytesPerSecond = 163840,
                .NativeFastPathFiles = 2,
                .ParallelNativeFastPathFiles = 3,
                .ManagedCopyFiles = 0,
                .NativeFallbackFiles = 1
            }
        }

        Dim json = IpcSerializer.SerializeEnvelope(IpcMessageTypes.WorkerJobResultEvent, payload)
        Dim envelope = IpcSerializer.DeserializeEnvelope(Of WorkerJobResultEvent)(json)

        Assert.Equal("job-789", envelope.Payload.JobId)
        Assert.True(envelope.Payload.Result.Succeeded)
        Assert.Equal(TransferEnginePolicy.NativeFast, envelope.Payload.Result.TransferEnginePolicy)
        Assert.Equal(250, envelope.Payload.Result.ElapsedMilliseconds)
        Assert.Equal(163840, envelope.Payload.Result.AverageBytesPerSecond)
        Assert.Equal(2, envelope.Payload.Result.NativeFastPathFiles)
        Assert.Equal(3, envelope.Payload.Result.ParallelNativeFastPathFiles)
        Assert.Equal(0, envelope.Payload.Result.ManagedCopyFiles)
        Assert.Equal(1, envelope.Payload.Result.NativeFallbackFiles)
    End Sub
End Class
