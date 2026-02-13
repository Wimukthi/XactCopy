Imports System.Diagnostics
Imports System.Runtime.InteropServices
Imports System.IO.Pipes
Imports System.Threading
Imports XactCopy.Ipc
Imports XactCopy.Ipc.Messages
Imports XactCopy.Models
Imports XactCopy.Services

Module Program
    Private ReadOnly SendLock As New SemaphoreSlim(1, 1)

    Private _activeJobId As String = String.Empty
    Private _lastProgressUtc As DateTimeOffset = DateTimeOffset.UtcNow
    Private _jobTask As Task = Task.CompletedTask
    Private _jobCancellation As CancellationTokenSource
    Private _jobExecutionControl As CopyExecutionControl
    Private _jobPaused As Boolean
    Private _shutdownRequested As Boolean

    Sub Main(args As String())
        Dim exitCode = MainAsync(args).GetAwaiter().GetResult()
        Environment.ExitCode = exitCode
    End Sub

    Private Async Function MainAsync(args As String()) As Task(Of Integer)
        Try
            Dim pipeName = ReadPipeName(args)
            If String.IsNullOrWhiteSpace(pipeName) Then
                NotifyManualLaunch()
                Return 2
            End If

            Await RunAsync(pipeName).ConfigureAwait(False)
            Return 0
        Catch ex As Exception
            Try
                Console.Error.WriteLine($"XactCopyExecutive fatal error: {ex}")
            Catch
            End Try
            Return 1
        End Try
    End Function

    Private Async Function RunAsync(pipeName As String) As Task
        Using lifetimeCts As New CancellationTokenSource()
            Using server As New NamedPipeServerStream(
                pipeName,
                PipeDirection.InOut,
                1,
                PipeTransmissionMode.Byte,
                PipeOptions.Asynchronous)

                Await server.WaitForConnectionAsync(lifetimeCts.Token).ConfigureAwait(False)
                Await SendLogAsync(server, String.Empty, "Worker connected to supervisor.", lifetimeCts.Token).ConfigureAwait(False)

                Dim heartbeatTask = RunHeartbeatLoopAsync(server, lifetimeCts.Token)
                Try
                    While server.IsConnected AndAlso Not lifetimeCts.IsCancellationRequested
                        Dim incoming = Await JsonMessagePipe.ReadMessageAsync(server, lifetimeCts.Token).ConfigureAwait(False)
                        If incoming Is Nothing Then
                            Exit While
                        End If

                        Await HandleIncomingMessageAsync(incoming, server, lifetimeCts.Token).ConfigureAwait(False)
                        If _shutdownRequested Then
                            Exit While
                        End If
                    End While
                Finally
                    lifetimeCts.Cancel()
                End Try

                Await SafeCancelJobAsync().ConfigureAwait(False)
                Await heartbeatTask.ConfigureAwait(False)
            End Using
        End Using
    End Function

    Private Sub NotifyManualLaunch()
        Dim message = "XactCopyExecutive is a background worker and must be started by XactCopy." &
                      Environment.NewLine &
                      "Use XactCopy.exe to start copy jobs."
        Try
            Console.WriteLine(message)
        Catch
        End Try

        Try
            MessageBoxW(IntPtr.Zero, message, "XactCopyExecutive", &H40UI)
        Catch
            ' Ignore message box failures.
        End Try
    End Sub

    Private Function ReadPipeName(args As String()) As String
        For index = 0 To args.Length - 2
            If StringComparer.OrdinalIgnoreCase.Equals(args(index), "--pipe") Then
                Return args(index + 1)
            End If
        Next

        Return String.Empty
    End Function

    Private Async Function HandleIncomingMessageAsync(
        incomingJson As String,
        stream As NamedPipeServerStream,
        cancellationToken As CancellationToken) As Task

        Dim messageType As String = String.Empty
        If Not IpcSerializer.TryReadMessageType(incomingJson, messageType) Then
            Await SendLogAsync(stream, _activeJobId, "Received malformed IPC message.", cancellationToken).ConfigureAwait(False)
            Return
        End If

        Select Case messageType
            Case IpcMessageTypes.StartJobCommand
                Dim envelope = IpcSerializer.DeserializeEnvelope(Of StartJobCommand)(incomingJson)
                Await StartJobAsync(envelope.Payload, stream, cancellationToken).ConfigureAwait(False)

            Case IpcMessageTypes.CancelJobCommand
                Dim envelope = IpcSerializer.DeserializeEnvelope(Of CancelJobCommand)(incomingJson)
                Await CancelJobAsync(envelope.Payload, stream, cancellationToken).ConfigureAwait(False)

            Case IpcMessageTypes.PauseJobCommand
                Dim envelope = IpcSerializer.DeserializeEnvelope(Of PauseJobCommand)(incomingJson)
                Await PauseJobAsync(envelope.Payload, stream, cancellationToken).ConfigureAwait(False)

            Case IpcMessageTypes.ResumeJobCommand
                Dim envelope = IpcSerializer.DeserializeEnvelope(Of ResumeJobCommand)(incomingJson)
                Await ResumeJobAsync(envelope.Payload, stream, cancellationToken).ConfigureAwait(False)

            Case IpcMessageTypes.PingCommand
                Await SendHeartbeatAsync(stream, cancellationToken).ConfigureAwait(False)

            Case IpcMessageTypes.ShutdownCommand
                _shutdownRequested = True
                Await SendLogAsync(stream, _activeJobId, "Shutdown command received.", cancellationToken).ConfigureAwait(False)

            Case Else
                Await SendLogAsync(stream, _activeJobId, $"Unsupported message type: {messageType}", cancellationToken).ConfigureAwait(False)
        End Select
    End Function

    Private Async Function StartJobAsync(
        command As StartJobCommand,
        stream As NamedPipeServerStream,
        cancellationToken As CancellationToken) As Task

        If _jobTask IsNot Nothing AndAlso Not _jobTask.IsCompleted Then
            Await SendLogAsync(stream, _activeJobId, "Job start ignored because another job is already running.", cancellationToken).ConfigureAwait(False)
            Return
        End If

        _activeJobId = command.JobId
        _lastProgressUtc = DateTimeOffset.UtcNow
        _jobCancellation = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken)
        _jobExecutionControl = New CopyExecutionControl()
        _jobPaused = False

        Dim copyService As New ResilientCopyService(command.Options, _jobExecutionControl)

        AddHandler copyService.ProgressChanged,
            Sub(sender, snapshot)
                _lastProgressUtc = DateTimeOffset.UtcNow
                Dim progressEventData As New WorkerProgressEvent() With {
                    .JobId = _activeJobId,
                    .Snapshot = snapshot,
                    .TimestampUtc = DateTimeOffset.UtcNow
                }
                Dim sendTask = SendEnvelopeAsync(stream, IpcMessageTypes.WorkerProgressEvent, progressEventData, _jobCancellation.Token)
                Dim continuationTask = sendTask.ContinueWith(Sub(task) HandleBackgroundSendFailure(task), TaskScheduler.Default)
                Dim unused = continuationTask
            End Sub

        AddHandler copyService.LogMessage,
            Sub(sender, message)
                Dim logEventData As New WorkerLogEvent() With {
                    .JobId = _activeJobId,
                    .Message = message,
                    .TimestampUtc = DateTimeOffset.UtcNow
                }
                Dim sendTask = SendEnvelopeAsync(stream, IpcMessageTypes.WorkerLogEvent, logEventData, _jobCancellation.Token)
                Dim continuationTask = sendTask.ContinueWith(Sub(task) HandleBackgroundSendFailure(task), TaskScheduler.Default)
                Dim unused = continuationTask
            End Sub

        _jobTask = Task.Run(
            Async Function() As Task
                Dim resultEventData As WorkerJobResultEvent = Nothing
                Dim fatalEventData As WorkerFatalEvent = Nothing

                Try
                    Dim result = Await copyService.RunAsync(_jobCancellation.Token).ConfigureAwait(False)
                    If _jobCancellation.IsCancellationRequested Then
                        result.Cancelled = True
                    End If

                    resultEventData = New WorkerJobResultEvent() With {
                        .JobId = _activeJobId,
                        .Result = result,
                        .TimestampUtc = DateTimeOffset.UtcNow
                    }
                Catch ex As OperationCanceledException
                    Dim result As New CopyJobResult() With {
                        .Succeeded = False,
                        .Cancelled = True,
                        .ErrorMessage = "Job cancelled by supervisor."
                    }

                    resultEventData = New WorkerJobResultEvent() With {
                        .JobId = _activeJobId,
                        .Result = result,
                        .TimestampUtc = DateTimeOffset.UtcNow
                    }
                Catch ex As Exception
                    fatalEventData = New WorkerFatalEvent() With {
                        .JobId = _activeJobId,
                        .ErrorMessage = ex.Message,
                        .StackTraceText = ex.ToString(),
                        .TimestampUtc = DateTimeOffset.UtcNow
                    }
                Finally
                    _jobExecutionControl = Nothing
                    _jobPaused = False
                    _activeJobId = String.Empty
                End Try

                If resultEventData IsNot Nothing Then
                    Await SendEnvelopeAsync(stream, IpcMessageTypes.WorkerJobResultEvent, resultEventData, CancellationToken.None).ConfigureAwait(False)
                End If

                If fatalEventData IsNot Nothing Then
                    Await SendEnvelopeAsync(stream, IpcMessageTypes.WorkerFatalEvent, fatalEventData, CancellationToken.None).ConfigureAwait(False)
                End If
            End Function)

        Await SendLogAsync(stream, _activeJobId, "Job accepted by worker.", cancellationToken).ConfigureAwait(False)
    End Function

    Private Async Function CancelJobAsync(
        command As CancelJobCommand,
        stream As NamedPipeServerStream,
        cancellationToken As CancellationToken) As Task

        If _jobCancellation Is Nothing OrElse _jobTask Is Nothing OrElse _jobTask.IsCompleted Then
            Await SendLogAsync(stream, command.JobId, "Cancel command ignored because no active job is running.", cancellationToken).ConfigureAwait(False)
            Return
        End If

        _jobExecutionControl?.Resume()
        _jobPaused = False
        _jobCancellation.Cancel()
        Await SendLogAsync(stream, command.JobId, $"Cancel requested: {command.Reason}", cancellationToken).ConfigureAwait(False)
    End Function

    Private Async Function PauseJobAsync(
        command As PauseJobCommand,
        stream As NamedPipeServerStream,
        cancellationToken As CancellationToken) As Task

        If _jobExecutionControl Is Nothing OrElse _jobTask Is Nothing OrElse _jobTask.IsCompleted Then
            Await SendLogAsync(stream, command.JobId, "Pause command ignored because no active job is running.", cancellationToken).ConfigureAwait(False)
            Return
        End If

        _jobExecutionControl.Pause()
        _jobPaused = True
        Await SendLogAsync(stream, command.JobId, $"Pause requested: {command.Reason}", cancellationToken).ConfigureAwait(False)
    End Function

    Private Async Function ResumeJobAsync(
        command As ResumeJobCommand,
        stream As NamedPipeServerStream,
        cancellationToken As CancellationToken) As Task

        If _jobExecutionControl Is Nothing OrElse _jobTask Is Nothing OrElse _jobTask.IsCompleted Then
            Await SendLogAsync(stream, command.JobId, "Resume command ignored because no active job is running.", cancellationToken).ConfigureAwait(False)
            Return
        End If

        _jobExecutionControl.Resume()
        _jobPaused = False
        Await SendLogAsync(stream, command.JobId, $"Resume requested: {command.Reason}", cancellationToken).ConfigureAwait(False)
    End Function

    Private Async Function SafeCancelJobAsync() As Task
        Try
            If _jobCancellation IsNot Nothing Then
                _jobExecutionControl?.Resume()
                _jobPaused = False
                _jobCancellation.Cancel()
            End If

            If _jobTask IsNot Nothing Then
                Await _jobTask.ConfigureAwait(False)
            End If
        Catch
            ' Suppress worker shutdown errors.
        End Try
    End Function

    Private Async Function RunHeartbeatLoopAsync(
        stream As NamedPipeServerStream,
        cancellationToken As CancellationToken) As Task

        While Not cancellationToken.IsCancellationRequested
            Try
                Await SendHeartbeatAsync(stream, cancellationToken).ConfigureAwait(False)
                Await Task.Delay(TimeSpan.FromSeconds(1), cancellationToken).ConfigureAwait(False)
            Catch ex As OperationCanceledException
                Exit While
            Catch
                Exit While
            End Try
        End While
    End Function

    Private Async Function SendHeartbeatAsync(stream As NamedPipeServerStream, cancellationToken As CancellationToken) As Task
        Dim heartbeat As New WorkerHeartbeatEvent() With {
            .WorkerProcessId = Process.GetCurrentProcess().Id,
            .IsJobRunning = (_jobTask IsNot Nothing AndAlso Not _jobTask.IsCompleted),
            .IsJobPaused = _jobPaused,
            .ActiveJobId = _activeJobId,
            .LastProgressUtc = _lastProgressUtc,
            .TimestampUtc = DateTimeOffset.UtcNow
        }

        Await SendEnvelopeAsync(stream, IpcMessageTypes.WorkerHeartbeatEvent, heartbeat, cancellationToken).ConfigureAwait(False)
    End Function

    Private Async Function SendLogAsync(
        stream As NamedPipeServerStream,
        jobId As String,
        message As String,
        cancellationToken As CancellationToken) As Task

        Dim logEventData As New WorkerLogEvent() With {
            .JobId = jobId,
            .Message = message,
            .TimestampUtc = DateTimeOffset.UtcNow
        }

        Await SendEnvelopeAsync(stream, IpcMessageTypes.WorkerLogEvent, logEventData, cancellationToken).ConfigureAwait(False)
    End Function

    Private Async Function SendEnvelopeAsync(
        stream As NamedPipeServerStream,
        messageType As String,
        payload As Object,
        cancellationToken As CancellationToken) As Task

        If stream Is Nothing OrElse Not stream.IsConnected Then
            Return
        End If

        Dim json = IpcSerializer.SerializeEnvelope(messageType, payload)
        Await SendLock.WaitAsync(cancellationToken).ConfigureAwait(False)
        Try
            Await JsonMessagePipe.WriteMessageAsync(stream, json, cancellationToken).ConfigureAwait(False)
        Finally
            SendLock.Release()
        End Try
    End Function

    Private Sub HandleBackgroundSendFailure(task As Task)
        If task Is Nothing OrElse task.IsCompletedSuccessfully Then
            Return
        End If

        Dim ignoredException = task.Exception
    End Sub

    <DllImport("user32.dll", CharSet:=CharSet.Unicode, SetLastError:=True)>
    Private Function MessageBoxW(hWnd As IntPtr, lpText As String, lpCaption As String, uType As UInteger) As Integer
    End Function
End Module
