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
    Private ReadOnly ProgressDispatchLock As New Object()
    Private ReadOnly LogRateLock As New Object()

    Private _activeJobId As String = String.Empty
    Private _lastProgressUtc As DateTimeOffset = DateTimeOffset.UtcNow
    Private _jobTask As Task = Task.CompletedTask
    Private _jobCancellation As CancellationTokenSource
    Private _jobExecutionControl As CopyExecutionControl
    Private _jobPaused As Boolean
    Private _shutdownRequested As Boolean
    Private _runtimeTelemetryProfile As WorkerTelemetryProfile = WorkerTelemetryProfile.Normal
    Private _runtimeProgressIntervalMs As Integer = 75
    Private _runtimeMaxLogsPerSecond As Integer = 100
    Private _progressDispatchScheduled As Integer
    Private _lastProgressSentTick As Long
    Private _pendingProgressSnapshot As CopyProgressSnapshot
    Private _progressReceivedCount As Long
    Private _progressSentCount As Long
    Private _progressCoalescedCount As Long
    Private _logWindowStartUtc As DateTimeOffset = DateTimeOffset.MinValue
    Private _logWindowSentCount As Integer
    Private _suppressedLogCount As Integer
    Private _suppressedLogTotal As Long
    Private _logSentCount As Long

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
        ConfigureRuntimeTelemetryPolicy(command.Options)

        Dim copyService As New ResilientCopyService(command.Options, _jobExecutionControl)

        AddHandler copyService.ProgressChanged,
            Sub(sender, snapshot)
                _lastProgressUtc = DateTimeOffset.UtcNow
                QueueProgressSnapshot(stream, snapshot)
            End Sub

        AddHandler copyService.LogMessage,
            Sub(sender, message)
                Dim sendTask = SendRuntimeLogAsync(stream, _activeJobId, message, _jobCancellation.Token)
                Dim continuationTask = sendTask.ContinueWith(Sub(task) HandleBackgroundSendFailure(task), TaskScheduler.Default)
                Dim unused = continuationTask
            End Sub

        _jobTask = Task.Run(
            Async Function() As Task
                Dim resultEventData As WorkerJobResultEvent = Nothing
                Dim fatalEventData As WorkerFatalEvent = Nothing
                Dim shouldFlushTelemetry = False
                Dim runJobId = _activeJobId

                Try
                    Dim result = Await copyService.RunAsync(_jobCancellation.Token).ConfigureAwait(False)
                    If _jobCancellation.IsCancellationRequested Then
                        result.Cancelled = True
                    End If

                    shouldFlushTelemetry = True

                    resultEventData = New WorkerJobResultEvent() With {
                        .JobId = runJobId,
                        .Result = result,
                        .TimestampUtc = DateTimeOffset.UtcNow
                    }
                Catch ex As OperationCanceledException
                    Dim result As New CopyJobResult() With {
                        .Succeeded = False,
                        .Cancelled = True,
                        .ErrorMessage = "Job cancelled by supervisor."
                    }

                    shouldFlushTelemetry = True

                    resultEventData = New WorkerJobResultEvent() With {
                        .JobId = runJobId,
                        .Result = result,
                        .TimestampUtc = DateTimeOffset.UtcNow
                    }
                Catch ex As Exception
                    shouldFlushTelemetry = True

                    fatalEventData = New WorkerFatalEvent() With {
                        .JobId = runJobId,
                        .ErrorMessage = ex.Message,
                        .StackTraceText = ex.ToString(),
                        .TimestampUtc = DateTimeOffset.UtcNow
                    }
                Finally
                    _jobExecutionControl = Nothing
                    _jobPaused = False
                    _activeJobId = String.Empty
                End Try

                If shouldFlushTelemetry Then
                    Try
                        Await FlushProgressAndTelemetryAsync(stream, runJobId).ConfigureAwait(False)
                    Catch
                    End Try
                End If

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

    Private Sub ConfigureRuntimeTelemetryPolicy(options As CopyJobOptions)
        Dim requestedProfile = WorkerTelemetryProfile.Normal
        Dim requestedProgressIntervalMs = 0
        Dim requestedMaxLogsPerSecond = -1

        If options IsNot Nothing Then
            requestedProfile = options.WorkerTelemetryProfile
            requestedProgressIntervalMs = options.WorkerProgressEmitIntervalMs
            requestedMaxLogsPerSecond = options.WorkerMaxLogsPerSecond
        End If

        Dim defaultProgressIntervalMs = 75
        Dim defaultMaxLogsPerSecond = 100
        ResolveTelemetryDefaults(requestedProfile, defaultProgressIntervalMs, defaultMaxLogsPerSecond)

        _runtimeTelemetryProfile = requestedProfile
        Dim baseProgressIntervalMs = Math.Max(20, Math.Min(1000, If(requestedProgressIntervalMs > 0, requestedProgressIntervalMs, defaultProgressIntervalMs)))
        Dim baseMaxLogsPerSecond = If(
            requestedMaxLogsPerSecond >= 0,
            Math.Min(5000, requestedMaxLogsPerSecond),
            defaultMaxLogsPerSecond)

        Select Case _runtimeTelemetryProfile
            Case WorkerTelemetryProfile.Verbose
                _runtimeProgressIntervalMs = Math.Max(20, CInt(Math.Round(baseProgressIntervalMs * 0.5R)))
                _runtimeMaxLogsPerSecond = If(baseMaxLogsPerSecond <= 0, 0, Math.Min(5000, baseMaxLogsPerSecond * 3))
            Case WorkerTelemetryProfile.Debug
                _runtimeProgressIntervalMs = Math.Max(20, CInt(Math.Round(baseProgressIntervalMs * 0.33R)))
                _runtimeMaxLogsPerSecond = 0
            Case Else
                _runtimeProgressIntervalMs = baseProgressIntervalMs
                _runtimeMaxLogsPerSecond = baseMaxLogsPerSecond
        End Select

        SyncLock ProgressDispatchLock
            _pendingProgressSnapshot = Nothing
        End SyncLock

        SyncLock LogRateLock
            _logWindowStartUtc = DateTimeOffset.MinValue
            _logWindowSentCount = 0
            _suppressedLogCount = 0
        End SyncLock

        Interlocked.Exchange(_progressDispatchScheduled, 0)
        _lastProgressSentTick = 0
        Interlocked.Exchange(_progressReceivedCount, 0)
        Interlocked.Exchange(_progressSentCount, 0)
        Interlocked.Exchange(_progressCoalescedCount, 0)
        Interlocked.Exchange(_logSentCount, 0)
        Interlocked.Exchange(_suppressedLogTotal, 0)
    End Sub

    Private Sub ResolveTelemetryDefaults(
        profile As WorkerTelemetryProfile,
        ByRef progressIntervalMs As Integer,
        ByRef maxLogsPerSecond As Integer)

        Select Case profile
            Case WorkerTelemetryProfile.Verbose
                progressIntervalMs = 40
                maxLogsPerSecond = 300
            Case WorkerTelemetryProfile.Debug
                progressIntervalMs = 20
                maxLogsPerSecond = 0
            Case Else
                progressIntervalMs = 75
                maxLogsPerSecond = 100
        End Select
    End Sub

    Private Sub QueueProgressSnapshot(stream As NamedPipeServerStream, snapshot As CopyProgressSnapshot)
        If snapshot Is Nothing Then
            Return
        End If

        Interlocked.Increment(_progressReceivedCount)

        SyncLock ProgressDispatchLock
            If _pendingProgressSnapshot IsNot Nothing Then
                Interlocked.Increment(_progressCoalescedCount)
            End If
            _pendingProgressSnapshot = snapshot
        End SyncLock

        If Interlocked.CompareExchange(_progressDispatchScheduled, 1, 0) <> 0 Then
            Return
        End If

        ScheduleProgressDrain(stream, 0)
    End Sub

    Private Sub ScheduleProgressDrain(stream As NamedPipeServerStream, delayMs As Integer)
        Dim safeDelayMs = Math.Max(0, delayMs)
        Dim dispatchToken As CancellationToken = If(_jobCancellation Is Nothing, CancellationToken.None, _jobCancellation.Token)

        Dim dispatchTask = Task.Run(
            Async Function() As Task
                Try
                    If safeDelayMs > 0 Then
                        Await Task.Delay(safeDelayMs, dispatchToken).ConfigureAwait(False)
                    End If

                    Await DrainPendingProgressAsync(stream, dispatchToken).ConfigureAwait(False)
                Catch ex As OperationCanceledException
                Catch
                Finally
                    Interlocked.Exchange(_progressDispatchScheduled, 0)

                    Dim hasPending As Boolean
                    SyncLock ProgressDispatchLock
                        hasPending = _pendingProgressSnapshot IsNot Nothing
                    End SyncLock

                    If hasPending AndAlso Interlocked.CompareExchange(_progressDispatchScheduled, 1, 0) = 0 Then
                        ScheduleProgressDrain(stream, 1)
                    End If
                End Try
            End Function)

        Dim continuationTask = dispatchTask.ContinueWith(Sub(task) HandleBackgroundSendFailure(task), TaskScheduler.Default)
        Dim unused = continuationTask
    End Sub

    Private Async Function DrainPendingProgressAsync(
        stream As NamedPipeServerStream,
        cancellationToken As CancellationToken) As Task

        Do
            Dim snapshot As CopyProgressSnapshot = Nothing
            SyncLock ProgressDispatchLock
                snapshot = _pendingProgressSnapshot
                _pendingProgressSnapshot = Nothing
            End SyncLock

            If snapshot Is Nothing Then
                Exit Do
            End If

            If _lastProgressSentTick > 0 Then
                Dim elapsed = Environment.TickCount64 - _lastProgressSentTick
                If elapsed < _runtimeProgressIntervalMs Then
                    Dim waitMs = CInt(Math.Max(1L, _runtimeProgressIntervalMs - elapsed))
                    SyncLock ProgressDispatchLock
                        If _pendingProgressSnapshot Is Nothing Then
                            _pendingProgressSnapshot = snapshot
                        Else
                            Interlocked.Increment(_progressCoalescedCount)
                        End If
                    End SyncLock

                    Await Task.Delay(waitMs, cancellationToken).ConfigureAwait(False)
                    Continue Do
                End If
            End If

            Await SendProgressSnapshotAsync(stream, snapshot, cancellationToken).ConfigureAwait(False)
            _lastProgressSentTick = Environment.TickCount64
        Loop
    End Function

    Private Async Function SendProgressSnapshotAsync(
        stream As NamedPipeServerStream,
        snapshot As CopyProgressSnapshot,
        cancellationToken As CancellationToken) As Task

        If snapshot Is Nothing Then
            Return
        End If

        Dim progressEventData As New WorkerProgressEvent() With {
            .JobId = _activeJobId,
            .Snapshot = snapshot,
            .TimestampUtc = DateTimeOffset.UtcNow
        }

        Await SendEnvelopeAsync(stream, IpcMessageTypes.WorkerProgressEvent, progressEventData, cancellationToken).ConfigureAwait(False)
        Interlocked.Increment(_progressSentCount)
    End Function

    Private Async Function SendRuntimeLogAsync(
        stream As NamedPipeServerStream,
        jobId As String,
        message As String,
        cancellationToken As CancellationToken) As Task

        If String.IsNullOrWhiteSpace(message) Then
            Return
        End If

        Dim suppressedSummaryCount = 0
        Dim shouldSend = False
        Dim nowUtc = DateTimeOffset.UtcNow

        SyncLock LogRateLock
            If _logWindowStartUtc = DateTimeOffset.MinValue OrElse (nowUtc - _logWindowStartUtc).TotalSeconds >= 1.0R Then
                suppressedSummaryCount = _suppressedLogCount
                _suppressedLogCount = 0
                _logWindowStartUtc = nowUtc
                _logWindowSentCount = 0
            End If

            If _runtimeMaxLogsPerSecond <= 0 OrElse _logWindowSentCount < _runtimeMaxLogsPerSecond Then
                _logWindowSentCount += 1
                shouldSend = True
            Else
                _suppressedLogCount += 1
                _suppressedLogTotal += 1
            End If
        End SyncLock

        If suppressedSummaryCount > 0 Then
            Await SendLogAsync(
                stream,
                jobId,
                $"[WorkerTelemetry] Suppressed {suppressedSummaryCount} log message(s) due to telemetry rate policy.",
                cancellationToken).ConfigureAwait(False)
        End If

        If Not shouldSend Then
            Return
        End If

        Await SendLogAsync(stream, jobId, message, cancellationToken).ConfigureAwait(False)
    End Function

    Private Async Function FlushProgressAndTelemetryAsync(stream As NamedPipeServerStream, jobId As String) As Task
        Dim finalSnapshot As CopyProgressSnapshot = Nothing
        SyncLock ProgressDispatchLock
            finalSnapshot = _pendingProgressSnapshot
            _pendingProgressSnapshot = Nothing
        End SyncLock

        If finalSnapshot IsNot Nothing Then
            Try
                Await SendProgressSnapshotAsync(stream, finalSnapshot, CancellationToken.None).ConfigureAwait(False)
            Catch
            End Try
        End If

        Try
            Await FlushSuppressedLogsAsync(stream, jobId).ConfigureAwait(False)
        Catch
        End Try

        Try
            Await SendLogAsync(stream, jobId, BuildTelemetrySummaryMessage(), CancellationToken.None).ConfigureAwait(False)
        Catch
        End Try
    End Function

    Private Async Function FlushSuppressedLogsAsync(stream As NamedPipeServerStream, jobId As String) As Task
        Dim suppressedCount = 0
        SyncLock LogRateLock
            suppressedCount = _suppressedLogCount
            _suppressedLogCount = 0
            _logWindowStartUtc = DateTimeOffset.MinValue
            _logWindowSentCount = 0
        End SyncLock

        If suppressedCount <= 0 Then
            Return
        End If

        Await SendLogAsync(
            stream,
            jobId,
            $"[WorkerTelemetry] Suppressed {suppressedCount} log message(s) due to telemetry rate policy.",
            CancellationToken.None).ConfigureAwait(False)
    End Function

    Private Function BuildTelemetrySummaryMessage() As String
        Dim received = Interlocked.Read(_progressReceivedCount)
        Dim emitted = Interlocked.Read(_progressSentCount)
        Dim coalesced = Interlocked.Read(_progressCoalescedCount)
        Dim logsSent = Interlocked.Read(_logSentCount)
        Dim logsSuppressed = Interlocked.Read(_suppressedLogTotal)

        Return $"[WorkerTelemetry] Profile={_runtimeTelemetryProfile}, progress interval={_runtimeProgressIntervalMs} ms, log cap={If(_runtimeMaxLogsPerSecond <= 0, "unlimited", _runtimeMaxLogsPerSecond & "/sec")}; progress recv={received}, sent={emitted}, coalesced={coalesced}; logs sent={logsSent}, suppressed={logsSuppressed}."
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
        Interlocked.Increment(_logSentCount)
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
