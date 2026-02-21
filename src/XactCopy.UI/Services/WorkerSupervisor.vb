Imports System.Diagnostics
Imports System.Collections.Generic
Imports System.IO
Imports System.IO.Pipes
Imports System.Linq
Imports System.Threading
Imports XactCopy.Ipc
Imports XactCopy.Ipc.Messages
Imports XactCopy.Models

Namespace Services
    ''' <summary>
    ''' Class WorkerSupervisor.
    ''' </summary>
    Public Class WorkerSupervisor
        Implements IDisposable

        Public Event LogMessage As EventHandler(Of String)
        Public Event ProgressChanged As EventHandler(Of CopyProgressSnapshot)
        Public Event JobCompleted As EventHandler(Of CopyJobResult)
        Public Event WorkerStateChanged As EventHandler(Of String)
        Public Event JobPauseStateChanged As EventHandler(Of Boolean)

        Private Shared ReadOnly HeartbeatCheckInterval As TimeSpan = TimeSpan.FromSeconds(2)
        Private Shared ReadOnly HeartbeatTimeout As TimeSpan = TimeSpan.FromSeconds(10)
        Private Shared ReadOnly MinimumJobActivityTimeout As TimeSpan = TimeSpan.FromSeconds(45)
        Private Shared ReadOnly MaximumJobActivityTimeout As TimeSpan = TimeSpan.FromMinutes(10)

        Private ReadOnly _restartLock As New SemaphoreSlim(1, 1)
        Private ReadOnly _sendLock As New SemaphoreSlim(1, 1)
        Private ReadOnly _disposeLock As New SemaphoreSlim(1, 1)
        Private ReadOnly _workerPidLock As New Object()
        Private ReadOnly _launchedWorkerPids As New HashSet(Of Integer)()

        Private _workerProcess As Process
        Private _pipeClient As NamedPipeClientStream
        Private _pipeCancellation As CancellationTokenSource
        Private _receiveLoopTask As Task = Task.CompletedTask
        Private _heartbeatTimer As Timer
        Private _lastHeartbeatUtc As DateTimeOffset = DateTimeOffset.MinValue
        Private _lastJobActivityUtc As DateTimeOffset = DateTimeOffset.MinValue
        Private _lastWorkerProgressUtc As DateTimeOffset = DateTimeOffset.MinValue

        Private _currentJobOptions As CopyJobOptions
        Private _currentJobId As String = String.Empty
        Private _isJobRunning As Boolean
        Private _isJobPaused As Boolean
        Private _cancelRequestedByUser As Boolean
        Private _autoRecoverJob As Boolean
        Private _isStopping As Boolean
        Private _isDisposed As Boolean
        Private _recoveryInFlight As Integer
        Private _consecutiveMalformedMessages As Integer

        ''' <summary>
        ''' Gets or sets IsJobRunning.
        ''' </summary>
        Public ReadOnly Property IsJobRunning As Boolean
            Get
                Return _isJobRunning
            End Get
        End Property

        ''' <summary>
        ''' Gets or sets IsJobPaused.
        ''' </summary>
        Public ReadOnly Property IsJobPaused As Boolean
            Get
                Return _isJobPaused
            End Get
        End Property

        ''' <summary>
        ''' Computes StartJobAsync.
        ''' </summary>
        Public Async Function StartJobAsync(options As CopyJobOptions, cancellationToken As CancellationToken) As Task
            ThrowIfDisposed()

            Await _restartLock.WaitAsync(cancellationToken).ConfigureAwait(False)
            Try
                If _isStopping Then
                    Throw New InvalidOperationException("Supervisor is stopping.")
                End If

                If _isJobRunning Then
                    Throw New InvalidOperationException("A copy job is already running.")
                End If

                _currentJobOptions = CloneOptions(options)
                _currentJobId = Guid.NewGuid().ToString("N")
                _isJobRunning = True
                _cancelRequestedByUser = False
                SetPausedState(False)
                _autoRecoverJob = True
                _consecutiveMalformedMessages = 0
                MarkJobActivity(markProgress:=True)

                Await EnsureWorkerConnectedAsync(cancellationToken).ConfigureAwait(False)
                Await SendStartJobCommandAsync(cancellationToken).ConfigureAwait(False)
                RaiseEvent WorkerStateChanged(Me, "Job running")
            Finally
                _restartLock.Release()
            End Try
        End Function

        ''' <summary>
        ''' Computes CancelJobAsync.
        ''' </summary>
        Public Async Function CancelJobAsync(reason As String) As Task
            ThrowIfDisposed()

            If Not _isJobRunning Then
                Return
            End If

            _autoRecoverJob = False
            _cancelRequestedByUser = True
            SetPausedState(False)
            Dim command As New CancelJobCommand() With {
                .JobId = _currentJobId,
                .Reason = If(String.IsNullOrWhiteSpace(reason), "User requested cancellation.", reason)
            }

            Await TrySendCommandAsync(IpcMessageTypes.CancelJobCommand, command, CancellationToken.None).ConfigureAwait(False)
        End Function

        ''' <summary>
        ''' Computes PauseJobAsync.
        ''' </summary>
        Public Async Function PauseJobAsync(reason As String) As Task
            ThrowIfDisposed()

            If Not _isJobRunning OrElse _isJobPaused Then
                Return
            End If

            Dim command As New PauseJobCommand() With {
                .JobId = _currentJobId,
                .Reason = If(String.IsNullOrWhiteSpace(reason), "User requested pause.", reason)
            }

            Await TrySendCommandAsync(IpcMessageTypes.PauseJobCommand, command, CancellationToken.None).ConfigureAwait(False)
            SetPausedState(True)
            RaiseEvent WorkerStateChanged(Me, "Job paused")
        End Function

        ''' <summary>
        ''' Computes ResumeJobAsync.
        ''' </summary>
        Public Async Function ResumeJobAsync(reason As String) As Task
            ThrowIfDisposed()

            If Not _isJobRunning OrElse Not _isJobPaused Then
                Return
            End If

            Dim command As New ResumeJobCommand() With {
                .JobId = _currentJobId,
                .Reason = If(String.IsNullOrWhiteSpace(reason), "User requested resume.", reason)
            }

            Await TrySendCommandAsync(IpcMessageTypes.ResumeJobCommand, command, CancellationToken.None).ConfigureAwait(False)
            SetPausedState(False)
            RaiseEvent WorkerStateChanged(Me, "Job running")
        End Function

        ''' <summary>
        ''' Computes StopAsync.
        ''' </summary>
        Public Async Function StopAsync() As Task
            If _isDisposed Then
                Return
            End If

            Await _disposeLock.WaitAsync().ConfigureAwait(False)
            Try
                If _isDisposed Then
                    Return
                End If

                _isStopping = True
                _autoRecoverJob = False
                _isJobRunning = False
                _cancelRequestedByUser = False
                SetPausedState(False)
                _currentJobId = String.Empty
                _currentJobOptions = Nothing
                _consecutiveMalformedMessages = 0

                Dim shutdown As New ShutdownCommand() With {
                    .Reason = "UI shutdown requested."
                }
                Await TrySendCommandAsync(IpcMessageTypes.ShutdownCommand, shutdown, CancellationToken.None).ConfigureAwait(False)

                Dim shutdownClean = Await ShutdownWorkerAsync(killIfStillRunning:=True).ConfigureAwait(False)
                If Not shutdownClean Then
                    RaiseEvent LogMessage(Me, "[Supervisor] One or more worker processes did not terminate during shutdown.")
                End If

                If _heartbeatTimer IsNot Nothing Then
                    Await _heartbeatTimer.DisposeAsync().ConfigureAwait(False)
                    _heartbeatTimer = Nothing
                End If

                _isDisposed = True
            Finally
                _disposeLock.Release()
            End Try
        End Function

        Private Async Function EnsureWorkerConnectedAsync(cancellationToken As CancellationToken) As Task
            If _pipeClient IsNot Nothing AndAlso _pipeClient.IsConnected Then
                Return
            End If

            Dim pipeName = $"xactcopy-{Guid.NewGuid():N}"
            Dim launch = ResolveWorkerLaunch(pipeName)
            Dim startInfo As New ProcessStartInfo(launch.FileName, launch.Arguments) With {
                .UseShellExecute = False,
                .CreateNoWindow = True,
                .WindowStyle = ProcessWindowStyle.Hidden
            }

            _workerProcess = Process.Start(startInfo)
            If _workerProcess Is Nothing Then
                Throw New InvalidOperationException("Failed to start worker process.")
            End If

            TrackWorkerProcess(_workerProcess.Id)
            _workerProcess.EnableRaisingEvents = True
            AddHandler _workerProcess.Exited, AddressOf WorkerProcess_Exited

            _pipeClient = New NamedPipeClientStream(
                ".",
                pipeName,
                PipeDirection.InOut,
                PipeOptions.Asynchronous)

            Await _pipeClient.ConnectAsync(CInt(TimeSpan.FromSeconds(15).TotalMilliseconds), cancellationToken).ConfigureAwait(False)

            _pipeCancellation = New CancellationTokenSource()
            _lastHeartbeatUtc = DateTimeOffset.UtcNow
            _receiveLoopTask = ReceiveLoopAsync(_pipeCancellation.Token)

            If _heartbeatTimer Is Nothing Then
                _heartbeatTimer = New Timer(AddressOf CheckHeartbeat, Nothing, HeartbeatCheckInterval, HeartbeatCheckInterval)
            End If

            RaiseEvent WorkerStateChanged(Me, $"Worker connected (PID {_workerProcess.Id})")
        End Function

        Private Async Function SendStartJobCommandAsync(cancellationToken As CancellationToken) As Task
            Dim command As New StartJobCommand() With {
                .JobId = _currentJobId,
                .Options = CloneOptions(_currentJobOptions)
            }

            Await TrySendCommandAsync(IpcMessageTypes.StartJobCommand, command, cancellationToken).ConfigureAwait(False)
        End Function

        Private Async Function TrySendCommandAsync(
            messageType As String,
            payload As Object,
            cancellationToken As CancellationToken) As Task

            If _pipeClient Is Nothing OrElse Not _pipeClient.IsConnected Then
                Return
            End If

            Dim json = IpcSerializer.SerializeEnvelope(messageType, payload)
            Await _sendLock.WaitAsync(cancellationToken).ConfigureAwait(False)
            Try
                Await JsonMessagePipe.WriteMessageAsync(_pipeClient, json, cancellationToken).ConfigureAwait(False)
            Catch ex As Exception
                RaiseEvent LogMessage(Me, $"[Supervisor] Send failed: {ex.Message}")
            Finally
                _sendLock.Release()
            End Try
        End Function

        Private Async Function ReceiveLoopAsync(cancellationToken As CancellationToken) As Task
            Try
                While Not cancellationToken.IsCancellationRequested
                    Dim rawMessage = Await JsonMessagePipe.ReadMessageAsync(_pipeClient, cancellationToken).ConfigureAwait(False)
                    If rawMessage Is Nothing Then
                        Exit While
                    End If

                    Await HandleIncomingMessageAsync(rawMessage).ConfigureAwait(False)
                End While
            Catch ex As OperationCanceledException
            Catch ex As Exception
                RaiseEvent LogMessage(Me, $"[Supervisor] Receive loop error: {ex.Message}")
            Finally
                If Not _isStopping Then
                    QueueRecovery("Worker disconnected.")
                End If
            End Try
        End Function

        Private Function HandleIncomingMessageAsync(rawMessage As String) As Task
            Dim messageType As String = String.Empty
            If Not IpcSerializer.TryReadMessageType(rawMessage, messageType) Then
                RaiseEvent LogMessage(Me, "[Supervisor] Received malformed message.")
                _consecutiveMalformedMessages += 1
                If _consecutiveMalformedMessages >= 3 AndAlso _isJobRunning Then
                    QueueRecovery("Worker IPC stream desynchronized (malformed messages).")
                End If
                Return Task.CompletedTask
            End If

            _consecutiveMalformedMessages = 0

            Select Case messageType
                Case IpcMessageTypes.WorkerHeartbeatEvent
                    Dim envelope = IpcSerializer.DeserializeEnvelope(Of WorkerHeartbeatEvent)(rawMessage)
                    _lastHeartbeatUtc = envelope.Payload.TimestampUtc
                    SetPausedState(envelope.Payload.IsJobPaused)
                    Dim reportedProgressUtc = NormalizeTimestampOrNow(envelope.Payload.LastProgressUtc, useNowWhenMinValue:=False)
                    If reportedProgressUtc <> DateTimeOffset.MinValue AndAlso reportedProgressUtc > _lastWorkerProgressUtc Then
                        _lastWorkerProgressUtc = reportedProgressUtc
                    End If

                Case IpcMessageTypes.WorkerProgressEvent
                    Dim envelope = IpcSerializer.DeserializeEnvelope(Of WorkerProgressEvent)(rawMessage)
                    _lastHeartbeatUtc = envelope.Payload.TimestampUtc
                    MarkJobActivity(envelope.Payload.TimestampUtc, markProgress:=True)
                    If Not _isJobRunning Then
                        Return Task.CompletedTask
                    End If

                    Dim messageJobId = If(envelope.Payload.JobId, String.Empty)
                    If Not String.IsNullOrWhiteSpace(_currentJobId) AndAlso
                        Not String.Equals(messageJobId, _currentJobId, StringComparison.OrdinalIgnoreCase) Then
                        Return Task.CompletedTask
                    End If

                    If envelope.Payload.Snapshot IsNot Nothing Then
                        RaiseEvent ProgressChanged(Me, envelope.Payload.Snapshot)
                    End If

                Case IpcMessageTypes.WorkerLogEvent
                    Dim envelope = IpcSerializer.DeserializeEnvelope(Of WorkerLogEvent)(rawMessage)
                    MarkJobActivity(envelope.Payload.TimestampUtc)
                    RaiseEvent LogMessage(Me, envelope.Payload.Message)

                Case IpcMessageTypes.WorkerJobResultEvent
                    Dim envelope = IpcSerializer.DeserializeEnvelope(Of WorkerJobResultEvent)(rawMessage)
                    _isJobRunning = False
                    SetPausedState(False)
                    _autoRecoverJob = False
                    _cancelRequestedByUser = False
                    _currentJobId = String.Empty
                    _currentJobOptions = Nothing
                    _lastJobActivityUtc = DateTimeOffset.MinValue
                    _lastWorkerProgressUtc = DateTimeOffset.MinValue
                    RaiseEvent JobCompleted(Me, envelope.Payload.Result)
                    RaiseEvent WorkerStateChanged(Me, "Job finished")

                Case IpcMessageTypes.WorkerFatalEvent
                    Dim envelope = IpcSerializer.DeserializeEnvelope(Of WorkerFatalEvent)(rawMessage)
                    Dim failureResult As New CopyJobResult() With {
                        .Succeeded = False,
                        .Cancelled = False,
                        .ErrorMessage = envelope.Payload.ErrorMessage
                    }

                    _isJobRunning = False
                    SetPausedState(False)
                    _autoRecoverJob = False
                    _cancelRequestedByUser = False
                    _currentJobId = String.Empty
                    _currentJobOptions = Nothing
                    _lastJobActivityUtc = DateTimeOffset.MinValue
                    _lastWorkerProgressUtc = DateTimeOffset.MinValue
                    RaiseEvent LogMessage(Me, $"[Worker Fatal] {envelope.Payload.ErrorMessage}")
                    RaiseEvent JobCompleted(Me, failureResult)
                    RaiseEvent WorkerStateChanged(Me, "Worker fatal error")
            End Select

            Return Task.CompletedTask
        End Function

        Private Sub CheckHeartbeat(state As Object)
            If _isStopping OrElse _isDisposed Then
                Return
            End If

            If _pipeClient Is Nothing OrElse Not _pipeClient.IsConnected Then
                Return
            End If

            Dim nowUtc = DateTimeOffset.UtcNow
            Dim elapsed = nowUtc - _lastHeartbeatUtc
            If elapsed <= HeartbeatTimeout Then
                CheckForJobActivityStall(nowUtc)
                Return
            End If

            QueueRecovery($"Worker heartbeat stale for {elapsed.TotalSeconds:0.0}s.")
        End Sub

        Private Sub WorkerProcess_Exited(sender As Object, e As EventArgs)
            Dim process = TryCast(sender, Process)
            If process IsNot Nothing Then
                UntrackWorkerProcess(process.Id)
            End If

            If _isStopping OrElse _isDisposed Then
                Return
            End If

            QueueRecovery("Worker process exited unexpectedly.")
        End Sub

        Private Sub QueueRecovery(reason As String)
            If Interlocked.Exchange(_recoveryInFlight, 1) = 1 Then
                Return
            End If

            Dim recoveryTask = Task.Run(
                Async Function() As Task
                    Try
                        Await RecoverWorkerAsync(reason).ConfigureAwait(False)
                    Finally
                        Interlocked.Exchange(_recoveryInFlight, 0)
                    End Try
                End Function)
            Dim ignoredTask = recoveryTask
        End Sub

        Private Async Function RecoverWorkerAsync(reason As String) As Task
            If Not _isJobRunning Then
                Return
            End If

            Await _restartLock.WaitAsync().ConfigureAwait(False)
            Try
                If _isStopping OrElse _isDisposed Then
                    Return
                End If

                If Not _autoRecoverJob Then
                    Dim shutdownClean = Await ShutdownWorkerAsync(killIfStillRunning:=True).ConfigureAwait(False)
                    Dim finalReason = If(
                        shutdownClean,
                        reason,
                        $"{reason} Worker process did not terminate cleanly.")
                    CompleteActiveJobAfterChannelLoss(finalReason)
                    Return
                End If

                RaiseEvent LogMessage(Me, $"[Supervisor] {reason} Restarting worker and resuming from journal.")

                Dim recoveredShutdown = Await ShutdownWorkerAsync(killIfStillRunning:=True).ConfigureAwait(False)
                If Not recoveredShutdown Then
                    _autoRecoverJob = False
                    CompleteActiveJobAfterChannelLoss(
                        $"{reason} Worker process did not terminate cleanly, so auto-recovery was aborted to prevent duplicate stuck workers.")
                    Return
                End If

                Await EnsureWorkerConnectedAsync(CancellationToken.None).ConfigureAwait(False)
                MarkJobActivity(markProgress:=True)

                If _autoRecoverJob AndAlso _isJobRunning AndAlso _currentJobOptions IsNot Nothing Then
                    If Not _currentJobOptions.ResumeFromJournal Then
                        _currentJobOptions.ResumeFromJournal = True
                        RaiseEvent LogMessage(Me, "[Supervisor] Forced resume-from-journal after worker recovery.")
                    End If

                    Await SendStartJobCommandAsync(CancellationToken.None).ConfigureAwait(False)
                    If _isJobPaused Then
                        Dim pauseCommand As New PauseJobCommand() With {
                            .JobId = _currentJobId,
                            .Reason = "Reapplying paused state after worker recovery."
                        }
                        Await TrySendCommandAsync(IpcMessageTypes.PauseJobCommand, pauseCommand, CancellationToken.None).ConfigureAwait(False)
                    End If
                End If
            Finally
                _restartLock.Release()
            End Try
        End Function

        Private Sub CompleteActiveJobAfterChannelLoss(reason As String)
            If Not _isJobRunning Then
                Return
            End If

            Dim normalizedReason = If(String.IsNullOrWhiteSpace(reason), "Worker channel was lost.", reason.Trim())
            Dim cancelled = _cancelRequestedByUser
            Dim message = If(
                cancelled,
                $"Cancellation was requested, but the worker channel closed before acknowledgement. {normalizedReason}",
                $"Worker channel lost: {normalizedReason}")

            Dim failureResult As New CopyJobResult() With {
                .Succeeded = False,
                .Cancelled = cancelled,
                .ErrorMessage = message
            }

            _isJobRunning = False
            SetPausedState(False)
            _autoRecoverJob = False
            _cancelRequestedByUser = False
            _currentJobId = String.Empty
            _currentJobOptions = Nothing
            _lastJobActivityUtc = DateTimeOffset.MinValue
            _lastWorkerProgressUtc = DateTimeOffset.MinValue
            _consecutiveMalformedMessages = 0

            RaiseEvent LogMessage(Me, $"[Supervisor] {message}")
            RaiseEvent JobCompleted(Me, failureResult)
            RaiseEvent WorkerStateChanged(Me, If(cancelled, "Job cancelled", "Worker disconnected"))
        End Sub

        Private Async Function ShutdownWorkerAsync(killIfStillRunning As Boolean) As Task(Of Boolean)
            Dim shutdownClean As Boolean = True

            If _pipeCancellation IsNot Nothing Then
                _pipeCancellation.Cancel()
            End If

            If _pipeClient IsNot Nothing Then
                Try
                    _pipeClient.Dispose()
                Catch
                End Try
                _pipeClient = Nothing
            End If

            If _receiveLoopTask IsNot Nothing Then
                Try
                    Await _receiveLoopTask.ConfigureAwait(False)
                Catch
                End Try
            End If

            If _workerProcess Is Nothing Then
                Return TryTerminateTrackedWorkers(killIfStillRunning, skipPid:=0)
            End If

            Try
                RemoveHandler _workerProcess.Exited, AddressOf WorkerProcess_Exited
            Catch
            End Try

            Dim workerPid = 0
            Try
                workerPid = _workerProcess.Id
            Catch
                workerPid = 0
            End Try

            Try
                If Not _workerProcess.HasExited Then
                    If killIfStillRunning Then
                        If Not _workerProcess.WaitForExit(CInt(TimeSpan.FromSeconds(2).TotalMilliseconds)) Then
                            Try
                                _workerProcess.Kill(entireProcessTree:=True)
                            Catch
                            End Try

                            If Not _workerProcess.WaitForExit(CInt(TimeSpan.FromSeconds(3).TotalMilliseconds)) AndAlso workerPid > 0 Then
                                Try
                                    Dim taskKillInfo As New ProcessStartInfo("taskkill.exe", $"/PID {workerPid} /T /F") With {
                                        .UseShellExecute = False,
                                        .CreateNoWindow = True,
                                        .WindowStyle = ProcessWindowStyle.Hidden
                                    }
                                    Using taskKillProcess = Process.Start(taskKillInfo)
                                        taskKillProcess?.WaitForExit(CInt(TimeSpan.FromSeconds(3).TotalMilliseconds))
                                    End Using
                                Catch
                                End Try
                            End If

                            If Not _workerProcess.HasExited Then
                                shutdownClean = False
                                RaiseEvent LogMessage(Me, $"[Supervisor] Worker PID {workerPid} did not terminate promptly; continuing with recovery.")
                            End If
                        End If
                    End If
                End If
            Catch
                shutdownClean = False
            Finally
                UntrackWorkerProcess(workerPid)
                _workerProcess.Dispose()
                _workerProcess = Nothing
            End Try

            If Not TryTerminateTrackedWorkers(killIfStillRunning, skipPid:=workerPid) Then
                shutdownClean = False
            End If

            Return shutdownClean
        End Function

        Private Sub TrackWorkerProcess(pid As Integer)
            If pid <= 0 Then
                Return
            End If

            SyncLock _workerPidLock
                _launchedWorkerPids.Add(pid)
            End SyncLock
        End Sub

        Private Sub UntrackWorkerProcess(pid As Integer)
            If pid <= 0 Then
                Return
            End If

            SyncLock _workerPidLock
                _launchedWorkerPids.Remove(pid)
            End SyncLock
        End Sub

        Private Function SnapshotTrackedWorkerPids() As List(Of Integer)
            SyncLock _workerPidLock
                Return _launchedWorkerPids.ToList()
            End SyncLock
        End Function

        Private Function TryTerminateTrackedWorkers(killIfStillRunning As Boolean, skipPid As Integer) As Boolean
            Dim allTerminated = True
            Dim trackedPids = SnapshotTrackedWorkerPids()
            If trackedPids.Count = 0 Then
                Return True
            End If

            For Each pid In trackedPids
                If pid <= 0 OrElse pid = skipPid Then
                    Continue For
                End If

                Dim trackedProcess As Process = Nothing
                Try
                    trackedProcess = Process.GetProcessById(pid)
                Catch
                    UntrackWorkerProcess(pid)
                    Continue For
                End Try

                Using trackedProcess
                    Try
                        If trackedProcess.HasExited Then
                            UntrackWorkerProcess(pid)
                            Continue For
                        End If

                        If killIfStillRunning Then
                            Try
                                trackedProcess.Kill(entireProcessTree:=True)
                            Catch
                            End Try

                            If Not trackedProcess.WaitForExit(CInt(TimeSpan.FromSeconds(2).TotalMilliseconds)) Then
                                Try
                                    Dim taskKillInfo As New ProcessStartInfo("taskkill.exe", $"/PID {pid} /T /F") With {
                                        .UseShellExecute = False,
                                        .CreateNoWindow = True,
                                        .WindowStyle = ProcessWindowStyle.Hidden
                                    }
                                    Using taskKillProcess = Process.Start(taskKillInfo)
                                        taskKillProcess?.WaitForExit(CInt(TimeSpan.FromSeconds(3).TotalMilliseconds))
                                    End Using
                                Catch
                                End Try
                            End If
                        End If

                        If trackedProcess.HasExited Then
                            UntrackWorkerProcess(pid)
                        Else
                            allTerminated = False
                            RaiseEvent LogMessage(Me, $"[Supervisor] Orphan worker PID {pid} is still alive after shutdown attempt.")
                        End If
                    Catch
                        allTerminated = False
                    End Try
                End Using
            Next

            Return allTerminated
        End Function

        Private Sub CheckForJobActivityStall(nowUtc As DateTimeOffset)
            If Not _isJobRunning OrElse _isJobPaused OrElse Not _autoRecoverJob Then
                Return
            End If

            Dim lastActivityUtc = _lastJobActivityUtc
            If lastActivityUtc = DateTimeOffset.MinValue Then
                lastActivityUtc = _lastWorkerProgressUtc
            End If

            If lastActivityUtc = DateTimeOffset.MinValue Then
                Return
            End If

            Dim activityTimeout = ResolveJobActivityTimeout()
            Dim activityElapsed = nowUtc - lastActivityUtc
            If activityElapsed <= activityTimeout Then
                Return
            End If

            Dim progressElapsed As TimeSpan
            If _lastWorkerProgressUtc = DateTimeOffset.MinValue Then
                progressElapsed = activityElapsed
            Else
                progressElapsed = nowUtc - _lastWorkerProgressUtc
            End If

            QueueRecovery(
                $"Worker activity stalled for {activityElapsed.TotalSeconds:0.0}s (last progress {progressElapsed.TotalSeconds:0.0}s ago).")
        End Sub

        Private Function ResolveJobActivityTimeout() As TimeSpan
            Dim operationTimeoutSeconds = 10.0R
            Dim maxRetryDelaySeconds = 2.0R

            If _currentJobOptions IsNot Nothing Then
                If _currentJobOptions.OperationTimeout > TimeSpan.Zero Then
                    operationTimeoutSeconds = _currentJobOptions.OperationTimeout.TotalSeconds
                End If

                If _currentJobOptions.MaxRetryDelay > TimeSpan.Zero Then
                    maxRetryDelaySeconds = _currentJobOptions.MaxRetryDelay.TotalSeconds
                End If
            End If

            Dim computedSeconds = (operationTimeoutSeconds * 4.0R) + (maxRetryDelaySeconds * 2.0R) + 10.0R
            Dim boundedSeconds = Math.Max(
                MinimumJobActivityTimeout.TotalSeconds,
                Math.Min(MaximumJobActivityTimeout.TotalSeconds, computedSeconds))
            Return TimeSpan.FromSeconds(boundedSeconds)
        End Function

        Private Sub MarkJobActivity(Optional timestampUtc As DateTimeOffset = Nothing, Optional markProgress As Boolean = False)
            Dim effectiveUtc = NormalizeTimestampOrNow(timestampUtc)
            _lastJobActivityUtc = effectiveUtc
            If markProgress Then
                _lastWorkerProgressUtc = effectiveUtc
            End If
        End Sub

        Private Shared Function NormalizeTimestampOrNow(
            timestampUtc As DateTimeOffset,
            Optional useNowWhenMinValue As Boolean = True) As DateTimeOffset

            If timestampUtc = DateTimeOffset.MinValue Then
                Return If(useNowWhenMinValue, DateTimeOffset.UtcNow, DateTimeOffset.MinValue)
            End If

            Return timestampUtc.ToUniversalTime()
        End Function

        Private Shared Function ResolveWorkerLaunch(pipeName As String) As WorkerLaunch
            Dim workerExecutablePath = ResolveWorkerExecutablePath()
            If Not String.IsNullOrWhiteSpace(workerExecutablePath) Then
                Return New WorkerLaunch(workerExecutablePath, $"--pipe ""{pipeName}""")
            End If

            Dim workerDllPath = ResolveWorkerDllPath()
            Return New WorkerLaunch("dotnet", $"""{workerDllPath}"" --pipe ""{pipeName}""")
        End Function

        Private Shared Function ResolveWorkerExecutablePath() As String
            Dim baseDirectory = AppContext.BaseDirectory
            Dim preferredPath = Path.Combine(baseDirectory, "XactCopyExecutive.exe")
            If File.Exists(preferredPath) Then
                Return preferredPath
            End If

            Dim legacyPath = Path.Combine(baseDirectory, "XactCopy.Worker.exe")
            If File.Exists(legacyPath) Then
                Return legacyPath
            End If

            Dim preferredMatch = Directory.EnumerateFiles(baseDirectory, "XactCopyExecutive.exe", SearchOption.AllDirectories).FirstOrDefault()
            If Not String.IsNullOrWhiteSpace(preferredMatch) Then
                Return preferredMatch
            End If

            Dim legacyMatch = Directory.EnumerateFiles(baseDirectory, "XactCopy.Worker.exe", SearchOption.AllDirectories).FirstOrDefault()
            Return legacyMatch
        End Function

        Private Shared Function ResolveWorkerDllPath() As String
            Dim baseDirectory = AppContext.BaseDirectory
            Dim preferredPath = Path.Combine(baseDirectory, "XactCopyExecutive.dll")
            If File.Exists(preferredPath) Then
                Return preferredPath
            End If

            Dim legacyPath = Path.Combine(baseDirectory, "XactCopy.Worker.dll")
            If File.Exists(legacyPath) Then
                Return legacyPath
            End If

            Dim preferredMatch = Directory.EnumerateFiles(baseDirectory, "XactCopyExecutive.dll", SearchOption.AllDirectories).FirstOrDefault()
            If Not String.IsNullOrWhiteSpace(preferredMatch) Then
                Return preferredMatch
            End If

            Dim legacyMatch = Directory.EnumerateFiles(baseDirectory, "XactCopy.Worker.dll", SearchOption.AllDirectories).FirstOrDefault()
            If Not String.IsNullOrWhiteSpace(legacyMatch) Then
                Return legacyMatch
            End If

            Throw New FileNotFoundException(
                "Unable to locate worker binary (XactCopyExecutive.exe/.dll). Build the solution so the worker output is available.")
        End Function

        Private Shared Function CloneOptions(options As CopyJobOptions) As CopyJobOptions
            If options Is Nothing Then
                Return New CopyJobOptions()
            End If

            ' Full deep clone protects active dispatch state from live UI mutations.
            Return New CopyJobOptions() With {
                .OperationMode = options.OperationMode,
                .SourceRoot = options.SourceRoot,
                .DestinationRoot = options.DestinationRoot,
                .ExpectedSourceIdentity = options.ExpectedSourceIdentity,
                .ExpectedDestinationIdentity = options.ExpectedDestinationIdentity,
                .UseBadRangeMap = options.UseBadRangeMap,
                .SkipKnownBadRanges = options.SkipKnownBadRanges,
                .UpdateBadRangeMapFromRun = options.UpdateBadRangeMapFromRun,
                .BadRangeMapMaxAgeDays = options.BadRangeMapMaxAgeDays,
                .UseExperimentalRawDiskScan = options.UseExperimentalRawDiskScan,
                .ResumeJournalPathHint = options.ResumeJournalPathHint,
                .AllowJournalRootRemap = options.AllowJournalRootRemap,
                .SelectedRelativePaths = New List(Of String)(If(options.SelectedRelativePaths, New List(Of String)())),
                .OverwritePolicy = options.OverwritePolicy,
                .SymlinkHandling = options.SymlinkHandling,
                .CopyEmptyDirectories = options.CopyEmptyDirectories,
                .BufferSizeBytes = options.BufferSizeBytes,
                .UseAdaptiveBufferSizing = options.UseAdaptiveBufferSizing,
                .MaxThroughputBytesPerSecond = options.MaxThroughputBytesPerSecond,
                .ParallelSmallFileWorkers = options.ParallelSmallFileWorkers,
                .SmallFileThresholdBytes = options.SmallFileThresholdBytes,
                .WaitForMediaAvailability = options.WaitForMediaAvailability,
                .WaitForFileLockRelease = options.WaitForFileLockRelease,
                .TreatAccessDeniedAsContention = options.TreatAccessDeniedAsContention,
                .LockContentionProbeInterval = options.LockContentionProbeInterval,
                .SourceMutationPolicy = options.SourceMutationPolicy,
                .FragileMediaMode = options.FragileMediaMode,
                .SkipFileOnFirstReadError = options.SkipFileOnFirstReadError,
                .PersistFragileSkipAcrossResume = options.PersistFragileSkipAcrossResume,
                .FragileFailureWindowSeconds = options.FragileFailureWindowSeconds,
                .FragileFailureThreshold = options.FragileFailureThreshold,
                .FragileCooldownSeconds = options.FragileCooldownSeconds,
                .MaxRetries = options.MaxRetries,
                .OperationTimeout = options.OperationTimeout,
                .PerFileTimeout = options.PerFileTimeout,
                .InitialRetryDelay = options.InitialRetryDelay,
                .MaxRetryDelay = options.MaxRetryDelay,
                .ResumeFromJournal = options.ResumeFromJournal,
                .VerifyAfterCopy = options.VerifyAfterCopy,
                .VerificationMode = options.VerificationMode,
                .VerificationHashAlgorithm = options.VerificationHashAlgorithm,
                .SampleVerificationChunkBytes = options.SampleVerificationChunkBytes,
                .SampleVerificationChunkCount = options.SampleVerificationChunkCount,
                .SalvageUnreadableBlocks = options.SalvageUnreadableBlocks,
                .SalvageFillPattern = options.SalvageFillPattern,
                .ContinueOnFileError = options.ContinueOnFileError,
                .PreserveTimestamps = options.PreserveTimestamps,
                .WorkerProcessPriorityClass = options.WorkerProcessPriorityClass,
                .WorkerTelemetryProfile = options.WorkerTelemetryProfile,
                .WorkerProgressEmitIntervalMs = options.WorkerProgressEmitIntervalMs,
                .WorkerMaxLogsPerSecond = options.WorkerMaxLogsPerSecond,
                .RescueFastScanChunkBytes = options.RescueFastScanChunkBytes,
                .RescueTrimChunkBytes = options.RescueTrimChunkBytes,
                .RescueScrapeChunkBytes = options.RescueScrapeChunkBytes,
                .RescueRetryChunkBytes = options.RescueRetryChunkBytes,
                .RescueSplitMinimumBytes = options.RescueSplitMinimumBytes,
                .RescueFastScanRetries = options.RescueFastScanRetries,
                .RescueTrimRetries = options.RescueTrimRetries,
                .RescueScrapeRetries = options.RescueScrapeRetries
            }
        End Function

        Private Sub SetPausedState(value As Boolean)
            If _isJobPaused = value Then
                Return
            End If

            _isJobPaused = value
            RaiseEvent JobPauseStateChanged(Me, value)
        End Sub

        Private Sub ThrowIfDisposed()
            If _isDisposed Then
                Throw New ObjectDisposedException(NameOf(WorkerSupervisor))
            End If
        End Sub

        ''' <summary>
        ''' Executes Dispose.
        ''' </summary>
        Public Sub Dispose() Implements IDisposable.Dispose
            StopAsync().GetAwaiter().GetResult()
            _restartLock.Dispose()
            _sendLock.Dispose()
            _disposeLock.Dispose()
        End Sub

        ''' <summary>
        ''' Class WorkerLaunch.
        ''' </summary>
        Private NotInheritable Class WorkerLaunch
            ''' <summary>
            ''' Initializes a new instance.
            ''' </summary>
            Public Sub New(fileName As String, arguments As String)
                Me.FileName = If(fileName, String.Empty)
                Me.Arguments = If(arguments, String.Empty)
            End Sub

            ''' <summary>
            ''' Gets or sets FileName.
            ''' </summary>
            Public ReadOnly Property FileName As String
            ''' <summary>
            ''' Gets or sets Arguments.
            ''' </summary>
            Public ReadOnly Property Arguments As String
        End Class
    End Class
End Namespace
