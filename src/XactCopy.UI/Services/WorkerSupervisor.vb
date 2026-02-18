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
    Public Class WorkerSupervisor
        Implements IDisposable

        Public Event LogMessage As EventHandler(Of String)
        Public Event ProgressChanged As EventHandler(Of CopyProgressSnapshot)
        Public Event JobCompleted As EventHandler(Of CopyJobResult)
        Public Event WorkerStateChanged As EventHandler(Of String)
        Public Event JobPauseStateChanged As EventHandler(Of Boolean)

        Private Shared ReadOnly HeartbeatCheckInterval As TimeSpan = TimeSpan.FromSeconds(2)
        Private Shared ReadOnly HeartbeatTimeout As TimeSpan = TimeSpan.FromSeconds(10)

        Private ReadOnly _restartLock As New SemaphoreSlim(1, 1)
        Private ReadOnly _sendLock As New SemaphoreSlim(1, 1)
        Private ReadOnly _disposeLock As New SemaphoreSlim(1, 1)

        Private _workerProcess As Process
        Private _pipeClient As NamedPipeClientStream
        Private _pipeCancellation As CancellationTokenSource
        Private _receiveLoopTask As Task = Task.CompletedTask
        Private _heartbeatTimer As Timer
        Private _lastHeartbeatUtc As DateTimeOffset = DateTimeOffset.MinValue

        Private _currentJobOptions As CopyJobOptions
        Private _currentJobId As String = String.Empty
        Private _isJobRunning As Boolean
        Private _isJobPaused As Boolean
        Private _autoRecoverJob As Boolean
        Private _isStopping As Boolean
        Private _isDisposed As Boolean
        Private _recoveryInFlight As Integer

        Public ReadOnly Property IsJobRunning As Boolean
            Get
                Return _isJobRunning
            End Get
        End Property

        Public ReadOnly Property IsJobPaused As Boolean
            Get
                Return _isJobPaused
            End Get
        End Property

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
                SetPausedState(False)
                _autoRecoverJob = True

                Await EnsureWorkerConnectedAsync(cancellationToken).ConfigureAwait(False)
                Await SendStartJobCommandAsync(cancellationToken).ConfigureAwait(False)
                RaiseEvent WorkerStateChanged(Me, "Job running")
            Finally
                _restartLock.Release()
            End Try
        End Function

        Public Async Function CancelJobAsync(reason As String) As Task
            ThrowIfDisposed()

            If Not _isJobRunning Then
                Return
            End If

            _autoRecoverJob = False
            SetPausedState(False)
            Dim command As New CancelJobCommand() With {
                .JobId = _currentJobId,
                .Reason = If(String.IsNullOrWhiteSpace(reason), "User requested cancellation.", reason)
            }

            Await TrySendCommandAsync(IpcMessageTypes.CancelJobCommand, command, CancellationToken.None).ConfigureAwait(False)
        End Function

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
                SetPausedState(False)
                _currentJobId = String.Empty
                _currentJobOptions = Nothing

                Dim shutdown As New ShutdownCommand() With {
                    .Reason = "UI shutdown requested."
                }
                Await TrySendCommandAsync(IpcMessageTypes.ShutdownCommand, shutdown, CancellationToken.None).ConfigureAwait(False)

                Await ShutdownWorkerAsync(killIfStillRunning:=True).ConfigureAwait(False)

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
                Return Task.CompletedTask
            End If

            Select Case messageType
                Case IpcMessageTypes.WorkerHeartbeatEvent
                    Dim envelope = IpcSerializer.DeserializeEnvelope(Of WorkerHeartbeatEvent)(rawMessage)
                    _lastHeartbeatUtc = envelope.Payload.TimestampUtc
                    SetPausedState(envelope.Payload.IsJobPaused)

                Case IpcMessageTypes.WorkerProgressEvent
                    Dim envelope = IpcSerializer.DeserializeEnvelope(Of WorkerProgressEvent)(rawMessage)
                    _lastHeartbeatUtc = envelope.Payload.TimestampUtc
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
                    RaiseEvent LogMessage(Me, envelope.Payload.Message)

                Case IpcMessageTypes.WorkerJobResultEvent
                    Dim envelope = IpcSerializer.DeserializeEnvelope(Of WorkerJobResultEvent)(rawMessage)
                    _isJobRunning = False
                    SetPausedState(False)
                    _autoRecoverJob = False
                    _currentJobId = String.Empty
                    _currentJobOptions = Nothing
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
                    _currentJobId = String.Empty
                    _currentJobOptions = Nothing
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

            Dim elapsed = DateTimeOffset.UtcNow - _lastHeartbeatUtc
            If elapsed <= HeartbeatTimeout Then
                Return
            End If

            QueueRecovery($"Worker heartbeat stale for {elapsed.TotalSeconds:0.0}s.")
        End Sub

        Private Sub WorkerProcess_Exited(sender As Object, e As EventArgs)
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
            If Not _autoRecoverJob OrElse Not _isJobRunning Then
                Return
            End If

            Await _restartLock.WaitAsync().ConfigureAwait(False)
            Try
                If _isStopping OrElse _isDisposed Then
                    Return
                End If

                RaiseEvent LogMessage(Me, $"[Supervisor] {reason} Restarting worker and resuming from journal.")

                Await ShutdownWorkerAsync(killIfStillRunning:=True).ConfigureAwait(False)
                Await EnsureWorkerConnectedAsync(CancellationToken.None).ConfigureAwait(False)

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

        Private Async Function ShutdownWorkerAsync(killIfStillRunning As Boolean) As Task
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
                Return
            End If

            Try
                RemoveHandler _workerProcess.Exited, AddressOf WorkerProcess_Exited
            Catch
            End Try

            Try
                If Not _workerProcess.HasExited Then
                    If killIfStillRunning Then
                        If Not _workerProcess.WaitForExit(CInt(TimeSpan.FromSeconds(2).TotalMilliseconds)) Then
                            _workerProcess.Kill(entireProcessTree:=True)
                            _workerProcess.WaitForExit(CInt(TimeSpan.FromSeconds(2).TotalMilliseconds))
                        End If
                    End If
                End If
            Catch
            Finally
                _workerProcess.Dispose()
                _workerProcess = Nothing
            End Try
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

        Public Sub Dispose() Implements IDisposable.Dispose
            StopAsync().GetAwaiter().GetResult()
            _restartLock.Dispose()
            _sendLock.Dispose()
            _disposeLock.Dispose()
        End Sub

        Private NotInheritable Class WorkerLaunch
            Public Sub New(fileName As String, arguments As String)
                Me.FileName = If(fileName, String.Empty)
                Me.Arguments = If(arguments, String.Empty)
            End Sub

            Public ReadOnly Property FileName As String
            Public ReadOnly Property Arguments As String
        End Class
    End Class
End Namespace
