Imports System.Collections.Generic
Imports System.Buffers
Imports System.IO
Imports System.Linq
Imports System.Globalization
Imports System.Runtime.InteropServices
Imports System.Security.Cryptography
Imports System.Threading
Imports Microsoft.Win32.SafeHandles
Imports XactCopy.Infrastructure
Imports XactCopy.Models

Namespace Services
    Public Class ResilientCopyService
        Public Event ProgressChanged As EventHandler(Of CopyProgressSnapshot)
        Public Event LogMessage As EventHandler(Of String)

        Private Const JournalFlushIntervalSeconds As Integer = 1
        Private Const RescueCoreName As String = "AegisRescueCore"
        Private Const MinimumRescueBlockSize As Integer = 4096
        Private Const MaximumIoBufferSize As Integer = 256 * 1024 * 1024
        Private Const IdentityMismatchLogIntervalSeconds As Integer = 5
        Private Const ErrorFileNotFound As Integer = 2
        Private Const ErrorPathNotFound As Integer = 3
        Private Const ErrorAccessDenied As Integer = 5
        Private Const ErrorNotReady As Integer = 21
        Private Const ErrorSharingViolation As Integer = 32
        Private Const ErrorLockViolation As Integer = 33
        Private Const ErrorHandleDiskFull As Integer = 39
        Private Const ErrorBadNetPath As Integer = 53
        Private Const ErrorNetNameDeleted As Integer = 64
        Private Const ErrorBadNetName As Integer = 67
        Private Const ErrorDiskFull As Integer = 112
        Private Const ErrorDeviceNotConnected As Integer = 1167

        Private ReadOnly _options As CopyJobOptions
        Private ReadOnly _executionControl As CopyExecutionControl
        Private ReadOnly _journalStore As JobJournalStore
        Private ReadOnly _badRangeMapStore As BadRangeMapStore
#If DEBUG Then
        Private ReadOnly _devFaultInjector As DevFaultInjector
#End If
        Private _lastJournalFlushUtc As DateTimeOffset = DateTimeOffset.MinValue
        Private _throttleWindowStartUtc As DateTimeOffset = DateTimeOffset.MinValue
        Private _throttleWindowBytes As Long
        Private _expectedSourceIdentity As String = String.Empty
        Private _expectedDestinationIdentity As String = String.Empty
        Private _lastSourceIdentityMismatchLogUtc As DateTimeOffset = DateTimeOffset.MinValue
        Private _lastDestinationIdentityMismatchLogUtc As DateTimeOffset = DateTimeOffset.MinValue
        Private _badRangeMap As BadRangeMap
        Private _badRangeMapPath As String = String.Empty
        Private _badRangeMapLoaded As Boolean
        Private _badRangeMapReadHintsEnabled As Boolean

        <DllImport("kernel32.dll", CharSet:=CharSet.Unicode, SetLastError:=True)>
        Private Shared Function GetVolumeInformation(
            lpRootPathName As String,
            lpVolumeNameBuffer As IntPtr,
            nVolumeNameSize As UInteger,
            ByRef lpVolumeSerialNumber As UInteger,
            ByRef lpMaximumComponentLength As UInteger,
            ByRef lpFileSystemFlags As UInteger,
            lpFileSystemNameBuffer As IntPtr,
            nFileSystemNameSize As UInteger) As Boolean
        End Function

        Public Sub New(options As CopyJobOptions, Optional executionControl As CopyExecutionControl = Nothing)
            _options = options
            _executionControl = If(executionControl, New CopyExecutionControl())
            _journalStore = New JobJournalStore()
            _badRangeMapStore = New BadRangeMapStore()
#If DEBUG Then
            _devFaultInjector = DevFaultInjector.TryCreateFromEnvironment()
#End If
        End Sub

        Public Async Function RunAsync(cancellationToken As CancellationToken) As Task(Of CopyJobResult)
            ValidateOptions()
#If DEBUG Then
            If _devFaultInjector IsNot Nothing Then
                EmitLog($"[DevFault] Enabled: {_devFaultInjector.Description}")
            End If
#End If

            Dim scanOnly = (_options.OperationMode = JobOperationMode.ScanOnly)
            Dim sourceRoot = Path.GetFullPath(_options.SourceRoot)
            Dim destinationRoot = Path.GetFullPath(_options.DestinationRoot)
            InitializeMediaIdentityExpectations(sourceRoot, destinationRoot)
            Await WaitForMediaAvailabilityAsync(sourceRoot, destinationRoot, cancellationToken).ConfigureAwait(False)
            EnsureMediaIdentityIntegrity(sourceRoot, destinationRoot)
            If Not scanOnly Then
                Directory.CreateDirectory(destinationRoot)
            End If

            Await InitializeBadRangeMapAsync(sourceRoot, cancellationToken).ConfigureAwait(False)

            EmitLog($"Scanning source: {sourceRoot}")
            Dim scanResult = DirectoryScanner.ScanSource(
                sourceRoot,
                _options.SelectedRelativePaths,
                _options.SymlinkHandling,
                _options.CopyEmptyDirectories,
                AddressOf EmitLog)
            Dim sourceFiles = scanResult.Files
            Dim totalBytes = sourceFiles.Sum(Function(fileDescriptor) fileDescriptor.Length)
            EmitLog($"Scan complete. {sourceFiles.Count} file(s), {FormatBytes(totalBytes)} total.")
            If scanOnly Then
                EmitLog("Operation mode: Scan only (read-only bad-block detection).")
            Else
                EmitDestinationCapacityWarning(destinationRoot, totalBytes)
            End If

            If Not scanOnly AndAlso _options.CopyEmptyDirectories AndAlso scanResult.Directories.Count > 0 Then
                EmitLog($"Preparing {scanResult.Directories.Count} destination directory path(s).")
                For Each relativeDirectory In scanResult.Directories
                    cancellationToken.ThrowIfCancellationRequested()
                    Await _executionControl.WaitIfPausedAsync(cancellationToken).ConfigureAwait(False)
                    Await WaitForMediaAvailabilityAsync(sourceRoot, destinationRoot, cancellationToken).ConfigureAwait(False)

                    Dim directoryPath = Path.Combine(destinationRoot, relativeDirectory)
                    Try
                        Directory.CreateDirectory(directoryPath)
                    Catch ex As Exception
                        EmitLog($"Destination directory skipped: {relativeDirectory} ({ex.Message})")
                    End Try
                Next
            End If

            Dim jobId = JobJournalStore.BuildJobId(sourceRoot, destinationRoot)
            Dim defaultJournalPath = JobJournalStore.GetDefaultJournalPath(jobId)
            Dim journalPath = ResolveWritableJournalPath(jobId, defaultJournalPath)

            Dim existingJournal As JobJournal = Nothing
            If _options.ResumeFromJournal Then
                For Each candidatePath In GetJournalLoadCandidates(defaultJournalPath, journalPath)
                    existingJournal = Await _journalStore.LoadAsync(candidatePath, cancellationToken).ConfigureAwait(False)
                    If existingJournal IsNot Nothing Then
                        EmitLog($"Loaded journal: {candidatePath}")
                        Exit For
                    End If
                Next

                If existingJournal Is Nothing Then
                    EmitLog("No journal found. Creating a fresh state.")
                End If
            Else
                EmitLog("Resume disabled. Starting with a fresh state.")
            End If

            Dim journal = MergeJournal(existingJournal, sourceFiles, sourceRoot, destinationRoot, jobId)
            Await _journalStore.SaveAsync(journalPath, journal, cancellationToken).ConfigureAwait(False)
            _lastJournalFlushUtc = DateTimeOffset.UtcNow

            Dim progress As New ProgressAccumulator() With {
                .TotalFiles = sourceFiles.Count,
                .TotalBytes = totalBytes
            }

            For Each descriptor In sourceFiles
                cancellationToken.ThrowIfCancellationRequested()
                Await _executionControl.WaitIfPausedAsync(cancellationToken).ConfigureAwait(False)
                Await WaitForMediaAvailabilityAsync(sourceRoot, destinationRoot, cancellationToken).ConfigureAwait(False)
                EnsureMediaIdentityIntegrity(sourceRoot, destinationRoot)

                Dim entry = journal.Files(descriptor.RelativePath)

                If Not scanOnly Then
                    Dim skipReason As String = String.Empty
                    If ShouldSkipByOverwritePolicy(descriptor, destinationRoot, skipReason) Then
                        progress.SkippedFiles += 1
                        progress.TotalBytesCopied += descriptor.Length
                        EmitLog($"Skipped: {descriptor.RelativePath} ({skipReason})")
                        EmitProgress(
                            progress,
                            descriptor.RelativePath,
                            descriptor.Length,
                            descriptor.Length,
                            lastChunkBytesTransferred:=0,
                            bufferSizeBytes:=ResolveBufferSizeForFile(descriptor.Length))
                        Continue For
                    End If

                    If IsAlreadyCompleted(entry, descriptor, destinationRoot) Then
                        progress.SkippedFiles += 1
                        progress.TotalBytesCopied += descriptor.Length
                        If entry.State = FileCopyState.CompletedWithRecovery Then
                            progress.RecoveredFiles += 1
                        End If

                        EmitProgress(
                            progress,
                            descriptor.RelativePath,
                            descriptor.Length,
                            descriptor.Length,
                            lastChunkBytesTransferred:=0,
                            bufferSizeBytes:=ResolveBufferSizeForFile(descriptor.Length))
                        Continue For
                    End If
                End If

                If scanOnly Then
                    If entry.State = FileCopyState.Failed Then
                        entry.State = FileCopyState.Pending
                        entry.LastError = String.Empty
                    End If

                    Dim scanException As Exception = Nothing
                    Dim scanDetectedBadRanges = False
                    Dim scanFileTimeoutCts As CancellationTokenSource = Nothing
                    Dim scanFileToken = cancellationToken

                    Try
                        If _options.PerFileTimeout > TimeSpan.Zero Then
                            scanFileTimeoutCts = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken)
                            scanFileTimeoutCts.CancelAfter(_options.PerFileTimeout)
                            scanFileToken = scanFileTimeoutCts.Token
                        End If

                        entry.State = FileCopyState.InProgress
                        entry.LastError = String.Empty
                        Await FlushJournalAsync(journalPath, journal, cancellationToken, force:=True).ConfigureAwait(False)

                        scanDetectedBadRanges = Await ScanSingleFileForBadRangesAsync(
                            descriptor,
                            entry,
                            progress,
                            journal,
                            journalPath,
                            scanFileToken).ConfigureAwait(False)
                    Catch ex As SourceMutationSkippedException
                        progress.SkippedFiles += 1
                        entry.State = FileCopyState.Pending
                        entry.LastError = ex.Message
                        EmitLog($"Skipped scan: {descriptor.RelativePath} ({ex.Message})")
                        FlushJournalAsync(journalPath, journal, cancellationToken, force:=True).GetAwaiter().GetResult()
                        Continue For
                    Catch ex As OperationCanceledException When cancellationToken.IsCancellationRequested
                        Throw
                    Catch ex As OperationCanceledException
                        scanException = New TimeoutException(
                            $"Per-file timeout ({CInt(Math.Round(_options.PerFileTimeout.TotalSeconds))} sec) reached while scanning {descriptor.RelativePath}.",
                            ex)
                    Catch ex As Exception
                        scanException = ex
                    Finally
                        If scanFileTimeoutCts IsNot Nothing Then
                            scanFileTimeoutCts.Dispose()
                        End If
                    End Try

                    If scanException Is Nothing Then
                        progress.CompletedFiles += 1
                        If scanDetectedBadRanges Then
                            progress.RecoveredFiles += 1
                        End If

                        entry.State = If(scanDetectedBadRanges, FileCopyState.CompletedWithRecovery, FileCopyState.Completed)
                        entry.BytesCopied = descriptor.Length
                        entry.LastError = String.Empty
                        Await TryPersistBadRangeMapEntryAsync(sourceRoot, descriptor, entry, cancellationToken).ConfigureAwait(False)
                        Await FlushJournalAsync(journalPath, journal, cancellationToken, force:=True).ConfigureAwait(False)
                        Continue For
                    End If

                    progress.FailedFiles += 1
                    entry.State = FileCopyState.Failed
                    entry.LastError = scanException.Message
                    EmitLog($"Scan failed: {descriptor.RelativePath} ({scanException.Message})")
                    Await TryPersistBadRangeMapEntryAsync(sourceRoot, descriptor, entry, cancellationToken).ConfigureAwait(False)
                    Await FlushJournalAsync(journalPath, journal, cancellationToken, force:=True).ConfigureAwait(False)

                    If Not _options.ContinueOnFileError Then
                        Return CreateResult(progress, journalPath, succeeded:=False, cancelled:=False, errorMessage:=scanException.Message)
                    End If

                    Continue For
                End If

                If entry.State = FileCopyState.Failed Then
                    entry.State = FileCopyState.Pending
                    entry.LastError = String.Empty
                End If

                Dim copyException As Exception = Nothing
                Dim recovered = False
                Dim fileTimeoutCts As CancellationTokenSource = Nothing
                Dim fileToken = cancellationToken

                Try
                    If _options.PerFileTimeout > TimeSpan.Zero Then
                        fileTimeoutCts = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken)
                        fileTimeoutCts.CancelAfter(_options.PerFileTimeout)
                        fileToken = fileTimeoutCts.Token
                    End If

                    entry.State = FileCopyState.InProgress
                    entry.LastError = String.Empty
                    Await FlushJournalAsync(journalPath, journal, cancellationToken, force:=True).ConfigureAwait(False)

                    recovered = Await CopySingleFileAsync(
                        descriptor,
                        entry,
                        destinationRoot,
                        progress,
                        journal,
                        journalPath,
                        fileToken).ConfigureAwait(False)

                Catch ex As SourceMutationSkippedException
                    progress.SkippedFiles += 1
                    entry.State = FileCopyState.Pending
                    entry.LastError = ex.Message
                    EmitLog($"Skipped: {descriptor.RelativePath} ({ex.Message})")
                    FlushJournalAsync(journalPath, journal, cancellationToken, force:=True).GetAwaiter().GetResult()
                    Continue For
                Catch ex As OperationCanceledException When cancellationToken.IsCancellationRequested
                    Throw
                Catch ex As OperationCanceledException
                    copyException = New TimeoutException(
                        $"Per-file timeout ({CInt(Math.Round(_options.PerFileTimeout.TotalSeconds))} sec) reached while copying {descriptor.RelativePath}.",
                        ex)
                Catch ex As Exception
                    copyException = ex
                Finally
                    If fileTimeoutCts IsNot Nothing Then
                        fileTimeoutCts.Dispose()
                    End If
                End Try

                If copyException Is Nothing Then
                    progress.CompletedFiles += 1
                    If recovered Then
                        progress.RecoveredFiles += 1
                    End If

                    entry.State = If(recovered, FileCopyState.CompletedWithRecovery, FileCopyState.Completed)
                    entry.BytesCopied = descriptor.Length
                    entry.LastError = String.Empty
                    Await TryPersistBadRangeMapEntryAsync(sourceRoot, descriptor, entry, cancellationToken).ConfigureAwait(False)
                    Await FlushJournalAsync(journalPath, journal, cancellationToken, force:=True).ConfigureAwait(False)
                    Continue For
                End If

                progress.FailedFiles += 1
                entry.State = FileCopyState.Failed
                entry.LastError = copyException.Message
                EmitLog($"Failed: {descriptor.RelativePath} ({copyException.Message})")
                Await TryPersistBadRangeMapEntryAsync(sourceRoot, descriptor, entry, cancellationToken).ConfigureAwait(False)
                Await FlushJournalAsync(journalPath, journal, cancellationToken, force:=True).ConfigureAwait(False)

                If Not _options.ContinueOnFileError Then
                    Return CreateResult(progress, journalPath, succeeded:=False, cancelled:=False, errorMessage:=copyException.Message)
                End If
            Next

            Await FlushBadRangeMapAsync(cancellationToken).ConfigureAwait(False)
            Await FlushJournalAsync(journalPath, journal, cancellationToken, force:=True).ConfigureAwait(False)
            Return CreateResult(progress, journalPath, succeeded:=(progress.FailedFiles = 0), cancelled:=False, errorMessage:=String.Empty)
        End Function

        Private Sub ValidateOptions()
            If String.IsNullOrWhiteSpace(_options.SourceRoot) Then
                Throw New InvalidOperationException("Source path is required.")
            End If

            Dim scanOnly = _options.OperationMode = JobOperationMode.ScanOnly
            If _options.OperationMode <> JobOperationMode.Copy AndAlso Not scanOnly Then
                Throw New InvalidOperationException("OperationMode is invalid.")
            End If

            If Not scanOnly AndAlso String.IsNullOrWhiteSpace(_options.DestinationRoot) Then
                Throw New InvalidOperationException("Destination path is required.")
            End If

            If scanOnly AndAlso String.IsNullOrWhiteSpace(_options.DestinationRoot) Then
                _options.DestinationRoot = _options.SourceRoot
            End If

            If Not _options.WaitForMediaAvailability AndAlso Not Directory.Exists(_options.SourceRoot) Then
                Throw New DirectoryNotFoundException($"Source directory not found: {_options.SourceRoot}")
            End If

            Dim sourceRoot = NormalizeRootPath(_options.SourceRoot)
            Dim destinationRoot = NormalizeRootPath(_options.DestinationRoot)

            If Not scanOnly AndAlso StringComparer.OrdinalIgnoreCase.Equals(sourceRoot, destinationRoot) Then
                Throw New InvalidOperationException("Source and destination cannot be the same path.")
            End If

            If _options.BufferSizeBytes < 4096 Then
                Throw New InvalidOperationException("BufferSizeBytes must be at least 4096.")
            End If

            If _options.MaxRetries < 0 Then
                Throw New InvalidOperationException("MaxRetries cannot be negative.")
            End If

            If _options.OperationTimeout <= TimeSpan.Zero Then
                Throw New InvalidOperationException("OperationTimeout must be greater than zero.")
            End If

            If _options.PerFileTimeout < TimeSpan.Zero Then
                Throw New InvalidOperationException("PerFileTimeout cannot be negative.")
            End If

            If _options.LockContentionProbeInterval <= TimeSpan.Zero Then
                Throw New InvalidOperationException("LockContentionProbeInterval must be greater than zero.")
            End If

            If _options.MaxThroughputBytesPerSecond < 0 Then
                Throw New InvalidOperationException("MaxThroughputBytesPerSecond cannot be negative.")
            End If

            If _options.SampleVerificationChunkBytes <= 0 Then
                Throw New InvalidOperationException("SampleVerificationChunkBytes must be greater than zero.")
            End If

            If _options.SampleVerificationChunkCount <= 0 Then
                Throw New InvalidOperationException("SampleVerificationChunkCount must be greater than zero.")
            End If

            If _options.BadRangeMapMaxAgeDays < 0 Then
                Throw New InvalidOperationException("BadRangeMapMaxAgeDays cannot be negative.")
            End If

            If _options.RescueFastScanChunkBytes < 0 OrElse
                _options.RescueTrimChunkBytes < 0 OrElse
                _options.RescueScrapeChunkBytes < 0 OrElse
                _options.RescueRetryChunkBytes < 0 OrElse
                _options.RescueSplitMinimumBytes < 0 Then
                Throw New InvalidOperationException("Rescue chunk/split tuning values cannot be negative.")
            End If

            If _options.RescueFastScanRetries < 0 OrElse
                _options.RescueTrimRetries < 0 OrElse
                _options.RescueScrapeRetries < 0 Then
                Throw New InvalidOperationException("Rescue pass retries cannot be negative.")
            End If

            Select Case _options.SourceMutationPolicy
                Case SourceMutationPolicy.FailFile, SourceMutationPolicy.SkipFile, SourceMutationPolicy.WaitForReappearance
                    ' Valid.
                Case Else
                    Throw New InvalidOperationException("SourceMutationPolicy is invalid.")
            End Select
        End Sub

        Private Function MergeJournal(
            existing As JobJournal,
            sourceFiles As IReadOnlyList(Of SourceFileDescriptor),
            sourceRoot As String,
            destinationRoot As String,
            jobId As String) As JobJournal

            Dim canReuseJournal = existing IsNot Nothing AndAlso
                (AreJournalRootsEquivalent(existing.SourceRoot, sourceRoot) AndAlso
                 AreJournalRootsEquivalent(existing.DestinationRoot, destinationRoot))

            If Not canReuseJournal AndAlso existing IsNot Nothing AndAlso _options.AllowJournalRootRemap Then
                canReuseJournal = True
                EmitLog(
                    $"Journal root remap enabled. Reusing existing journal rooted at source='{existing.SourceRoot}', destination='{existing.DestinationRoot}'.")
            End If

            Dim journal As JobJournal
            If canReuseJournal Then
                journal = existing
            Else
                journal = New JobJournal() With {
                    .JobId = jobId,
                    .SourceRoot = sourceRoot,
                    .DestinationRoot = destinationRoot,
                    .CreatedUtc = DateTimeOffset.UtcNow,
                    .UpdatedUtc = DateTimeOffset.UtcNow,
                    .Files = New Dictionary(Of String, JournalFileEntry)(StringComparer.OrdinalIgnoreCase)
                }
            End If

            If journal.Files Is Nothing Then
                journal.Files = New Dictionary(Of String, JournalFileEntry)(StringComparer.OrdinalIgnoreCase)
            End If

            journal.JobId = jobId
            journal.SourceRoot = sourceRoot
            journal.DestinationRoot = destinationRoot

            Dim discoveredPaths As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

            For Each descriptor In sourceFiles
                discoveredPaths.Add(descriptor.RelativePath)

                Dim entry As JournalFileEntry = Nothing
                If Not journal.Files.TryGetValue(descriptor.RelativePath, entry) Then
                    entry = New JournalFileEntry() With {
                        .RelativePath = descriptor.RelativePath,
                        .SourceLength = descriptor.Length,
                        .SourceLastWriteUtcTicks = descriptor.LastWriteTimeUtc.Ticks,
                        .BytesCopied = 0,
                        .State = FileCopyState.Pending,
                        .LastError = String.Empty,
                        .RecoveredRanges = New List(Of ByteRange)(),
                        .RescueRanges = New List(Of RescueRange)(),
                        .LastRescuePass = String.Empty
                    }
                    journal.Files(descriptor.RelativePath) = entry
                    Continue For
                End If

                Dim sourceChanged = entry.SourceLength <> descriptor.Length OrElse
                    entry.SourceLastWriteUtcTicks <> descriptor.LastWriteTimeUtc.Ticks

                entry.SourceLength = descriptor.Length
                entry.SourceLastWriteUtcTicks = descriptor.LastWriteTimeUtc.Ticks

                If entry.RecoveredRanges Is Nothing Then
                    entry.RecoveredRanges = New List(Of ByteRange)()
                End If

                If entry.RescueRanges Is Nothing Then
                    entry.RescueRanges = New List(Of RescueRange)()
                End If

                If sourceChanged Then
                    entry.BytesCopied = 0
                    entry.State = FileCopyState.Pending
                    entry.LastError = String.Empty
                    entry.RecoveredRanges.Clear()
                    entry.RescueRanges.Clear()
                    entry.LastRescuePass = String.Empty
                End If
            Next

            Dim stalePaths = journal.Files.Keys.Where(Function(path) Not discoveredPaths.Contains(path)).ToList()
            For Each stalePath In stalePaths
                journal.Files.Remove(stalePath)
            Next

            Return journal
        End Function

        Private Shared Function AreJournalRootsEquivalent(leftPath As String, rightPath As String) As Boolean
            Try
                Return StringComparer.OrdinalIgnoreCase.Equals(Path.GetFullPath(leftPath), rightPath)
            Catch
                Return StringComparer.OrdinalIgnoreCase.Equals(
                    If(leftPath, String.Empty).Trim(),
                    If(rightPath, String.Empty).Trim())
            End Try
        End Function

        Private Function IsAlreadyCompleted(entry As JournalFileEntry, descriptor As SourceFileDescriptor, destinationRoot As String) As Boolean
            If entry Is Nothing Then
                Return False
            End If

            If entry.State <> FileCopyState.Completed AndAlso entry.State <> FileCopyState.CompletedWithRecovery Then
                Return False
            End If

            Dim destinationPath = Path.Combine(destinationRoot, descriptor.RelativePath)
            If Not File.Exists(destinationPath) Then
                Return False
            End If

            Try
                Dim destinationInfo As New FileInfo(destinationPath)
                Return destinationInfo.Length = descriptor.Length
            Catch
                Return False
            End Try
        End Function

        Private Function ShouldSkipByOverwritePolicy(
            descriptor As SourceFileDescriptor,
            destinationRoot As String,
            ByRef reason As String) As Boolean

            reason = String.Empty
            Dim destinationPath = Path.Combine(destinationRoot, descriptor.RelativePath)
            If Not File.Exists(destinationPath) Then
                Return False
            End If

            Select Case _options.OverwritePolicy
                Case OverwritePolicy.SkipExisting
                    reason = "destination already exists"
                    Return True

                Case OverwritePolicy.OverwriteIfSourceNewer
                    Try
                        Dim destinationInfo As New FileInfo(destinationPath)
                        If descriptor.LastWriteTimeUtc <= destinationInfo.LastWriteTimeUtc Then
                            reason = "destination is newer or same age"
                            Return True
                        End If
                    Catch
                        Return False
                    End Try

                Case OverwritePolicy.Ask
                    Return False

                Case Else
                    Return False
            End Select

            Return False
        End Function

        Private Async Function CopySingleFileAsync(
            descriptor As SourceFileDescriptor,
            entry As JournalFileEntry,
            destinationRoot As String,
            progress As ProgressAccumulator,
            journal As JobJournal,
            journalPath As String,
            cancellationToken As CancellationToken) As Task(Of Boolean)

            Dim destinationPath = Path.Combine(destinationRoot, descriptor.RelativePath)
            Dim destinationDirectory = Path.GetDirectoryName(destinationPath)
            If Not String.IsNullOrWhiteSpace(destinationDirectory) Then
                Directory.CreateDirectory(destinationDirectory)
            End If

            Dim destinationLength = GetExistingFileLength(destinationPath)
            Dim hasPersistedCoverage =
                entry IsNot Nothing AndAlso
                entry.RescueRanges IsNot Nothing AndAlso
                entry.RescueRanges.Any(
                    Function(range)
                        If range Is Nothing OrElse range.Length <= 0 Then
                            Return False
                        End If

                        Dim state = NormalizeRescueRangeState(range.State)
                        Return state = RescueRangeState.Good OrElse state = RescueRangeState.Recovered
                    End Function)
            Dim shouldPreserveExisting = _options.ResumeFromJournal AndAlso
                destinationLength > 0 AndAlso
                ((entry IsNot Nothing AndAlso entry.BytesCopied > 0) OrElse hasPersistedCoverage)

            PrepareDestinationFile(destinationPath, descriptor.Length, destinationLength, shouldPreserveExisting)
            destinationLength = GetExistingFileLength(destinationPath)

            Dim activeBufferSize = ResolveBufferSizeForFile(descriptor.Length)
            Dim ioBuffer = ArrayPool(Of Byte).Shared.Rent(Math.Max(activeBufferSize, MinimumRescueBlockSize))

            Try
                Using transferSession As New FileTransferSession(descriptor.FullPath, destinationPath, activeBufferSize)
                    entry.RescueRanges = BuildRescueRanges(entry, descriptor.Length, destinationLength)
                    Dim mappedUnreadableBytes = ApplyKnownBadRangesFromMap(sourceRoot:=_options.SourceRoot, descriptor:=descriptor, entry:=entry)
                    entry.LastRescuePass = "Init"

                    Dim alreadySatisfiedBytes = GetRescueRangeBytes(entry.RescueRanges, RescueRangeState.Good, RescueRangeState.Recovered)
                    entry.BytesCopied = alreadySatisfiedBytes

                    progress.TotalBytesCopied += alreadySatisfiedBytes
                    EmitProgress(
                        progress,
                        descriptor.RelativePath,
                        alreadySatisfiedBytes,
                        descriptor.Length,
                        lastChunkBytesTransferred:=0,
                        bufferSizeBytes:=activeBufferSize,
                        rescuePass:=entry.LastRescuePass,
                        rescueBadRegionCount:=GetUnreadableRegionCount(entry.RescueRanges),
                        rescueRemainingBytes:=GetRescueRangeBytes(entry.RescueRanges, RescueRangeState.Pending, RescueRangeState.Bad, RescueRangeState.KnownBad))

                    Dim recoveredAny = False
                    Dim useSmallFileFastPath = ShouldUseSmallFileFastPath(descriptor.Length, hasPersistedCoverage) AndAlso mappedUnreadableBytes <= 0
                    If useSmallFileFastPath Then
                        entry.LastRescuePass = "SmallFileFast"
                        recoveredAny = Await CopySmallFileFastAsync(
                            descriptor,
                            entry,
                            destinationPath,
                            transferSession,
                            ioBuffer,
                            activeBufferSize,
                            progress,
                            journal,
                            journalPath,
                            cancellationToken).ConfigureAwait(False)
                    Else
                        Dim passPlan = BuildRescuePassPlan(activeBufferSize).ToList()
                        Dim badDensity = ComputeRescueBadDensity(entry.RescueRanges, descriptor.Length)
                        For passIndex = 0 To passPlan.Count - 1
                            Dim passDefinition = AdaptRescuePassDefinitionForDensity(passPlan(passIndex), badDensity, activeBufferSize)
                            Dim targetCount = GetRescueRegionCount(entry.RescueRanges, passDefinition.TargetState)
                            If targetCount <= 0 Then
                                Continue For
                            End If

                            entry.LastRescuePass = passDefinition.Name

                            Dim passOutcome = Await ExecuteRescuePassAsync(
                                passDefinition,
                                descriptor,
                                entry,
                                destinationPath,
                                transferSession,
                                ioBuffer,
                                progress,
                                journal,
                                journalPath,
                                cancellationToken).ConfigureAwait(False)

                            Dim remainingBadRegions = GetUnreadableRegionCount(entry.RescueRanges)
                            Dim passRemainingBadBytes = GetUnreadableRangeBytes(entry.RescueRanges)
                            badDensity = ComputeRescueBadDensity(entry.RescueRanges, descriptor.Length)
                            If ShouldLogRescuePassSummary(passDefinition, passOutcome, remainingBadRegions) Then
                                EmitLog(
                                    $"[{RescueCoreName}] {passDefinition.Name} {descriptor.RelativePath}: attempts {passOutcome.AttemptedSegments}, recovered {FormatBytes(passOutcome.RecoveredBytes)}, failed {passOutcome.FailedSegments}, remaining bad {remainingBadRegions} ({FormatBytes(passRemainingBadBytes)}), density {badDensity:P1}.")
                            End If

                            SyncEntryBytesCopied(entry)
                            Await FlushJournalAsync(journalPath, journal, cancellationToken, force:=True).ConfigureAwait(False)
                        Next

                        recoveredAny = GetRescueRegionCount(entry.RescueRanges, RescueRangeState.Recovered) > 0
                        Dim remainingBadBytes = GetUnreadableRangeBytes(entry.RescueRanges)
                        If remainingBadBytes > 0 Then
                            If Not _options.SalvageUnreadableBlocks Then
                                Throw New IOException(
                                    $"[{RescueCoreName}] Unrecoverable regions remain on {descriptor.RelativePath}: {FormatBytes(remainingBadBytes)} in {GetUnreadableRegionCount(entry.RescueRanges)} block(s).")
                            End If

                            entry.LastRescuePass = "SalvageFill"
                            EmitLog($"[{RescueCoreName}] Salvaging remaining unreadable regions on {descriptor.RelativePath}.")
                            recoveredAny = Await SalvageRemainingRangesAsync(
                                descriptor,
                                entry,
                                destinationPath,
                                transferSession,
                                ioBuffer,
                                activeBufferSize,
                                progress,
                                journal,
                                journalPath,
                                cancellationToken).ConfigureAwait(False) OrElse recoveredAny
                            SyncEntryBytesCopied(entry)
                            entry.RecoveredRanges = MergeByteRanges(entry.RecoveredRanges)
                            Await FlushJournalAsync(journalPath, journal, cancellationToken, force:=True).ConfigureAwait(False)
                        End If
                    End If

                    If _options.PreserveTimestamps Then
                        File.SetLastWriteTimeUtc(destinationPath, descriptor.LastWriteTimeUtc)
                    End If

                    Dim verificationMode = ResolveVerificationMode()
                    If verificationMode <> VerificationMode.None Then
                        If recoveredAny Then
                            EmitLog($"Verification skipped for recovered file: {descriptor.RelativePath}")
                        Else
                            Await VerifyFileWithRetriesAsync(
                                descriptor.FullPath,
                                destinationPath,
                                descriptor.RelativePath,
                                verificationMode,
                                cancellationToken).ConfigureAwait(False)
                        End If
                    End If

                    Return recoveredAny
                End Using
            Finally
                ArrayPool(Of Byte).Shared.Return(ioBuffer, clearArray:=False)
            End Try
        End Function

        Private Async Function ScanSingleFileForBadRangesAsync(
            descriptor As SourceFileDescriptor,
            entry As JournalFileEntry,
            progress As ProgressAccumulator,
            journal As JobJournal,
            journalPath As String,
            cancellationToken As CancellationToken) As Task(Of Boolean)

            Dim activeBufferSize = ResolveBufferSizeForFile(descriptor.Length)
            Dim ioBuffer = ArrayPool(Of Byte).Shared.Rent(Math.Max(activeBufferSize, MinimumRescueBlockSize))

            Try
                Using transferSession As New FileTransferSession(descriptor.FullPath, String.Empty, activeBufferSize)
                    entry.RescueRanges = New List(Of RescueRange)()
                    AppendRescueRange(entry.RescueRanges, 0L, descriptor.Length, RescueRangeState.Pending)
                    If entry.RecoveredRanges Is Nothing Then
                        entry.RecoveredRanges = New List(Of ByteRange)()
                    Else
                        entry.RecoveredRanges.Clear()
                    End If

                    Dim mappedUnreadableBytes = ApplyKnownBadRangesFromMap(sourceRoot:=_options.SourceRoot, descriptor:=descriptor, entry:=entry)
                    entry.LastRescuePass = "Init"
                    SyncEntryBytesCopied(entry)
                    progress.TotalBytesCopied += entry.BytesCopied

                    EmitProgress(
                        progress,
                        descriptor.RelativePath,
                        entry.BytesCopied,
                        descriptor.Length,
                        lastChunkBytesTransferred:=0,
                        bufferSizeBytes:=activeBufferSize,
                        rescuePass:=entry.LastRescuePass,
                        rescueBadRegionCount:=GetUnreadableRegionCount(entry.RescueRanges),
                        rescueRemainingBytes:=GetRescueRangeBytes(entry.RescueRanges, RescueRangeState.Pending, RescueRangeState.Bad, RescueRangeState.KnownBad))

                    If mappedUnreadableBytes > 0 Then
                        EmitLog($"[{RescueCoreName}] Scan seeded with map hints: {FormatBytes(mappedUnreadableBytes)} known unreadable on {descriptor.RelativePath}.")
                    End If

                    Dim passPlan = BuildRescuePassPlan(activeBufferSize).ToList()
                    Dim badDensity = ComputeRescueBadDensity(entry.RescueRanges, descriptor.Length)
                    For passIndex = 0 To passPlan.Count - 1
                        Dim passDefinition = AdaptRescuePassDefinitionForDensity(passPlan(passIndex), badDensity, activeBufferSize)
                        Dim targetCount = GetRescueRegionCount(entry.RescueRanges, passDefinition.TargetState)
                        If targetCount <= 0 Then
                            Continue For
                        End If

                        entry.LastRescuePass = passDefinition.Name

                        Dim passOutcome = Await ExecuteRescuePassAsync(
                            passDefinition,
                            descriptor,
                            entry,
                            destinationPath:=String.Empty,
                            transferSession:=transferSession,
                            ioBuffer:=ioBuffer,
                            progress:=progress,
                            journal:=journal,
                            journalPath:=journalPath,
                            cancellationToken:=cancellationToken,
                            writeRecovered:=False,
                            countFailedAsProcessed:=True).ConfigureAwait(False)

                        Dim remainingUnreadableRegions = GetUnreadableRegionCount(entry.RescueRanges)
                        Dim remainingUnreadableBytes = GetUnreadableRangeBytes(entry.RescueRanges)
                        badDensity = ComputeRescueBadDensity(entry.RescueRanges, descriptor.Length)
                        If ShouldLogRescuePassSummary(passDefinition, passOutcome, remainingUnreadableRegions) Then
                            EmitLog(
                                $"[{RescueCoreName}] {passDefinition.Name} {descriptor.RelativePath}: attempts {passOutcome.AttemptedSegments}, recovered {FormatBytes(passOutcome.RecoveredBytes)}, failed {passOutcome.FailedSegments}, remaining bad {remainingUnreadableRegions} ({FormatBytes(remainingUnreadableBytes)}), density {badDensity:P1}.")
                        End If

                        SyncEntryBytesCopied(entry)
                        Await FlushJournalAsync(journalPath, journal, cancellationToken, force:=True).ConfigureAwait(False)
                    Next

                    Dim finalUnreadableBytes = GetUnreadableRangeBytes(entry.RescueRanges)
                    Dim finalUnreadableRegions = GetUnreadableRegionCount(entry.RescueRanges)
                    entry.RecoveredRanges = MergeByteRanges(
                        SnapshotRangesByState(
                            entry.RescueRanges,
                            RescueRangeState.Bad,
                            RescueRangeState.KnownBad))
                    SyncEntryBytesCopied(entry)

                    EmitProgress(
                        progress,
                        descriptor.RelativePath,
                        entry.BytesCopied,
                        descriptor.Length,
                        lastChunkBytesTransferred:=0,
                        bufferSizeBytes:=activeBufferSize,
                        rescuePass:="ScanComplete",
                        rescueBadRegionCount:=finalUnreadableRegions,
                        rescueRemainingBytes:=finalUnreadableBytes)

                    If finalUnreadableBytes > 0 Then
                        EmitLog($"[{RescueCoreName}] Scan detected unreadable ranges on {descriptor.RelativePath}: {FormatBytes(finalUnreadableBytes)} in {finalUnreadableRegions} segment(s).")
                    End If

                    Return finalUnreadableBytes > 0
                End Using
            Finally
                ArrayPool(Of Byte).Shared.Return(ioBuffer, clearArray:=False)
            End Try
        End Function

        Private Function ShouldTrackBadRangeMap() As Boolean
            Return _options.OperationMode = JobOperationMode.ScanOnly OrElse
                _options.UseBadRangeMap OrElse
                _options.UpdateBadRangeMapFromRun
        End Function

        Private Async Function InitializeBadRangeMapAsync(sourceRoot As String, cancellationToken As CancellationToken) As Task
            _badRangeMap = Nothing
            _badRangeMapPath = String.Empty
            _badRangeMapLoaded = False
            _badRangeMapReadHintsEnabled = False

            If Not ShouldTrackBadRangeMap() Then
                Return
            End If

            Dim normalizedSourceRoot = NormalizeRootPath(sourceRoot)
            _badRangeMapPath = BadRangeMapStore.GetDefaultMapPath(normalizedSourceRoot)

            Dim loadedMap As BadRangeMap = Nothing
            Try
                loadedMap = Await _badRangeMapStore.LoadAsync(_badRangeMapPath, cancellationToken).ConfigureAwait(False)
            Catch ex As OperationCanceledException
                Throw
            Catch ex As Exception
                EmitLog($"Bad-range map load failed, starting fresh: {ex.Message}")
                loadedMap = Nothing
            End Try

            Dim allowReadHints = False
            If loadedMap IsNot Nothing Then
                Dim mapSource = NormalizeRootPath(loadedMap.SourceRoot)
                If mapSource.Length > 0 AndAlso
                    Not String.Equals(mapSource, normalizedSourceRoot, StringComparison.OrdinalIgnoreCase) Then
                    EmitLog("Bad-range map source root mismatch; ignoring existing map payload.")
                    loadedMap = Nothing
                End If
            End If

            If loadedMap Is Nothing Then
                loadedMap = New BadRangeMap() With {
                    .SchemaVersion = 1,
                    .SourceRoot = normalizedSourceRoot,
                    .SourceIdentity = If(_expectedSourceIdentity, String.Empty),
                    .UpdatedUtc = DateTimeOffset.UtcNow,
                    .Files = New Dictionary(Of String, BadRangeMapFileEntry)(StringComparer.OrdinalIgnoreCase)
                }
                EmitLog("Bad-range map: initialized new map for this source.")
            Else
                If loadedMap.Files Is Nothing Then
                    loadedMap.Files = New Dictionary(Of String, BadRangeMapFileEntry)(StringComparer.OrdinalIgnoreCase)
                End If

                loadedMap.SourceRoot = normalizedSourceRoot
                If loadedMap.SchemaVersion <= 0 Then
                    loadedMap.SchemaVersion = 1
                End If

                loadedMap.SourceIdentity = If(_expectedSourceIdentity, loadedMap.SourceIdentity)
                allowReadHints = IsBadRangeMapFresh(loadedMap, _options.BadRangeMapMaxAgeDays)

                If allowReadHints Then
                    EmitLog($"Bad-range map loaded: {loadedMap.Files.Count} file entr{If(loadedMap.Files.Count = 1, "y", "ies")} available.")
                ElseIf _options.UseBadRangeMap AndAlso _options.SkipKnownBadRanges Then
                    EmitLog($"Bad-range map is stale (>{_options.BadRangeMapMaxAgeDays} days). Skip hints disabled for this run.")
                End If
            End If

            _badRangeMap = loadedMap
            _badRangeMapLoaded = True
            _badRangeMapReadHintsEnabled = _options.UseBadRangeMap AndAlso _options.SkipKnownBadRanges AndAlso allowReadHints
        End Function

        Private Shared Function IsBadRangeMapFresh(map As BadRangeMap, maxAgeDays As Integer) As Boolean
            If map Is Nothing Then
                Return False
            End If

            If maxAgeDays <= 0 Then
                Return True
            End If

            If map.UpdatedUtc = DateTimeOffset.MinValue Then
                Return False
            End If

            Dim age = DateTimeOffset.UtcNow - map.UpdatedUtc
            Return age <= TimeSpan.FromDays(maxAgeDays)
        End Function

        Private Function ApplyKnownBadRangesFromMap(
            sourceRoot As String,
            descriptor As SourceFileDescriptor,
            entry As JournalFileEntry) As Long

            If Not _badRangeMapReadHintsEnabled OrElse descriptor Is Nothing OrElse entry Is Nothing Then
                Return 0L
            End If

            Dim mapEntry As BadRangeMapFileEntry = Nothing
            If Not TryGetBadRangeMapEntry(sourceRoot, descriptor, mapEntry) Then
                Return 0L
            End If

            Dim totalMappedBytes = 0L
            For Each badRange In mapEntry.BadRanges
                If badRange Is Nothing OrElse badRange.Length <= 0 Then
                    Continue For
                End If

                Dim rangeStart = Math.Max(0L, badRange.Offset)
                If rangeStart >= descriptor.Length Then
                    Continue For
                End If

                Dim rangeLength = CInt(Math.Min(CLng(Integer.MaxValue), Math.Min(CLng(badRange.Length), descriptor.Length - rangeStart)))
                If rangeLength <= 0 Then
                    Continue For
                End If

                SetRescueRangeState(entry.RescueRanges, rangeStart, rangeLength, RescueRangeState.KnownBad)
                totalMappedBytes += rangeLength
            Next

            If totalMappedBytes > 0 Then
                EmitLog($"[{RescueCoreName}] Applied bad-range map hints to {descriptor.RelativePath}: {FormatBytes(totalMappedBytes)}.")
            End If

            Return totalMappedBytes
        End Function

        Private Function TryGetBadRangeMapEntry(
            sourceRoot As String,
            descriptor As SourceFileDescriptor,
            ByRef mapEntry As BadRangeMapFileEntry) As Boolean

            mapEntry = Nothing
            If _badRangeMap Is Nothing OrElse _badRangeMap.Files Is Nothing OrElse descriptor Is Nothing Then
                Return False
            End If

            Dim sourceRootKey = NormalizeRootPath(sourceRoot)
            Dim mapSourceRootKey = NormalizeRootPath(_badRangeMap.SourceRoot)
            If mapSourceRootKey.Length > 0 AndAlso
                sourceRootKey.Length > 0 AndAlso
                Not String.Equals(sourceRootKey, mapSourceRootKey, StringComparison.OrdinalIgnoreCase) Then
                Return False
            End If

            Dim relativePath = NormalizeRelativePath(descriptor.RelativePath)
            If relativePath.Length = 0 Then
                Return False
            End If

            Dim candidate As BadRangeMapFileEntry = Nothing
            If Not _badRangeMap.Files.TryGetValue(relativePath, candidate) Then
                Return False
            End If

            If candidate Is Nothing OrElse candidate.BadRanges Is Nothing OrElse candidate.BadRanges.Count = 0 Then
                Return False
            End If

            If Not IsBadRangeMapEntryCompatible(descriptor, candidate) Then
                Return False
            End If

            mapEntry = candidate
            Return True
        End Function

        Private Shared Function IsBadRangeMapEntryCompatible(descriptor As SourceFileDescriptor, mapEntry As BadRangeMapFileEntry) As Boolean
            If descriptor Is Nothing OrElse mapEntry Is Nothing Then
                Return False
            End If

            If mapEntry.SourceLength > 0 AndAlso mapEntry.SourceLength <> descriptor.Length Then
                Return False
            End If

            If mapEntry.LastWriteUtcTicks > 0 AndAlso mapEntry.LastWriteUtcTicks <> descriptor.LastWriteTimeUtc.Ticks Then
                Return False
            End If

            Dim fingerprint = If(mapEntry.FileFingerprint, String.Empty).Trim()
            If fingerprint.Length = 0 Then
                Return True
            End If

            Return String.Equals(
                fingerprint,
                BuildFileFingerprint(descriptor),
                StringComparison.OrdinalIgnoreCase)
        End Function

        Private Async Function TryPersistBadRangeMapEntryAsync(
            sourceRoot As String,
            descriptor As SourceFileDescriptor,
            entry As JournalFileEntry,
            cancellationToken As CancellationToken) As Task

            If descriptor Is Nothing OrElse entry Is Nothing Then
                Return
            End If

            If Not _badRangeMapLoaded OrElse _badRangeMapStore Is Nothing OrElse String.IsNullOrWhiteSpace(_badRangeMapPath) Then
                Return
            End If

            If Not _options.UpdateBadRangeMapFromRun AndAlso _options.OperationMode <> JobOperationMode.ScanOnly Then
                Return
            End If

            If _badRangeMap Is Nothing Then
                Return
            End If

            If _badRangeMap.Files Is Nothing Then
                _badRangeMap.Files = New Dictionary(Of String, BadRangeMapFileEntry)(StringComparer.OrdinalIgnoreCase)
            End If

            Dim relativePath = NormalizeRelativePath(descriptor.RelativePath)
            If relativePath.Length = 0 Then
                Return
            End If

            Dim unreadableRanges = MergeByteRanges(
                SnapshotRangesByState(
                    entry.RescueRanges,
                    RescueRangeState.Bad,
                    RescueRangeState.KnownBad,
                    RescueRangeState.Recovered))

            If unreadableRanges.Count = 0 Then
                _badRangeMap.Files.Remove(relativePath)
            Else
                Dim mapEntry As New BadRangeMapFileEntry() With {
                    .RelativePath = relativePath,
                    .SourceLength = descriptor.Length,
                    .LastWriteUtcTicks = descriptor.LastWriteTimeUtc.Ticks,
                    .FileFingerprint = BuildFileFingerprint(descriptor),
                    .BadRanges = unreadableRanges,
                    .LastScanUtc = DateTimeOffset.UtcNow,
                    .LastError = If(entry.LastError, String.Empty)
                }
                _badRangeMap.Files(relativePath) = mapEntry
            End If

            _badRangeMap.SourceRoot = NormalizeRootPath(sourceRoot)
            _badRangeMap.SourceIdentity = If(_expectedSourceIdentity, String.Empty)
            _badRangeMap.UpdatedUtc = DateTimeOffset.UtcNow
            If _badRangeMap.SchemaVersion <= 0 Then
                _badRangeMap.SchemaVersion = 1
            End If

            Try
                Await _badRangeMapStore.SaveAsync(_badRangeMapPath, _badRangeMap, cancellationToken).ConfigureAwait(False)
            Catch ex As OperationCanceledException
                If cancellationToken.IsCancellationRequested Then
                    Return
                End If
                EmitLog($"Bad-range map save canceled: {ex.Message}")
            Catch ex As Exception
                EmitLog($"Bad-range map save failed: {ex.Message}")
            End Try
        End Function

        Private Async Function FlushBadRangeMapAsync(cancellationToken As CancellationToken) As Task
            If Not _badRangeMapLoaded OrElse _badRangeMap Is Nothing OrElse String.IsNullOrWhiteSpace(_badRangeMapPath) Then
                Return
            End If

            If Not _options.UpdateBadRangeMapFromRun AndAlso _options.OperationMode <> JobOperationMode.ScanOnly Then
                Return
            End If

            Try
                _badRangeMap.UpdatedUtc = DateTimeOffset.UtcNow
                Await _badRangeMapStore.SaveAsync(_badRangeMapPath, _badRangeMap, cancellationToken).ConfigureAwait(False)
            Catch ex As OperationCanceledException
                If cancellationToken.IsCancellationRequested Then
                    Return
                End If
                EmitLog($"Bad-range map final flush canceled: {ex.Message}")
            Catch ex As Exception
                EmitLog($"Bad-range map final flush failed: {ex.Message}")
            End Try
        End Function

        Private Shared Function BuildFileFingerprint(descriptor As SourceFileDescriptor) As String
            If descriptor Is Nothing Then
                Return String.Empty
            End If

            Return $"{descriptor.Length:X16}:{descriptor.LastWriteTimeUtc.Ticks:X16}"
        End Function

        Private Shared Function NormalizeRootPath(pathValue As String) As String
            If String.IsNullOrWhiteSpace(pathValue) Then
                Return String.Empty
            End If

            Try
                Dim fullPath = Path.GetFullPath(pathValue)
                Return TrimTrailingSeparatorsPreservingRoot(fullPath)
            Catch
                Return pathValue.Trim()
            End Try
        End Function

        Private Shared Function TrimTrailingSeparatorsPreservingRoot(pathValue As String) As String
            If String.IsNullOrWhiteSpace(pathValue) Then
                Return String.Empty
            End If

            Dim trimmed = pathValue.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar)
            If trimmed.Length = 0 Then
                Return pathValue
            End If

            Dim root = Path.GetPathRoot(pathValue)
            If String.IsNullOrWhiteSpace(root) Then
                Return trimmed
            End If

            Dim normalizedRoot = root.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar)
            If normalizedRoot.Length = 0 Then
                normalizedRoot = root
            End If

            If String.Equals(trimmed, normalizedRoot, StringComparison.OrdinalIgnoreCase) Then
                Return root
            End If

            Return trimmed
        End Function

        Private Shared Function NormalizeRelativePath(pathValue As String) As String
            Dim value = If(pathValue, String.Empty).Trim()
            If value.Length = 0 Then
                Return String.Empty
            End If

            value = value.Replace(Path.AltDirectorySeparatorChar, Path.DirectorySeparatorChar)
            If value = "." Then
                Return String.Empty
            End If

            Return value.Trim(Path.DirectorySeparatorChar)
        End Function

        Private Shared Sub PrepareDestinationFile(
            destinationPath As String,
            expectedLength As Long,
            existingLength As Long,
            preserveExisting As Boolean)
            If existingLength < 0 Then
                existingLength = 0
            End If

            If existingLength = 0 OrElse Not preserveExisting Then
                Using resetStream = New FileStream(destinationPath, FileMode.Create, FileAccess.Write, FileShare.Read)
                End Using
                Return
            End If

            If existingLength > expectedLength Then
                Using trimStream = New FileStream(destinationPath, FileMode.OpenOrCreate, FileAccess.Write, FileShare.Read)
                    trimStream.SetLength(expectedLength)
                End Using
            End If
        End Sub

        Private Function ShouldUseSmallFileFastPath(fileLength As Long, hasPersistedCoverage As Boolean) As Boolean
            Dim thresholdBytes = Math.Max(MinimumRescueBlockSize, _options.SmallFileThresholdBytes)
            If thresholdBytes <= 0 Then
                Return False
            End If

            If hasPersistedCoverage Then
                Return False
            End If

            Return fileLength > 0 AndAlso fileLength <= thresholdBytes
        End Function

        Private Async Function CopySmallFileFastAsync(
            descriptor As SourceFileDescriptor,
            entry As JournalFileEntry,
            destinationPath As String,
            transferSession As FileTransferSession,
            ioBuffer As Byte(),
            ioBufferSize As Integer,
            progress As ProgressAccumulator,
            journal As JobJournal,
            journalPath As String,
            cancellationToken As CancellationToken) As Task(Of Boolean)

            entry.RescueRanges = New List(Of RescueRange)()
            AppendRescueRange(entry.RescueRanges, 0L, descriptor.Length, RescueRangeState.Pending)
            If entry.RecoveredRanges Is Nothing Then
                entry.RecoveredRanges = New List(Of ByteRange)()
            Else
                entry.RecoveredRanges.Clear()
            End If
            entry.LastRescuePass = "SmallFileFast"
            entry.BytesCopied = 0

            Dim recoveredAny = False
            Dim offset = 0L
            Dim remaining = descriptor.Length
            While remaining > 0
                cancellationToken.ThrowIfCancellationRequested()
                Await _executionControl.WaitIfPausedAsync(cancellationToken).ConfigureAwait(False)
                Await WaitForMediaAvailabilityAsync(_options.SourceRoot, _options.DestinationRoot, cancellationToken).ConfigureAwait(False)

                Dim chunkLength = CInt(Math.Min(CLng(ioBufferSize), remaining))
                Dim bytesRead = Await ReadChunkWithRetriesAsync(
                    descriptor.FullPath,
                    descriptor.RelativePath,
                    offset,
                    chunkLength,
                    ioBufferSize,
                    transferSession,
                    ioBuffer,
                    cancellationToken,
                    maxRetries:=Math.Max(0, _options.MaxRetries),
                    allowSalvage:=False).ConfigureAwait(False)

                Dim segmentState = RescueRangeState.Good
                If bytesRead <= 0 Then
                    If Not _options.SalvageUnreadableBlocks Then
                        Throw New IOException($"Read failed on {descriptor.RelativePath} at offset {offset}.")
                    End If

                    bytesRead = chunkLength
                    FillSalvageBuffer(ioBuffer, bytesRead)
                    segmentState = RescueRangeState.Recovered
                    recoveredAny = True
                    entry.RecoveredRanges.Add(New ByteRange() With {
                        .Offset = offset,
                        .Length = bytesRead
                    })
                    Dim fillDescription = DescribeSalvageFillPattern(_options.SalvageFillPattern)
                    EmitLog($"Recovered unreadable block on {descriptor.RelativePath} at {FormatBytes(offset)} ({bytesRead} bytes {fillDescription}-filled).")
                End If

                Await WriteChunkWithRetriesAsync(
                    destinationPath,
                    descriptor.RelativePath,
                    offset,
                    ioBuffer,
                    bytesRead,
                    ioBufferSize,
                    transferSession,
                    cancellationToken).ConfigureAwait(False)
                Await ApplyThroughputThrottleAsync(bytesRead, cancellationToken).ConfigureAwait(False)

                SetRescueRangeState(entry.RescueRanges, offset, bytesRead, segmentState)
                progress.TotalBytesCopied += bytesRead
                SyncEntryBytesCopied(entry)
                EmitProgress(
                    progress,
                    descriptor.RelativePath,
                    entry.BytesCopied,
                    descriptor.Length,
                    lastChunkBytesTransferred:=bytesRead,
                    bufferSizeBytes:=ioBufferSize,
                    rescuePass:=entry.LastRescuePass,
                    rescueBadRegionCount:=GetUnreadableRegionCount(entry.RescueRanges),
                    rescueRemainingBytes:=GetRescueRangeBytes(entry.RescueRanges, RescueRangeState.Pending, RescueRangeState.Bad, RescueRangeState.KnownBad))

                offset += bytesRead
                remaining -= bytesRead
            End While

            entry.RecoveredRanges = MergeByteRanges(entry.RecoveredRanges)
            SyncEntryBytesCopied(entry)
            Await FlushJournalAsync(journalPath, journal, cancellationToken, force:=True).ConfigureAwait(False)

            Return recoveredAny
        End Function

        Private Function BuildRescueRanges(entry As JournalFileEntry, fileLength As Long, destinationLength As Long) As List(Of RescueRange)
            If fileLength <= 0 Then
                Return New List(Of RescueRange)()
            End If

            Dim ranges As New List(Of RescueRange)()

            If entry IsNot Nothing AndAlso entry.RescueRanges IsNot Nothing AndAlso entry.RescueRanges.Count > 0 Then
                For Each range In entry.RescueRanges
                    If range Is Nothing Then
                        Continue For
                    End If

                    ranges.Add(New RescueRange() With {
                        .Offset = range.Offset,
                        .Length = range.Length,
                        .State = NormalizeRescueRangeState(range.State)
                    })
                Next
            Else
                Dim legacyCopiedBytes = 0L
                If _options.ResumeFromJournal AndAlso entry IsNot Nothing Then
                    legacyCopiedBytes = Math.Max(0L, Math.Min(fileLength, Math.Min(destinationLength, entry.BytesCopied)))
                End If

                If legacyCopiedBytes > 0 Then
                    AppendRescueRange(ranges, 0L, legacyCopiedBytes, RescueRangeState.Good)
                End If

                If legacyCopiedBytes < fileLength Then
                    AppendRescueRange(ranges, legacyCopiedBytes, fileLength - legacyCopiedBytes, RescueRangeState.Pending)
                End If
            End If

            Dim normalized = NormalizeRescueRanges(ranges, fileLength)
            Return DowngradeRescueCoverageBeyondDestination(normalized, destinationLength, fileLength)
        End Function

        Private Shared Function NormalizeRescueRangeState(value As RescueRangeState) As RescueRangeState
            Select Case value
                Case RescueRangeState.Pending, RescueRangeState.Good, RescueRangeState.Bad, RescueRangeState.Recovered, RescueRangeState.KnownBad
                    Return value
                Case Else
                    Return RescueRangeState.Pending
            End Select
        End Function

        Private Shared Function NormalizeRescueRanges(ranges As IEnumerable(Of RescueRange), fileLength As Long) As List(Of RescueRange)
            Dim normalized As New List(Of RescueRange)()
            If fileLength <= 0 Then
                Return normalized
            End If

            Dim orderedRanges = ranges.
                Where(Function(item) item IsNot Nothing AndAlso item.Length > 0).
                Select(
                    Function(item)
                        Dim startOffset = Math.Max(0L, item.Offset)
                        Dim endOffset = Math.Min(fileLength, startOffset + CLng(item.Length))
                        Return New RescueRange() With {
                            .Offset = startOffset,
                            .Length = CInt(Math.Max(0L, endOffset - startOffset)),
                            .State = NormalizeRescueRangeState(item.State)
                        }
                    End Function).
                Where(Function(item) item.Length > 0).
                OrderBy(Function(item) item.Offset).
                ToList()

            Dim cursor = 0L
            For Each item In orderedRanges
                Dim itemStart = Math.Max(cursor, item.Offset)
                Dim itemEnd = Math.Min(fileLength, item.Offset + CLng(item.Length))
                If itemEnd <= itemStart Then
                    Continue For
                End If

                If itemStart > cursor Then
                    AppendRescueRange(normalized, cursor, itemStart - cursor, RescueRangeState.Pending)
                End If

                AppendRescueRange(normalized, itemStart, itemEnd - itemStart, item.State)
                cursor = itemEnd
            Next

            If cursor < fileLength Then
                AppendRescueRange(normalized, cursor, fileLength - cursor, RescueRangeState.Pending)
            End If

            Return MergeRescueRanges(normalized)
        End Function

        Private Shared Function DowngradeRescueCoverageBeyondDestination(
            ranges As List(Of RescueRange),
            destinationLength As Long,
            fileLength As Long) As List(Of RescueRange)

            If fileLength <= 0 Then
                Return New List(Of RescueRange)()
            End If

            Dim effectiveDestinationLength = Math.Max(0L, Math.Min(destinationLength, fileLength))
            Dim adjusted As New List(Of RescueRange)()

            For Each item In ranges
                Dim itemStart = Math.Max(0L, item.Offset)
                Dim itemEnd = Math.Min(fileLength, itemStart + CLng(item.Length))
                If itemEnd <= itemStart Then
                    Continue For
                End If

                Dim state = NormalizeRescueRangeState(item.State)
                If (state = RescueRangeState.Good OrElse state = RescueRangeState.Recovered) AndAlso itemStart >= effectiveDestinationLength Then
                    AppendRescueRange(adjusted, itemStart, itemEnd - itemStart, RescueRangeState.Pending)
                    Continue For
                End If

                If (state = RescueRangeState.Good OrElse state = RescueRangeState.Recovered) AndAlso itemEnd > effectiveDestinationLength Then
                    If effectiveDestinationLength > itemStart Then
                        AppendRescueRange(adjusted, itemStart, effectiveDestinationLength - itemStart, state)
                    End If
                    AppendRescueRange(adjusted, effectiveDestinationLength, itemEnd - effectiveDestinationLength, RescueRangeState.Pending)
                    Continue For
                End If

                AppendRescueRange(adjusted, itemStart, itemEnd - itemStart, state)
            Next

            Return NormalizeRescueRanges(adjusted, fileLength)
        End Function

        Private Shared Sub AppendRescueRange(target As List(Of RescueRange), offset As Long, length As Long, state As RescueRangeState)
            If target Is Nothing OrElse length <= 0 Then
                Return
            End If

            Dim remaining = length
            Dim currentOffset = offset
            While remaining > 0
                Dim segmentLength = CInt(Math.Min(CLng(Integer.MaxValue), remaining))
                target.Add(New RescueRange() With {
                    .Offset = currentOffset,
                    .Length = segmentLength,
                    .State = NormalizeRescueRangeState(state)
                })
                currentOffset += segmentLength
                remaining -= segmentLength
            End While
        End Sub

        Private Shared Function MergeRescueRanges(ranges As IEnumerable(Of RescueRange)) As List(Of RescueRange)
            Dim merged As New List(Of RescueRange)()
            For Each item In ranges.
                Where(Function(value) value IsNot Nothing AndAlso value.Length > 0).
                OrderBy(Function(value) value.Offset)

                Dim state = NormalizeRescueRangeState(item.State)
                Dim start = item.Offset
                Dim [end] = start + CLng(item.Length)
                If [end] <= start Then
                    Continue For
                End If

                If merged.Count = 0 Then
                    merged.Add(New RescueRange() With {
                        .Offset = start,
                        .Length = item.Length,
                        .State = state
                    })
                    Continue For
                End If

                Dim previous = merged(merged.Count - 1)
                Dim previousEnd = previous.Offset + CLng(previous.Length)
                If previous.State = state AndAlso previousEnd = start AndAlso CLng(previous.Length) + CLng(item.Length) <= Integer.MaxValue Then
                    previous.Length += item.Length
                ElseIf previousEnd > start Then
                    Dim overlapAdjustedStart = previousEnd
                    If [end] > overlapAdjustedStart Then
                        AppendRescueRange(merged, overlapAdjustedStart, [end] - overlapAdjustedStart, state)
                    End If
                Else
                    merged.Add(New RescueRange() With {
                        .Offset = start,
                        .Length = item.Length,
                        .State = state
                    })
                End If
            Next

            Return merged
        End Function

        Private Shared Function GetRescueRangeBytes(ranges As IEnumerable(Of RescueRange), ParamArray states() As RescueRangeState) As Long
            If ranges Is Nothing OrElse states Is Nothing OrElse states.Length = 0 Then
                Return 0L
            End If

            Dim allowedStates = New HashSet(Of RescueRangeState)(states.Select(AddressOf NormalizeRescueRangeState))
            Dim totalBytes = 0L
            For Each item In ranges
                If item Is Nothing OrElse item.Length <= 0 Then
                    Continue For
                End If

                If allowedStates.Contains(NormalizeRescueRangeState(item.State)) Then
                    totalBytes += item.Length
                End If
            Next

            Return totalBytes
        End Function

        Private Shared Function GetRescueRegionCount(ranges As IEnumerable(Of RescueRange), ParamArray states() As RescueRangeState) As Integer
            If ranges Is Nothing OrElse states Is Nothing OrElse states.Length = 0 Then
                Return 0
            End If

            Dim stateSet = New HashSet(Of RescueRangeState)(states.Select(AddressOf NormalizeRescueRangeState))
            Return ranges.Count(
                Function(item)
                    Return item IsNot Nothing AndAlso
                        item.Length > 0 AndAlso
                        stateSet.Contains(NormalizeRescueRangeState(item.State))
                End Function)
        End Function

        Private Shared Function SnapshotRangesByState(ranges As IEnumerable(Of RescueRange), ParamArray states() As RescueRangeState) As List(Of ByteRange)
            Dim snapshots As New List(Of ByteRange)()
            If ranges Is Nothing OrElse states Is Nothing OrElse states.Length = 0 Then
                Return snapshots
            End If

            Dim stateSet = New HashSet(Of RescueRangeState)(states.Select(AddressOf NormalizeRescueRangeState))
            For Each item In ranges
                If item Is Nothing OrElse item.Length <= 0 Then
                    Continue For
                End If

                If Not stateSet.Contains(NormalizeRescueRangeState(item.State)) Then
                    Continue For
                End If

                snapshots.Add(New ByteRange() With {
                    .Offset = item.Offset,
                    .Length = item.Length
                })
            Next

            Return snapshots
        End Function

        Private Shared Sub SetRescueRangeState(ranges As List(Of RescueRange), offset As Long, length As Integer, newState As RescueRangeState)
            If ranges Is Nothing OrElse ranges.Count = 0 OrElse length <= 0 Then
                Return
            End If

            Dim targetStart = offset
            Dim targetEnd = offset + CLng(length)
            If targetEnd <= targetStart Then
                Return
            End If

            Dim updated As New List(Of RescueRange)()
            For Each item In ranges
                If item Is Nothing OrElse item.Length <= 0 Then
                    Continue For
                End If

                Dim itemStart = item.Offset
                Dim itemEnd = itemStart + CLng(item.Length)
                If itemEnd <= targetStart OrElse itemStart >= targetEnd Then
                    AppendRescueRange(updated, itemStart, itemEnd - itemStart, item.State)
                    Continue For
                End If

                If itemStart < targetStart Then
                    AppendRescueRange(updated, itemStart, targetStart - itemStart, item.State)
                End If

                Dim overlapStart = Math.Max(itemStart, targetStart)
                Dim overlapEnd = Math.Min(itemEnd, targetEnd)
                AppendRescueRange(updated, overlapStart, overlapEnd - overlapStart, newState)

                If itemEnd > targetEnd Then
                    AppendRescueRange(updated, targetEnd, itemEnd - targetEnd, item.State)
                End If
            Next

            ranges.Clear()
            ranges.AddRange(MergeRescueRanges(updated))
        End Sub

        Private Shared Function GetUnreadableRangeBytes(ranges As IEnumerable(Of RescueRange)) As Long
            Return GetRescueRangeBytes(ranges, RescueRangeState.Bad, RescueRangeState.KnownBad)
        End Function

        Private Shared Function GetUnreadableRegionCount(ranges As IEnumerable(Of RescueRange)) As Integer
            Return GetRescueRegionCount(ranges, RescueRangeState.Bad, RescueRangeState.KnownBad)
        End Function

        Private Sub SyncEntryBytesCopied(entry As JournalFileEntry)
            If entry Is Nothing Then
                Return
            End If

            If _options.OperationMode = JobOperationMode.ScanOnly Then
                entry.BytesCopied = GetRescueRangeBytes(
                    entry.RescueRanges,
                    RescueRangeState.Good,
                    RescueRangeState.Recovered,
                    RescueRangeState.Bad,
                    RescueRangeState.KnownBad)
                Return
            End If

            entry.BytesCopied = GetRescueRangeBytes(entry.RescueRanges, RescueRangeState.Good, RescueRangeState.Recovered)
        End Sub

        Private Function BuildRescuePassPlan(activeBufferSize As Integer) As IReadOnlyList(Of RescuePassDefinition)
            Dim fastAuto = NormalizeIoBufferSize(activeBufferSize)
            Dim trimAuto = NormalizeIoBufferSize(Math.Max(MinimumRescueBlockSize * 16, fastAuto \ 8))
            Dim scrapeAuto = NormalizeIoBufferSize(Math.Max(MinimumRescueBlockSize * 2, trimAuto \ 4))
            Dim retryAuto = NormalizeIoBufferSize(MinimumRescueBlockSize)

            Dim fastChunkSize = ResolveRescueChunkSize(_options.RescueFastScanChunkBytes, fastAuto)
            Dim trimChunkSize = ResolveRescueChunkSize(_options.RescueTrimChunkBytes, trimAuto)
            Dim scrapeChunkSize = ResolveRescueChunkSize(_options.RescueScrapeChunkBytes, scrapeAuto)
            Dim retryChunkSize = ResolveRescueChunkSize(_options.RescueRetryChunkBytes, retryAuto)

            Dim splitMinimum = ResolveRescueSplitMinimum(_options.RescueSplitMinimumBytes, retryChunkSize)
            Dim trimSplitMinimum = CapSplitMinimumForPass(splitMinimum, trimChunkSize)
            Dim scrapeSplitMinimum = CapSplitMinimumForPass(splitMinimum, scrapeChunkSize)

            Dim fastRetries = ResolveRescueRetries(_options.RescueFastScanRetries, 0)
            Dim trimRetries = ResolveRescueRetries(_options.RescueTrimRetries, 1)
            Dim scrapeRetries = ResolveRescueRetries(_options.RescueScrapeRetries, 2)

            Return New List(Of RescuePassDefinition) From {
                New RescuePassDefinition(
                    name:="FastScan",
                    targetState:=RescueRangeState.Pending,
                    chunkSizeBytes:=fastChunkSize,
                    maxReadRetries:=fastRetries,
                    splitOnFailure:=False,
                    minimumSplitBytes:=trimSplitMinimum,
                    processDescending:=False),
                New RescuePassDefinition(
                    name:="TrimSweep",
                    targetState:=RescueRangeState.Bad,
                    chunkSizeBytes:=trimChunkSize,
                    maxReadRetries:=trimRetries,
                    splitOnFailure:=True,
                    minimumSplitBytes:=trimSplitMinimum,
                    processDescending:=False),
                New RescuePassDefinition(
                    name:="TrimSweepReverse",
                    targetState:=RescueRangeState.Bad,
                    chunkSizeBytes:=trimChunkSize,
                    maxReadRetries:=Math.Max(trimRetries, scrapeRetries),
                    splitOnFailure:=True,
                    minimumSplitBytes:=trimSplitMinimum,
                    processDescending:=True),
                New RescuePassDefinition(
                    name:="Scrape",
                    targetState:=RescueRangeState.Bad,
                    chunkSizeBytes:=scrapeChunkSize,
                    maxReadRetries:=scrapeRetries,
                    splitOnFailure:=True,
                    minimumSplitBytes:=scrapeSplitMinimum,
                    processDescending:=False),
                New RescuePassDefinition(
                    name:="RetryBad",
                    targetState:=RescueRangeState.Bad,
                    chunkSizeBytes:=retryChunkSize,
                    maxReadRetries:=Math.Max(0, _options.MaxRetries),
                    splitOnFailure:=False,
                    minimumSplitBytes:=retryChunkSize,
                    processDescending:=True)
            }
        End Function

        Private Shared Function ComputeRescueBadDensity(ranges As IEnumerable(Of RescueRange), totalBytes As Long) As Double
            If totalBytes <= 0 Then
                Return 0.0R
            End If

            Dim badBytes = GetRescueRangeBytes(ranges, RescueRangeState.Bad, RescueRangeState.KnownBad)
            Return Math.Clamp(CDbl(badBytes) / CDbl(totalBytes), 0.0R, 1.0R)
        End Function

        Private Function AdaptRescuePassDefinitionForDensity(
            passDefinition As RescuePassDefinition,
            badDensity As Double,
            activeBufferSize As Integer) As RescuePassDefinition

            If passDefinition Is Nothing Then
                Throw New ArgumentNullException(NameOf(passDefinition))
            End If

            If passDefinition.TargetState <> RescueRangeState.Bad Then
                Return passDefinition
            End If

            Dim tunedChunkSize = passDefinition.ChunkSizeBytes
            Dim tunedRetries = passDefinition.MaxReadRetries

            If badDensity >= 0.25R Then
                tunedChunkSize = Math.Max(
                    passDefinition.MinimumSplitBytes,
                    AlignToBlockSize(Math.Max(MinimumRescueBlockSize, passDefinition.ChunkSizeBytes \ 2), passDefinition.MinimumSplitBytes))
                tunedRetries = Math.Min(32, passDefinition.MaxReadRetries + 2)
            ElseIf badDensity <= 0.02R Then
                Dim passCeiling = NormalizeIoBufferSize(Math.Max(MinimumRescueBlockSize, activeBufferSize))
                Dim boosted = NormalizeIoBufferSize(CInt(Math.Min(CLng(Integer.MaxValue), CLng(passDefinition.ChunkSizeBytes) * 2L)))
                tunedChunkSize = Math.Min(passCeiling, boosted)
            End If

            If tunedChunkSize = passDefinition.ChunkSizeBytes AndAlso tunedRetries = passDefinition.MaxReadRetries Then
                Return passDefinition
            End If

            Return New RescuePassDefinition(
                name:=passDefinition.Name,
                targetState:=passDefinition.TargetState,
                chunkSizeBytes:=tunedChunkSize,
                maxReadRetries:=tunedRetries,
                splitOnFailure:=passDefinition.SplitOnFailure,
                minimumSplitBytes:=passDefinition.MinimumSplitBytes,
                processDescending:=passDefinition.ProcessDescending)
        End Function

        Private Shared Function ResolveRescueChunkSize(configuredValue As Integer, autoValue As Integer) As Integer
            If configuredValue <= 0 Then
                Return NormalizeIoBufferSize(autoValue)
            End If

            Return NormalizeIoBufferSize(configuredValue)
        End Function

        Private Shared Function ResolveRescueSplitMinimum(configuredValue As Integer, fallbackValue As Integer) As Integer
            If configuredValue <= 0 Then
                Return NormalizeIoBufferSize(fallbackValue)
            End If

            Return NormalizeIoBufferSize(configuredValue)
        End Function

        Private Shared Function ResolveRescueRetries(configuredValue As Integer, fallbackValue As Integer) As Integer
            If configuredValue < 0 Then
                Return Math.Max(0, fallbackValue)
            End If

            Return configuredValue
        End Function

        Private Shared Function CapSplitMinimumForPass(splitMinimum As Integer, passChunkSize As Integer) As Integer
            Dim halfPassChunk = Math.Max(MinimumRescueBlockSize, passChunkSize \ 2)
            Return Math.Max(MinimumRescueBlockSize, Math.Min(splitMinimum, halfPassChunk))
        End Function

        Private Shared Function AlignToBlockSize(value As Integer, blockSize As Integer) As Integer
            Dim normalizedBlock = Math.Max(MinimumRescueBlockSize, blockSize)
            Dim aligned = (value \ normalizedBlock) * normalizedBlock
            If aligned <= 0 Then
                Return normalizedBlock
            End If
            Return aligned
        End Function

        Private Async Function ExecuteRescuePassAsync(
            passDefinition As RescuePassDefinition,
            descriptor As SourceFileDescriptor,
            entry As JournalFileEntry,
            destinationPath As String,
            transferSession As FileTransferSession,
            ioBuffer As Byte(),
            progress As ProgressAccumulator,
            journal As JobJournal,
            journalPath As String,
            cancellationToken As CancellationToken,
            Optional writeRecovered As Boolean = True,
            Optional countFailedAsProcessed As Boolean = False) As Task(Of RescuePassOutcome)

            Dim outcome As New RescuePassOutcome()
            Dim targetSnapshots = SnapshotRangesByState(entry.RescueRanges, passDefinition.TargetState)
            If passDefinition.ProcessDescending Then
                targetSnapshots.Reverse()
            End If

            For Each targetRange In targetSnapshots
                cancellationToken.ThrowIfCancellationRequested()
                Await _executionControl.WaitIfPausedAsync(cancellationToken).ConfigureAwait(False)
                Await WaitForMediaAvailabilityAsync(_options.SourceRoot, _options.DestinationRoot, cancellationToken).ConfigureAwait(False)

                Dim work As New Stack(Of ByteRange)()
                work.Push(New ByteRange() With {
                    .Offset = targetRange.Offset,
                    .Length = targetRange.Length
                })

                While work.Count > 0
                    cancellationToken.ThrowIfCancellationRequested()
                    Await _executionControl.WaitIfPausedAsync(cancellationToken).ConfigureAwait(False)
                    Await WaitForMediaAvailabilityAsync(_options.SourceRoot, _options.DestinationRoot, cancellationToken).ConfigureAwait(False)

                    Dim segment = work.Pop()
                    If segment Is Nothing OrElse segment.Length <= 0 Then
                        Continue While
                    End If

                    Dim segmentLength = segment.Length
                    If segmentLength > passDefinition.ChunkSizeBytes Then
                        If passDefinition.ProcessDescending Then
                            Dim emitted = 0
                            While emitted < segmentLength
                                Dim chunkLength = Math.Min(passDefinition.ChunkSizeBytes, segmentLength - emitted)
                                work.Push(New ByteRange() With {
                                    .Offset = segment.Offset + emitted,
                                    .Length = chunkLength
                                })
                                emitted += chunkLength
                            End While
                        Else
                            Dim remaining = segmentLength
                            While remaining > 0
                                Dim chunkLength = Math.Min(passDefinition.ChunkSizeBytes, remaining)
                                work.Push(New ByteRange() With {
                                    .Offset = segment.Offset + (remaining - chunkLength),
                                    .Length = chunkLength
                                })
                                remaining -= chunkLength
                            End While
                        End If
                        Continue While
                    End If

                    outcome.AttemptedSegments += 1
                    Dim bytesRead = Await ReadChunkWithRetriesAsync(
                        descriptor.FullPath,
                        descriptor.RelativePath,
                        segment.Offset,
                        segmentLength,
                        passDefinition.ChunkSizeBytes,
                        transferSession,
                        ioBuffer,
                        cancellationToken,
                        maxRetries:=passDefinition.MaxReadRetries,
                        allowSalvage:=False).ConfigureAwait(False)

                    If bytesRead > 0 Then
                        If writeRecovered Then
                            Await WriteChunkWithRetriesAsync(
                                destinationPath,
                                descriptor.RelativePath,
                                segment.Offset,
                                ioBuffer,
                                bytesRead,
                                passDefinition.ChunkSizeBytes,
                                transferSession,
                                cancellationToken).ConfigureAwait(False)
                            Await ApplyThroughputThrottleAsync(bytesRead, cancellationToken).ConfigureAwait(False)
                        Else
                            Await ApplyThroughputThrottleAsync(bytesRead, cancellationToken).ConfigureAwait(False)
                        End If

                        SetRescueRangeState(entry.RescueRanges, segment.Offset, bytesRead, RescueRangeState.Good)
                        progress.TotalBytesCopied += bytesRead
                        SyncEntryBytesCopied(entry)
                        EmitProgress(
                            progress,
                            descriptor.RelativePath,
                            entry.BytesCopied,
                            descriptor.Length,
                            lastChunkBytesTransferred:=bytesRead,
                            bufferSizeBytes:=passDefinition.ChunkSizeBytes,
                            rescuePass:=passDefinition.Name,
                            rescueBadRegionCount:=GetUnreadableRegionCount(entry.RescueRanges),
                            rescueRemainingBytes:=GetRescueRangeBytes(entry.RescueRanges, RescueRangeState.Pending, RescueRangeState.Bad, RescueRangeState.KnownBad))
                        Await FlushJournalAsync(journalPath, journal, cancellationToken, force:=False).ConfigureAwait(False)
                        outcome.RecoveredBytes += bytesRead
                        Continue While
                    End If

                    If passDefinition.SplitOnFailure AndAlso segmentLength >= passDefinition.MinimumSplitBytes * 2 Then
                        Dim halfLength = Math.Max(passDefinition.MinimumSplitBytes, AlignToBlockSize(segmentLength \ 2, passDefinition.MinimumSplitBytes))
                        halfLength = Math.Min(segmentLength - passDefinition.MinimumSplitBytes, halfLength)
                        If halfLength > 0 AndAlso halfLength < segmentLength Then
                            Dim secondLength = segmentLength - halfLength
                            If passDefinition.ProcessDescending Then
                                work.Push(New ByteRange() With {
                                    .Offset = segment.Offset,
                                    .Length = halfLength
                                })
                                work.Push(New ByteRange() With {
                                    .Offset = segment.Offset + halfLength,
                                    .Length = secondLength
                                })
                            Else
                                work.Push(New ByteRange() With {
                                    .Offset = segment.Offset + halfLength,
                                    .Length = secondLength
                                })
                                work.Push(New ByteRange() With {
                                    .Offset = segment.Offset,
                                    .Length = halfLength
                                })
                            End If
                            Continue While
                        End If
                    End If

                    SetRescueRangeState(entry.RescueRanges, segment.Offset, segmentLength, RescueRangeState.Bad)
                    If countFailedAsProcessed Then
                        progress.TotalBytesCopied += segmentLength
                    End If
                    SyncEntryBytesCopied(entry)
                    EmitProgress(
                        progress,
                        descriptor.RelativePath,
                        entry.BytesCopied,
                        descriptor.Length,
                        lastChunkBytesTransferred:=0,
                        bufferSizeBytes:=passDefinition.ChunkSizeBytes,
                        rescuePass:=passDefinition.Name,
                        rescueBadRegionCount:=GetUnreadableRegionCount(entry.RescueRanges),
                        rescueRemainingBytes:=GetRescueRangeBytes(entry.RescueRanges, RescueRangeState.Pending, RescueRangeState.Bad, RescueRangeState.KnownBad))
                    Await FlushJournalAsync(journalPath, journal, cancellationToken, force:=False).ConfigureAwait(False)
                    outcome.FailedSegments += 1
                End While
            Next

            Return outcome
        End Function

        Private Shared Function ShouldLogRescuePassSummary(
            passDefinition As RescuePassDefinition,
            outcome As RescuePassOutcome,
            remainingBadRegions As Integer) As Boolean

            If passDefinition Is Nothing OrElse outcome Is Nothing Then
                Return False
            End If

            If String.Equals(passDefinition.Name, "FastScan", StringComparison.OrdinalIgnoreCase) Then
                Return outcome.FailedSegments > 0
            End If

            Return outcome.AttemptedSegments > 0 OrElse
                outcome.RecoveredBytes > 0 OrElse
                remainingBadRegions > 0
        End Function

        Private Async Function SalvageRemainingRangesAsync(
            descriptor As SourceFileDescriptor,
            entry As JournalFileEntry,
            destinationPath As String,
            transferSession As FileTransferSession,
            ioBuffer As Byte(),
            ioBufferSize As Integer,
            progress As ProgressAccumulator,
            journal As JobJournal,
            journalPath As String,
            cancellationToken As CancellationToken) As Task(Of Boolean)

            Dim recoveredAny = False
            Dim badSnapshots = SnapshotRangesByState(entry.RescueRanges, RescueRangeState.Bad, RescueRangeState.KnownBad)
            For Each badRange In badSnapshots
                Dim offset = badRange.Offset
                Dim remaining = badRange.Length
                While remaining > 0
                    cancellationToken.ThrowIfCancellationRequested()
                    Await _executionControl.WaitIfPausedAsync(cancellationToken).ConfigureAwait(False)
                    Await WaitForMediaAvailabilityAsync(_options.SourceRoot, _options.DestinationRoot, cancellationToken).ConfigureAwait(False)

                    Dim chunkLength = Math.Min(ioBufferSize, remaining)
                    FillSalvageBuffer(ioBuffer, chunkLength)
                    Await WriteChunkWithRetriesAsync(
                        destinationPath,
                        descriptor.RelativePath,
                        offset,
                        ioBuffer,
                        chunkLength,
                        ioBufferSize,
                        transferSession,
                        cancellationToken).ConfigureAwait(False)
                    Await ApplyThroughputThrottleAsync(chunkLength, cancellationToken).ConfigureAwait(False)

                    SetRescueRangeState(entry.RescueRanges, offset, chunkLength, RescueRangeState.Recovered)
                    entry.RecoveredRanges.Add(New ByteRange() With {
                        .Offset = offset,
                        .Length = chunkLength
                    })

                    progress.TotalBytesCopied += chunkLength
                    SyncEntryBytesCopied(entry)
                    EmitProgress(
                        progress,
                        descriptor.RelativePath,
                        entry.BytesCopied,
                        descriptor.Length,
                        lastChunkBytesTransferred:=chunkLength,
                        bufferSizeBytes:=ioBufferSize,
                        rescuePass:="SalvageFill",
                        rescueBadRegionCount:=GetUnreadableRegionCount(entry.RescueRanges),
                        rescueRemainingBytes:=GetRescueRangeBytes(entry.RescueRanges, RescueRangeState.Pending, RescueRangeState.Bad, RescueRangeState.KnownBad))
                    Await FlushJournalAsync(journalPath, journal, cancellationToken, force:=False).ConfigureAwait(False)

                    offset += chunkLength
                    remaining -= chunkLength
                    recoveredAny = True
                End While
            Next

            Return recoveredAny
        End Function

        Private Shared Function MergeByteRanges(ranges As IEnumerable(Of ByteRange)) As List(Of ByteRange)
            Dim merged As New List(Of ByteRange)()
            If ranges Is Nothing Then
                Return merged
            End If

            For Each item In ranges.
                Where(Function(value) value IsNot Nothing AndAlso value.Length > 0).
                OrderBy(Function(value) value.Offset)

                Dim start = item.Offset
                Dim [end] = start + CLng(item.Length)
                If [end] <= start Then
                    Continue For
                End If

                If merged.Count = 0 Then
                    merged.Add(New ByteRange() With {.Offset = start, .Length = item.Length})
                    Continue For
                End If

                Dim last = merged(merged.Count - 1)
                Dim lastEnd = last.Offset + CLng(last.Length)
                If start <= lastEnd AndAlso [end] > lastEnd Then
                    last.Length = CInt(Math.Min(CLng(Integer.MaxValue), [end] - last.Offset))
                ElseIf start > lastEnd Then
                    merged.Add(New ByteRange() With {.Offset = start, .Length = item.Length})
                End If
            Next

            Return merged
        End Function

        Private NotInheritable Class RescuePassDefinition
            Public Sub New(
                name As String,
                targetState As RescueRangeState,
                chunkSizeBytes As Integer,
                maxReadRetries As Integer,
                splitOnFailure As Boolean,
                minimumSplitBytes As Integer,
                processDescending As Boolean)

                Me.Name = If(String.IsNullOrWhiteSpace(name), "Pass", name.Trim())
                Me.TargetState = NormalizeRescueRangeState(targetState)
                Me.ChunkSizeBytes = NormalizeIoBufferSize(chunkSizeBytes)
                Me.MaxReadRetries = Math.Max(0, maxReadRetries)
                Me.SplitOnFailure = splitOnFailure
                Me.MinimumSplitBytes = NormalizeIoBufferSize(Math.Max(MinimumRescueBlockSize, minimumSplitBytes))
                Me.ProcessDescending = processDescending
            End Sub

            Public ReadOnly Property Name As String
            Public ReadOnly Property TargetState As RescueRangeState
            Public ReadOnly Property ChunkSizeBytes As Integer
            Public ReadOnly Property MaxReadRetries As Integer
            Public ReadOnly Property SplitOnFailure As Boolean
            Public ReadOnly Property MinimumSplitBytes As Integer
            Public ReadOnly Property ProcessDescending As Boolean
        End Class

        Private NotInheritable Class RescuePassOutcome
            Public Property AttemptedSegments As Integer
            Public Property RecoveredBytes As Long
            Public Property FailedSegments As Integer
        End Class

        Private Shared Function GetExistingFileLength(path As String) As Long
            If Not File.Exists(path) Then
                Return 0
            End If

            Try
                Dim info As New FileInfo(path)
                Return info.Length
            Catch
                Return 0
            End Try
        End Function

        Private Async Function ReadChunkWithRetriesAsync(
            sourcePath As String,
            relativePath As String,
            offset As Long,
            length As Integer,
            ioBufferSize As Integer,
            transferSession As FileTransferSession,
            buffer As Byte(),
            cancellationToken As CancellationToken,
            Optional maxRetries As Integer = -1,
            Optional allowSalvage As Boolean = True) As Task(Of Integer)

            If buffer Is Nothing Then
                Throw New ArgumentNullException(NameOf(buffer))
            End If

            If length <= 0 OrElse length > buffer.Length Then
                Throw New ArgumentOutOfRangeException(NameOf(length))
            End If

            Dim attempt = 0
            Dim lastError As Exception = Nothing
            Dim effectiveMaxRetries = If(maxRetries >= 0, maxRetries, _options.MaxRetries)

            While attempt <= effectiveMaxRetries
                cancellationToken.ThrowIfCancellationRequested()
                ThrowIfMediaIdentityMismatch(sourcePath, _expectedSourceIdentity, isSource:=True)

                Try
#If DEBUG Then
                    Dim injectedReadFault As Exception = Nothing
                    If _devFaultInjector IsNot Nothing Then
                        injectedReadFault = _devFaultInjector.CreateReadFault(relativePath, offset, length)
                    End If
                    If injectedReadFault IsNot Nothing Then
                        Throw injectedReadFault
                    End If
#End If
                    Using linkedCts = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken)
                        linkedCts.CancelAfter(_options.OperationTimeout)
                        Dim bytesRead = Await ReadChunkOnceAsync(
                            sourcePath,
                            offset,
                            buffer,
                            length,
                            ioBufferSize,
                            transferSession,
                            linkedCts.Token).ConfigureAwait(False)

                        If bytesRead <> length Then
                            Throw New IOException($"Short read at offset {offset}. Expected {length}, got {bytesRead}.")
                        End If

                        Return bytesRead
                    End Using
                Catch ex As OperationCanceledException When Not cancellationToken.IsCancellationRequested
                    lastError = New TimeoutException($"Read timeout at offset {offset}.", ex)
                Catch ex As Exception
                    lastError = ex
                End Try

                If IsSourceFileMissingException(lastError) Then
                    Select Case _options.SourceMutationPolicy
                        Case SourceMutationPolicy.SkipFile
                            Throw New SourceMutationSkippedException(
                                $"Source file disappeared during copy: {relativePath}.",
                                lastError)
                        Case SourceMutationPolicy.WaitForReappearance
                            transferSession?.InvalidateSource()
                            EmitLog($"Source file disappeared during copy of {relativePath}. Waiting for it to reappear.")
                            Await WaitForSourceFileAsync(sourcePath, cancellationToken, allowWithoutMediaWait:=True).ConfigureAwait(False)
                            attempt = 0
                            Continue While
                        Case Else
                            Throw New IOException($"Source file disappeared during copy: {relativePath}.", lastError)
                    End Select
                End If

                If IsReadContentionException(lastError) AndAlso _options.WaitForFileLockRelease Then
                    transferSession?.InvalidateSource()
                    EmitLog($"Source contention detected on {relativePath}; waiting for lock release.")
                    Await WaitForSourceReadAccessAsync(sourcePath, cancellationToken).ConfigureAwait(False)
                    attempt = 0
                    Continue While
                End If

                If IsFatalReadException(lastError, _options.TreatAccessDeniedAsContention) Then
                    Exit While
                End If

                If _options.WaitForMediaAvailability AndAlso IsAvailabilityRelatedException(lastError, includeFileNotFound:=False) Then
                    transferSession?.InvalidateSource()
                    EmitLog($"Source unavailable during read of {relativePath}. Waiting for media to return.")
                    Await WaitForSourceFileAsync(sourcePath, cancellationToken).ConfigureAwait(False)
                    attempt = 0
                    Continue While
                End If

                transferSession?.InvalidateSource()

                If IsReadContentionException(lastError) Then
                    EmitLog($"Source file contention detected on {relativePath}; retrying.")
                End If

                attempt += 1
                If attempt > effectiveMaxRetries Then
                    Exit While
                End If

                EmitLog($"Read retry {attempt}/{effectiveMaxRetries} on {relativePath} at {FormatBytes(offset)}: {lastError.Message}")
                Await DelayForRetryAsync(attempt, cancellationToken).ConfigureAwait(False)
            End While

            If IsFatalReadException(lastError, _options.TreatAccessDeniedAsContention) Then
                Throw New IOException($"Read failed on {relativePath} at offset {offset}.", lastError)
            End If

            If IsAvailabilityRelatedException(lastError) Then
                Throw New IOException($"Read failed on {relativePath} at offset {offset} because source is unavailable.", lastError)
            End If

            If allowSalvage AndAlso _options.SalvageUnreadableBlocks Then
                FillSalvageBuffer(buffer, length)
                Dim fillDescription = DescribeSalvageFillPattern(_options.SalvageFillPattern)
                EmitLog($"Recovered unreadable block on {relativePath} at {FormatBytes(offset)} ({length} bytes {fillDescription}-filled).")
                Return length
            End If

            Return -1
        End Function

        Private Async Function ReadChunkOnceAsync(
            sourcePath As String,
            offset As Long,
            buffer As Byte(),
            count As Integer,
            ioBufferSize As Integer,
            transferSession As FileTransferSession,
            cancellationToken As CancellationToken) As Task(Of Integer)

            If transferSession Is Nothing Then
                Using stream = New FileStream(
                    sourcePath,
                    FileMode.Open,
                    FileAccess.Read,
                    FileShare.ReadWrite Or FileShare.Delete,
                    NormalizeIoBufferSize(ioBufferSize),
                    FileOptions.Asynchronous Or FileOptions.SequentialScan)

                    stream.Seek(offset, SeekOrigin.Begin)
                    Return Await ReadExactlyAsync(stream, buffer, count, cancellationToken).ConfigureAwait(False)
                End Using
            End If

            Dim sourceHandle = transferSession.GetSourceHandle()
            Return Await ReadExactlyAsync(sourceHandle, offset, buffer, count, cancellationToken).ConfigureAwait(False)
        End Function

        Private Shared Async Function ReadExactlyAsync(
            stream As FileStream,
            buffer As Byte(),
            count As Integer,
            cancellationToken As CancellationToken) As Task(Of Integer)

            Dim totalRead = 0
            While totalRead < count
                Dim readNow = Await stream.ReadAsync(buffer.AsMemory(totalRead, count - totalRead), cancellationToken).ConfigureAwait(False)
                If readNow = 0 Then
                    Exit While
                End If

                totalRead += readNow
            End While

            Return totalRead
        End Function

        Private Shared Async Function ReadExactlyAsync(
            handle As SafeFileHandle,
            offset As Long,
            buffer As Byte(),
            count As Integer,
            cancellationToken As CancellationToken) As Task(Of Integer)

            Dim totalRead = 0
            While totalRead < count
                Dim readNow = Await RandomAccess.ReadAsync(
                    handle,
                    buffer.AsMemory(totalRead, count - totalRead),
                    offset + totalRead,
                    cancellationToken).ConfigureAwait(False)
                If readNow = 0 Then
                    Exit While
                End If

                totalRead += readNow
            End While

            Return totalRead
        End Function

        Private Async Function WriteChunkWithRetriesAsync(
            destinationPath As String,
            relativePath As String,
            offset As Long,
            buffer As Byte(),
            count As Integer,
            ioBufferSize As Integer,
            transferSession As FileTransferSession,
            cancellationToken As CancellationToken) As Task

            Dim attempt = 0
            Dim lastError As Exception = Nothing

            While attempt <= _options.MaxRetries
                cancellationToken.ThrowIfCancellationRequested()
                ThrowIfMediaIdentityMismatch(destinationPath, _expectedDestinationIdentity, isSource:=False)

                Try
#If DEBUG Then
                    Dim injectedWriteFault As Exception = Nothing
                    If _devFaultInjector IsNot Nothing Then
                        injectedWriteFault = _devFaultInjector.CreateWriteFault(relativePath, offset, count)
                    End If
                    If injectedWriteFault IsNot Nothing Then
                        Throw injectedWriteFault
                    End If
#End If
                    Using linkedCts = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken)
                        linkedCts.CancelAfter(_options.OperationTimeout)
                        Await WriteChunkOnceAsync(
                            destinationPath,
                            offset,
                            buffer,
                            count,
                            ioBufferSize,
                            transferSession,
                            linkedCts.Token).ConfigureAwait(False)
                        Return
                    End Using
                Catch ex As OperationCanceledException When Not cancellationToken.IsCancellationRequested
                    lastError = New TimeoutException($"Write timeout at offset {offset}.", ex)
                Catch ex As Exception
                    lastError = ex
                End Try

                If IsDiskFullException(lastError) Then
                    transferSession?.InvalidateDestination()
                    Throw New IOException(
                        $"Destination is out of free space while writing {relativePath} at offset {offset}.",
                        lastError)
                End If

                If IsAccessDeniedException(lastError) AndAlso Not _options.TreatAccessDeniedAsContention Then
                    transferSession?.InvalidateDestination()
                    Throw New IOException(
                        $"Destination access denied while writing {relativePath} at offset {offset}.",
                        lastError)
                End If

                If IsWriteContentionException(lastError) AndAlso _options.WaitForFileLockRelease Then
                    transferSession?.InvalidateDestination()
                    EmitLog($"Destination contention detected on {relativePath}; waiting for lock release.")
                    Await WaitForDestinationWriteAccessAsync(destinationPath, cancellationToken).ConfigureAwait(False)
                    attempt = 0
                    Continue While
                End If

                If _options.WaitForMediaAvailability AndAlso IsAvailabilityRelatedException(lastError) Then
                    transferSession?.InvalidateDestination()
                    EmitLog($"Destination unavailable during write of {relativePath}. Waiting for media to return.")
                    Await WaitForDestinationPathAsync(destinationPath, cancellationToken).ConfigureAwait(False)
                    attempt = 0
                    Continue While
                End If

                transferSession?.InvalidateDestination()

                If IsWriteContentionException(lastError) Then
                    EmitLog($"Destination file contention detected on {relativePath}; retrying.")
                End If

                attempt += 1
                If attempt > _options.MaxRetries Then
                    Exit While
                End If

                EmitLog($"Write retry {attempt}/{_options.MaxRetries} on {relativePath} at {FormatBytes(offset)}: {lastError.Message}")
                Await DelayForRetryAsync(attempt, cancellationToken).ConfigureAwait(False)
            End While

            Throw New IOException($"Write failed on {relativePath} at offset {offset}.", lastError)
        End Function

        Private Async Function WriteChunkOnceAsync(
            destinationPath As String,
            offset As Long,
            buffer As Byte(),
            count As Integer,
            ioBufferSize As Integer,
            transferSession As FileTransferSession,
            cancellationToken As CancellationToken) As Task

            If transferSession Is Nothing Then
                Using stream = New FileStream(
                    destinationPath,
                    FileMode.OpenOrCreate,
                    FileAccess.Write,
                    FileShare.Read,
                    NormalizeIoBufferSize(ioBufferSize),
                    FileOptions.Asynchronous Or FileOptions.SequentialScan)

                    stream.Seek(offset, SeekOrigin.Begin)
                    Await stream.WriteAsync(buffer.AsMemory(0, count), cancellationToken).ConfigureAwait(False)
                End Using
                Return
            End If

            Dim destinationHandle = transferSession.GetDestinationHandle()
            Await RandomAccess.WriteAsync(
                destinationHandle,
                buffer.AsMemory(0, count),
                offset,
                cancellationToken).ConfigureAwait(False)
        End Function

        Private Async Function VerifyFileWithRetriesAsync(
            sourcePath As String,
            destinationPath As String,
            relativePath As String,
            verificationMode As VerificationMode,
            cancellationToken As CancellationToken) As Task

            Select Case verificationMode
                Case VerificationMode.Sampled
                    EmitLog($"Verifying (sampled): {relativePath}")
                    Await VerifySampledFileWithRetriesAsync(sourcePath, destinationPath, relativePath, cancellationToken).ConfigureAwait(False)
                    Return

                Case VerificationMode.Full
                    EmitLog($"Verifying (full hash): {relativePath}")
                    Dim sourceHash = Await ComputeHashWithRetriesAsync(sourcePath, $"source {relativePath}", cancellationToken).ConfigureAwait(False)
                    Dim destinationHash = Await ComputeHashWithRetriesAsync(destinationPath, $"destination {relativePath}", cancellationToken).ConfigureAwait(False)

                    If Not sourceHash.SequenceEqual(destinationHash) Then
                        Throw New IOException($"Verification hash mismatch: {relativePath}")
                    End If
            End Select
        End Function

        Private Async Function VerifySampledFileWithRetriesAsync(
            sourcePath As String,
            destinationPath As String,
            relativePath As String,
            cancellationToken As CancellationToken) As Task

            Dim sourceLength = New FileInfo(sourcePath).Length
            Dim destinationLength = New FileInfo(destinationPath).Length
            If sourceLength <> destinationLength Then
                Throw New IOException($"Verification length mismatch: {relativePath}")
            End If

            If sourceLength <= 0 Then
                Return
            End If

            Dim chunkSize = CInt(Math.Min(sourceLength, CLng(Math.Max(4096, _options.SampleVerificationChunkBytes))))
            Dim chunkCount = Math.Max(1, _options.SampleVerificationChunkCount)
            Dim sampleOffsets = BuildSampleOffsets(sourceLength, chunkSize, chunkCount)

            For Each sampleOffset In sampleOffsets
                cancellationToken.ThrowIfCancellationRequested()

                Dim sampleLength = CInt(Math.Min(CLng(chunkSize), sourceLength - sampleOffset))
                Dim sourceSampleHash = Await ReadAndHashChunkWithRetriesAsync(
                    sourcePath,
                    sampleOffset,
                    sampleLength,
                    $"source sample {relativePath} at {FormatBytes(sampleOffset)}",
                    cancellationToken).ConfigureAwait(False)

                Dim destinationSampleHash = Await ReadAndHashChunkWithRetriesAsync(
                    destinationPath,
                    sampleOffset,
                    sampleLength,
                    $"destination sample {relativePath} at {FormatBytes(sampleOffset)}",
                    cancellationToken).ConfigureAwait(False)

                If Not sourceSampleHash.SequenceEqual(destinationSampleHash) Then
                    Throw New IOException($"Sample verification mismatch: {relativePath} at offset {sampleOffset}.")
                End If
            Next
        End Function

        Private Shared Function BuildSampleOffsets(fileLength As Long, chunkSize As Integer, chunkCount As Integer) As List(Of Long)
            Dim offsets As New List(Of Long)()
            Dim effectiveChunkSize = Math.Max(1, chunkSize)
            Dim effectiveCount = Math.Max(1, chunkCount)

            If fileLength <= effectiveChunkSize OrElse effectiveCount = 1 Then
                offsets.Add(0)
                Return offsets
            End If

            Dim maxOffset = fileLength - effectiveChunkSize
            Dim stepSize = CDbl(maxOffset) / CDbl(effectiveCount - 1)
            For sampleIndex = 0 To effectiveCount - 1
                Dim offsetValue = CLng(Math.Round(stepSize * sampleIndex))
                offsets.Add(Math.Max(0L, Math.Min(maxOffset, offsetValue)))
            Next

            Return offsets.Distinct().OrderBy(Function(value) value).ToList()
        End Function

        Private Async Function ReadAndHashChunkWithRetriesAsync(
            filePath As String,
            offset As Long,
            count As Integer,
            context As String,
            cancellationToken As CancellationToken) As Task(Of Byte())

            Dim attempt = 0
            Dim lastError As Exception = Nothing
            Dim rentedBuffer = ArrayPool(Of Byte).Shared.Rent(Math.Max(count, MinimumRescueBlockSize))

            Try
                While attempt <= _options.MaxRetries
                    cancellationToken.ThrowIfCancellationRequested()

                    Try
                        Using linkedCts = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken)
                            linkedCts.CancelAfter(_options.OperationTimeout)

                            Dim bytesRead = Await ReadChunkOnceAsync(
                                filePath,
                                offset,
                                rentedBuffer,
                                count,
                                _options.BufferSizeBytes,
                                Nothing,
                                linkedCts.Token).ConfigureAwait(False)

                            If bytesRead <> count Then
                                Throw New IOException($"Short read while sampling {context}. Expected {count}, got {bytesRead}.")
                            End If

                            Return ComputeHash(rentedBuffer, bytesRead)
                        End Using
                    Catch ex As OperationCanceledException When Not cancellationToken.IsCancellationRequested
                        lastError = New TimeoutException($"Verification timeout for {context}.", ex)
                    Catch ex As Exception
                        lastError = ex
                    End Try

                    If IsReadContentionException(lastError) AndAlso _options.WaitForFileLockRelease Then
                        EmitLog($"Verification contention detected on {context}; waiting for lock release.")
                        Await WaitForSourceReadAccessAsync(filePath, cancellationToken).ConfigureAwait(False)
                        attempt = 0
                        Continue While
                    End If

                    attempt += 1
                    If attempt > _options.MaxRetries Then
                        Exit While
                    End If

                    EmitLog($"Verification retry {attempt}/{_options.MaxRetries} on {context}: {lastError.Message}")
                    Await DelayForRetryAsync(attempt, cancellationToken).ConfigureAwait(False)
                End While
            Finally
                ArrayPool(Of Byte).Shared.Return(rentedBuffer, clearArray:=False)
            End Try

            Throw New IOException($"Unable to verify sampled chunk for {context}.", lastError)
        End Function

        Private Async Function ComputeHashWithRetriesAsync(
            filePath As String,
            context As String,
            cancellationToken As CancellationToken) As Task(Of Byte())

            Dim attempt = 0
            Dim lastError As Exception = Nothing

            While attempt <= _options.MaxRetries
                cancellationToken.ThrowIfCancellationRequested()

                Try
                    Dim dynamicTimeout = EstimateHashTimeout(filePath)

                    Using linkedCts = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken)
                        linkedCts.CancelAfter(dynamicTimeout)

                        Using stream = New FileStream(
                            filePath,
                            FileMode.Open,
                            FileAccess.Read,
                            FileShare.ReadWrite Or FileShare.Delete,
                            _options.BufferSizeBytes,
                            FileOptions.Asynchronous Or FileOptions.SequentialScan)

                            Return Await ComputeStreamHashAsync(stream, linkedCts.Token).ConfigureAwait(False)
                        End Using
                    End Using
                Catch ex As OperationCanceledException When Not cancellationToken.IsCancellationRequested
                    lastError = New TimeoutException($"Hash timeout for {context}.", ex)
                Catch ex As Exception
                    lastError = ex
                End Try

                If IsReadContentionException(lastError) AndAlso _options.WaitForFileLockRelease Then
                    EmitLog($"Hash contention detected on {context}; waiting for lock release.")
                    Await WaitForSourceReadAccessAsync(filePath, cancellationToken).ConfigureAwait(False)
                    attempt = 0
                    Continue While
                End If

                attempt += 1
                If attempt > _options.MaxRetries Then
                    Exit While
                End If

                EmitLog($"Hash retry {attempt}/{_options.MaxRetries} on {context}: {lastError.Message}")
                Await DelayForRetryAsync(attempt, cancellationToken).ConfigureAwait(False)
            End While

            Throw New IOException($"Unable to compute hash for {context}.", lastError)
        End Function

        Private Function ComputeHash(buffer As Byte(), count As Integer) As Byte()
            If count < 0 Then
                Throw New ArgumentOutOfRangeException(NameOf(count))
            End If

            If count = 0 Then
                Select Case _options.VerificationHashAlgorithm
                    Case VerificationHashAlgorithm.Sha512
                        Return SHA512.HashData(Array.Empty(Of Byte)())
                    Case Else
                        Return SHA256.HashData(Array.Empty(Of Byte)())
                End Select
            End If

            Select Case _options.VerificationHashAlgorithm
                Case VerificationHashAlgorithm.Sha512
                    Return SHA512.HashData(buffer.AsSpan(0, count))
                Case Else
                    Return SHA256.HashData(buffer.AsSpan(0, count))
            End Select
        End Function

        Private Async Function ComputeStreamHashAsync(stream As Stream, cancellationToken As CancellationToken) As Task(Of Byte())
            Select Case _options.VerificationHashAlgorithm
                Case VerificationHashAlgorithm.Sha512
                    Return Await SHA512.HashDataAsync(stream, cancellationToken).ConfigureAwait(False)
                Case Else
                    Return Await SHA256.HashDataAsync(stream, cancellationToken).ConfigureAwait(False)
            End Select
        End Function

        Private Shared Function EstimateHashTimeout(filePath As String) As TimeSpan
            Dim fileLength As Long
            Try
                fileLength = New FileInfo(filePath).Length
            Catch
                fileLength = 0
            End Try

            Const assumedBytesPerSecond As Long = 8L * 1024L * 1024L
            Dim estimatedSeconds = Math.Ceiling(CDbl(fileLength) / assumedBytesPerSecond)
            Dim timeoutSeconds = Math.Max(30.0R, Math.Min(3600.0R, estimatedSeconds + 10.0R))
            Return TimeSpan.FromSeconds(timeoutSeconds)
        End Function

        Private Async Function DelayForRetryAsync(attempt As Integer, cancellationToken As CancellationToken) As Task
            Dim boundedAttempt = Math.Min(attempt, 20)
            Dim baseDelayMs = _options.InitialRetryDelay.TotalMilliseconds * Math.Pow(2.0R, boundedAttempt - 1)
            Dim delayMs = CInt(Math.Min(_options.MaxRetryDelay.TotalMilliseconds, baseDelayMs))
            Await Task.Delay(delayMs, cancellationToken).ConfigureAwait(False)
        End Function

        Private Sub EmitDestinationCapacityWarning(destinationRoot As String, totalBytes As Long)
            If totalBytes <= 0 OrElse String.IsNullOrWhiteSpace(destinationRoot) Then
                Return
            End If

            Try
                Dim fullPath = Path.GetFullPath(destinationRoot)
                Dim root = Path.GetPathRoot(fullPath)
                If String.IsNullOrWhiteSpace(root) OrElse root.StartsWith("\\", StringComparison.OrdinalIgnoreCase) Then
                    Return
                End If

                Dim drive = New DriveInfo(root)
                If Not drive.IsReady Then
                    Return
                End If

                If drive.AvailableFreeSpace < totalBytes Then
                    EmitLog(
                        $"Destination free-space warning: estimated {FormatBytes(totalBytes)} required, {FormatBytes(drive.AvailableFreeSpace)} available.")
                End If
            Catch
                ' Ignore preflight failures and continue copy flow.
            End Try
        End Sub

        Private Iterator Function GetJournalLoadCandidates(defaultPath As String, activePath As String) As IEnumerable(Of String)
            Dim seen As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

            Dim hintedPath = If(_options.ResumeJournalPathHint, String.Empty).Trim()
            If hintedPath.Length > 0 AndAlso seen.Add(hintedPath) Then
                Yield hintedPath
            End If

            If Not String.IsNullOrWhiteSpace(activePath) AndAlso seen.Add(activePath) Then
                Yield activePath
            End If

            If Not String.IsNullOrWhiteSpace(defaultPath) AndAlso seen.Add(defaultPath) Then
                Yield defaultPath
            End If
        End Function

        Private Function ResolveWritableJournalPath(jobId As String, defaultJournalPath As String) As String
            If IsJournalPathWritable(defaultJournalPath) Then
                Return defaultJournalPath
            End If

            Dim fallbackDirectory = Path.Combine(Path.GetTempPath(), "XactCopy", "journals")
            Dim fallbackPath = Path.Combine(fallbackDirectory, $"job-{jobId}.json")
            If IsJournalPathWritable(fallbackPath) Then
                EmitLog($"Journal path fallback active: {fallbackPath}")
                Return fallbackPath
            End If

            Throw New IOException(
                $"Unable to write journal at default path '{defaultJournalPath}' or fallback '{fallbackPath}'.")
        End Function

        Private Shared Function IsJournalPathWritable(journalPath As String) As Boolean
            If String.IsNullOrWhiteSpace(journalPath) Then
                Return False
            End If

            Try
                Dim directoryPath = Path.GetDirectoryName(journalPath)
                If String.IsNullOrWhiteSpace(directoryPath) Then
                    Return False
                End If

                Directory.CreateDirectory(directoryPath)
                Dim probePath = Path.Combine(directoryPath, $".xactcopy-probe-{Guid.NewGuid():N}.tmp")
                File.WriteAllText(probePath, "ok")
                File.Delete(probePath)
                Return True
            Catch
                Return False
            End Try
        End Function

        Private Async Function FlushJournalAsync(
            journalPath As String,
            journal As JobJournal,
            cancellationToken As CancellationToken,
            force As Boolean) As Task

            Dim now = DateTimeOffset.UtcNow
            If Not force AndAlso (now - _lastJournalFlushUtc).TotalSeconds < JournalFlushIntervalSeconds Then
                Return
            End If

            Await _journalStore.SaveAsync(journalPath, journal, cancellationToken).ConfigureAwait(False)
            _lastJournalFlushUtc = now
        End Function

        Private Sub EmitProgress(
            progress As ProgressAccumulator,
            currentFile As String,
            currentFileBytes As Long,
            currentFileTotal As Long,
            lastChunkBytesTransferred As Integer,
            bufferSizeBytes As Integer,
            Optional rescuePass As String = "",
            Optional rescueBadRegionCount As Integer = 0,
            Optional rescueRemainingBytes As Long = 0)
            Dim snapshot As New CopyProgressSnapshot() With {
                .CurrentFile = currentFile,
                .CurrentFileBytesCopied = currentFileBytes,
                .CurrentFileBytesTotal = currentFileTotal,
                .TotalBytesCopied = progress.TotalBytesCopied,
                .TotalBytes = progress.TotalBytes,
                .LastChunkBytesTransferred = lastChunkBytesTransferred,
                .BufferSizeBytes = bufferSizeBytes,
                .CompletedFiles = progress.CompletedFiles,
                .FailedFiles = progress.FailedFiles,
                .RecoveredFiles = progress.RecoveredFiles,
                .SkippedFiles = progress.SkippedFiles,
                .TotalFiles = progress.TotalFiles,
                .RescuePass = If(rescuePass, String.Empty),
                .RescueBadRegionCount = Math.Max(0, rescueBadRegionCount),
                .RescueRemainingBytes = Math.Max(0L, rescueRemainingBytes)
            }
            RaiseEvent ProgressChanged(Me, snapshot)
        End Sub

        Private Sub InitializeMediaIdentityExpectations(sourceRoot As String, destinationRoot As String)
            _expectedSourceIdentity = NormalizeIdentity(_options.ExpectedSourceIdentity)
            _expectedDestinationIdentity = NormalizeIdentity(_options.ExpectedDestinationIdentity)

            If _expectedSourceIdentity.Length = 0 Then
                Dim sourceIdentity = ResolveMediaIdentity(sourceRoot)
                If sourceIdentity.Length > 0 Then
                    _expectedSourceIdentity = sourceIdentity
                    _options.ExpectedSourceIdentity = sourceIdentity
                    EmitLog("Source media identity baseline captured.")
                End If
            End If

            If _expectedDestinationIdentity.Length = 0 Then
                Dim destinationIdentity = ResolveMediaIdentity(destinationRoot)
                If destinationIdentity.Length > 0 Then
                    _expectedDestinationIdentity = destinationIdentity
                    _options.ExpectedDestinationIdentity = destinationIdentity
                    EmitLog("Destination media identity baseline captured.")
                End If
            End If
        End Sub

        Private Sub EnsureMediaIdentityIntegrity(sourceRoot As String, destinationRoot As String)
            ThrowIfMediaIdentityMismatch(sourceRoot, _expectedSourceIdentity, isSource:=True)
            ThrowIfMediaIdentityMismatch(destinationRoot, _expectedDestinationIdentity, isSource:=False)
        End Sub

        Private Function IsSourceAvailable(sourceRoot As String) As Boolean
            If String.IsNullOrWhiteSpace(sourceRoot) Then
                Return False
            End If

            If Not Directory.Exists(sourceRoot) Then
                Return False
            End If

            Return IsMediaIdentityAccepted(sourceRoot, _expectedSourceIdentity, isSource:=True, allowAutoCapture:=True)
        End Function

        Private Function IsMediaIdentityAccepted(pathValue As String, ByRef expectedIdentity As String, isSource As Boolean, allowAutoCapture As Boolean) As Boolean
            expectedIdentity = NormalizeIdentity(expectedIdentity)
            Dim currentIdentity = ResolveMediaIdentity(pathValue)
            If currentIdentity.Length = 0 Then
                Return expectedIdentity.Length = 0
            End If

            If expectedIdentity.Length = 0 Then
                If allowAutoCapture Then
                    expectedIdentity = currentIdentity
                    If isSource Then
                        _expectedSourceIdentity = currentIdentity
                        _options.ExpectedSourceIdentity = currentIdentity
                    Else
                        _expectedDestinationIdentity = currentIdentity
                        _options.ExpectedDestinationIdentity = currentIdentity
                    End If
                    EmitLog($"{If(isSource, "Source", "Destination")} media identity baseline captured.")
                End If

                Return True
            End If

            If AreMediaIdentitiesEquivalent(expectedIdentity, currentIdentity) Then
                Return True
            End If

            EmitMediaIdentityMismatch(isSource, expectedIdentity, currentIdentity, pathValue)
            Return False
        End Function

        Private Sub ThrowIfMediaIdentityMismatch(pathValue As String, expectedIdentity As String, isSource As Boolean)
            Dim expected = NormalizeIdentity(expectedIdentity)
            If expected.Length = 0 Then
                Return
            End If

            Dim currentIdentity = ResolveMediaIdentity(pathValue)
            If currentIdentity.Length = 0 Then
                Return
            End If

            If AreMediaIdentitiesEquivalent(expected, currentIdentity) Then
                Return
            End If

            Throw New IOException(
                $"{If(isSource, "Source", "Destination")} media identity mismatch. Expected '{expected}', found '{currentIdentity}'.")
        End Sub

        Private Function ResolveMediaIdentity(pathValue As String) As String
            If String.IsNullOrWhiteSpace(pathValue) Then
                Return String.Empty
            End If

            Dim fullPath As String
            Try
                fullPath = System.IO.Path.GetFullPath(pathValue)
            Catch
                Return String.Empty
            End Try

            Dim root = System.IO.Path.GetPathRoot(fullPath)
            If String.IsNullOrWhiteSpace(root) Then
                Return String.Empty
            End If

            Dim normalizedRoot = root.TrimEnd(System.IO.Path.DirectorySeparatorChar, System.IO.Path.AltDirectorySeparatorChar)
            If normalizedRoot.StartsWith("\\", StringComparison.OrdinalIgnoreCase) Then
                Dim uncShare = NormalizeUncShareRoot(normalizedRoot)
                If uncShare.Length = 0 Then
                    Return String.Empty
                End If

                Return "unc:" & uncShare.ToUpperInvariant()
            End If

            Dim serial As UInteger
            If Not TryGetVolumeSerial(root, serial) Then
                Return String.Empty
            End If

            Return $"vol:{serial:X8}"
        End Function

        Private Shared Function AreMediaIdentitiesEquivalent(expectedIdentity As String, currentIdentity As String) As Boolean
            Dim expected = NormalizeIdentity(expectedIdentity)
            Dim current = NormalizeIdentity(currentIdentity)
            If expected.Length = 0 OrElse current.Length = 0 Then
                Return expected.Length = current.Length
            End If

            If String.Equals(expected, current, StringComparison.OrdinalIgnoreCase) Then
                Return True
            End If

            Dim expectedSerial As String = Nothing
            Dim currentSerial As String = Nothing
            If TryExtractVolumeSerial(expected, expectedSerial) AndAlso
                TryExtractVolumeSerial(current, currentSerial) Then
                Return String.Equals(expectedSerial, currentSerial, StringComparison.OrdinalIgnoreCase)
            End If

            Return False
        End Function

        Private Shared Function TryExtractVolumeSerial(identity As String, ByRef serial As String) As Boolean
            serial = String.Empty
            Dim normalized = NormalizeIdentity(identity)
            If Not normalized.StartsWith("vol:", StringComparison.OrdinalIgnoreCase) Then
                Return False
            End If

            Dim parts = normalized.Split(":"c, StringSplitOptions.RemoveEmptyEntries)
            If parts.Length < 2 Then
                Return False
            End If

            Dim candidate = parts(parts.Length - 1).Trim()
            Dim parsed As UInteger
            If Not UInteger.TryParse(candidate, NumberStyles.HexNumber, CultureInfo.InvariantCulture, parsed) Then
                Return False
            End If

            serial = parsed.ToString("X8", CultureInfo.InvariantCulture)
            Return True
        End Function

        Private Shared Function NormalizeUncShareRoot(value As String) As String
            If String.IsNullOrWhiteSpace(value) Then
                Return String.Empty
            End If

            Dim trimmed = value.Trim().Trim("\"c)
            If trimmed.Length = 0 Then
                Return String.Empty
            End If

            Dim parts = trimmed.Split("\"c, StringSplitOptions.RemoveEmptyEntries)
            If parts.Length < 2 Then
                Return String.Empty
            End If

            Return "\\" & parts(0) & "\" & parts(1)
        End Function

        Private Shared Function TryGetVolumeSerial(rootPath As String, ByRef serial As UInteger) As Boolean
            serial = 0UI
            If String.IsNullOrWhiteSpace(rootPath) Then
                Return False
            End If

            Dim normalizedRoot = rootPath
            If Not normalizedRoot.EndsWith(Path.DirectorySeparatorChar.ToString(), StringComparison.Ordinal) AndAlso
                Not normalizedRoot.EndsWith(Path.AltDirectorySeparatorChar.ToString(), StringComparison.Ordinal) Then
                normalizedRoot &= Path.DirectorySeparatorChar
            End If

            Dim maxComponentLength As UInteger = 0UI
            Dim fileSystemFlags As UInteger = 0UI

            Return GetVolumeInformation(
                normalizedRoot,
                IntPtr.Zero,
                0UI,
                serial,
                maxComponentLength,
                fileSystemFlags,
                IntPtr.Zero,
                0UI)
        End Function

        Private Shared Function NormalizeIdentity(value As String) As String
            Return If(value, String.Empty).Trim()
        End Function

        Private Sub EmitMediaIdentityMismatch(isSource As Boolean, expectedIdentity As String, currentIdentity As String, path As String)
            Dim nowUtc = DateTimeOffset.UtcNow
            If isSource Then
                If _lastSourceIdentityMismatchLogUtc <> DateTimeOffset.MinValue AndAlso
                    (nowUtc - _lastSourceIdentityMismatchLogUtc) < TimeSpan.FromSeconds(IdentityMismatchLogIntervalSeconds) Then
                    Return
                End If

                _lastSourceIdentityMismatchLogUtc = nowUtc
            Else
                If _lastDestinationIdentityMismatchLogUtc <> DateTimeOffset.MinValue AndAlso
                    (nowUtc - _lastDestinationIdentityMismatchLogUtc) < TimeSpan.FromSeconds(IdentityMismatchLogIntervalSeconds) Then
                    Return
                End If

                _lastDestinationIdentityMismatchLogUtc = nowUtc
            End If

            EmitLog(
                $"{If(isSource, "Source", "Destination")} media identity mismatch on '{path}'. Expected {expectedIdentity}, found {currentIdentity}.")
        End Sub

        Private Async Function WaitForMediaAvailabilityAsync(
            sourceRoot As String,
            destinationRoot As String,
            cancellationToken As CancellationToken) As Task

            If Not _options.WaitForMediaAvailability Then
                Return
            End If

            Dim lastLogUtc = DateTimeOffset.MinValue
            Do
                cancellationToken.ThrowIfCancellationRequested()
                Await _executionControl.WaitIfPausedAsync(cancellationToken).ConfigureAwait(False)

                Dim sourceAvailable = IsSourceAvailable(sourceRoot)
                Dim destinationAvailable = IsDestinationAvailable(destinationRoot)
                If sourceAvailable AndAlso destinationAvailable Then
                    Return
                End If

                Dim nowUtc = DateTimeOffset.UtcNow
                If (nowUtc - lastLogUtc).TotalSeconds >= 5 Then
                    EmitLog($"Waiting for media availability. Source={If(sourceAvailable, "online", "offline")} Destination={If(destinationAvailable, "online", "offline")}")
                    lastLogUtc = nowUtc
                End If

                Await Task.Delay(TimeSpan.FromSeconds(1), cancellationToken).ConfigureAwait(False)
            Loop
        End Function

        Private Async Function WaitForSourceFileAsync(
            sourcePath As String,
            cancellationToken As CancellationToken,
            Optional allowWithoutMediaWait As Boolean = False) As Task

            If Not _options.WaitForMediaAvailability AndAlso Not allowWithoutMediaWait Then
                Return
            End If

            Dim lastLogUtc = DateTimeOffset.MinValue
            Do
                cancellationToken.ThrowIfCancellationRequested()
                Await _executionControl.WaitIfPausedAsync(cancellationToken).ConfigureAwait(False)

                Dim directoryPath = Path.GetDirectoryName(sourcePath)
                Dim fileAvailable = File.Exists(sourcePath)
                Dim directoryAvailable = Not String.IsNullOrWhiteSpace(directoryPath) AndAlso Directory.Exists(directoryPath)
                Dim identityAccepted = IsMediaIdentityAccepted(sourcePath, _expectedSourceIdentity, isSource:=True, allowAutoCapture:=True)

                If fileAvailable AndAlso identityAccepted Then
                    Return
                End If

                Dim nowUtc = DateTimeOffset.UtcNow
                If (nowUtc - lastLogUtc).TotalSeconds >= 5 Then
                    EmitLog($"Waiting for source path: {sourcePath} (directory {(If(directoryAvailable, "online", "offline"))}, identity {(If(identityAccepted, "ok", "mismatch/offline"))})")
                    lastLogUtc = nowUtc
                End If

                Await Task.Delay(TimeSpan.FromSeconds(1), cancellationToken).ConfigureAwait(False)
            Loop
        End Function

        Private Async Function WaitForDestinationPathAsync(destinationPath As String, cancellationToken As CancellationToken) As Task
            If Not _options.WaitForMediaAvailability Then
                Return
            End If

            Dim lastLogUtc = DateTimeOffset.MinValue
            Dim directoryPath = Path.GetDirectoryName(destinationPath)
            Dim probePath = If(String.IsNullOrWhiteSpace(directoryPath), destinationPath, directoryPath)

            Do
                cancellationToken.ThrowIfCancellationRequested()
                Await _executionControl.WaitIfPausedAsync(cancellationToken).ConfigureAwait(False)

                If IsDestinationAvailable(probePath) Then
                    Return
                End If

                Dim nowUtc = DateTimeOffset.UtcNow
                If (nowUtc - lastLogUtc).TotalSeconds >= 5 Then
                    EmitLog($"Waiting for destination path: {probePath}")
                    lastLogUtc = nowUtc
                End If

                Await Task.Delay(TimeSpan.FromSeconds(1), cancellationToken).ConfigureAwait(False)
            Loop
        End Function

        Private Async Function WaitForSourceReadAccessAsync(sourcePath As String, cancellationToken As CancellationToken) As Task
            Await WaitForPathAccessAsync(
                Function()
                    If Not File.Exists(sourcePath) Then
                        Return False
                    End If

                    Using stream = New FileStream(
                        sourcePath,
                        FileMode.Open,
                        FileAccess.Read,
                        FileShare.ReadWrite Or FileShare.Delete,
                        MinimumRescueBlockSize,
                        FileOptions.None)
                    End Using
                    Return True
                End Function,
                $"Waiting for source lock/access release: {sourcePath}",
                cancellationToken).ConfigureAwait(False)
        End Function

        Private Async Function WaitForDestinationWriteAccessAsync(destinationPath As String, cancellationToken As CancellationToken) As Task
            Await WaitForPathAccessAsync(
                Function()
                    Dim directoryPath = Path.GetDirectoryName(destinationPath)
                    If Not String.IsNullOrWhiteSpace(directoryPath) Then
                        Directory.CreateDirectory(directoryPath)
                    End If

                    Using stream = New FileStream(
                        destinationPath,
                        FileMode.OpenOrCreate,
                        FileAccess.Write,
                        FileShare.Read,
                        MinimumRescueBlockSize,
                        FileOptions.None)
                    End Using
                    Return True
                End Function,
                $"Waiting for destination lock/access release: {destinationPath}",
                cancellationToken).ConfigureAwait(False)
        End Function

        Private Async Function WaitForPathAccessAsync(
            probe As Func(Of Boolean),
            logText As String,
            cancellationToken As CancellationToken) As Task

            Dim lastLogUtc = DateTimeOffset.MinValue
            Dim waitInterval = _options.LockContentionProbeInterval
            If waitInterval <= TimeSpan.Zero Then
                waitInterval = TimeSpan.FromMilliseconds(500)
            End If

            Do
                cancellationToken.ThrowIfCancellationRequested()
                Await _executionControl.WaitIfPausedAsync(cancellationToken).ConfigureAwait(False)

                Dim ready = False
                Try
                    ready = probe()
                Catch
                    ready = False
                End Try

                If ready Then
                    Return
                End If

                If _options.WaitForMediaAvailability Then
                    Try
                        Await WaitForMediaAvailabilityAsync(_options.SourceRoot, _options.DestinationRoot, cancellationToken).ConfigureAwait(False)
                    Catch
                    End Try
                End If

                Dim nowUtc = DateTimeOffset.UtcNow
                If (nowUtc - lastLogUtc).TotalSeconds >= 5 Then
                    EmitLog(logText)
                    lastLogUtc = nowUtc
                End If

                Await Task.Delay(waitInterval, cancellationToken).ConfigureAwait(False)
            Loop
        End Function

        Private Function IsDestinationAvailable(targetPath As String) As Boolean
            If String.IsNullOrWhiteSpace(targetPath) Then
                Return False
            End If

            Try
                Dim fullPath = Path.GetFullPath(targetPath)
                Dim root = Path.GetPathRoot(fullPath)
                If String.IsNullOrWhiteSpace(root) OrElse Not Directory.Exists(root) Then
                    Return False
                End If

                If Not IsMediaIdentityAccepted(fullPath, _expectedDestinationIdentity, isSource:=False, allowAutoCapture:=True) Then
                    Return False
                End If

                Dim probePath = fullPath
                If File.Exists(probePath) Then
                    Return True
                End If

                If Directory.Exists(probePath) Then
                    Return True
                End If

                Dim parentPath = Path.GetDirectoryName(probePath)
                While Not String.IsNullOrWhiteSpace(parentPath)
                    If Directory.Exists(parentPath) Then
                        Return True
                    End If

                    Dim nextParent = Path.GetDirectoryName(parentPath)
                    If String.Equals(nextParent, parentPath, StringComparison.OrdinalIgnoreCase) Then
                        Exit While
                    End If

                    parentPath = nextParent
                End While
            Catch
            End Try

            Return False
        End Function

        Private Shared Function IsAvailabilityRelatedException(ex As Exception, Optional includeFileNotFound As Boolean = True) As Boolean
            If ex Is Nothing Then
                Return False
            End If

            For Each candidate In EnumerateExceptionChain(ex)
                If TypeOf candidate Is DriveNotFoundException OrElse TypeOf candidate Is DirectoryNotFoundException Then
                    Return True
                End If

                Dim win32Code As Integer
                If TryGetWin32ErrorCode(candidate, win32Code) Then
                    Select Case win32Code
                        Case ErrorPathNotFound, ErrorNotReady, ErrorBadNetPath, ErrorNetNameDeleted, ErrorBadNetName, ErrorDeviceNotConnected
                            Return True
                        Case ErrorFileNotFound
                            If includeFileNotFound Then
                                Return True
                            End If
                    End Select
                End If

                Dim ioEx = TryCast(candidate, IOException)
                If ioEx IsNot Nothing Then
                    Dim message = ioEx.Message.ToLowerInvariant()
                    If message.Contains("device is not ready") OrElse
                        message.Contains("network path") OrElse
                        message.Contains("network name") OrElse
                        message.Contains("specified network") OrElse
                        message.Contains("media identity mismatch") Then
                        Return True
                    End If

                    If includeFileNotFound AndAlso
                        (message.Contains("cannot find the path") OrElse message.Contains("file not found")) Then
                        Return True
                    End If
                End If
            Next

            Return False
        End Function

        Private Shared Function IsDiskFullException(ex As Exception) As Boolean
            If ex Is Nothing Then
                Return False
            End If

            For Each candidate In EnumerateExceptionChain(ex)
                Dim win32Code As Integer
                If TryGetWin32ErrorCode(candidate, win32Code) Then
                    If win32Code = ErrorDiskFull OrElse win32Code = ErrorHandleDiskFull Then
                        Return True
                    End If
                End If

                Dim ioEx = TryCast(candidate, IOException)
                If ioEx IsNot Nothing Then
                    Dim message = ioEx.Message.ToLowerInvariant()
                    If message.Contains("not enough space") OrElse
                        message.Contains("disk full") OrElse
                        message.Contains("there is not enough space") Then
                        Return True
                    End If
                End If
            Next

            Return False
        End Function

        Private Shared Function IsAccessDeniedException(ex As Exception) As Boolean
            If ex Is Nothing Then
                Return False
            End If

            For Each candidate In EnumerateExceptionChain(ex)
                If TypeOf candidate Is UnauthorizedAccessException Then
                    Return True
                End If

                Dim win32Code As Integer
                If TryGetWin32ErrorCode(candidate, win32Code) AndAlso win32Code = ErrorAccessDenied Then
                    Return True
                End If
            Next

            Return False
        End Function

        Private Shared Function IsFileLockException(ex As Exception) As Boolean
            If ex Is Nothing Then
                Return False
            End If

            For Each candidate In EnumerateExceptionChain(ex)
                Dim win32Code As Integer
                If TryGetWin32ErrorCode(candidate, win32Code) Then
                    If win32Code = ErrorSharingViolation OrElse win32Code = ErrorLockViolation Then
                        Return True
                    End If
                End If
            Next

            Return False
        End Function

        Private Shared Function IsSourceFileMissingException(ex As Exception) As Boolean
            If ex Is Nothing Then
                Return False
            End If

            For Each candidate In EnumerateExceptionChain(ex)
                If TypeOf candidate Is FileNotFoundException Then
                    Return True
                End If

                Dim win32Code As Integer
                If TryGetWin32ErrorCode(candidate, win32Code) Then
                    If win32Code = ErrorFileNotFound OrElse win32Code = ErrorPathNotFound Then
                        Return True
                    End If
                End If
            Next

            Return False
        End Function

        Private Function IsReadContentionException(ex As Exception) As Boolean
            Return IsFileLockException(ex) OrElse (_options.TreatAccessDeniedAsContention AndAlso IsAccessDeniedException(ex))
        End Function

        Private Function IsWriteContentionException(ex As Exception) As Boolean
            Return IsFileLockException(ex) OrElse (_options.TreatAccessDeniedAsContention AndAlso IsAccessDeniedException(ex))
        End Function

        Private Shared Iterator Function EnumerateExceptionChain(ex As Exception) As IEnumerable(Of Exception)
            Dim current = ex
            While current IsNot Nothing
                Yield current
                current = current.InnerException
            End While
        End Function

        Private Shared Function TryGetWin32ErrorCode(ex As Exception, ByRef code As Integer) As Boolean
            code = 0
            If ex Is Nothing Then
                Return False
            End If

            Dim hresult = ex.HResult
            Dim facility = (hresult >> 16) And &H1FFF
            If facility = 7 Then
                code = hresult And &HFFFF
                Return True
            End If

            Return False
        End Function

        Private Shared Function IsFatalReadException(ex As Exception, treatAccessDeniedAsContention As Boolean) As Boolean
            If ex Is Nothing Then
                Return False
            End If

            If Not treatAccessDeniedAsContention AndAlso IsAccessDeniedException(ex) Then
                Return True
            End If

            For Each candidate In EnumerateExceptionChain(ex)
                If TypeOf candidate Is NotSupportedException OrElse
                    TypeOf candidate Is System.Security.SecurityException OrElse
                    TypeOf candidate Is ArgumentException Then
                    Return True
                End If
            Next

            Return False
        End Function

        Private Function ResolveVerificationMode() As VerificationMode
            If _options.VerificationMode <> VerificationMode.None Then
                Return _options.VerificationMode
            End If

            If _options.VerifyAfterCopy Then
                Return VerificationMode.Full
            End If

            Return VerificationMode.None
        End Function

        Private Async Function ApplyThroughputThrottleAsync(bytesTransferred As Integer, cancellationToken As CancellationToken) As Task
            If _options.MaxThroughputBytesPerSecond <= 0 OrElse bytesTransferred <= 0 Then
                Return
            End If

            Dim nowUtc = DateTimeOffset.UtcNow
            If _throttleWindowStartUtc = DateTimeOffset.MinValue Then
                _throttleWindowStartUtc = nowUtc
                _throttleWindowBytes = 0
            End If

            _throttleWindowBytes += bytesTransferred

            Dim expectedElapsedSeconds = CDbl(_throttleWindowBytes) / CDbl(_options.MaxThroughputBytesPerSecond)
            Dim actualElapsedSeconds = (nowUtc - _throttleWindowStartUtc).TotalSeconds
            If expectedElapsedSeconds > actualElapsedSeconds Then
                Dim delayMilliseconds = CInt(Math.Ceiling((expectedElapsedSeconds - actualElapsedSeconds) * 1000.0R))
                If delayMilliseconds > 0 Then
                    Await Task.Delay(delayMilliseconds, cancellationToken).ConfigureAwait(False)
                End If
            End If

            Dim elapsedAfterDelaySeconds = (DateTimeOffset.UtcNow - _throttleWindowStartUtc).TotalSeconds
            If elapsedAfterDelaySeconds >= 5.0R Then
                _throttleWindowStartUtc = DateTimeOffset.UtcNow
                _throttleWindowBytes = 0
            End If
        End Function

        Private Sub FillSalvageBuffer(buffer As Byte(), count As Integer)
            If buffer Is Nothing Then
                Throw New ArgumentNullException(NameOf(buffer))
            End If

            If count <= 0 OrElse count > buffer.Length Then
                Throw New ArgumentOutOfRangeException(NameOf(count))
            End If

            Select Case _options.SalvageFillPattern
                Case SalvageFillPattern.Ones
                    buffer.AsSpan(0, count).Fill(&HFF)
                Case SalvageFillPattern.Random
                    RandomNumberGenerator.Fill(buffer.AsSpan(0, count))
                Case Else
                    Array.Clear(buffer, 0, count)
            End Select
        End Sub

        Private Shared Function DescribeSalvageFillPattern(pattern As SalvageFillPattern) As String
            Select Case pattern
                Case SalvageFillPattern.Ones
                    Return "0xFF"
                Case SalvageFillPattern.Random
                    Return "random"
                Case Else
                    Return "zero"
            End Select
        End Function

        Private Function ResolveBufferSizeForFile(fileLength As Long) As Integer
            Dim configured = NormalizeIoBufferSize(_options.BufferSizeBytes)
            If Not _options.UseAdaptiveBufferSizing Then
                Return configured
            End If

            Dim minimumAdaptive = Math.Min(32 * 1024, configured)
            Dim suggested = Math.Max(CLng(minimumAdaptive), fileLength \ 16L)
            Dim clamped = Math.Min(CLng(configured), suggested)
            Dim aligned = Math.Max(4096L, (clamped \ 4096L) * 4096L)
            Return NormalizeIoBufferSize(CInt(aligned))
        End Function

        Private Shared Function NormalizeIoBufferSize(value As Integer) As Integer
            Return Math.Min(MaximumIoBufferSize, Math.Max(4096, value))
        End Function

        Private Sub EmitLog(message As String)
            RaiseEvent LogMessage(Me, $"[{DateTimeOffset.Now:HH:mm:ss}] {message}")
        End Sub

        Private Shared Function FormatBytes(value As Long) As String
            If value < 1024 Then
                Return $"{value} B"
            End If

            Dim units = {"KB", "MB", "GB", "TB", "PB"}
            Dim size = CDbl(value)
            Dim unitIndex = -1

            Do
                size /= 1024.0R
                unitIndex += 1
            Loop While size >= 1024.0R AndAlso unitIndex < units.Length - 1

            Return $"{size:0.##} {units(unitIndex)}"
        End Function

        Private Shared Function CreateResult(
            progress As ProgressAccumulator,
            journalPath As String,
            succeeded As Boolean,
            cancelled As Boolean,
            errorMessage As String) As CopyJobResult

            Return New CopyJobResult() With {
                .Succeeded = succeeded,
                .Cancelled = cancelled,
                .TotalFiles = progress.TotalFiles,
                .CompletedFiles = progress.CompletedFiles,
                .FailedFiles = progress.FailedFiles,
                .RecoveredFiles = progress.RecoveredFiles,
                .SkippedFiles = progress.SkippedFiles,
                .TotalBytes = progress.TotalBytes,
                .CopiedBytes = progress.TotalBytesCopied,
                .JournalPath = journalPath,
                .ErrorMessage = errorMessage
            }
        End Function

        Private NotInheritable Class FileTransferSession
            Implements IDisposable

            Private ReadOnly _sourcePath As String
            Private ReadOnly _destinationPath As String
            Private ReadOnly _ioBufferSize As Integer
            Private _sourceStream As FileStream
            Private _destinationStream As FileStream
            Private _disposed As Boolean

            Public Sub New(sourcePath As String, destinationPath As String, ioBufferSize As Integer)
                _sourcePath = sourcePath
                _destinationPath = destinationPath
                _ioBufferSize = NormalizeIoBufferSize(ioBufferSize)
            End Sub

            Public Function GetSourceHandle() As SafeFileHandle
                ThrowIfDisposed()
                If _sourceStream Is Nothing Then
                    _sourceStream = New FileStream(
                        _sourcePath,
                        FileMode.Open,
                        FileAccess.Read,
                        FileShare.ReadWrite Or FileShare.Delete,
                        _ioBufferSize,
                        FileOptions.Asynchronous Or FileOptions.SequentialScan)
                End If

                Return _sourceStream.SafeFileHandle
            End Function

            Public Function GetDestinationHandle() As SafeFileHandle
                ThrowIfDisposed()
                If _destinationStream Is Nothing Then
                    _destinationStream = New FileStream(
                        _destinationPath,
                        FileMode.OpenOrCreate,
                        FileAccess.Write,
                        FileShare.Read,
                        _ioBufferSize,
                        FileOptions.Asynchronous Or FileOptions.SequentialScan)
                End If

                Return _destinationStream.SafeFileHandle
            End Function

            Public Sub InvalidateSource()
                If _sourceStream Is Nothing Then
                    Return
                End If

                Try
                    _sourceStream.Dispose()
                Catch
                Finally
                    _sourceStream = Nothing
                End Try
            End Sub

            Public Sub InvalidateDestination()
                If _destinationStream Is Nothing Then
                    Return
                End If

                Try
                    _destinationStream.Dispose()
                Catch
                Finally
                    _destinationStream = Nothing
                End Try
            End Sub

            Public Sub Dispose() Implements IDisposable.Dispose
                If _disposed Then
                    Return
                End If

                _disposed = True
                InvalidateSource()
                InvalidateDestination()
            End Sub

            Private Sub ThrowIfDisposed()
                If _disposed Then
                    Throw New ObjectDisposedException(NameOf(FileTransferSession))
                End If
            End Sub
        End Class

#If DEBUG Then
        Private NotInheritable Class DevFaultInjector
            Private Const RulesEnvVar As String = "XACTCOPY_DEV_FAULT_RULES"
            Private Const SeedEnvVar As String = "XACTCOPY_DEV_FAULT_SEED"

            Private ReadOnly _rules As List(Of DevFaultRule)
            Private ReadOnly _random As Random
            Private ReadOnly _syncRoot As New Object()

            Private Sub New(rules As IEnumerable(Of DevFaultRule), seed As Integer)
                _rules = rules.ToList()
                _random = New Random(seed)
            End Sub

            Public ReadOnly Property Description As String
                Get
                    Return $"{_rules.Count} rule(s) via {RulesEnvVar}"
                End Get
            End Property

            Public Shared Function TryCreateFromEnvironment() As DevFaultInjector
                Dim rawRules = Environment.GetEnvironmentVariable(RulesEnvVar)
                If String.IsNullOrWhiteSpace(rawRules) Then
                    Return Nothing
                End If

                Dim parsedRules As New List(Of DevFaultRule)()
                For Each token In rawRules.Split(";"c)
                    Dim candidate = token.Trim()
                    If candidate.Length = 0 Then
                        Continue For
                    End If

                    Dim parsedRule As DevFaultRule = Nothing
                    If DevFaultRule.TryParse(candidate, parsedRule) Then
                        parsedRules.Add(parsedRule)
                    End If
                Next

                If parsedRules.Count = 0 Then
                    Return Nothing
                End If

                Dim seed = 1337
                Dim rawSeed = Environment.GetEnvironmentVariable(SeedEnvVar)
                If Not String.IsNullOrWhiteSpace(rawSeed) Then
                    Dim parsedSeed As Integer
                    If Integer.TryParse(rawSeed, NumberStyles.Integer, CultureInfo.InvariantCulture, parsedSeed) Then
                        seed = parsedSeed
                    End If
                End If

                Return New DevFaultInjector(parsedRules, seed)
            End Function

            Public Function CreateReadFault(relativePath As String, offset As Long, length As Integer) As Exception
                Return CreateFault(DevFaultOperation.ReadOperation, relativePath, offset, length)
            End Function

            Public Function CreateWriteFault(relativePath As String, offset As Long, length As Integer) As Exception
                Return CreateFault(DevFaultOperation.WriteOperation, relativePath, offset, length)
            End Function

            Private Function CreateFault(
                operationType As DevFaultOperation,
                relativePath As String,
                offset As Long,
                length As Integer) As Exception

                SyncLock _syncRoot
                    For Each rule In _rules
                        Dim injected = rule.TryCreateFault(operationType, relativePath, offset, length, _random)
                        If injected IsNot Nothing Then
                            Return injected
                        End If
                    Next
                End SyncLock

                Return Nothing
            End Function
        End Class

        Private Enum DevFaultOperation
            ReadOperation = 0
            WriteOperation = 1
        End Enum

        Private Enum DevFaultTriggerMode
            Once = 0
            Always = 1
            Percent = 2
            Every = 3
        End Enum

        Private Enum DevFaultErrorKind
            Io = 0
            Timeout = 1
            Offline = 2
        End Enum

        Private NotInheritable Class DevFaultRule
            Private ReadOnly _operation As DevFaultOperation
            Private ReadOnly _offset As Long
            Private ReadOnly _length As Long
            Private ReadOnly _trigger As DevFaultTriggerMode
            Private ReadOnly _errorKind As DevFaultErrorKind
            Private ReadOnly _parameter As Integer
            Private _matchCount As Integer
            Private _fireCount As Integer

            Private Sub New(
                operationType As DevFaultOperation,
                offset As Long,
                length As Long,
                trigger As DevFaultTriggerMode,
                errorKind As DevFaultErrorKind,
                parameter As Integer)

                _operation = operationType
                _offset = Math.Max(0L, offset)
                _length = Math.Max(0L, length)
                _trigger = trigger
                _errorKind = errorKind
                _parameter = Math.Max(0, parameter)
            End Sub

            Public Shared Function TryParse(text As String, ByRef rule As DevFaultRule) As Boolean
                rule = Nothing
                If String.IsNullOrWhiteSpace(text) Then
                    Return False
                End If

                ' Rule format:
                ' op,offset,length,mode,error[,param]
                ' Example:
                ' read,4096,4096,always,io
                ' write,0,0,percent,io,30
                Dim parts = text.Split(","c)
                If parts.Length < 5 Then
                    Return False
                End If

                Dim operationType As DevFaultOperation
                Select Case parts(0).Trim().ToLowerInvariant()
                    Case "read"
                        operationType = DevFaultOperation.ReadOperation
                    Case "write"
                        operationType = DevFaultOperation.WriteOperation
                    Case Else
                        Return False
                End Select

                Dim offset As Long
                If Not Long.TryParse(parts(1).Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, offset) Then
                    Return False
                End If

                Dim length As Long
                If Not Long.TryParse(parts(2).Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, length) Then
                    Return False
                End If

                Dim trigger As DevFaultTriggerMode
                Select Case parts(3).Trim().ToLowerInvariant()
                    Case "once"
                        trigger = DevFaultTriggerMode.Once
                    Case "always"
                        trigger = DevFaultTriggerMode.Always
                    Case "percent"
                        trigger = DevFaultTriggerMode.Percent
                    Case "every"
                        trigger = DevFaultTriggerMode.Every
                    Case Else
                        Return False
                End Select

                Dim errorKind As DevFaultErrorKind
                Select Case parts(4).Trim().ToLowerInvariant()
                    Case "io"
                        errorKind = DevFaultErrorKind.Io
                    Case "timeout"
                        errorKind = DevFaultErrorKind.Timeout
                    Case "offline"
                        errorKind = DevFaultErrorKind.Offline
                    Case Else
                        Return False
                End Select

                Dim parameter = 0
                If parts.Length >= 6 Then
                    Integer.TryParse(parts(5).Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, parameter)
                End If

                If trigger = DevFaultTriggerMode.Percent Then
                    parameter = Math.Clamp(parameter, 1, 100)
                ElseIf trigger = DevFaultTriggerMode.Every Then
                    parameter = Math.Max(1, parameter)
                End If

                rule = New DevFaultRule(operationType, offset, length, trigger, errorKind, parameter)
                Return True
            End Function

            Public Function TryCreateFault(
                operationType As DevFaultOperation,
                relativePath As String,
                offset As Long,
                length As Integer,
                random As Random) As Exception

                If operationType <> _operation Then
                    Return Nothing
                End If

                If length <= 0 Then
                    Return Nothing
                End If

                If Not Overlaps(offset, length) Then
                    Return Nothing
                End If

                _matchCount += 1
                If Not ShouldTrigger(random) Then
                    Return Nothing
                End If

                _fireCount += 1
                Return BuildException(relativePath, offset, length)
            End Function

            Private Function Overlaps(offset As Long, length As Integer) As Boolean
                Dim requestStart = Math.Max(0L, offset)
                Dim requestEnd = requestStart + Math.Max(0L, CLng(length))
                If requestEnd <= requestStart Then
                    Return False
                End If

                Dim ruleStart = _offset
                Dim ruleEnd = If(_length <= 0, Long.MaxValue, ruleStart + _length)
                Return requestStart < ruleEnd AndAlso requestEnd > ruleStart
            End Function

            Private Function ShouldTrigger(random As Random) As Boolean
                Select Case _trigger
                    Case DevFaultTriggerMode.Once
                        Return _fireCount = 0
                    Case DevFaultTriggerMode.Always
                        Return True
                    Case DevFaultTriggerMode.Percent
                        Dim chance = Math.Clamp(_parameter, 1, 100)
                        Return random.Next(0, 100) < chance
                    Case DevFaultTriggerMode.Every
                        Dim interval = Math.Max(1, _parameter)
                        Return (_matchCount Mod interval) = 0
                    Case Else
                        Return False
                End Select
            End Function

            Private Function BuildException(relativePath As String, offset As Long, length As Integer) As Exception
                Dim operationLabel = If(_operation = DevFaultOperation.ReadOperation, "read", "write")
                Dim pathLabel = If(String.IsNullOrWhiteSpace(relativePath), "?", relativePath)
                Dim message = $"[DevFault] Injected {operationLabel} fault on {pathLabel} at {FormatBytes(offset)} ({length} B)."

                Select Case _errorKind
                    Case DevFaultErrorKind.Timeout
                        Return New TimeoutException(message)
                    Case DevFaultErrorKind.Offline
                        Return New DriveNotFoundException(message)
                    Case Else
                        Return New IOException(message)
                End Select
            End Function
        End Class
#End If

        Private NotInheritable Class ProgressAccumulator
            Public Property TotalFiles As Integer
            Public Property TotalBytes As Long
            Public Property TotalBytesCopied As Long
            Public Property CompletedFiles As Integer
            Public Property FailedFiles As Integer
            Public Property RecoveredFiles As Integer
            Public Property SkippedFiles As Integer
        End Class

        Private NotInheritable Class SourceMutationSkippedException
            Inherits IOException

            Public Sub New(message As String, innerException As Exception)
                MyBase.New(message, innerException)
            End Sub
        End Class
    End Class
End Namespace
