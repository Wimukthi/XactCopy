Imports System.Collections.Generic
Imports System.Buffers
Imports System.IO
Imports System.Linq
Imports System.Globalization
Imports System.Runtime.InteropServices
Imports System.Security.Principal
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
        Private Const BadRangeMapFlushIntervalSeconds As Integer = 2
        Private Const MediaIdentityProbeIntervalMilliseconds As Integer = 1000
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
        Private Const ErrorNotSupported As Integer = 50
        Private Const ErrorBadNetPath As Integer = 53
        Private Const ErrorNetNameDeleted As Integer = 64
        Private Const ErrorBadNetName As Integer = 67
        Private Const ErrorInvalidParameter As Integer = 87
        Private Const ErrorDiskFull As Integer = 112
        Private Const ErrorDeviceNotConnected As Integer = 1167
        Private Const NativeCopyFastPathMaxBytes As Long = 64L * 1024L * 1024L

        Private ReadOnly _options As CopyJobOptions
        Private ReadOnly _executionControl As CopyExecutionControl
        Private ReadOnly _journalStore As JobJournalStore
        Private ReadOnly _badRangeMapStore As BadRangeMapStore
#If DEBUG Then
        Private ReadOnly _devFaultInjector As DevFaultInjector
#End If
        Private _lastJournalFlushUtc As DateTimeOffset = DateTimeOffset.MinValue
        Private _lastBadRangeMapFlushUtc As DateTimeOffset = DateTimeOffset.MinValue
        Private _throttleWindowStartUtc As DateTimeOffset = DateTimeOffset.MinValue
        Private _throttleWindowBytes As Long
        Private _expectedSourceIdentity As String = String.Empty
        Private _expectedDestinationIdentity As String = String.Empty
        Private _lastSourceIdentityMismatchLogUtc As DateTimeOffset = DateTimeOffset.MinValue
        Private _lastDestinationIdentityMismatchLogUtc As DateTimeOffset = DateTimeOffset.MinValue
        Private _lastSourceIdentityProbeUtc As DateTimeOffset = DateTimeOffset.MinValue
        Private _lastDestinationIdentityProbeUtc As DateTimeOffset = DateTimeOffset.MinValue
        Private _badRangeMap As BadRangeMap
        Private _badRangeMapPath As String = String.Empty
        Private _badRangeMapLoaded As Boolean
        Private _badRangeMapReadHintsEnabled As Boolean
        Private _rawDiskScanContext As RawDiskScanContext
        Private ReadOnly _rawScanFallbackLoggedPaths As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
        Private ReadOnly _fragileFailureTimestamps As New Queue(Of DateTimeOffset)()

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

        <DllImport("kernel32.dll", SetLastError:=True)>
        Private Shared Function CancelIoEx(
            hFile As SafeFileHandle,
            lpOverlapped As IntPtr) As Boolean
        End Function

        <Flags>
        Private Enum CopyFileFlags As UInteger
            None = 0UI
            Restartable = &H2UI
            OpenSourceForWrite = &H4UI
            AllowDecryptedDestination = &H8UI
            CopySymlink = &H800UI
            NoBuffering = &H1000UI
            RequestSecurityPrivileges = &H2000UI
            ResumeFromPause = &H4000UI
            NoOffload = &H40000000UI
        End Enum

        Private Enum CopyProgressCallbackReason As UInteger
            ChunkFinished = 0UI
            StreamSwitch = 1UI
        End Enum

        Private Enum CopyProgressResult As UInteger
            ContinueCopy = 0UI
            CancelCopy = 1UI
            StopCopy = 2UI
            Quiet = 3UI
        End Enum

        <UnmanagedFunctionPointer(CallingConvention.Winapi)>
        Private Delegate Function CopyProgressRoutine(
            totalFileSize As Long,
            totalBytesTransferred As Long,
            streamSize As Long,
            streamBytesTransferred As Long,
            dwStreamNumber As UInteger,
            dwCallbackReason As CopyProgressCallbackReason,
            hSourceFile As IntPtr,
            hDestinationFile As IntPtr,
            lpData As IntPtr) As CopyProgressResult

        <DllImport("kernel32.dll", CharSet:=CharSet.Unicode, SetLastError:=True)>
        Private Shared Function CopyFileEx(
            lpExistingFileName As String,
            lpNewFileName As String,
            lpProgressRoutine As CopyProgressRoutine,
            lpData As IntPtr,
            ByRef pbCancel As Integer,
            dwCopyFlags As CopyFileFlags) As Boolean
        End Function

        Private Enum GetFileExInfoLevels As Integer
            GetFileExInfoStandard = 0
        End Enum

        <StructLayout(LayoutKind.Sequential)>
        Private Structure Win32FileAttributeData
            Public FileAttributes As UInteger
            Public CreationTimeLow As UInteger
            Public CreationTimeHigh As UInteger
            Public LastAccessTimeLow As UInteger
            Public LastAccessTimeHigh As UInteger
            Public LastWriteTimeLow As UInteger
            Public LastWriteTimeHigh As UInteger
            Public FileSizeHigh As UInteger
            Public FileSizeLow As UInteger
        End Structure

        <DllImport("kernel32.dll", CharSet:=CharSet.Unicode, SetLastError:=True, EntryPoint:="GetFileAttributesExW")>
        Private Shared Function GetFileAttributesEx(
            lpFileName As String,
            fInfoLevelId As GetFileExInfoLevels,
            ByRef lpFileInformation As Win32FileAttributeData) As Boolean
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
            Try
                _rawScanFallbackLoggedPaths.Clear()
                _fragileFailureTimestamps.Clear()

                Dim scanOnly = (_options.OperationMode = JobOperationMode.ScanOnly)
                Dim sourceRoot = Path.GetFullPath(_options.SourceRoot)
                Dim destinationRoot = Path.GetFullPath(_options.DestinationRoot)
                InitializeMediaIdentityExpectations(sourceRoot, destinationRoot)
                Await WaitForMediaAvailabilityAsync(sourceRoot, destinationRoot, cancellationToken).ConfigureAwait(False)
                EnsureMediaIdentityIntegrity(sourceRoot, destinationRoot, includeDestination:=Not scanOnly, force:=True)
                InitializeRawDiskScanContext(scanOnly, sourceRoot)
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
                    EnsureMediaIdentityIntegrity(sourceRoot, destinationRoot, includeDestination:=Not scanOnly, force:=False)

                    Dim entry = journal.Files(descriptor.RelativePath)

                ' Overwrite/existing-destination semantics are copy-only. Scan mode must always
                ' attempt source reads to discover bad ranges.
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
                        If ShouldSkipFailedEntryForFragileResume(entry) Then
                            progress.SkippedFiles += 1
                            progress.TotalBytesCopied += Math.Max(0L, descriptor.Length - Math.Max(0L, entry.BytesCopied))
                            EmitLog($"Skipped scan: {descriptor.RelativePath} (persisted fragile skip).")
                            EmitProgress(
                                progress,
                                descriptor.RelativePath,
                                descriptor.Length,
                                descriptor.Length,
                                lastChunkBytesTransferred:=0,
                                bufferSizeBytes:=ResolveBufferSizeForFile(descriptor.Length))
                            Continue For
                        End If

                        If entry.State = FileCopyState.Failed Then
                            entry.State = FileCopyState.Pending
                            entry.LastError = String.Empty
                            entry.DoNotRetry = False
                        End If

                        Dim scanException As Exception = Nothing
                        Dim fragileScanSkip As FragileReadSkipException = Nothing
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
                            entry.DoNotRetry = False
                            Await FlushJournalAsync(journalPath, journal, cancellationToken, force:=False).ConfigureAwait(False)

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
                            entry.DoNotRetry = False
                            EmitLog($"Skipped scan: {descriptor.RelativePath} ({ex.Message})")
                            FlushJournalAsync(journalPath, journal, cancellationToken, force:=True).GetAwaiter().GetResult()
                            Continue For
                        Catch ex As FragileReadSkipException
                            fragileScanSkip = ex
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

                        If fragileScanSkip IsNot Nothing Then
                            Await HandleFragileReadSkipAsync(
                                "scan",
                                descriptor,
                                entry,
                                progress,
                                journal,
                                journalPath,
                                sourceRoot,
                                fragileScanSkip,
                                cancellationToken).ConfigureAwait(False)
                            Continue For
                        End If

                        If scanException Is Nothing Then
                            progress.CompletedFiles += 1
                            If scanDetectedBadRanges Then
                                progress.RecoveredFiles += 1
                            End If

                            entry.State = If(scanDetectedBadRanges, FileCopyState.CompletedWithRecovery, FileCopyState.Completed)
                            entry.BytesCopied = descriptor.Length
                            entry.LastError = String.Empty
                            entry.DoNotRetry = False
                            Await TryPersistBadRangeMapEntryAsync(sourceRoot, descriptor, entry, cancellationToken, forceSave:=False).ConfigureAwait(False)
                            Await FlushJournalAsync(journalPath, journal, cancellationToken, force:=False).ConfigureAwait(False)
                            Continue For
                        End If

                        progress.FailedFiles += 1
                        entry.State = FileCopyState.Failed
                        entry.LastError = scanException.Message
                        entry.DoNotRetry = False
                        EmitLog($"Scan failed: {descriptor.RelativePath} ({scanException.Message})")
                        Await RegisterFragileFailureAndMaybeCooldownAsync(descriptor.RelativePath, scanException.Message, cancellationToken).ConfigureAwait(False)
                        Await TryPersistBadRangeMapEntryAsync(sourceRoot, descriptor, entry, cancellationToken, forceSave:=False).ConfigureAwait(False)
                        Await FlushJournalAsync(journalPath, journal, cancellationToken, force:=False).ConfigureAwait(False)

                        If Not _options.ContinueOnFileError Then
                            Await FlushJournalAsync(journalPath, journal, cancellationToken, force:=True).ConfigureAwait(False)
                            Return CreateResult(progress, journalPath, succeeded:=False, cancelled:=False, errorMessage:=scanException.Message)
                        End If

                        Continue For
                    End If

                    If ShouldSkipFailedEntryForFragileResume(entry) Then
                        progress.SkippedFiles += 1
                        progress.TotalBytesCopied += Math.Max(0L, descriptor.Length - Math.Max(0L, entry.BytesCopied))
                        EmitLog($"Skipped: {descriptor.RelativePath} (persisted fragile skip).")
                        EmitProgress(
                            progress,
                            descriptor.RelativePath,
                            descriptor.Length,
                            descriptor.Length,
                            lastChunkBytesTransferred:=0,
                            bufferSizeBytes:=ResolveBufferSizeForFile(descriptor.Length))
                        Continue For
                    End If

                    If entry.State = FileCopyState.Failed Then
                        entry.State = FileCopyState.Pending
                        entry.LastError = String.Empty
                        entry.DoNotRetry = False
                    End If

                    Dim copyException As Exception = Nothing
                    Dim fragileCopySkip As FragileReadSkipException = Nothing
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
                        entry.DoNotRetry = False
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
                        entry.DoNotRetry = False
                        EmitLog($"Skipped: {descriptor.RelativePath} ({ex.Message})")
                        FlushJournalAsync(journalPath, journal, cancellationToken, force:=True).GetAwaiter().GetResult()
                        Continue For
                    Catch ex As FragileReadSkipException
                        fragileCopySkip = ex
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

                    If fragileCopySkip IsNot Nothing Then
                        Await HandleFragileReadSkipAsync(
                            "copy",
                            descriptor,
                            entry,
                            progress,
                            journal,
                            journalPath,
                            sourceRoot,
                            fragileCopySkip,
                            cancellationToken).ConfigureAwait(False)
                        Continue For
                    End If

                    If copyException Is Nothing Then
                        progress.CompletedFiles += 1
                        If recovered Then
                            progress.RecoveredFiles += 1
                        End If

                        entry.State = If(recovered, FileCopyState.CompletedWithRecovery, FileCopyState.Completed)
                        entry.BytesCopied = descriptor.Length
                        entry.LastError = String.Empty
                        entry.DoNotRetry = False
                        Await TryPersistBadRangeMapEntryAsync(sourceRoot, descriptor, entry, cancellationToken).ConfigureAwait(False)
                        Await FlushJournalAsync(journalPath, journal, cancellationToken, force:=True).ConfigureAwait(False)
                        Continue For
                    End If

                    progress.FailedFiles += 1
                    entry.State = FileCopyState.Failed
                    entry.LastError = copyException.Message
                    entry.DoNotRetry = False
                    EmitLog($"Failed: {descriptor.RelativePath} ({copyException.Message})")
                    Await RegisterFragileFailureAndMaybeCooldownAsync(descriptor.RelativePath, copyException.Message, cancellationToken).ConfigureAwait(False)
                    Await TryPersistBadRangeMapEntryAsync(sourceRoot, descriptor, entry, cancellationToken).ConfigureAwait(False)
                    Await FlushJournalAsync(journalPath, journal, cancellationToken, force:=True).ConfigureAwait(False)

                    If Not _options.ContinueOnFileError Then
                        Return CreateResult(progress, journalPath, succeeded:=False, cancelled:=False, errorMessage:=copyException.Message)
                    End If
                Next

                Await FlushBadRangeMapAsync(cancellationToken).ConfigureAwait(False)
                Await FlushJournalAsync(journalPath, journal, cancellationToken, force:=True).ConfigureAwait(False)
                Return CreateResult(progress, journalPath, succeeded:=(progress.FailedFiles = 0), cancelled:=False, errorMessage:=String.Empty)
            Finally
                DisposeRawDiskScanContext()
            End Try
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
                ' Scan mode is read-only; worker keeps destination aligned to source for journal IDs.
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

            If _options.FragileFailureWindowSeconds <= 0 Then
                Throw New InvalidOperationException("FragileFailureWindowSeconds must be greater than zero.")
            End If

            If _options.FragileFailureThreshold <= 0 Then
                Throw New InvalidOperationException("FragileFailureThreshold must be greater than zero.")
            End If

            If _options.FragileCooldownSeconds < 0 Then
                Throw New InvalidOperationException("FragileCooldownSeconds cannot be negative.")
            End If
        End Sub

        Private Sub InitializeRawDiskScanContext(scanOnly As Boolean, sourceRoot As String)
            DisposeRawDiskScanContext()
            _rawScanFallbackLoggedPaths.Clear()

            If Not scanOnly OrElse Not _options.UseExperimentalRawDiskScan Then
                Return
            End If

            If Not IsProcessElevated() Then
                EmitLog("Scan backend: Raw disk requested but worker is not elevated. Using standard file reads. Run XactCopy as Administrator to enable raw scan.")
                Return
            End If

            Dim context As RawDiskScanContext = Nothing
            Dim reason As String = String.Empty
            If RawDiskScanContext.TryCreate(sourceRoot, context, reason) Then
                _rawDiskScanContext = context
                EmitLog($"Scan backend: Raw disk (experimental) enabled. Cluster size {FormatBytes(context.ClusterSizeBytes)}.")
                Return
            End If

            Dim normalizedReason = If(String.IsNullOrWhiteSpace(reason), "unsupported source/media", reason)
            EmitLog($"Scan backend: Raw disk unavailable ({normalizedReason}); using standard file reads.")
        End Sub

        Private Shared Function IsProcessElevated() As Boolean
            If Not OperatingSystem.IsWindows() Then
                Return False
            End If

            Try
                Using identity = WindowsIdentity.GetCurrent()
                    If identity Is Nothing Then
                        Return False
                    End If

                    Dim principal As New WindowsPrincipal(identity)
                    Return principal.IsInRole(WindowsBuiltInRole.Administrator)
                End Using
            Catch
                Return False
            End Try
        End Function

        Private Sub DisposeRawDiskScanContext()
            If _rawDiskScanContext Is Nothing Then
                Return
            End If

            Try
                _rawDiskScanContext.Dispose()
            Catch
            Finally
                _rawDiskScanContext = Nothing
            End Try
        End Sub

        Private Function ShouldUseRawDiskReadBackend() As Boolean
            Return _options.OperationMode = JobOperationMode.ScanOnly AndAlso
                _options.UseExperimentalRawDiskScan AndAlso
                _rawDiskScanContext IsNot Nothing
        End Function

        Private Sub EmitRawDiskFallbackOnce(sourcePath As String, reason As String)
            Dim key = NormalizeRootPath(sourcePath)
            If key.Length = 0 Then
                key = If(sourcePath, String.Empty)
            End If

            If key.Length = 0 Then
                key = "(unknown)"
            End If

            If Not _rawScanFallbackLoggedPaths.Add(key) Then
                Return
            End If

            Dim normalizedReason = If(String.IsNullOrWhiteSpace(reason), "unsupported source layout", reason)
            EmitLog($"Raw disk scan fallback for '{key}': {normalizedReason}.")
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
                        .DoNotRetry = False,
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
                    entry.DoNotRetry = False
                    entry.RecoveredRanges.Clear()
                    entry.RescueRanges.Clear()
                    entry.LastRescuePass = String.Empty
                    Continue For
                End If

                If entry.State = FileCopyState.InProgress Then
                    If _options.FragileMediaMode AndAlso _options.SkipFileOnFirstReadError Then
                        entry.State = FileCopyState.Failed
                        entry.LastError = "Fragile media guard: previous attempt stalled while this file was in progress."
                        entry.DoNotRetry = True
                        EmitLog($"[Fragile] Promoted stale in-progress entry to non-retry: {descriptor.RelativePath}")
                    Else
                        entry.State = FileCopyState.Pending
                        If String.IsNullOrWhiteSpace(entry.LastError) Then
                            entry.LastError = "Previous attempt ended while file was in progress."
                        End If
                        entry.DoNotRetry = False
                    End If
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
            Dim destinationMetadata As ExistingFileMetadata = Nothing
            If Not TryGetExistingFileMetadata(destinationPath, destinationMetadata) Then
                Return False
            End If

            Return destinationMetadata.Length = descriptor.Length
        End Function

        Private Function ShouldSkipFailedEntryForFragileResume(entry As JournalFileEntry) As Boolean
            If entry Is Nothing Then
                Return False
            End If

            Return entry.State = FileCopyState.Failed AndAlso
                entry.DoNotRetry AndAlso
                (_options.PersistFragileSkipAcrossResume OrElse _options.FragileMediaMode)
        End Function

        Private Async Function HandleFragileReadSkipAsync(
            operationLabel As String,
            descriptor As SourceFileDescriptor,
            entry As JournalFileEntry,
            progress As ProgressAccumulator,
            journal As JobJournal,
            journalPath As String,
            sourceRoot As String,
            ex As FragileReadSkipException,
            cancellationToken As CancellationToken) As Task

            If descriptor Is Nothing OrElse entry Is Nothing OrElse progress Is Nothing Then
                Return
            End If

            Dim remainingBytes = Math.Max(0L, descriptor.Length - Math.Max(0L, entry.BytesCopied))
            progress.SkippedFiles += 1
            progress.TotalBytesCopied += remainingBytes

            entry.State = FileCopyState.Failed
            entry.LastError = ex.Message
            entry.DoNotRetry = _options.PersistFragileSkipAcrossResume

            Dim modeText = If(String.IsNullOrWhiteSpace(operationLabel), "copy", operationLabel.Trim().ToLowerInvariant())
            EmitLog($"Skipped {modeText}: {descriptor.RelativePath} ({ex.Message})")
            EmitProgress(
                progress,
                descriptor.RelativePath,
                descriptor.Length,
                descriptor.Length,
                lastChunkBytesTransferred:=0,
                bufferSizeBytes:=ResolveBufferSizeForFile(descriptor.Length))

            Await RegisterFragileFailureAndMaybeCooldownAsync(descriptor.RelativePath, ex.Message, cancellationToken).ConfigureAwait(False)
            Await TryPersistBadRangeMapEntryAsync(sourceRoot, descriptor, entry, cancellationToken, forceSave:=False).ConfigureAwait(False)
            Await FlushJournalAsync(journalPath, journal, cancellationToken, force:=True).ConfigureAwait(False)
        End Function

        Private Async Function RegisterFragileFailureAndMaybeCooldownAsync(
            relativePath As String,
            reason As String,
            cancellationToken As CancellationToken) As Task

            If Not _options.FragileMediaMode Then
                Return
            End If

            Dim nowUtc = DateTimeOffset.UtcNow
            Dim windowSeconds = Math.Max(1, _options.FragileFailureWindowSeconds)
            Dim threshold = Math.Max(1, _options.FragileFailureThreshold)
            Dim cooldownSeconds = Math.Max(0, _options.FragileCooldownSeconds)

            _fragileFailureTimestamps.Enqueue(nowUtc)
            Dim cutoffUtc = nowUtc - TimeSpan.FromSeconds(windowSeconds)
            While _fragileFailureTimestamps.Count > 0 AndAlso _fragileFailureTimestamps.Peek() < cutoffUtc
                _fragileFailureTimestamps.Dequeue()
            End While

            If _fragileFailureTimestamps.Count < threshold Then
                Return
            End If

            Dim pathLabel = If(String.IsNullOrWhiteSpace(relativePath), "?", relativePath)
            Dim reasonSuffix = If(String.IsNullOrWhiteSpace(reason), String.Empty, $" ({reason})")
            EmitLog($"[Fragile] Failure threshold reached after {pathLabel}{reasonSuffix}.")
            _fragileFailureTimestamps.Clear()

            If cooldownSeconds <= 0 Then
                Return
            End If

            EmitLog($"[Fragile] Cooling down for {cooldownSeconds} second(s) to reduce stress on unstable media.")
            Await Task.Delay(TimeSpan.FromSeconds(cooldownSeconds), cancellationToken).ConfigureAwait(False)
        End Function

        Private Function ShouldSkipByOverwritePolicy(
            descriptor As SourceFileDescriptor,
            destinationRoot As String,
            ByRef reason As String) As Boolean

            reason = String.Empty
            Dim destinationPath = Path.Combine(destinationRoot, descriptor.RelativePath)
            Dim destinationMetadata As ExistingFileMetadata = Nothing
            If Not TryGetExistingFileMetadata(destinationPath, destinationMetadata) Then
                Return False
            End If

            Select Case _options.OverwritePolicy
                Case OverwritePolicy.SkipExisting
                    reason = "destination already exists"
                    Return True

                Case OverwritePolicy.OverwriteIfSourceNewer
                    If descriptor.LastWriteTimeUtc <= destinationMetadata.LastWriteTimeUtc Then
                        reason = "destination is newer or same age"
                        Return True
                    End If

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

            Dim nativeFastPathBaselineBytes = progress.TotalBytesCopied
            Dim nativeFastPathAttempt = Await TryCopyWithNativeFastPathAsync(
                descriptor,
                entry,
                destinationPath,
                destinationLength,
                hasPersistedCoverage,
                progress,
                journal,
                journalPath,
                cancellationToken).ConfigureAwait(False)

            If nativeFastPathAttempt.Attempted Then
                If nativeFastPathAttempt.Succeeded Then
                    If _options.PreserveTimestamps Then
                        File.SetLastWriteTimeUtc(destinationPath, descriptor.LastWriteTimeUtc)
                    End If

                    Dim nativeVerificationMode = ResolveVerificationMode()
                    If nativeVerificationMode <> VerificationMode.None Then
                        Await VerifyFileWithRetriesAsync(
                            descriptor.FullPath,
                            destinationPath,
                            descriptor.RelativePath,
                            nativeVerificationMode,
                            cancellationToken).ConfigureAwait(False)
                    End If

                    Return False
                End If

                progress.TotalBytesCopied = nativeFastPathBaselineBytes
                entry.BytesCopied = 0
                entry.LastRescuePass = "Init"

                If nativeFastPathAttempt.FallbackReason.Length > 0 Then
                    EmitLog($"Native fast-path fallback on {descriptor.RelativePath}: {nativeFastPathAttempt.FallbackReason}")
                End If

                PrepareDestinationFile(
                    destinationPath,
                    descriptor.Length,
                    GetExistingFileLength(destinationPath),
                    preserveExisting:=False)
                destinationLength = GetExistingFileLength(destinationPath)
            End If

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

                    Dim useSmallFileFastPath = ShouldUseSmallFileFastPath(descriptor.Length, hasPersistedCoverage:=False)
                    If useSmallFileFastPath Then
                        entry.LastRescuePass = "ScanSmallFast"
                        Await ScanSmallFileFastAsync(
                            descriptor,
                            entry,
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
                    End If

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

        Private Async Function ScanSmallFileFastAsync(
            descriptor As SourceFileDescriptor,
            entry As JournalFileEntry,
            transferSession As FileTransferSession,
            ioBuffer As Byte(),
            ioBufferSize As Integer,
            progress As ProgressAccumulator,
            journal As JobJournal,
            journalPath As String,
            cancellationToken As CancellationToken) As Task

            If descriptor Is Nothing OrElse entry Is Nothing OrElse ioBuffer Is Nothing Then
                Return
            End If

            Dim pendingSnapshots = SnapshotRangesByState(entry.RescueRanges, RescueRangeState.Pending)
            For Each pendingRange In pendingSnapshots
                Dim offset = pendingRange.Offset
                Dim remaining = CLng(pendingRange.Length)
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

                    If bytesRead > 0 Then
                        SetRescueRangeState(entry.RescueRanges, offset, bytesRead, RescueRangeState.Good)
                        progress.TotalBytesCopied += bytesRead
                    Else
                        SetRescueRangeState(entry.RescueRanges, offset, chunkLength, RescueRangeState.Bad)
                        progress.TotalBytesCopied += chunkLength
                    End If

                    SyncEntryBytesCopied(entry)
                    EmitProgress(
                        progress,
                        descriptor.RelativePath,
                        entry.BytesCopied,
                        descriptor.Length,
                        lastChunkBytesTransferred:=Math.Max(0, bytesRead),
                        bufferSizeBytes:=ioBufferSize,
                        rescuePass:="ScanSmallFast",
                        rescueBadRegionCount:=GetUnreadableRegionCount(entry.RescueRanges),
                        rescueRemainingBytes:=GetRescueRangeBytes(entry.RescueRanges, RescueRangeState.Pending, RescueRangeState.Bad, RescueRangeState.KnownBad))
                    Await FlushJournalAsync(journalPath, journal, cancellationToken, force:=False).ConfigureAwait(False)

                    offset += chunkLength
                    remaining -= chunkLength
                End While
            Next
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
            cancellationToken As CancellationToken,
            Optional forceSave As Boolean = True) As Task

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

            Dim shouldSave = forceSave
            If Not shouldSave Then
                If _lastBadRangeMapFlushUtc = DateTimeOffset.MinValue Then
                    shouldSave = True
                Else
                    shouldSave = (DateTimeOffset.UtcNow - _lastBadRangeMapFlushUtc) >= TimeSpan.FromSeconds(BadRangeMapFlushIntervalSeconds)
                End If
            End If

            If Not shouldSave Then
                Return
            End If

            Try
                Await _badRangeMapStore.SaveAsync(_badRangeMapPath, _badRangeMap, cancellationToken).ConfigureAwait(False)
                _lastBadRangeMapFlushUtc = DateTimeOffset.UtcNow
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
                _lastBadRangeMapFlushUtc = DateTimeOffset.UtcNow
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
            ' Preserve roots such as "D:\" to avoid accidental drive-relative normalization ("D:").
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

        Private Function ShouldUseNativeCopyFastPath(
            fileLength As Long,
            destinationLength As Long,
            hasPersistedCoverage As Boolean) As Boolean

            If fileLength <= 0 Then
                Return False
            End If

            If fileLength > NativeCopyFastPathMaxBytes Then
                Return False
            End If

            If destinationLength > 0 OrElse hasPersistedCoverage Then
                Return False
            End If

            If _options.MaxThroughputBytesPerSecond > 0 Then
                Return False
            End If

            If _options.UseBadRangeMap AndAlso _options.SkipKnownBadRanges Then
                Return False
            End If

            Return _options.OperationMode = JobOperationMode.Copy
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

        Private Async Function TryCopyWithNativeFastPathAsync(
            descriptor As SourceFileDescriptor,
            entry As JournalFileEntry,
            destinationPath As String,
            destinationLength As Long,
            hasPersistedCoverage As Boolean,
            progress As ProgressAccumulator,
            journal As JobJournal,
            journalPath As String,
            cancellationToken As CancellationToken) As Task(Of NativeFastPathAttemptResult)

            Dim result As New NativeFastPathAttemptResult()
            If descriptor Is Nothing OrElse entry Is Nothing Then
                Return result
            End If

            If Not ShouldUseNativeCopyFastPath(descriptor.Length, destinationLength, hasPersistedCoverage) Then
                Return result
            End If

            result.Attempted = True
            Dim baselineBytes = progress.TotalBytesCopied
            Dim lastTransferred = 0L
            Dim cancelledForPause = False
            Dim progressCallback As CopyProgressRoutine =
                Function(
                    totalFileSize As Long,
                    totalBytesTransferred As Long,
                    streamSize As Long,
                    streamBytesTransferred As Long,
                    dwStreamNumber As UInteger,
                    dwCallbackReason As CopyProgressCallbackReason,
                    hSourceFile As IntPtr,
                    hDestinationFile As IntPtr,
                    lpData As IntPtr) As CopyProgressResult

                    If cancellationToken.IsCancellationRequested Then
                        Return CopyProgressResult.CancelCopy
                    End If

                    If _executionControl IsNot Nothing AndAlso _executionControl.IsPaused Then
                        cancelledForPause = True
                        Return CopyProgressResult.CancelCopy
                    End If

                    Dim boundedTransferred = Math.Max(0L, Math.Min(descriptor.Length, totalBytesTransferred))
                    Dim chunkBytes = boundedTransferred - lastTransferred
                    If chunkBytes < 0L Then
                        chunkBytes = 0L
                    End If

                    lastTransferred = boundedTransferred
                    result.BytesTransferred = boundedTransferred
                    entry.BytesCopied = boundedTransferred
                    progress.TotalBytesCopied = baselineBytes + boundedTransferred
                    EmitProgress(
                        progress,
                        descriptor.RelativePath,
                        boundedTransferred,
                        descriptor.Length,
                        lastChunkBytesTransferred:=CInt(Math.Min(CLng(Integer.MaxValue), chunkBytes)),
                        bufferSizeBytes:=ResolveBufferSizeForFile(descriptor.Length),
                        rescuePass:="NativeFastPath",
                        rescueBadRegionCount:=0,
                        rescueRemainingBytes:=0L)

                    Return CopyProgressResult.ContinueCopy
                End Function

            Dim cancelFlag = 0
            Dim copySucceeded = False
            Dim lastWin32Error = 0
            Try
                copySucceeded = CopyFileEx(
                    descriptor.FullPath,
                    destinationPath,
                    progressCallback,
                    IntPtr.Zero,
                    cancelFlag,
                    CopyFileFlags.AllowDecryptedDestination)
                If Not copySucceeded Then
                    lastWin32Error = Marshal.GetLastWin32Error()
                End If
            Catch ex As Exception
                result.FallbackReason = ex.Message
                progress.TotalBytesCopied = baselineBytes
                entry.BytesCopied = 0
                Return result
            End Try

            If copySucceeded Then
                result.Succeeded = True
                result.BytesTransferred = descriptor.Length
                progress.TotalBytesCopied = baselineBytes + descriptor.Length
                entry.BytesCopied = descriptor.Length
                entry.LastError = String.Empty
                entry.LastRescuePass = "NativeFastPath"
                If entry.RecoveredRanges Is Nothing Then
                    entry.RecoveredRanges = New List(Of ByteRange)()
                Else
                    entry.RecoveredRanges.Clear()
                End If

                entry.RescueRanges = New List(Of RescueRange)()
                AppendRescueRange(entry.RescueRanges, 0L, descriptor.Length, RescueRangeState.Good)
                SyncEntryBytesCopied(entry)

                EmitProgress(
                    progress,
                    descriptor.RelativePath,
                    entry.BytesCopied,
                    descriptor.Length,
                    lastChunkBytesTransferred:=CInt(Math.Min(CLng(Integer.MaxValue), Math.Max(0L, descriptor.Length - lastTransferred))),
                    bufferSizeBytes:=ResolveBufferSizeForFile(descriptor.Length),
                    rescuePass:="NativeFastPath",
                    rescueBadRegionCount:=0,
                    rescueRemainingBytes:=0L)

                Await FlushJournalAsync(journalPath, journal, cancellationToken, force:=False).ConfigureAwait(False)
                Return result
            End If

            progress.TotalBytesCopied = baselineBytes
            entry.BytesCopied = 0
            result.Succeeded = False

            If cancellationToken.IsCancellationRequested Then
                Throw New OperationCanceledException(cancellationToken)
            End If

            If cancelledForPause Then
                result.FallbackReason = "paused by user during native fast-path copy."
                Return result
            End If

            result.FallbackReason = $"CopyFileEx failed ({lastWin32Error})."
            Return result
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

            Dim normalizedState = NormalizeRescueRangeState(newState)

            Dim startIndex = 0
            While startIndex < ranges.Count
                Dim candidate = ranges(startIndex)
                If candidate Is Nothing OrElse candidate.Length <= 0 Then
                    ranges.RemoveAt(startIndex)
                    Continue While
                End If

                Dim candidateEnd = candidate.Offset + CLng(candidate.Length)
                If candidateEnd > targetStart Then
                    Exit While
                End If

                startIndex += 1
            End While

            If startIndex >= ranges.Count Then
                Return
            End If

            Dim replacement As New List(Of RescueRange)()
            Dim index = startIndex
            Dim affectedCount = 0

            While index < ranges.Count
                Dim item = ranges(index)
                If item Is Nothing OrElse item.Length <= 0 Then
                    index += 1
                    Continue While
                End If

                Dim itemStart = item.Offset
                If itemStart >= targetEnd Then
                    Exit While
                End If

                Dim itemEnd = itemStart + CLng(item.Length)
                If itemEnd <= targetStart Then
                    startIndex += 1
                    index += 1
                    Continue While
                End If

                affectedCount += 1

                If itemStart < targetStart Then
                    AppendRescueRange(replacement, itemStart, targetStart - itemStart, item.State)
                End If

                Dim overlapStart = Math.Max(itemStart, targetStart)
                Dim overlapEnd = Math.Min(itemEnd, targetEnd)
                If overlapEnd > overlapStart Then
                    AppendRescueRange(replacement, overlapStart, overlapEnd - overlapStart, normalizedState)
                End If

                If itemEnd > targetEnd Then
                    AppendRescueRange(replacement, targetEnd, itemEnd - targetEnd, item.State)
                End If

                index += 1
            End While

            If affectedCount <= 0 Then
                Return
            End If

            ranges.RemoveRange(startIndex, affectedCount)
            If replacement.Count > 0 Then
                ranges.InsertRange(startIndex, replacement)
            End If

            MergeAdjacentRescueRangesInPlace(ranges, Math.Max(0, startIndex - 1))
        End Sub

        Private Shared Sub MergeAdjacentRescueRangesInPlace(ranges As List(Of RescueRange), startIndex As Integer)
            If ranges Is Nothing OrElse ranges.Count <= 1 Then
                Return
            End If

            Dim index = Math.Max(0, startIndex)
            If index >= ranges.Count Then
                index = ranges.Count - 1
            End If

            If index > 0 Then
                index -= 1
            End If

            While index < ranges.Count - 1
                Dim left = ranges(index)
                If left Is Nothing OrElse left.Length <= 0 Then
                    ranges.RemoveAt(index)
                    If index > 0 Then
                        index -= 1
                    End If
                    Continue While
                End If

                Dim right = ranges(index + 1)
                If right Is Nothing OrElse right.Length <= 0 Then
                    ranges.RemoveAt(index + 1)
                    Continue While
                End If

                Dim leftState = NormalizeRescueRangeState(left.State)
                Dim rightState = NormalizeRescueRangeState(right.State)
                Dim leftEnd = left.Offset + CLng(left.Length)
                Dim rightEnd = right.Offset + CLng(right.Length)

                If leftEnd < right.Offset Then
                    left.State = leftState
                    right.State = rightState
                    ranges(index) = left
                    ranges(index + 1) = right
                    index += 1
                    Continue While
                End If

                If leftState = rightState Then
                    Dim mergedEnd = Math.Max(leftEnd, rightEnd)
                    left.Length = CInt(Math.Min(CLng(Integer.MaxValue), Math.Max(0L, mergedEnd - left.Offset)))
                    left.State = leftState
                    ranges(index) = left
                    ranges.RemoveAt(index + 1)
                    Continue While
                End If

                If rightEnd <= leftEnd Then
                    ranges.RemoveAt(index + 1)
                    Continue While
                End If

                right.Offset = leftEnd
                right.Length = CInt(Math.Min(CLng(Integer.MaxValue), rightEnd - leftEnd))
                right.State = rightState
                ranges(index + 1) = right
                index += 1
            End While
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

        Private NotInheritable Class ExistingFileMetadata
            Public Property Length As Long
            Public Property LastWriteTimeUtc As DateTime
        End Class

        Private Shared Function GetExistingFileLength(path As String) As Long
            Dim metadata As ExistingFileMetadata = Nothing
            If Not TryGetExistingFileMetadata(path, metadata) Then
                Return 0
            End If

            Return metadata.Length
        End Function

        Private Shared Function TryGetExistingFileMetadata(path As String, ByRef metadata As ExistingFileMetadata) As Boolean
            metadata = Nothing
            If String.IsNullOrWhiteSpace(path) Then
                Return False
            End If

            Dim attributeData As New Win32FileAttributeData()
            If GetFileAttributesEx(path, GetFileExInfoLevels.GetFileExInfoStandard, attributeData) Then
                Dim attributes = CType(attributeData.FileAttributes, FileAttributes)
                If (attributes And FileAttributes.Directory) = FileAttributes.Directory Then
                    Return False
                End If

                Dim rawLength = (CULng(attributeData.FileSizeHigh) << 32) Or CULng(attributeData.FileSizeLow)
                Dim boundedLength = If(rawLength > CULng(Long.MaxValue), Long.MaxValue, CLng(rawLength))
                Dim lastWriteTicks = (CULng(attributeData.LastWriteTimeHigh) << 32) Or CULng(attributeData.LastWriteTimeLow)
                metadata = New ExistingFileMetadata() With {
                    .Length = Math.Max(0L, boundedLength),
                    .LastWriteTimeUtc = FileTimeTicksToUtc(lastWriteTicks)
                }
                Return True
            End If

            Dim win32Error = Marshal.GetLastWin32Error()
            If win32Error = ErrorFileNotFound OrElse win32Error = ErrorPathNotFound Then
                Return False
            End If

            Try
                If Not File.Exists(path) Then
                    Return False
                End If

                Dim fileInfo As New FileInfo(path)
                metadata = New ExistingFileMetadata() With {
                    .Length = Math.Max(0L, fileInfo.Length),
                    .LastWriteTimeUtc = fileInfo.LastWriteTimeUtc
                }
                Return True
            Catch
                Return False
            End Try
        End Function

        Private Shared Function FileTimeTicksToUtc(fileTimeTicks As ULong) As DateTime
            If fileTimeTicks = 0UL Then
                Return DateTime.MinValue
            End If

            Try
                Return DateTime.FromFileTimeUtc(CLng(fileTimeTicks))
            Catch
                Return DateTime.MinValue
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
                ThrowIfMediaIdentityMismatchThrottled(sourcePath, _expectedSourceIdentity, isSource:=True)

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
                            relativePath,
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
                    CancelPendingIo(transferSession, isSource:=True)
                    lastError = New TimeoutException($"Read timeout at offset {offset}.", ex)
                Catch ex As Exception
                    lastError = ex
                End Try

                If _options.FragileMediaMode AndAlso _options.SkipFileOnFirstReadError Then
                    transferSession?.InvalidateSource()
                    Dim errorText = If(lastError Is Nothing, "unknown read error", lastError.Message)
                    Throw New FragileReadSkipException(
                        $"Fragile mode: first read failure at {FormatBytes(offset)} ({errorText}).",
                        lastError)
                End If

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
            relativePath As String,
            offset As Long,
            buffer As Byte(),
            count As Integer,
            ioBufferSize As Integer,
            transferSession As FileTransferSession,
            cancellationToken As CancellationToken) As Task(Of Integer)

            If ShouldUseRawDiskReadBackend() Then
                Dim rawOutcome = Await TryReadChunkViaRawDiskAsync(
                    sourcePath,
                    relativePath,
                    offset,
                    buffer,
                    count,
                    cancellationToken).ConfigureAwait(False)

                If rawOutcome.Handled Then
                    If rawOutcome.ReadError IsNot Nothing Then
                        Throw rawOutcome.ReadError
                    End If

                    Return rawOutcome.BytesRead
                End If

                If rawOutcome.FallbackReason.Length > 0 Then
                    EmitRawDiskFallbackOnce(sourcePath, rawOutcome.FallbackReason)
                End If
            End If

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

        Private Async Function TryReadChunkViaRawDiskAsync(
            sourcePath As String,
            relativePath As String,
            offset As Long,
            buffer As Byte(),
            count As Integer,
            cancellationToken As CancellationToken) As Task(Of RawReadAttemptOutcome)

            Dim outcome As New RawReadAttemptOutcome()

            If _rawDiskScanContext Is Nothing Then
                Return outcome
            End If

            Dim plan As IReadOnlyList(Of RawDiskReadSegment) = Nothing
            If Not _rawDiskScanContext.TryCreateReadPlan(sourcePath, offset, count, plan, outcome.FallbackReason) Then
                Return outcome
            End If

            outcome.Handled = True
            Try
                Dim destinationOffset = 0
                For Each segment In plan
                    cancellationToken.ThrowIfCancellationRequested()
                    If segment Is Nothing OrElse segment.Length <= 0 Then
                        Continue For
                    End If

                    If segment.IsSparse Then
                        Array.Clear(buffer, destinationOffset, segment.Length)
                        destinationOffset += segment.Length
                        Continue For
                    End If

                    Await ReadRawSegmentAlignedAsync(
                        _rawDiskScanContext.VolumeHandle,
                        segment.VolumeOffsetBytes,
                        buffer,
                        destinationOffset,
                        segment.Length,
                        _rawDiskScanContext.SectorSizeBytes,
                        cancellationToken).ConfigureAwait(False)
                    destinationOffset += segment.Length
                Next

                outcome.BytesRead = destinationOffset
                If outcome.BytesRead <> count Then
                    Throw New IOException(
                        $"Raw disk read length mismatch on {If(String.IsNullOrWhiteSpace(relativePath), sourcePath, relativePath)} at offset {offset}. Expected {count}, got {outcome.BytesRead}.")
                End If

                Return outcome
            Catch ex As Exception
                If IsRawReadFallbackException(ex) Then
                    outcome.Handled = False
                    outcome.ReadError = Nothing
                    outcome.FallbackReason = $"raw read unavailable ({ex.Message})"
                    Try
                        _rawDiskScanContext.MarkPathUnsupported(sourcePath, outcome.FallbackReason)
                    Catch
                    End Try
                    Return outcome
                End If

                outcome.ReadError = ex
                Return outcome
            End Try
        End Function

        Private Shared Async Function ReadRawSegmentAlignedAsync(
            handle As SafeFileHandle,
            volumeOffset As Long,
            destinationBuffer As Byte(),
            destinationOffset As Integer,
            segmentLength As Integer,
            sectorSizeBytes As Integer,
            cancellationToken As CancellationToken) As Task

            If handle Is Nothing OrElse handle.IsInvalid Then
                Throw New IOException("Raw volume handle is invalid.")
            End If

            If destinationBuffer Is Nothing Then
                Throw New ArgumentNullException(NameOf(destinationBuffer))
            End If

            If destinationOffset < 0 OrElse segmentLength < 0 OrElse destinationOffset + segmentLength > destinationBuffer.Length Then
                Throw New ArgumentOutOfRangeException(NameOf(destinationOffset))
            End If

            If segmentLength = 0 Then
                Return
            End If

            Dim sector = Math.Max(512, sectorSizeBytes)
            Dim alignedStart = (volumeOffset \ sector) * sector
            Dim endOffset = volumeOffset + CLng(segmentLength)
            Dim alignedEnd = ((endOffset + sector - 1) \ sector) * sector
            Dim alignedLengthLong = alignedEnd - alignedStart
            If alignedLengthLong <= 0 OrElse alignedLengthLong > Integer.MaxValue Then
                Throw New IOException("Raw read alignment produced an invalid length.")
            End If

            Dim alignedLength = CInt(alignedLengthLong)
            If alignedStart = volumeOffset AndAlso alignedLength = segmentLength Then
                Dim directRead = Await ReadExactlyAsync(
                    handle,
                    volumeOffset,
                    destinationBuffer,
                    destinationOffset,
                    segmentLength,
                    cancellationToken).ConfigureAwait(False)
                If directRead <> segmentLength Then
                    Throw New IOException($"Raw disk short read. Expected {segmentLength}, got {directRead}.")
                End If
                Return
            End If

            Dim alignedBuffer = ArrayPool(Of Byte).Shared.Rent(alignedLength)
            Try
                Dim alignedRead = Await ReadExactlyAsync(
                    handle,
                    alignedStart,
                    alignedBuffer,
                    bufferOffset:=0,
                    count:=alignedLength,
                    cancellationToken:=cancellationToken).ConfigureAwait(False)
                If alignedRead <> alignedLength Then
                    Throw New IOException($"Raw disk short aligned read. Expected {alignedLength}, got {alignedRead}.")
                End If

                Dim copyOffset = CInt(volumeOffset - alignedStart)
                Buffer.BlockCopy(alignedBuffer, copyOffset, destinationBuffer, destinationOffset, segmentLength)
            Finally
                ArrayPool(Of Byte).Shared.Return(alignedBuffer, clearArray:=False)
            End Try
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

            Return Await ReadExactlyAsync(
                handle,
                offset,
                buffer,
                bufferOffset:=0,
                count:=count,
                cancellationToken:=cancellationToken).ConfigureAwait(False)
        End Function

        Private Shared Async Function ReadExactlyAsync(
            handle As SafeFileHandle,
            offset As Long,
            buffer As Byte(),
            bufferOffset As Integer,
            count As Integer,
            cancellationToken As CancellationToken) As Task(Of Integer)

            If buffer Is Nothing Then
                Throw New ArgumentNullException(NameOf(buffer))
            End If

            If bufferOffset < 0 OrElse count < 0 OrElse bufferOffset + count > buffer.Length Then
                Throw New ArgumentOutOfRangeException(NameOf(bufferOffset))
            End If

            Dim totalRead = 0
            While totalRead < count
                Dim readNow = Await RandomAccess.ReadAsync(
                    handle,
                    buffer.AsMemory(bufferOffset + totalRead, count - totalRead),
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
                ThrowIfMediaIdentityMismatchThrottled(destinationPath, _expectedDestinationIdentity, isSource:=False)

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
                    CancelPendingIo(transferSession, isSource:=False)
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

        Private Sub CancelPendingIo(transferSession As FileTransferSession, isSource As Boolean)
            If transferSession Is Nothing Then
                Return
            End If

            Dim handle As SafeFileHandle = If(
                isSource,
                transferSession.TryGetOpenSourceHandle(),
                transferSession.TryGetOpenDestinationHandle())
            SafeCancelPendingIo(handle)
        End Sub

        Private Shared Sub SafeCancelPendingIo(handle As SafeFileHandle)
            If handle Is Nothing OrElse handle.IsInvalid OrElse handle.IsClosed Then
                Return
            End If

            Try
                Dim ignored = CancelIoEx(handle, IntPtr.Zero)
            Catch
                ' Best effort only.
            End Try
        End Sub

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
                                context,
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

        Private Sub EnsureMediaIdentityIntegrity(
            sourceRoot As String,
            destinationRoot As String,
            Optional includeDestination As Boolean = True,
            Optional force As Boolean = False)

            ThrowIfMediaIdentityMismatchThrottled(sourceRoot, _expectedSourceIdentity, isSource:=True, force:=force)
            If Not includeDestination Then
                Return
            End If

            ThrowIfMediaIdentityMismatchThrottled(destinationRoot, _expectedDestinationIdentity, isSource:=False, force:=force)
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

        Private Sub ThrowIfMediaIdentityMismatchThrottled(
            pathValue As String,
            expectedIdentity As String,
            isSource As Boolean,
            Optional force As Boolean = False)

            If Not force AndAlso Not ShouldProbeMediaIdentity(isSource) Then
                Return
            End If

            ThrowIfMediaIdentityMismatch(pathValue, expectedIdentity, isSource)
        End Sub

        Private Function ShouldProbeMediaIdentity(isSource As Boolean) As Boolean
            Dim nowUtc = DateTimeOffset.UtcNow
            Dim lastProbeUtc = If(isSource, _lastSourceIdentityProbeUtc, _lastDestinationIdentityProbeUtc)
            If lastProbeUtc = DateTimeOffset.MinValue OrElse
                (nowUtc - lastProbeUtc).TotalMilliseconds >= MediaIdentityProbeIntervalMilliseconds Then
                If isSource Then
                    _lastSourceIdentityProbeUtc = nowUtc
                Else
                    _lastDestinationIdentityProbeUtc = nowUtc
                End If
                Return True
            End If

            Return False
        End Function

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

        Private Shared Function IsRawReadFallbackException(ex As Exception) As Boolean
            If ex Is Nothing Then
                Return False
            End If

            For Each candidate In EnumerateExceptionChain(ex)
                If TypeOf candidate Is ArgumentException Then
                    Return True
                End If

                Dim win32Code As Integer
                If TryGetWin32ErrorCode(candidate, win32Code) Then
                    If win32Code = ErrorInvalidParameter OrElse win32Code = ErrorNotSupported Then
                        Return True
                    End If
                End If

                Dim ioEx = TryCast(candidate, IOException)
                If ioEx IsNot Nothing Then
                    Dim message = ioEx.Message.ToLowerInvariant()
                    If message.Contains("parameter is incorrect") OrElse
                        message.Contains("invalid parameter") OrElse
                        message.Contains("request is not supported") Then
                        Return True
                    End If
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

        Private NotInheritable Class NativeFastPathAttemptResult
            Public Property Attempted As Boolean
            Public Property Succeeded As Boolean
            Public Property BytesTransferred As Long
            Public Property FallbackReason As String = String.Empty
        End Class

        Private NotInheritable Class RawReadAttemptOutcome
            Public Property Handled As Boolean
            Public Property BytesRead As Integer
            Public Property ReadError As Exception
            Public Property FallbackReason As String = String.Empty
        End Class

        Private NotInheritable Class RawDiskScanContext
            Implements IDisposable

            Private Const FsctlGetRetrievalPointers As UInteger = &H90073UI
            Private Const ErrorMoreData As Integer = 234

            Private ReadOnly _sourceRoot As String
            Private ReadOnly _sourceVolumeRoot As String
            Private ReadOnly _volumeHandle As SafeFileHandle
            Private ReadOnly _sectorSizeBytes As Integer
            Private ReadOnly _clusterSizeBytes As Integer
            Private ReadOnly _layoutCache As New Dictionary(Of String, RawDiskFileLayout)(StringComparer.OrdinalIgnoreCase)
            Private _disposed As Boolean

            <DllImport("kernel32.dll", CharSet:=CharSet.Unicode, SetLastError:=True)>
            Private Shared Function GetDiskFreeSpace(
                lpRootPathName As String,
                ByRef lpSectorsPerCluster As UInteger,
                ByRef lpBytesPerSector As UInteger,
                ByRef lpNumberOfFreeClusters As UInteger,
                ByRef lpTotalNumberOfClusters As UInteger) As Boolean
            End Function

            <DllImport("kernel32.dll", SetLastError:=True)>
            Private Shared Function DeviceIoControl(
                hDevice As SafeFileHandle,
                dwIoControlCode As UInteger,
                lpInBuffer As IntPtr,
                nInBufferSize As Integer,
                lpOutBuffer As IntPtr,
                nOutBufferSize As Integer,
                ByRef lpBytesReturned As Integer,
                lpOverlapped As IntPtr) As Boolean
            End Function

            <StructLayout(LayoutKind.Sequential)>
            Private Structure StartingVcnInputBuffer
                Public StartingVcn As Long
            End Structure

            <StructLayout(LayoutKind.Sequential)>
            Private Structure RetrievalPointersBufferHeader
                Public ExtentCount As UInteger
                Public StartingVcn As Long
            End Structure

            <StructLayout(LayoutKind.Sequential)>
            Private Structure RetrievalPointersExtent
                Public NextVcn As Long
                Public Lcn As Long
            End Structure

            Private Sub New(
                sourceRoot As String,
                sourceVolumeRoot As String,
                volumeHandle As SafeFileHandle,
                sectorSizeBytes As Integer,
                clusterSizeBytes As Integer)

                _sourceRoot = NormalizePathForCache(sourceRoot)
                _sourceVolumeRoot = NormalizeVolumeRoot(sourceVolumeRoot)
                _volumeHandle = volumeHandle
                _sectorSizeBytes = Math.Max(512, sectorSizeBytes)
                _clusterSizeBytes = Math.Max(MinimumRescueBlockSize, clusterSizeBytes)
            End Sub

            Public ReadOnly Property VolumeHandle As SafeFileHandle
                Get
                    Return _volumeHandle
                End Get
            End Property

            Public ReadOnly Property ClusterSizeBytes As Integer
                Get
                    Return _clusterSizeBytes
                End Get
            End Property

            Public ReadOnly Property SectorSizeBytes As Integer
                Get
                    Return _sectorSizeBytes
                End Get
            End Property

            Public Shared Function TryCreate(sourceRoot As String, ByRef context As RawDiskScanContext, ByRef reason As String) As Boolean
                context = Nothing
                reason = String.Empty

                Dim normalizedSourceRoot = NormalizePathForCache(sourceRoot)
                If String.IsNullOrWhiteSpace(normalizedSourceRoot) Then
                    reason = "source path is empty."
                    Return False
                End If

                Dim volumeRoot = NormalizeVolumeRoot(Path.GetPathRoot(normalizedSourceRoot))
                If String.IsNullOrWhiteSpace(volumeRoot) Then
                    reason = "unable to resolve source volume."
                    Return False
                End If

                If volumeRoot.StartsWith("\\", StringComparison.OrdinalIgnoreCase) Then
                    reason = "UNC/network roots are not supported."
                    Return False
                End If

                Dim drive As DriveInfo = Nothing
                Try
                    drive = New DriveInfo(volumeRoot)
                Catch
                    reason = "unable to inspect source drive."
                    Return False
                End Try

                Try
                    If Not drive.IsReady Then
                        reason = "source drive is not ready."
                        Return False
                    End If
                Catch
                    reason = "source drive readiness could not be determined."
                    Return False
                End Try

                If drive.DriveType = DriveType.Network Then
                    reason = "network drives are not supported."
                    Return False
                End If

                Dim format = String.Empty
                Try
                    format = If(drive.DriveFormat, String.Empty)
                Catch
                    format = String.Empty
                End Try

                If Not String.Equals(format, "NTFS", StringComparison.OrdinalIgnoreCase) Then
                    reason = $"drive format '{format}' is unsupported (NTFS required)."
                    Return False
                End If

                Dim sectorSizeBytes = 0
                Dim clusterSizeBytes = 0
                If Not TryGetDiskGeometry(volumeRoot, sectorSizeBytes, clusterSizeBytes, reason) Then
                    Return False
                End If

                Dim volumePath = BuildRawVolumePath(volumeRoot)
                If volumePath.Length = 0 Then
                    reason = "unable to build raw volume path."
                    Return False
                End If

                Dim volumeHandle As SafeFileHandle = Nothing
                Try
                    volumeHandle = File.OpenHandle(
                        volumePath,
                        FileMode.Open,
                        FileAccess.Read,
                        FileShare.ReadWrite Or FileShare.Delete,
                        FileOptions.Asynchronous)
                Catch ex As UnauthorizedAccessException
                    reason = "access denied opening raw volume (run elevated to enable raw scan backend)."
                    Return False
                Catch ex As Exception
                    reason = ex.Message
                    Return False
                End Try

                If volumeHandle Is Nothing OrElse volumeHandle.IsInvalid Then
                    reason = "raw volume handle is invalid."
                    Try
                        volumeHandle?.Dispose()
                    Catch
                    End Try
                    Return False
                End If

                context = New RawDiskScanContext(normalizedSourceRoot, volumeRoot, volumeHandle, sectorSizeBytes, clusterSizeBytes)
                Return True
            End Function

            Public Function TryCreateReadPlan(
                sourcePath As String,
                offset As Long,
                count As Integer,
                ByRef plan As IReadOnlyList(Of RawDiskReadSegment),
                ByRef reason As String) As Boolean

                plan = Nothing
                reason = String.Empty
                If _disposed OrElse count <= 0 OrElse offset < 0 Then
                    reason = "invalid read request."
                    Return False
                End If

                Dim normalizedPath = NormalizePathForCache(sourcePath)
                If normalizedPath.Length = 0 Then
                    reason = "source path is invalid."
                    Return False
                End If

                Dim layoutReason As String = String.Empty
                Dim layout = GetOrCreateLayout(normalizedPath, layoutReason)
                If layout Is Nothing OrElse Not layout.IsSupported Then
                    reason = If(layoutReason, "file layout is unsupported.")
                    Return False
                End If

                Dim endOffset = offset + CLng(count)
                If endOffset < offset OrElse endOffset > layout.FileLength Then
                    reason = "requested raw read range exceeds file bounds."
                    Return False
                End If

                Dim segments As New List(Of RawDiskReadSegment)()
                Dim extentIndex = 0
                While extentIndex < layout.Extents.Count AndAlso
                    (layout.Extents(extentIndex).FileOffsetBytes + layout.Extents(extentIndex).LengthBytes) <= offset
                    extentIndex += 1
                End While

                Dim cursor = offset
                While cursor < endOffset
                    If extentIndex < layout.Extents.Count Then
                        Dim extent = layout.Extents(extentIndex)
                        Dim extentStart = extent.FileOffsetBytes
                        Dim extentEnd = extentStart + extent.LengthBytes

                        If cursor < extentStart Then
                            Dim sparseLength = CInt(Math.Min(CLng(Integer.MaxValue), Math.Min(endOffset - cursor, extentStart - cursor)))
                            segments.Add(RawDiskReadSegment.CreateSparse(sparseLength))
                            cursor += sparseLength
                            Continue While
                        End If

                        If cursor >= extentEnd Then
                            extentIndex += 1
                            Continue While
                        End If

                        Dim mappedLength = CInt(Math.Min(CLng(Integer.MaxValue), Math.Min(endOffset - cursor, extentEnd - cursor)))
                        Dim volumeOffset = extent.VolumeOffsetBytes + (cursor - extentStart)
                        segments.Add(RawDiskReadSegment.CreateMapped(volumeOffset, mappedLength))
                        cursor += mappedLength
                        Continue While
                    End If

                    Dim trailingSparseLength = CInt(Math.Min(CLng(Integer.MaxValue), endOffset - cursor))
                    segments.Add(RawDiskReadSegment.CreateSparse(trailingSparseLength))
                    cursor += trailingSparseLength
                End While

                plan = segments
                Return True
            End Function

            Private Function GetOrCreateLayout(sourcePath As String, ByRef reason As String) As RawDiskFileLayout
                reason = String.Empty
                Dim cached As RawDiskFileLayout = Nothing
                SyncLock _layoutCache
                    If _layoutCache.TryGetValue(sourcePath, cached) Then
                        reason = cached.UnsupportedReason
                        Return cached
                    End If
                End SyncLock

                Dim builtReason As String = String.Empty
                Dim builtLayout = BuildLayout(sourcePath, builtReason)
                If builtLayout Is Nothing Then
                    builtLayout = RawDiskFileLayout.CreateUnsupported(0, If(builtReason, "file layout unavailable."))
                End If

                SyncLock _layoutCache
                    _layoutCache(sourcePath) = builtLayout
                End SyncLock
                reason = builtLayout.UnsupportedReason
                Return builtLayout
            End Function

            Public Sub MarkPathUnsupported(sourcePath As String, reason As String)
                Dim normalizedPath = NormalizePathForCache(sourcePath)
                If normalizedPath.Length = 0 Then
                    Return
                End If

                Dim fileLength = 0L
                Try
                    fileLength = Math.Max(0L, New FileInfo(normalizedPath).Length)
                Catch
                    fileLength = 0L
                End Try

                Dim unsupported = RawDiskFileLayout.CreateUnsupported(fileLength, reason)
                SyncLock _layoutCache
                    _layoutCache(normalizedPath) = unsupported
                End SyncLock
            End Sub

            Private Function BuildLayout(sourcePath As String, ByRef reason As String) As RawDiskFileLayout
                reason = String.Empty

                If String.IsNullOrWhiteSpace(sourcePath) Then
                    reason = "source path is empty."
                    Return RawDiskFileLayout.CreateUnsupported(0, reason)
                End If

                Dim volumeRoot = NormalizeVolumeRoot(Path.GetPathRoot(sourcePath))
                If volumeRoot.Length = 0 OrElse
                    Not String.Equals(volumeRoot, _sourceVolumeRoot, StringComparison.OrdinalIgnoreCase) Then
                    reason = "file is outside the source volume."
                    Return RawDiskFileLayout.CreateUnsupported(0, reason)
                End If

                Dim fileLength = 0L
                Try
                    Dim info As New FileInfo(sourcePath)
                    fileLength = Math.Max(0L, info.Length)
                Catch ex As Exception
                    reason = ex.Message
                    Return RawDiskFileLayout.CreateUnsupported(0, reason)
                End Try

                If fileLength = 0 Then
                    Return RawDiskFileLayout.CreateSupported(fileLength, New List(Of RawDiskExtent)())
                End If

                Dim attributes As FileAttributes
                Try
                    attributes = File.GetAttributes(sourcePath)
                Catch ex As Exception
                    reason = ex.Message
                    Return RawDiskFileLayout.CreateUnsupported(fileLength, reason)
                End Try

                If (attributes And FileAttributes.Compressed) <> 0 Then
                    reason = "compressed files are not supported."
                    Return RawDiskFileLayout.CreateUnsupported(fileLength, reason)
                End If

                If (attributes And FileAttributes.Encrypted) <> 0 Then
                    reason = "encrypted files are not supported."
                    Return RawDiskFileLayout.CreateUnsupported(fileLength, reason)
                End If

                Dim extents As List(Of RawDiskExtent) = Nothing
                If Not TryGetRetrievalPointerExtents(sourcePath, fileLength, extents, reason) Then
                    Return RawDiskFileLayout.CreateUnsupported(fileLength, reason)
                End If

                Return RawDiskFileLayout.CreateSupported(fileLength, extents)
            End Function

            Private Function TryGetRetrievalPointerExtents(
                sourcePath As String,
                fileLength As Long,
                ByRef extents As List(Of RawDiskExtent),
                ByRef reason As String) As Boolean

                extents = New List(Of RawDiskExtent)()
                reason = String.Empty

                Const outputBufferSize As Integer = 64 * 1024
                Dim inputSize = Marshal.SizeOf(Of StartingVcnInputBuffer)()
                Dim headerSize = Marshal.SizeOf(Of RetrievalPointersBufferHeader)()
                Dim extentSize = Marshal.SizeOf(Of RetrievalPointersExtent)()

                Dim outputBuffer = Marshal.AllocHGlobal(outputBufferSize)
                Dim inputBuffer = Marshal.AllocHGlobal(inputSize)
                Try
                    Using fileStream As New FileStream(
                        sourcePath,
                        FileMode.Open,
                        FileAccess.Read,
                        FileShare.ReadWrite Or FileShare.Delete,
                        MinimumRescueBlockSize,
                        FileOptions.None)

                        Dim fileHandle = fileStream.SafeFileHandle
                        Dim startVcn = 0L
                        Dim lastStartVcn = Long.MinValue
                        Do
                            Dim input As New StartingVcnInputBuffer() With {
                                .StartingVcn = startVcn
                            }
                            Marshal.StructureToPtr(input, inputBuffer, fDeleteOld:=False)

                            Dim bytesReturned = 0
                            Dim success = DeviceIoControl(
                                fileHandle,
                                FsctlGetRetrievalPointers,
                                inputBuffer,
                                inputSize,
                                outputBuffer,
                                outputBufferSize,
                                bytesReturned,
                                IntPtr.Zero)

                            Dim errorCode = If(success, 0, Marshal.GetLastWin32Error())
                            If Not success AndAlso errorCode <> ErrorMoreData Then
                                reason = $"FSCTL_GET_RETRIEVAL_POINTERS failed ({errorCode})."
                                Return False
                            End If

                            If bytesReturned < headerSize Then
                                reason = "retrieval-pointers response is too short."
                                Return False
                            End If

                            Dim header = Marshal.PtrToStructure(Of RetrievalPointersBufferHeader)(outputBuffer)
                            Dim currentVcn = header.StartingVcn
                            For index = 0 To CInt(header.ExtentCount) - 1
                                Dim extentPtr = IntPtr.Add(outputBuffer, headerSize + (index * extentSize))
                                If IntPtr.Add(extentPtr, extentSize).ToInt64() > IntPtr.Add(outputBuffer, bytesReturned).ToInt64() Then
                                    Exit For
                                End If

                                Dim extent = Marshal.PtrToStructure(Of RetrievalPointersExtent)(extentPtr)
                                Dim nextVcn = extent.NextVcn
                                If nextVcn <= currentVcn Then
                                    Continue For
                                End If

                                Dim runClusters = nextVcn - currentVcn
                                Dim runFileOffsetBytes = currentVcn * CLng(_clusterSizeBytes)
                                Dim runLengthBytes = runClusters * CLng(_clusterSizeBytes)
                                If extent.Lcn >= 0 Then
                                    Dim volumeOffsetBytes = extent.Lcn * CLng(_clusterSizeBytes)
                                    extents.Add(New RawDiskExtent(runFileOffsetBytes, volumeOffsetBytes, runLengthBytes))
                                End If

                                currentVcn = nextVcn
                            Next

                            If success Then
                                Exit Do
                            End If

                            startVcn = currentVcn
                            If startVcn <= lastStartVcn Then
                                reason = "retrieval-pointers walk did not advance."
                                Return False
                            End If
                            lastStartVcn = startVcn
                        Loop
                    End Using
                Catch ex As Exception
                    reason = ex.Message
                    Return False
                Finally
                    Marshal.FreeHGlobal(inputBuffer)
                    Marshal.FreeHGlobal(outputBuffer)
                End Try

                extents = NormalizeAndMergeExtents(extents, fileLength)
                Return True
            End Function

            Private Shared Function NormalizeAndMergeExtents(extents As IEnumerable(Of RawDiskExtent), fileLength As Long) As List(Of RawDiskExtent)
                Dim normalized As New List(Of RawDiskExtent)()
                If extents Is Nothing OrElse fileLength <= 0 Then
                    Return normalized
                End If

                Dim ordered = extents.
                    Where(Function(item) item IsNot Nothing AndAlso item.LengthBytes > 0).
                    OrderBy(Function(item) item.FileOffsetBytes).
                    ToList()

                For Each item In ordered
                    Dim startOffset = Math.Max(0L, item.FileOffsetBytes)
                    If startOffset >= fileLength Then
                        Continue For
                    End If

                    Dim length = Math.Min(item.LengthBytes, fileLength - startOffset)
                    If length <= 0 Then
                        Continue For
                    End If

                    Dim candidate As New RawDiskExtent(startOffset, item.VolumeOffsetBytes + (startOffset - item.FileOffsetBytes), length)
                    If normalized.Count = 0 Then
                        normalized.Add(candidate)
                        Continue For
                    End If

                    Dim previous = normalized(normalized.Count - 1)
                    Dim previousFileEnd = previous.FileOffsetBytes + previous.LengthBytes
                    Dim previousVolumeEnd = previous.VolumeOffsetBytes + previous.LengthBytes
                    If previousFileEnd = candidate.FileOffsetBytes AndAlso previousVolumeEnd = candidate.VolumeOffsetBytes Then
                        normalized(normalized.Count - 1) = New RawDiskExtent(previous.FileOffsetBytes, previous.VolumeOffsetBytes, previous.LengthBytes + candidate.LengthBytes)
                    Else
                        normalized.Add(candidate)
                    End If
                Next

                Return normalized
            End Function

            Private Shared Function NormalizePathForCache(pathValue As String) As String
                If String.IsNullOrWhiteSpace(pathValue) Then
                    Return String.Empty
                End If

                Try
                    Return Path.GetFullPath(pathValue).Trim()
                Catch
                    Return pathValue.Trim()
                End Try
            End Function

            Private Shared Function NormalizeVolumeRoot(value As String) As String
                If String.IsNullOrWhiteSpace(value) Then
                    Return String.Empty
                End If

                Dim fullRoot = value.Trim()
                If fullRoot.Length >= 2 AndAlso fullRoot(1) = ":"c Then
                    Return fullRoot.Substring(0, 2).ToUpperInvariant() & Path.DirectorySeparatorChar
                End If

                Return fullRoot.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar)
            End Function

            Private Shared Function BuildRawVolumePath(volumeRoot As String) As String
                Dim normalized = NormalizeVolumeRoot(volumeRoot)
                If normalized.Length < 2 OrElse normalized(1) <> ":"c Then
                    Return String.Empty
                End If

                Dim driveLetter = Char.ToUpperInvariant(normalized(0))
                Return $"\\.\{driveLetter}:"
            End Function

            Private Shared Function TryGetDiskGeometry(
                volumeRoot As String,
                ByRef sectorSizeBytes As Integer,
                ByRef clusterSizeBytes As Integer,
                ByRef reason As String) As Boolean

                sectorSizeBytes = 0
                clusterSizeBytes = 0
                reason = String.Empty

                Dim sectorsPerCluster As UInteger = 0UI
                Dim bytesPerSector As UInteger = 0UI
                Dim freeClusters As UInteger = 0UI
                Dim totalClusters As UInteger = 0UI
                Dim success = GetDiskFreeSpace(
                    NormalizeVolumeRoot(volumeRoot),
                    sectorsPerCluster,
                    bytesPerSector,
                    freeClusters,
                    totalClusters)
                If Not success Then
                    reason = $"GetDiskFreeSpace failed ({Marshal.GetLastWin32Error()})."
                    Return False
                End If

                If bytesPerSector < 512UI OrElse bytesPerSector > CUInt(Integer.MaxValue) Then
                    reason = "invalid sector size reported by volume."
                    Return False
                End If

                Dim computedSectorSize = CInt(bytesPerSector)
                Dim computedCluster = CLng(sectorsPerCluster) * CLng(bytesPerSector)
                If computedCluster < MinimumRescueBlockSize OrElse computedCluster > Integer.MaxValue Then
                    reason = "invalid cluster size reported by volume."
                    Return False
                End If

                sectorSizeBytes = computedSectorSize
                clusterSizeBytes = CInt(computedCluster)
                Return True
            End Function

            Public Sub Dispose() Implements IDisposable.Dispose
                If _disposed Then
                    Return
                End If

                _disposed = True
                SyncLock _layoutCache
                    _layoutCache.Clear()
                End SyncLock
                Try
                    _volumeHandle?.Dispose()
                Catch
                End Try
            End Sub
        End Class

        Private NotInheritable Class RawDiskFileLayout
            Private Sub New(isSupported As Boolean, fileLength As Long, extents As List(Of RawDiskExtent), unsupportedReason As String)
                Me.IsSupported = isSupported
                Me.FileLength = Math.Max(0L, fileLength)
                Me.Extents = If(extents, New List(Of RawDiskExtent)())
                Me.UnsupportedReason = If(unsupportedReason, String.Empty).Trim()
            End Sub

            Public ReadOnly Property IsSupported As Boolean
            Public ReadOnly Property FileLength As Long
            Public ReadOnly Property Extents As List(Of RawDiskExtent)
            Public ReadOnly Property UnsupportedReason As String

            Public Shared Function CreateSupported(fileLength As Long, extents As List(Of RawDiskExtent)) As RawDiskFileLayout
                Return New RawDiskFileLayout(True, fileLength, extents, String.Empty)
            End Function

            Public Shared Function CreateUnsupported(fileLength As Long, reason As String) As RawDiskFileLayout
                Return New RawDiskFileLayout(False, fileLength, New List(Of RawDiskExtent)(), reason)
            End Function
        End Class

        Private NotInheritable Class RawDiskExtent
            Public Sub New(fileOffsetBytes As Long, volumeOffsetBytes As Long, lengthBytes As Long)
                Me.FileOffsetBytes = Math.Max(0L, fileOffsetBytes)
                Me.VolumeOffsetBytes = Math.Max(0L, volumeOffsetBytes)
                Me.LengthBytes = Math.Max(0L, lengthBytes)
            End Sub

            Public ReadOnly Property FileOffsetBytes As Long
            Public ReadOnly Property VolumeOffsetBytes As Long
            Public ReadOnly Property LengthBytes As Long
        End Class

        Private NotInheritable Class RawDiskReadSegment
            Private Sub New(isSparse As Boolean, volumeOffsetBytes As Long, length As Integer)
                Me.IsSparse = isSparse
                Me.VolumeOffsetBytes = Math.Max(0L, volumeOffsetBytes)
                Me.Length = Math.Max(0, length)
            End Sub

            Public ReadOnly Property IsSparse As Boolean
            Public ReadOnly Property VolumeOffsetBytes As Long
            Public ReadOnly Property Length As Integer

            Public Shared Function CreateMapped(volumeOffsetBytes As Long, length As Integer) As RawDiskReadSegment
                Return New RawDiskReadSegment(False, volumeOffsetBytes, length)
            End Function

            Public Shared Function CreateSparse(length As Integer) As RawDiskReadSegment
                Return New RawDiskReadSegment(True, 0L, length)
            End Function
        End Class

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

            Public Function TryGetOpenSourceHandle() As SafeFileHandle
                ThrowIfDisposed()
                If _sourceStream Is Nothing Then
                    Return Nothing
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

            Public Function TryGetOpenDestinationHandle() As SafeFileHandle
                ThrowIfDisposed()
                If _destinationStream Is Nothing Then
                    Return Nothing
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

        Private NotInheritable Class FragileReadSkipException
            Inherits IOException

            Public Sub New(message As String, innerException As Exception)
                MyBase.New(message, innerException)
            End Sub
        End Class
    End Class
End Namespace
