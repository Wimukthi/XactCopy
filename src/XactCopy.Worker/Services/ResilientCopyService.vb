Imports System.Collections.Generic
Imports System.IO
Imports System.Linq
Imports System.Security.Cryptography
Imports System.Threading
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

        Private ReadOnly _options As CopyJobOptions
        Private ReadOnly _executionControl As CopyExecutionControl
        Private ReadOnly _journalStore As JobJournalStore
        Private _lastJournalFlushUtc As DateTimeOffset = DateTimeOffset.MinValue
        Private _throttleWindowStartUtc As DateTimeOffset = DateTimeOffset.MinValue
        Private _throttleWindowBytes As Long

        Public Sub New(options As CopyJobOptions, Optional executionControl As CopyExecutionControl = Nothing)
            _options = options
            _executionControl = If(executionControl, New CopyExecutionControl())
            _journalStore = New JobJournalStore()
        End Sub

        Public Async Function RunAsync(cancellationToken As CancellationToken) As Task(Of CopyJobResult)
            ValidateOptions()

            Dim sourceRoot = Path.GetFullPath(_options.SourceRoot)
            Dim destinationRoot = Path.GetFullPath(_options.DestinationRoot)
            Await WaitForMediaAvailabilityAsync(sourceRoot, destinationRoot, cancellationToken).ConfigureAwait(False)
            Directory.CreateDirectory(destinationRoot)

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

            If _options.CopyEmptyDirectories AndAlso scanResult.Directories.Count > 0 Then
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
            Dim journalPath = JobJournalStore.GetDefaultJournalPath(jobId)

            Dim existingJournal As JobJournal = Nothing
            If _options.ResumeFromJournal Then
                existingJournal = Await _journalStore.LoadAsync(journalPath, cancellationToken).ConfigureAwait(False)
                If existingJournal IsNot Nothing Then
                    EmitLog($"Loaded journal: {journalPath}")
                Else
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

                Dim entry = journal.Files(descriptor.RelativePath)

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
                    Await FlushJournalAsync(journalPath, journal, cancellationToken, force:=True).ConfigureAwait(False)
                    Continue For
                End If

                progress.FailedFiles += 1
                entry.State = FileCopyState.Failed
                entry.LastError = copyException.Message
                EmitLog($"Failed: {descriptor.RelativePath} ({copyException.Message})")
                Await FlushJournalAsync(journalPath, journal, cancellationToken, force:=True).ConfigureAwait(False)

                If Not _options.ContinueOnFileError Then
                    Return CreateResult(progress, journalPath, succeeded:=False, cancelled:=False, errorMessage:=copyException.Message)
                End If
            Next

            Await FlushJournalAsync(journalPath, journal, cancellationToken, force:=True).ConfigureAwait(False)
            Return CreateResult(progress, journalPath, succeeded:=(progress.FailedFiles = 0), cancelled:=False, errorMessage:=String.Empty)
        End Function

        Private Sub ValidateOptions()
            If String.IsNullOrWhiteSpace(_options.SourceRoot) Then
                Throw New InvalidOperationException("Source path is required.")
            End If

            If String.IsNullOrWhiteSpace(_options.DestinationRoot) Then
                Throw New InvalidOperationException("Destination path is required.")
            End If

            If Not _options.WaitForMediaAvailability AndAlso Not Directory.Exists(_options.SourceRoot) Then
                Throw New DirectoryNotFoundException($"Source directory not found: {_options.SourceRoot}")
            End If

            Dim sourceRoot = Path.GetFullPath(_options.SourceRoot).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar)
            Dim destinationRoot = Path.GetFullPath(_options.DestinationRoot).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar)

            If StringComparer.OrdinalIgnoreCase.Equals(sourceRoot, destinationRoot) Then
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

            If _options.MaxThroughputBytesPerSecond < 0 Then
                Throw New InvalidOperationException("MaxThroughputBytesPerSecond cannot be negative.")
            End If

            If _options.SampleVerificationChunkBytes <= 0 Then
                Throw New InvalidOperationException("SampleVerificationChunkBytes must be greater than zero.")
            End If

            If _options.SampleVerificationChunkCount <= 0 Then
                Throw New InvalidOperationException("SampleVerificationChunkCount must be greater than zero.")
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
        End Sub

        Private Function MergeJournal(
            existing As JobJournal,
            sourceFiles As IReadOnlyList(Of SourceFileDescriptor),
            sourceRoot As String,
            destinationRoot As String,
            jobId As String) As JobJournal

            Dim canReuseJournal = existing IsNot Nothing AndAlso
                StringComparer.OrdinalIgnoreCase.Equals(Path.GetFullPath(existing.SourceRoot), sourceRoot) AndAlso
                StringComparer.OrdinalIgnoreCase.Equals(Path.GetFullPath(existing.DestinationRoot), destinationRoot)

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

            entry.RescueRanges = BuildRescueRanges(entry, descriptor.Length, destinationLength)
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
                rescueBadRegionCount:=GetRescueRegionCount(entry.RescueRanges, RescueRangeState.Bad),
                rescueRemainingBytes:=GetRescueRangeBytes(entry.RescueRanges, RescueRangeState.Pending, RescueRangeState.Bad))

            Dim passPlan = BuildRescuePassPlan(activeBufferSize)
            For Each passDefinition In passPlan
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
                    progress,
                    journal,
                    journalPath,
                    cancellationToken).ConfigureAwait(False)

                Dim remainingBadRegions = GetRescueRegionCount(entry.RescueRanges, RescueRangeState.Bad)
                Dim passRemainingBadBytes = GetRescueRangeBytes(entry.RescueRanges, RescueRangeState.Bad)
                If ShouldLogRescuePassSummary(passDefinition, passOutcome, remainingBadRegions) Then
                    EmitLog(
                        $"[{RescueCoreName}] {passDefinition.Name} {descriptor.RelativePath}: attempts {passOutcome.AttemptedSegments}, recovered {FormatBytes(passOutcome.RecoveredBytes)}, failed {passOutcome.FailedSegments}, remaining bad {remainingBadRegions} ({FormatBytes(passRemainingBadBytes)}).")
                End If

                SyncEntryBytesCopied(entry)
                Await FlushJournalAsync(journalPath, journal, cancellationToken, force:=True).ConfigureAwait(False)
            Next

            Dim recoveredAny = GetRescueRegionCount(entry.RescueRanges, RescueRangeState.Recovered) > 0
            Dim remainingBadBytes = GetRescueRangeBytes(entry.RescueRanges, RescueRangeState.Bad)
            If remainingBadBytes > 0 Then
                If Not _options.SalvageUnreadableBlocks Then
                    Throw New IOException(
                        $"[{RescueCoreName}] Unrecoverable regions remain on {descriptor.RelativePath}: {FormatBytes(remainingBadBytes)} in {GetRescueRegionCount(entry.RescueRanges, RescueRangeState.Bad)} block(s).")
                End If

                entry.LastRescuePass = "SalvageFill"
                EmitLog($"[{RescueCoreName}] Salvaging remaining unreadable regions on {descriptor.RelativePath}.")
                recoveredAny = Await SalvageRemainingRangesAsync(
                    descriptor,
                    entry,
                    destinationPath,
                    activeBufferSize,
                    progress,
                    journal,
                    journalPath,
                    cancellationToken).ConfigureAwait(False) OrElse recoveredAny
                SyncEntryBytesCopied(entry)
                entry.RecoveredRanges = MergeByteRanges(entry.RecoveredRanges)
                Await FlushJournalAsync(journalPath, journal, cancellationToken, force:=True).ConfigureAwait(False)
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
                Case RescueRangeState.Pending, RescueRangeState.Good, RescueRangeState.Bad, RescueRangeState.Recovered
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

        Private Shared Function GetRescueRegionCount(ranges As IEnumerable(Of RescueRange), state As RescueRangeState) As Integer
            If ranges Is Nothing Then
                Return 0
            End If

            Dim targetState = NormalizeRescueRangeState(state)
            Return ranges.Count(Function(item) item IsNot Nothing AndAlso item.Length > 0 AndAlso NormalizeRescueRangeState(item.State) = targetState)
        End Function

        Private Shared Function SnapshotRangesByState(ranges As IEnumerable(Of RescueRange), state As RescueRangeState) As List(Of ByteRange)
            Dim snapshots As New List(Of ByteRange)()
            If ranges Is Nothing Then
                Return snapshots
            End If

            Dim targetState = NormalizeRescueRangeState(state)
            For Each item In ranges
                If item Is Nothing OrElse item.Length <= 0 Then
                    Continue For
                End If

                If NormalizeRescueRangeState(item.State) <> targetState Then
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

        Private Sub SyncEntryBytesCopied(entry As JournalFileEntry)
            If entry Is Nothing Then
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
                    minimumSplitBytes:=trimSplitMinimum),
                New RescuePassDefinition(
                    name:="TrimSweep",
                    targetState:=RescueRangeState.Bad,
                    chunkSizeBytes:=trimChunkSize,
                    maxReadRetries:=trimRetries,
                    splitOnFailure:=True,
                    minimumSplitBytes:=trimSplitMinimum),
                New RescuePassDefinition(
                    name:="Scrape",
                    targetState:=RescueRangeState.Bad,
                    chunkSizeBytes:=scrapeChunkSize,
                    maxReadRetries:=scrapeRetries,
                    splitOnFailure:=True,
                    minimumSplitBytes:=scrapeSplitMinimum),
                New RescuePassDefinition(
                    name:="RetryBad",
                    targetState:=RescueRangeState.Bad,
                    chunkSizeBytes:=retryChunkSize,
                    maxReadRetries:=Math.Max(0, _options.MaxRetries),
                    splitOnFailure:=False,
                    minimumSplitBytes:=retryChunkSize)
            }
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
            progress As ProgressAccumulator,
            journal As JobJournal,
            journalPath As String,
            cancellationToken As CancellationToken) As Task(Of RescuePassOutcome)

            Dim outcome As New RescuePassOutcome()
            Dim targetSnapshots = SnapshotRangesByState(entry.RescueRanges, passDefinition.TargetState)
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
                        Dim remaining = segmentLength
                        While remaining > 0
                            Dim chunkLength = Math.Min(passDefinition.ChunkSizeBytes, remaining)
                            work.Push(New ByteRange() With {
                                .Offset = segment.Offset + (remaining - chunkLength),
                                .Length = chunkLength
                            })
                            remaining -= chunkLength
                        End While
                        Continue While
                    End If

                    outcome.AttemptedSegments += 1
                    Dim readResult = Await ReadChunkWithRetriesAsync(
                        descriptor.FullPath,
                        descriptor.RelativePath,
                        segment.Offset,
                        segmentLength,
                        passDefinition.ChunkSizeBytes,
                        cancellationToken,
                        maxRetries:=passDefinition.MaxReadRetries,
                        allowSalvage:=False).ConfigureAwait(False)

                    If readResult IsNot Nothing Then
                        Await WriteChunkWithRetriesAsync(
                            destinationPath,
                            descriptor.RelativePath,
                            segment.Offset,
                            readResult.Buffer,
                            readResult.Count,
                            passDefinition.ChunkSizeBytes,
                            cancellationToken).ConfigureAwait(False)
                        Await ApplyThroughputThrottleAsync(readResult.Count, cancellationToken).ConfigureAwait(False)

                        SetRescueRangeState(entry.RescueRanges, segment.Offset, readResult.Count, RescueRangeState.Good)
                        progress.TotalBytesCopied += readResult.Count
                        SyncEntryBytesCopied(entry)
                        EmitProgress(
                            progress,
                            descriptor.RelativePath,
                            entry.BytesCopied,
                            descriptor.Length,
                            lastChunkBytesTransferred:=readResult.Count,
                            bufferSizeBytes:=passDefinition.ChunkSizeBytes,
                            rescuePass:=passDefinition.Name,
                            rescueBadRegionCount:=GetRescueRegionCount(entry.RescueRanges, RescueRangeState.Bad),
                            rescueRemainingBytes:=GetRescueRangeBytes(entry.RescueRanges, RescueRangeState.Pending, RescueRangeState.Bad))
                        Await FlushJournalAsync(journalPath, journal, cancellationToken, force:=False).ConfigureAwait(False)
                        outcome.RecoveredBytes += readResult.Count
                        Continue While
                    End If

                    If passDefinition.SplitOnFailure AndAlso segmentLength >= passDefinition.MinimumSplitBytes * 2 Then
                        Dim halfLength = Math.Max(passDefinition.MinimumSplitBytes, AlignToBlockSize(segmentLength \ 2, passDefinition.MinimumSplitBytes))
                        halfLength = Math.Min(segmentLength - passDefinition.MinimumSplitBytes, halfLength)
                        If halfLength > 0 AndAlso halfLength < segmentLength Then
                            Dim secondLength = segmentLength - halfLength
                            work.Push(New ByteRange() With {
                                .Offset = segment.Offset + halfLength,
                                .Length = secondLength
                            })
                            work.Push(New ByteRange() With {
                                .Offset = segment.Offset,
                                .Length = halfLength
                            })
                            Continue While
                        End If
                    End If

                    SetRescueRangeState(entry.RescueRanges, segment.Offset, segmentLength, RescueRangeState.Bad)
                    SyncEntryBytesCopied(entry)
                    EmitProgress(
                        progress,
                        descriptor.RelativePath,
                        entry.BytesCopied,
                        descriptor.Length,
                        lastChunkBytesTransferred:=0,
                        bufferSizeBytes:=passDefinition.ChunkSizeBytes,
                        rescuePass:=passDefinition.Name,
                        rescueBadRegionCount:=GetRescueRegionCount(entry.RescueRanges, RescueRangeState.Bad),
                        rescueRemainingBytes:=GetRescueRangeBytes(entry.RescueRanges, RescueRangeState.Pending, RescueRangeState.Bad))
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
            ioBufferSize As Integer,
            progress As ProgressAccumulator,
            journal As JobJournal,
            journalPath As String,
            cancellationToken As CancellationToken) As Task(Of Boolean)

            Dim recoveredAny = False
            Dim badSnapshots = SnapshotRangesByState(entry.RescueRanges, RescueRangeState.Bad)
            For Each badRange In badSnapshots
                Dim offset = badRange.Offset
                Dim remaining = badRange.Length
                While remaining > 0
                    cancellationToken.ThrowIfCancellationRequested()
                    Await _executionControl.WaitIfPausedAsync(cancellationToken).ConfigureAwait(False)
                    Await WaitForMediaAvailabilityAsync(_options.SourceRoot, _options.DestinationRoot, cancellationToken).ConfigureAwait(False)

                    Dim chunkLength = Math.Min(ioBufferSize, remaining)
                    Dim salvageBuffer = CreateSalvageBuffer(chunkLength)
                    Await WriteChunkWithRetriesAsync(
                        destinationPath,
                        descriptor.RelativePath,
                        offset,
                        salvageBuffer,
                        chunkLength,
                        ioBufferSize,
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
                        rescueBadRegionCount:=GetRescueRegionCount(entry.RescueRanges, RescueRangeState.Bad),
                        rescueRemainingBytes:=GetRescueRangeBytes(entry.RescueRanges, RescueRangeState.Pending, RescueRangeState.Bad))
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
                minimumSplitBytes As Integer)

                Me.Name = If(String.IsNullOrWhiteSpace(name), "Pass", name.Trim())
                Me.TargetState = NormalizeRescueRangeState(targetState)
                Me.ChunkSizeBytes = NormalizeIoBufferSize(chunkSizeBytes)
                Me.MaxReadRetries = Math.Max(0, maxReadRetries)
                Me.SplitOnFailure = splitOnFailure
                Me.MinimumSplitBytes = NormalizeIoBufferSize(Math.Max(MinimumRescueBlockSize, minimumSplitBytes))
            End Sub

            Public ReadOnly Property Name As String
            Public ReadOnly Property TargetState As RescueRangeState
            Public ReadOnly Property ChunkSizeBytes As Integer
            Public ReadOnly Property MaxReadRetries As Integer
            Public ReadOnly Property SplitOnFailure As Boolean
            Public ReadOnly Property MinimumSplitBytes As Integer
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
            cancellationToken As CancellationToken,
            Optional maxRetries As Integer = -1,
            Optional allowSalvage As Boolean = True) As Task(Of ChunkReadResult)

            Dim buffer(length - 1) As Byte
            Dim attempt = 0
            Dim lastError As Exception = Nothing
            Dim effectiveMaxRetries = If(maxRetries >= 0, maxRetries, _options.MaxRetries)

            While attempt <= effectiveMaxRetries
                cancellationToken.ThrowIfCancellationRequested()

                Try
                    Using linkedCts = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken)
                        linkedCts.CancelAfter(_options.OperationTimeout)
                        Dim bytesRead = Await ReadChunkOnceAsync(
                            sourcePath,
                            offset,
                            buffer,
                            length,
                            ioBufferSize,
                            linkedCts.Token).ConfigureAwait(False)

                        If bytesRead <> length Then
                            Throw New IOException($"Short read at offset {offset}. Expected {length}, got {bytesRead}.")
                        End If

                        Return New ChunkReadResult(buffer, bytesRead, usedRecovery:=False)
                    End Using
                Catch ex As OperationCanceledException When Not cancellationToken.IsCancellationRequested
                    lastError = New TimeoutException($"Read timeout at offset {offset}.", ex)
                Catch ex As Exception
                    lastError = ex
                End Try

                If _options.WaitForMediaAvailability AndAlso IsAvailabilityRelatedException(lastError) Then
                    EmitLog($"Source unavailable during read of {relativePath}. Waiting for media to return.")
                    Await WaitForSourceFileAsync(sourcePath, cancellationToken).ConfigureAwait(False)
                    attempt = 0
                    Continue While
                End If

                attempt += 1
                If attempt > effectiveMaxRetries Then
                    Exit While
                End If

                EmitLog($"Read retry {attempt}/{effectiveMaxRetries} on {relativePath} at {FormatBytes(offset)}: {lastError.Message}")
                Await DelayForRetryAsync(attempt, cancellationToken).ConfigureAwait(False)
            End While

            If IsFatalReadException(lastError) Then
                Throw New IOException($"Read failed on {relativePath} at offset {offset}.", lastError)
            End If

            If allowSalvage AndAlso _options.SalvageUnreadableBlocks Then
                Dim recoveredBuffer = CreateSalvageBuffer(length)
                Dim fillDescription = DescribeSalvageFillPattern(_options.SalvageFillPattern)
                EmitLog($"Recovered unreadable block on {relativePath} at {FormatBytes(offset)} ({length} bytes {fillDescription}-filled).")
                Return New ChunkReadResult(recoveredBuffer, length, usedRecovery:=True)
            End If

            Return Nothing
        End Function

        Private Async Function ReadChunkOnceAsync(
            sourcePath As String,
            offset As Long,
            buffer As Byte(),
            count As Integer,
            ioBufferSize As Integer,
            cancellationToken As CancellationToken) As Task(Of Integer)

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

        Private Async Function WriteChunkWithRetriesAsync(
            destinationPath As String,
            relativePath As String,
            offset As Long,
            buffer As Byte(),
            count As Integer,
            ioBufferSize As Integer,
            cancellationToken As CancellationToken) As Task

            Dim attempt = 0
            Dim lastError As Exception = Nothing

            While attempt <= _options.MaxRetries
                cancellationToken.ThrowIfCancellationRequested()

                Try
                    Using linkedCts = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken)
                        linkedCts.CancelAfter(_options.OperationTimeout)
                        Await WriteChunkOnceAsync(
                            destinationPath,
                            offset,
                            buffer,
                            count,
                            ioBufferSize,
                            linkedCts.Token).ConfigureAwait(False)
                        Return
                    End Using
                Catch ex As OperationCanceledException When Not cancellationToken.IsCancellationRequested
                    lastError = New TimeoutException($"Write timeout at offset {offset}.", ex)
                Catch ex As Exception
                    lastError = ex
                End Try

                If _options.WaitForMediaAvailability AndAlso IsAvailabilityRelatedException(lastError) Then
                    EmitLog($"Destination unavailable during write of {relativePath}. Waiting for media to return.")
                    Await WaitForDestinationPathAsync(destinationPath, cancellationToken).ConfigureAwait(False)
                    attempt = 0
                    Continue While
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
            cancellationToken As CancellationToken) As Task

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

            While attempt <= _options.MaxRetries
                cancellationToken.ThrowIfCancellationRequested()

                Try
                    Using linkedCts = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken)
                        linkedCts.CancelAfter(_options.OperationTimeout)

                        Dim buffer(count - 1) As Byte
                        Dim bytesRead = Await ReadChunkOnceAsync(
                            filePath,
                            offset,
                            buffer,
                            count,
                            _options.BufferSizeBytes,
                            linkedCts.Token).ConfigureAwait(False)

                        If bytesRead <> count Then
                            Throw New IOException($"Short read while sampling {context}. Expected {count}, got {bytesRead}.")
                        End If

                        Return ComputeHash(buffer, bytesRead)
                    End Using
                Catch ex As OperationCanceledException When Not cancellationToken.IsCancellationRequested
                    lastError = New TimeoutException($"Verification timeout for {context}.", ex)
                Catch ex As Exception
                    lastError = ex
                End Try

                attempt += 1
                If attempt > _options.MaxRetries Then
                    Exit While
                End If

                EmitLog($"Verification retry {attempt}/{_options.MaxRetries} on {context}: {lastError.Message}")
                Await DelayForRetryAsync(attempt, cancellationToken).ConfigureAwait(False)
            End While

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

                Dim sourceAvailable = Directory.Exists(sourceRoot)
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

        Private Async Function WaitForSourceFileAsync(sourcePath As String, cancellationToken As CancellationToken) As Task
            If Not _options.WaitForMediaAvailability Then
                Return
            End If

            Dim lastLogUtc = DateTimeOffset.MinValue
            Do
                cancellationToken.ThrowIfCancellationRequested()
                Await _executionControl.WaitIfPausedAsync(cancellationToken).ConfigureAwait(False)

                Dim directoryPath = Path.GetDirectoryName(sourcePath)
                Dim fileAvailable = File.Exists(sourcePath)
                Dim directoryAvailable = Not String.IsNullOrWhiteSpace(directoryPath) AndAlso Directory.Exists(directoryPath)

                If fileAvailable Then
                    Return
                End If

                Dim nowUtc = DateTimeOffset.UtcNow
                If (nowUtc - lastLogUtc).TotalSeconds >= 5 Then
                    EmitLog($"Waiting for source path: {sourcePath} (directory {(If(directoryAvailable, "online", "offline"))})")
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

        Private Shared Function IsDestinationAvailable(targetPath As String) As Boolean
            If String.IsNullOrWhiteSpace(targetPath) Then
                Return False
            End If

            Try
                Dim fullPath = Path.GetFullPath(targetPath)
                Dim root = Path.GetPathRoot(fullPath)
                If String.IsNullOrWhiteSpace(root) OrElse Not Directory.Exists(root) Then
                    Return False
                End If

                Directory.CreateDirectory(fullPath)
                Return True
            Catch
                Return False
            End Try
        End Function

        Private Shared Function IsAvailabilityRelatedException(ex As Exception) As Boolean
            If ex Is Nothing Then
                Return False
            End If

            If TypeOf ex Is DriveNotFoundException OrElse TypeOf ex Is DirectoryNotFoundException Then
                Return True
            End If

            Dim ioEx = TryCast(ex, IOException)
            If ioEx Is Nothing Then
                Return False
            End If

            Dim message = ioEx.Message.ToLowerInvariant()
            Return message.Contains("device is not ready") OrElse
                message.Contains("network path") OrElse
                message.Contains("network name") OrElse
                message.Contains("cannot find the path") OrElse
                message.Contains("specified network")
        End Function

        Private Shared Function IsFatalReadException(ex As Exception) As Boolean
            If ex Is Nothing Then
                Return False
            End If

            Return TypeOf ex Is UnauthorizedAccessException OrElse
                TypeOf ex Is NotSupportedException OrElse
                TypeOf ex Is System.Security.SecurityException OrElse
                TypeOf ex Is ArgumentException
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

        Private Function CreateSalvageBuffer(length As Integer) As Byte()
            Dim buffer(length - 1) As Byte
            Select Case _options.SalvageFillPattern
                Case SalvageFillPattern.Ones
                    For index = 0 To buffer.Length - 1
                        buffer(index) = &HFF
                    Next
                Case SalvageFillPattern.Random
                    RandomNumberGenerator.Fill(buffer)
                Case Else
                    ' Zero is default byte state.
            End Select

            Return buffer
        End Function

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

        Private NotInheritable Class ChunkReadResult
            Public Sub New(buffer As Byte(), count As Integer, usedRecovery As Boolean)
                Me.Buffer = buffer
                Me.Count = count
                Me.UsedRecovery = usedRecovery
            End Sub

            Public ReadOnly Property Buffer As Byte()
            Public ReadOnly Property Count As Integer
            Public ReadOnly Property UsedRecovery As Boolean
        End Class

        Private NotInheritable Class ProgressAccumulator
            Public Property TotalFiles As Integer
            Public Property TotalBytes As Long
            Public Property TotalBytesCopied As Long
            Public Property CompletedFiles As Integer
            Public Property FailedFiles As Integer
            Public Property RecoveredFiles As Integer
            Public Property SkippedFiles As Integer
        End Class
    End Class
End Namespace
