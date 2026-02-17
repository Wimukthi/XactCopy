Imports System.IO
Imports System.Linq
Imports System.Security.Cryptography
Imports System.Text
Imports System.Text.Json
Imports System.Text.Json.Serialization
Imports System.Threading
Imports XactCopy.Models

Namespace Infrastructure
    Public Class JobJournalStore
        Private Const BackupGenerationCount As Integer = 3
        Private Const LedgerMagic As UInteger = &H58434A4CUL ' XCJL
        Private Const LedgerVersion As Byte = 1
        Private Const LedgerFrameHeaderSize As Integer = 9
        Private Const LedgerFrameTrailerSize As Integer = 4
        Private Const LedgerMaxBytes As Long = 8L * 1024L * 1024L
        Private Const LedgerCompactRecordCount As Integer = 2048
        Private Const HmacKeySizeBytes As Integer = 32
        Private Const SecurityKeyFileName As String = "journal-hmac.key"

        Private Shared ReadOnly HmacKeySync As New Object()
        Private Shared CachedHmacKey As Byte()

        Private ReadOnly _seedStateSync As New Object()
        Private ReadOnly _seedStateByPath As New Dictionary(Of String, LedgerSeedState)(StringComparer.OrdinalIgnoreCase)
        Private ReadOnly _serializerOptions As JsonSerializerOptions

        Public Sub New()
            _serializerOptions = New JsonSerializerOptions() With {
                .WriteIndented = True
            }
            _serializerOptions.Converters.Add(New JsonStringEnumConverter())
        End Sub

        Public Shared Function BuildJobId(sourceRoot As String, destinationRoot As String) As String
            Dim payload = $"{sourceRoot.Trim().ToUpperInvariant()}|{destinationRoot.Trim().ToUpperInvariant()}"
            Dim hash = SHA256.HashData(Encoding.UTF8.GetBytes(payload))
            Return Convert.ToHexString(hash).Substring(0, 20).ToLowerInvariant()
        End Function

        Public Shared Function GetDefaultJournalPath(jobId As String) As String
            Dim root = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "XactCopy",
                "journals")
            Directory.CreateDirectory(root)
            Return Path.Combine(root, $"job-{jobId}.json")
        End Function

        Public Async Function LoadAsync(journalPath As String, cancellationToken As CancellationToken) As Task(Of JobJournal)
            If String.IsNullOrWhiteSpace(journalPath) Then
                Return Nothing
            End If

            Dim candidates = BuildSnapshotCandidates(journalPath)
            If candidates.Count = 0 Then
                Return Nothing
            End If

            Dim anySnapshotExists = False
            For Each path In candidates
                If File.Exists(path) Then
                    anySnapshotExists = True
                    Exit For
                End If
            Next

            If Not anySnapshotExists Then
                Return Nothing
            End If

            Dim securityContext = CreateSecurityContext(journalPath)
            Dim trustedHashes = Await LoadTrustedSnapshotHashesAsync(journalPath, securityContext, cancellationToken).ConfigureAwait(False)
            If trustedHashes.Count > 0 Then
                Dim trustedSnapshots As New Dictionary(Of String, JobJournal)(StringComparer.OrdinalIgnoreCase)
                For Each candidatePath In candidates
                    Dim snapshot = Await TryReadSnapshotAsync(candidatePath, cancellationToken).ConfigureAwait(False)
                    If snapshot IsNot Nothing AndAlso trustedHashes.ContainsKey(snapshot.SnapshotHash) AndAlso Not trustedSnapshots.ContainsKey(snapshot.SnapshotHash) Then
                        trustedSnapshots(snapshot.SnapshotHash) = snapshot.Journal
                    End If
                Next

                Dim orderedHashes = trustedHashes.OrderByDescending(Function(item) item.Value).Select(Function(item) item.Key)
                For Each snapshotHash In orderedHashes
                    Dim trustedJournal As JobJournal = Nothing
                    If trustedSnapshots.TryGetValue(snapshotHash, trustedJournal) Then
                        Return trustedJournal
                    End If
                Next
            End If

            ' Compatibility fallback for legacy/plain snapshots.
            For Each candidatePath In candidates
                Dim snapshot = Await TryReadSnapshotAsync(candidatePath, cancellationToken).ConfigureAwait(False)
                If snapshot IsNot Nothing Then
                    Return snapshot.Journal
                End If
            Next

            Return Nothing
        End Function

        Public Async Function SaveAsync(journalPath As String, journal As JobJournal, cancellationToken As CancellationToken) As Task
            If String.IsNullOrWhiteSpace(journalPath) Then
                Throw New ArgumentException("Journal path is required.", NameOf(journalPath))
            End If

            If journal Is Nothing Then
                Throw New ArgumentNullException(NameOf(journal))
            End If

            Dim directoryPath = Path.GetDirectoryName(journalPath)
            If Not String.IsNullOrWhiteSpace(directoryPath) Then
                Directory.CreateDirectory(directoryPath)
            End If

            journal = NormalizeJournal(journal)
            journal.UpdatedUtc = DateTimeOffset.UtcNow

            Dim payload = JsonSerializer.SerializeToUtf8Bytes(journal, _serializerOptions)
            Dim snapshotHash = ComputeSha256Hex(payload)

            Dim mirrorPath = BuildMirrorJournalPath(journalPath)
            EnsureDirectoryForPath(mirrorPath)

            Await SaveSnapshotSetAsync(journalPath, payload, cancellationToken).ConfigureAwait(False)
            Await SaveSnapshotSetAsync(mirrorPath, payload, cancellationToken).ConfigureAwait(False)

            Dim securityContext = CreateSecurityContext(journalPath)
            Dim primaryAnchorPath = GetAnchorPath(journalPath)
            Dim mirrorAnchorPath = GetAnchorPath(mirrorPath)
            Dim primaryLedgerPath = GetLedgerPath(journalPath)
            Dim mirrorLedgerPath = GetLedgerPath(mirrorPath)
            Dim seedCacheKey = NormalizePath(journalPath)

            Dim seedState = TryGetCachedSeedState(seedCacheKey)
            If seedState Is Nothing Then
                seedState = Await GetLedgerSeedStateAsync(
                    primaryAnchorPath,
                    mirrorAnchorPath,
                    primaryLedgerPath,
                    mirrorLedgerPath,
                    securityContext,
                    cancellationToken).ConfigureAwait(False)
                SetCachedSeedState(seedCacheKey, seedState)
            End If

            Dim record As New JournalLedgerRecord() With {
                .Sequence = seedState.Sequence + 1,
                .UpdatedUtcTicks = DateTimeOffset.UtcNow.UtcTicks,
                .SnapshotHash = snapshotHash,
                .SnapshotLength = payload.LongLength,
                .PreviousRecordHash = seedState.LastRecordHash
            }
            record.RecordHash = ComputeRecordHash(record)
            record.Signature = ComputeRecordSignature(record, securityContext)

            Await AppendLedgerRecordAsync(primaryLedgerPath, record, cancellationToken).ConfigureAwait(False)
            Await AppendLedgerRecordAsync(mirrorLedgerPath, record, cancellationToken).ConfigureAwait(False)

            Dim anchor As New JournalAnchor() With {
                .Sequence = record.Sequence,
                .LastRecordHash = record.RecordHash,
                .UpdatedUtcTicks = record.UpdatedUtcTicks
            }
            anchor.Signature = ComputeAnchorSignature(anchor, securityContext)

            Await WriteAnchorAsync(primaryAnchorPath, anchor, cancellationToken).ConfigureAwait(False)
            Await WriteAnchorAsync(mirrorAnchorPath, anchor, cancellationToken).ConfigureAwait(False)

            Await CompactLedgerIfNeededAsync(primaryLedgerPath, securityContext, cancellationToken).ConfigureAwait(False)
            Await CompactLedgerIfNeededAsync(mirrorLedgerPath, securityContext, cancellationToken).ConfigureAwait(False)
            SetCachedSeedState(seedCacheKey, New LedgerSeedState() With {
                .Sequence = anchor.Sequence,
                .LastRecordHash = anchor.LastRecordHash
            })
        End Function

        Private Async Function LoadTrustedSnapshotHashesAsync(
            journalPath As String,
            context As JournalSecurityContext,
            cancellationToken As CancellationToken) As Task(Of Dictionary(Of String, Long))

            Dim trustedHashes As New Dictionary(Of String, Long)(StringComparer.OrdinalIgnoreCase)
            Dim mirrorPath = BuildMirrorJournalPath(journalPath)
            Dim primaryLedgerRecords = Await ReadLedgerRecordsAsync(GetLedgerPath(journalPath), context, cancellationToken).ConfigureAwait(False)
            Dim mirrorLedgerRecords = Await ReadLedgerRecordsAsync(GetLedgerPath(mirrorPath), context, cancellationToken).ConfigureAwait(False)

            For Each record In primaryLedgerRecords
                If Not String.IsNullOrWhiteSpace(record.SnapshotHash) Then
                    MergeTrustedHash(trustedHashes, record)
                End If
            Next

            For Each record In mirrorLedgerRecords
                If Not String.IsNullOrWhiteSpace(record.SnapshotHash) Then
                    MergeTrustedHash(trustedHashes, record)
                End If
            Next

            Return trustedHashes
        End Function

        Private Async Function TryReadSnapshotAsync(
            snapshotPath As String,
            cancellationToken As CancellationToken) As Task(Of SnapshotReadResult)
            If String.IsNullOrWhiteSpace(snapshotPath) OrElse Not File.Exists(snapshotPath) Then
                Return Nothing
            End If

            Try
                Dim payload = Await File.ReadAllBytesAsync(snapshotPath, cancellationToken).ConfigureAwait(False)
                If payload.Length = 0 Then
                    Return Nothing
                End If

                Dim snapshotHash = ComputeSha256Hex(payload)
                Dim journal = JsonSerializer.Deserialize(Of JobJournal)(payload, _serializerOptions)
                journal = NormalizeJournal(journal)
                If journal Is Nothing Then
                    Return Nothing
                End If

                Return New SnapshotReadResult() With {
                    .SnapshotHash = snapshotHash,
                    .Journal = journal
                }
            Catch ex As OperationCanceledException
                Throw
            Catch
                Return Nothing
            End Try
        End Function

        Private Shared Sub MergeTrustedHash(target As Dictionary(Of String, Long), record As JournalLedgerRecord)
            If target Is Nothing OrElse record Is Nothing OrElse String.IsNullOrWhiteSpace(record.SnapshotHash) Then
                Return
            End If

            Dim currentSequence As Long = 0
            If target.TryGetValue(record.SnapshotHash, currentSequence) Then
                If record.Sequence > currentSequence Then
                    target(record.SnapshotHash) = record.Sequence
                End If

                Return
            End If

            target(record.SnapshotHash) = record.Sequence
        End Sub

        Private Async Function SaveSnapshotSetAsync(
            snapshotPath As String,
            payload As Byte(),
            cancellationToken As CancellationToken) As Task

            EnsureDirectoryForPath(snapshotPath)
            RotateBackups(snapshotPath)
            Await WriteAtomicBytesAsync(snapshotPath, payload, cancellationToken).ConfigureAwait(False)
        End Function

        Private Sub RotateBackups(snapshotPath As String)
            For generation = BackupGenerationCount To 1 Step -1
                Dim destinationPath = GetBackupPath(snapshotPath, generation)
                Dim sourcePath = If(generation = 1, snapshotPath, GetBackupPath(snapshotPath, generation - 1))
                If File.Exists(destinationPath) Then
                    File.Delete(destinationPath)
                End If

                If File.Exists(sourcePath) Then
                    File.Copy(sourcePath, destinationPath, overwrite:=True)
                End If
            Next
        End Sub

        Private Shared Async Function WriteAtomicBytesAsync(path As String, payload As Byte(), cancellationToken As CancellationToken) As Task
            Dim tempPath = $"{path}.tmp.{Guid.NewGuid():N}"
            Try
                Using stream = New FileStream(tempPath, FileMode.CreateNew, FileAccess.Write, FileShare.None, 65536, FileOptions.Asynchronous Or FileOptions.WriteThrough)
                    Await stream.WriteAsync(payload.AsMemory(0, payload.Length), cancellationToken).ConfigureAwait(False)
                    Await stream.FlushAsync(cancellationToken).ConfigureAwait(False)
                    stream.Flush(flushToDisk:=True)
                End Using

                File.Move(tempPath, path, overwrite:=True)
            Catch
                If File.Exists(tempPath) Then
                    Try
                        File.Delete(tempPath)
                    Catch
                    End Try
                End If

                Throw
            End Try
        End Function

        Private Async Function GetLedgerSeedStateAsync(
            primaryAnchorPath As String,
            mirrorAnchorPath As String,
            primaryLedgerPath As String,
            mirrorLedgerPath As String,
            context As JournalSecurityContext,
            cancellationToken As CancellationToken) As Task(Of LedgerSeedState)

            Dim candidateStates As New List(Of LedgerSeedState)()

            Dim primaryAnchor = Await ReadAnchorAsync(primaryAnchorPath, context, cancellationToken).ConfigureAwait(False)
            If primaryAnchor IsNot Nothing Then
                candidateStates.Add(New LedgerSeedState() With {
                    .Sequence = primaryAnchor.Sequence,
                    .LastRecordHash = primaryAnchor.LastRecordHash
                })
            End If

            Dim mirrorAnchor = Await ReadAnchorAsync(mirrorAnchorPath, context, cancellationToken).ConfigureAwait(False)
            If mirrorAnchor IsNot Nothing Then
                candidateStates.Add(New LedgerSeedState() With {
                    .Sequence = mirrorAnchor.Sequence,
                    .LastRecordHash = mirrorAnchor.LastRecordHash
                })
            End If

            Dim primaryRecords = Await ReadLedgerRecordsAsync(primaryLedgerPath, context, cancellationToken).ConfigureAwait(False)
            If primaryRecords.Count > 0 Then
                Dim lastRecord = primaryRecords(primaryRecords.Count - 1)
                candidateStates.Add(New LedgerSeedState() With {
                    .Sequence = lastRecord.Sequence,
                    .LastRecordHash = lastRecord.RecordHash
                })
            End If

            Dim mirrorRecords = Await ReadLedgerRecordsAsync(mirrorLedgerPath, context, cancellationToken).ConfigureAwait(False)
            If mirrorRecords.Count > 0 Then
                Dim lastRecord = mirrorRecords(mirrorRecords.Count - 1)
                candidateStates.Add(New LedgerSeedState() With {
                    .Sequence = lastRecord.Sequence,
                    .LastRecordHash = lastRecord.RecordHash
                })
            End If

            If candidateStates.Count = 0 Then
                Return New LedgerSeedState() With {.Sequence = 0, .LastRecordHash = String.Empty}
            End If

            Dim bestState As LedgerSeedState = candidateStates(0)
            For i = 1 To candidateStates.Count - 1
                Dim candidate = candidateStates(i)
                If candidate.Sequence > bestState.Sequence Then
                    bestState = candidate
                End If
            Next

            Return bestState
        End Function

        Private Function TryGetCachedSeedState(seedCacheKey As String) As LedgerSeedState
            If String.IsNullOrWhiteSpace(seedCacheKey) Then
                Return Nothing
            End If

            SyncLock _seedStateSync
                Dim cached As LedgerSeedState = Nothing
                If _seedStateByPath.TryGetValue(seedCacheKey, cached) AndAlso cached IsNot Nothing Then
                    Return New LedgerSeedState() With {
                        .Sequence = cached.Sequence,
                        .LastRecordHash = cached.LastRecordHash
                    }
                End If
            End SyncLock

            Return Nothing
        End Function

        Private Sub SetCachedSeedState(seedCacheKey As String, state As LedgerSeedState)
            If String.IsNullOrWhiteSpace(seedCacheKey) OrElse state Is Nothing Then
                Return
            End If

            SyncLock _seedStateSync
                _seedStateByPath(seedCacheKey) = New LedgerSeedState() With {
                    .Sequence = state.Sequence,
                    .LastRecordHash = state.LastRecordHash
                }
            End SyncLock
        End Sub

        Private Async Function AppendLedgerRecordAsync(
            ledgerPath As String,
            record As JournalLedgerRecord,
            cancellationToken As CancellationToken) As Task

            EnsureDirectoryForPath(ledgerPath)

            Dim payload = JsonSerializer.SerializeToUtf8Bytes(record, _serializerOptions)
            Dim checksum = ComputeChecksum32(payload)
            Dim frameLength = LedgerFrameHeaderSize + payload.Length + LedgerFrameTrailerSize
            Dim frame(frameLength - 1) As Byte

            WriteUInt32(frame, 0, LedgerMagic)
            frame(4) = LedgerVersion
            WriteInt32(frame, 5, payload.Length)
            Buffer.BlockCopy(payload, 0, frame, LedgerFrameHeaderSize, payload.Length)
            WriteUInt32(frame, LedgerFrameHeaderSize + payload.Length, checksum)

            Using stream = New FileStream(ledgerPath, FileMode.Append, FileAccess.Write, FileShare.Read, 4096, FileOptions.Asynchronous Or FileOptions.WriteThrough)
                Await stream.WriteAsync(frame.AsMemory(0, frame.Length), cancellationToken).ConfigureAwait(False)
                Await stream.FlushAsync(cancellationToken).ConfigureAwait(False)
                stream.Flush(flushToDisk:=True)
            End Using
        End Function

        Private Async Function ReadLedgerRecordsAsync(
            ledgerPath As String,
            context As JournalSecurityContext,
            cancellationToken As CancellationToken) As Task(Of List(Of JournalLedgerRecord))

            Dim records As New List(Of JournalLedgerRecord)()
            If Not File.Exists(ledgerPath) Then
                Return records
            End If

            Dim lastAccepted As JournalLedgerRecord = Nothing

            Try
                Using stream = New FileStream(ledgerPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite, 4096, FileOptions.Asynchronous)
                    While stream.Position < stream.Length
                        cancellationToken.ThrowIfCancellationRequested()

                        Dim header(LedgerFrameHeaderSize - 1) As Byte
                        Dim headerRead = Await stream.ReadAsync(header.AsMemory(0, header.Length), cancellationToken).ConfigureAwait(False)
                        If headerRead = 0 Then
                            Exit While
                        End If

                        If headerRead < LedgerFrameHeaderSize Then
                            Exit While
                        End If

                        Dim magic = ReadUInt32(header, 0)
                        If magic <> LedgerMagic Then
                            Exit While
                        End If

                        Dim version = header(4)
                        If version <> LedgerVersion Then
                            Exit While
                        End If

                        Dim payloadLength = ReadInt32(header, 5)
                        If payloadLength <= 0 OrElse payloadLength > 1024 * 1024 Then
                            Exit While
                        End If

                        Dim payload(payloadLength - 1) As Byte
                        If Not Await ReadExactAsync(stream, payload, cancellationToken).ConfigureAwait(False) Then
                            Exit While
                        End If

                        Dim trailer(LedgerFrameTrailerSize - 1) As Byte
                        If Not Await ReadExactAsync(stream, trailer, cancellationToken).ConfigureAwait(False) Then
                            Exit While
                        End If

                        Dim expectedChecksum = ReadUInt32(trailer, 0)
                        Dim actualChecksum = ComputeChecksum32(payload)
                        If expectedChecksum <> actualChecksum Then
                            Exit While
                        End If

                        Dim record = JsonSerializer.Deserialize(Of JournalLedgerRecord)(payload, _serializerOptions)
                        If Not IsLedgerRecordValid(record, lastAccepted, context) Then
                            Exit While
                        End If

                        records.Add(record)
                        lastAccepted = record
                    End While
                End Using
            Catch ex As OperationCanceledException
                Throw
            Catch
                Return New List(Of JournalLedgerRecord)()
            End Try

            Return records
        End Function

        Private Async Function CompactLedgerIfNeededAsync(
            ledgerPath As String,
            context As JournalSecurityContext,
            cancellationToken As CancellationToken) As Task

            If Not File.Exists(ledgerPath) Then
                Return
            End If

            Dim info As New FileInfo(ledgerPath)
            If info.Length <= LedgerMaxBytes Then
                Return
            End If

            Dim records = Await ReadLedgerRecordsAsync(ledgerPath, context, cancellationToken).ConfigureAwait(False)
            If records.Count <= LedgerCompactRecordCount Then
                Return
            End If

            Dim retained = records.Skip(Math.Max(0, records.Count - LedgerCompactRecordCount)).ToList()
            Await RewriteLedgerAsync(ledgerPath, retained, cancellationToken).ConfigureAwait(False)
        End Function

        Private Async Function RewriteLedgerAsync(
            ledgerPath As String,
            records As List(Of JournalLedgerRecord),
            cancellationToken As CancellationToken) As Task

            EnsureDirectoryForPath(ledgerPath)
            Dim tempPath = $"{ledgerPath}.tmp.{Guid.NewGuid():N}"

            Try
                Using stream = New FileStream(tempPath, FileMode.CreateNew, FileAccess.Write, FileShare.None, 4096, FileOptions.Asynchronous Or FileOptions.WriteThrough)
                    For Each record In records
                        cancellationToken.ThrowIfCancellationRequested()

                        Dim payload = JsonSerializer.SerializeToUtf8Bytes(record, _serializerOptions)
                        Dim checksum = ComputeChecksum32(payload)
                        Dim frameLength = LedgerFrameHeaderSize + payload.Length + LedgerFrameTrailerSize
                        Dim frame(frameLength - 1) As Byte

                        WriteUInt32(frame, 0, LedgerMagic)
                        frame(4) = LedgerVersion
                        WriteInt32(frame, 5, payload.Length)
                        Buffer.BlockCopy(payload, 0, frame, LedgerFrameHeaderSize, payload.Length)
                        WriteUInt32(frame, LedgerFrameHeaderSize + payload.Length, checksum)

                        Await stream.WriteAsync(frame.AsMemory(0, frame.Length), cancellationToken).ConfigureAwait(False)
                    Next

                    Await stream.FlushAsync(cancellationToken).ConfigureAwait(False)
                    stream.Flush(flushToDisk:=True)
                End Using

                File.Move(tempPath, ledgerPath, overwrite:=True)
            Catch
                If File.Exists(tempPath) Then
                    Try
                        File.Delete(tempPath)
                    Catch
                    End Try
                End If

                Throw
            End Try
        End Function

        Private Async Function ReadAnchorAsync(
            anchorPath As String,
            context As JournalSecurityContext,
            cancellationToken As CancellationToken) As Task(Of JournalAnchor)

            If Not File.Exists(anchorPath) Then
                Return Nothing
            End If

            Try
                Dim payload = Await File.ReadAllBytesAsync(anchorPath, cancellationToken).ConfigureAwait(False)
                If payload.Length = 0 Then
                    Return Nothing
                End If

                Dim anchor = JsonSerializer.Deserialize(Of JournalAnchor)(payload, _serializerOptions)
                If Not IsAnchorValid(anchor, context) Then
                    Return Nothing
                End If

                Return anchor
            Catch ex As OperationCanceledException
                Throw
            Catch
                Return Nothing
            End Try
        End Function

        Private Async Function WriteAnchorAsync(
            anchorPath As String,
            anchor As JournalAnchor,
            cancellationToken As CancellationToken) As Task

            EnsureDirectoryForPath(anchorPath)
            Dim payload = JsonSerializer.SerializeToUtf8Bytes(anchor, _serializerOptions)
            Await WriteAtomicBytesAsync(anchorPath, payload, cancellationToken).ConfigureAwait(False)
        End Function

        Private Function IsLedgerRecordValid(
            record As JournalLedgerRecord,
            previousRecord As JournalLedgerRecord,
            context As JournalSecurityContext) As Boolean

            If record Is Nothing Then
                Return False
            End If

            If record.Sequence <= 0 Then
                Return False
            End If

            If String.IsNullOrWhiteSpace(record.SnapshotHash) OrElse record.SnapshotHash.Length <> 64 Then
                Return False
            End If

            If String.IsNullOrWhiteSpace(record.RecordHash) OrElse record.RecordHash.Length <> 64 Then
                Return False
            End If

            Dim expectedHash = ComputeRecordHash(record)
            If Not SecureEquals(record.RecordHash, expectedHash) Then
                Return False
            End If

            Dim expectedSignature = ComputeRecordSignature(record, context)
            If Not SecureEquals(record.Signature, expectedSignature) Then
                Return False
            End If

            If previousRecord IsNot Nothing Then
                If record.Sequence <= previousRecord.Sequence Then
                    Return False
                End If

                If Not SecureEquals(record.PreviousRecordHash, previousRecord.RecordHash) Then
                    Return False
                End If
            End If

            Return True
        End Function

        Private Function IsAnchorValid(anchor As JournalAnchor, context As JournalSecurityContext) As Boolean
            If anchor Is Nothing Then
                Return False
            End If

            If anchor.Sequence < 0 Then
                Return False
            End If

            If anchor.Sequence > 0 AndAlso (String.IsNullOrWhiteSpace(anchor.LastRecordHash) OrElse anchor.LastRecordHash.Length <> 64) Then
                Return False
            End If

            Dim expectedSignature = ComputeAnchorSignature(anchor, context)
            Return SecureEquals(anchor.Signature, expectedSignature)
        End Function

        Private Function ComputeRecordHash(record As JournalLedgerRecord) As String
            Dim payload = $"{record.Sequence}|{record.UpdatedUtcTicks}|{record.SnapshotHash}|{record.SnapshotLength}|{record.PreviousRecordHash}"
            Return ComputeSha256Hex(Encoding.UTF8.GetBytes(payload))
        End Function

        Private Function ComputeRecordSignature(record As JournalLedgerRecord, context As JournalSecurityContext) As String
            Dim payload = $"{record.RecordHash}|{context.PathFingerprint}"
            Return ComputeHmacBase64(context.HmacKey, payload)
        End Function

        Private Function ComputeAnchorSignature(anchor As JournalAnchor, context As JournalSecurityContext) As String
            Dim payload = $"{anchor.Sequence}|{anchor.LastRecordHash}|{anchor.UpdatedUtcTicks}|{context.PathFingerprint}"
            Return ComputeHmacBase64(context.HmacKey, payload)
        End Function

        Private Function CreateSecurityContext(journalPath As String) As JournalSecurityContext
            Dim normalizedPath = NormalizePath(journalPath)
            Dim fingerprint = ComputeSha256Hex(Encoding.UTF8.GetBytes(normalizedPath))
            Return New JournalSecurityContext() With {
                .PathFingerprint = fingerprint,
                .HmacKey = GetOrCreateHmacKey()
            }
        End Function

        Private Shared Function NormalizePath(path As String) As String
            If String.IsNullOrWhiteSpace(path) Then
                Return String.Empty
            End If

            Try
                Return IO.Path.GetFullPath(path).Trim().ToUpperInvariant()
            Catch
                Return path.Trim().ToUpperInvariant()
            End Try
        End Function

        Private Shared Function BuildSnapshotCandidates(journalPath As String) As List(Of String)
            Dim candidates As New List(Of String)()
            If String.IsNullOrWhiteSpace(journalPath) Then
                Return candidates
            End If

            AddUniquePath(candidates, journalPath)
            For generation = 1 To BackupGenerationCount
                AddUniquePath(candidates, GetBackupPath(journalPath, generation))
            Next

            Dim mirrorPath = BuildMirrorJournalPath(journalPath)
            AddUniquePath(candidates, mirrorPath)
            For generation = 1 To BackupGenerationCount
                AddUniquePath(candidates, GetBackupPath(mirrorPath, generation))
            Next

            Return candidates
        End Function

        Private Shared Sub AddUniquePath(paths As List(Of String), candidate As String)
            If String.IsNullOrWhiteSpace(candidate) Then
                Return
            End If

            For Each existing In paths
                If StringComparer.OrdinalIgnoreCase.Equals(existing, candidate) Then
                    Return
                End If
            Next

            paths.Add(candidate)
        End Sub

        Private Shared Function BuildMirrorJournalPath(journalPath As String) As String
            Dim root = IO.Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "XactCopy",
                "journals-mirror")

            Dim normalized = NormalizePath(journalPath)
            Dim pathHash = ComputeSha256Hex(Encoding.UTF8.GetBytes(normalized))
            Dim fileName = IO.Path.GetFileName(journalPath)
            If String.IsNullOrWhiteSpace(fileName) Then
                fileName = "journal.json"
            End If

            Dim extension = IO.Path.GetExtension(fileName)
            If String.IsNullOrWhiteSpace(extension) Then
                extension = ".json"
            End If

            Dim baseName = IO.Path.GetFileNameWithoutExtension(fileName)
            If String.IsNullOrWhiteSpace(baseName) Then
                baseName = "journal"
            End If

            Directory.CreateDirectory(root)
            Return IO.Path.Combine(root, $"{baseName}-{pathHash.Substring(0, 16)}{extension}")
        End Function

        Private Shared Function GetBackupPath(snapshotPath As String, generation As Integer) As String
            Return $"{snapshotPath}.bak{generation}"
        End Function

        Private Shared Function GetLedgerPath(snapshotPath As String) As String
            Return $"{snapshotPath}.ledger"
        End Function

        Private Shared Function GetAnchorPath(snapshotPath As String) As String
            Return $"{snapshotPath}.anchor"
        End Function

        Private Shared Sub EnsureDirectoryForPath(path As String)
            Dim directoryPath = IO.Path.GetDirectoryName(path)
            If Not String.IsNullOrWhiteSpace(directoryPath) Then
                Directory.CreateDirectory(directoryPath)
            End If
        End Sub

        Private Shared Function ComputeSha256Hex(bytes As Byte()) As String
            Dim hash = SHA256.HashData(bytes)
            Return Convert.ToHexString(hash).ToLowerInvariant()
        End Function

        Private Shared Function ComputeHmacBase64(key As Byte(), payload As String) As String
            Dim data = Encoding.UTF8.GetBytes(payload)
            Using hmac As New HMACSHA256(key)
                Return Convert.ToBase64String(hmac.ComputeHash(data))
            End Using
        End Function

        Private Shared Function GetOrCreateHmacKey() As Byte()
            SyncLock HmacKeySync
                If CachedHmacKey IsNot Nothing AndAlso CachedHmacKey.Length = HmacKeySizeBytes Then
                    Return DirectCast(CachedHmacKey.Clone(), Byte())
                End If

                Dim securityRoot = IO.Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "XactCopy",
                    "security")
                Directory.CreateDirectory(securityRoot)

                Dim keyPath = IO.Path.Combine(securityRoot, SecurityKeyFileName)
                Dim loadedKey As Byte() = Nothing
                If File.Exists(keyPath) Then
                    Try
                        loadedKey = File.ReadAllBytes(keyPath)
                    Catch
                        loadedKey = Nothing
                    End Try
                End If

                If loadedKey Is Nothing OrElse loadedKey.Length <> HmacKeySizeBytes Then
                    loadedKey = New Byte(HmacKeySizeBytes - 1) {}
                    RandomNumberGenerator.Fill(loadedKey)

                    Dim tempPath = $"{keyPath}.tmp.{Guid.NewGuid():N}"
                    File.WriteAllBytes(tempPath, loadedKey)
                    Try
                        If File.Exists(keyPath) Then
                            File.Delete(keyPath)
                        End If

                        File.Move(tempPath, keyPath)
                    Catch
                        If File.Exists(tempPath) Then
                            Try
                                File.Delete(tempPath)
                            Catch
                            End Try
                        End If

                        If File.Exists(keyPath) Then
                            loadedKey = File.ReadAllBytes(keyPath)
                        End If
                    End Try

                    Try
                        File.SetAttributes(keyPath, File.GetAttributes(keyPath) Or FileAttributes.Hidden)
                    Catch
                    End Try
                End If

                CachedHmacKey = loadedKey
                Return DirectCast(CachedHmacKey.Clone(), Byte())
            End SyncLock
        End Function

        Private Shared Async Function ReadExactAsync(stream As Stream, buffer As Byte(), cancellationToken As CancellationToken) As Task(Of Boolean)
            Dim offset = 0
            Dim remaining = buffer.Length
            While remaining > 0
                Dim bytesRead = Await stream.ReadAsync(buffer.AsMemory(offset, remaining), cancellationToken).ConfigureAwait(False)
                If bytesRead <= 0 Then
                    Return False
                End If

                offset += bytesRead
                remaining -= bytesRead
            End While

            Return True
        End Function

        Private Shared Function ComputeChecksum32(payload As Byte()) As UInteger
            Dim hash As UInteger = &H811C9DC5UI
            For Each value In payload
                Dim mixed = hash Xor value
                hash = CUInt((CULng(mixed) * &H1000193UL) And &HFFFFFFFFUL)
            Next

            Return hash
        End Function

        Private Shared Sub WriteUInt32(buffer As Byte(), offset As Integer, value As UInteger)
            buffer(offset) = CByte(value And &HFFUI)
            buffer(offset + 1) = CByte((value >> 8) And &HFFUI)
            buffer(offset + 2) = CByte((value >> 16) And &HFFUI)
            buffer(offset + 3) = CByte((value >> 24) And &HFFUI)
        End Sub

        Private Shared Sub WriteInt32(buffer As Byte(), offset As Integer, value As Integer)
            WriteUInt32(buffer, offset, CUInt(value))
        End Sub

        Private Shared Function ReadUInt32(buffer As Byte(), offset As Integer) As UInteger
            Return CUInt(buffer(offset)) Or
                (CUInt(buffer(offset + 1)) << 8) Or
                (CUInt(buffer(offset + 2)) << 16) Or
                (CUInt(buffer(offset + 3)) << 24)
        End Function

        Private Shared Function ReadInt32(buffer As Byte(), offset As Integer) As Integer
            Return CInt(ReadUInt32(buffer, offset))
        End Function

        Private Shared Function SecureEquals(left As String, right As String) As Boolean
            If left Is Nothing OrElse right Is Nothing Then
                Return False
            End If

            Dim leftBytes = Encoding.UTF8.GetBytes(left)
            Dim rightBytes = Encoding.UTF8.GetBytes(right)
            Return CryptographicOperations.FixedTimeEquals(leftBytes, rightBytes)
        End Function

        Private Shared Function NormalizeJournal(journal As JobJournal) As JobJournal
            If journal Is Nothing Then
                Return Nothing
            End If

            If String.IsNullOrWhiteSpace(journal.JobId) Then
                journal.JobId = Guid.NewGuid().ToString("N")
            End If

            If journal.CreatedUtc = DateTimeOffset.MinValue Then
                journal.CreatedUtc = DateTimeOffset.UtcNow
            End If

            If journal.UpdatedUtc = DateTimeOffset.MinValue Then
                journal.UpdatedUtc = journal.CreatedUtc
            End If

            Dim normalizedFiles As New Dictionary(Of String, JournalFileEntry)(StringComparer.OrdinalIgnoreCase)
            If journal.Files IsNot Nothing Then
                For Each pair In journal.Files
                    If String.IsNullOrWhiteSpace(pair.Key) Then
                        Continue For
                    End If

                    Dim entry = If(pair.Value, New JournalFileEntry() With {.RelativePath = pair.Key})
                    If String.IsNullOrWhiteSpace(entry.RelativePath) Then
                        entry.RelativePath = pair.Key
                    End If

                    normalizedFiles(pair.Key) = entry
                Next
            End If

            journal.Files = normalizedFiles
            Return journal
        End Function

        Private NotInheritable Class JournalSecurityContext
            Public Property PathFingerprint As String = String.Empty
            Public Property HmacKey As Byte() = Array.Empty(Of Byte)()
        End Class

        Private NotInheritable Class LedgerSeedState
            Public Property Sequence As Long
            Public Property LastRecordHash As String = String.Empty
        End Class

        Private NotInheritable Class JournalLedgerRecord
            Public Property Sequence As Long
            Public Property UpdatedUtcTicks As Long
            Public Property SnapshotHash As String = String.Empty
            Public Property SnapshotLength As Long
            Public Property PreviousRecordHash As String = String.Empty
            Public Property RecordHash As String = String.Empty
            Public Property Signature As String = String.Empty
        End Class

        Private NotInheritable Class JournalAnchor
            Public Property Sequence As Long
            Public Property LastRecordHash As String = String.Empty
            Public Property UpdatedUtcTicks As Long
            Public Property Signature As String = String.Empty
        End Class

        Private NotInheritable Class SnapshotReadResult
            Public Property SnapshotHash As String = String.Empty
            Public Property Journal As JobJournal
        End Class
    End Class
End Namespace
