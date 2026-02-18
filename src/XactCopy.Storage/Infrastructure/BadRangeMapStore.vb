Imports System.IO
Imports System.Linq
Imports System.Security.Cryptography
Imports System.Text
Imports System.Text.Json
Imports System.Text.Json.Serialization
Imports System.Threading
Imports XactCopy.Models

Namespace Infrastructure
    ''' <summary>
    ''' Durable bad-range map persistence with atomic writes, rotating backups, mirrored snapshots,
    ''' and signed envelopes to detect tampering/corruption.
    ''' </summary>
    Public Class BadRangeMapStore
        Private Const BackupGenerationCount As Integer = 2
        Private Const SecurityKeyFileName As String = "badmap-hmac.key"
        Private Const HmacKeySizeBytes As Integer = 32
        Private Const CurrentEnvelopeSchemaVersion As Integer = 1

        Private Shared ReadOnly HmacKeySync As New Object()
        Private Shared CachedHmacKey As Byte()

        Private ReadOnly _serializerOptions As JsonSerializerOptions

        Public Sub New()
            _serializerOptions = New JsonSerializerOptions() With {
                .WriteIndented = True
            }
            _serializerOptions.Converters.Add(New JsonStringEnumConverter())
        End Sub

        Public Shared Function GetDefaultMapPath(sourceRoot As String) As String
            Dim normalizedSource = NormalizePath(sourceRoot)
            Dim hash = ComputeSha256Hex(Encoding.UTF8.GetBytes(normalizedSource))
            Dim root = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "XactCopy",
                "badmaps")
            Directory.CreateDirectory(root)
            Return Path.Combine(root, $"map-{hash.Substring(0, 20)}.json")
        End Function

        Public Async Function LoadAsync(mapPath As String, cancellationToken As CancellationToken) As Task(Of BadRangeMap)
            If String.IsNullOrWhiteSpace(mapPath) Then
                Return Nothing
            End If

            Dim candidates = BuildSnapshotCandidates(mapPath)
            Dim context = CreateSecurityContext(mapPath)

            For Each candidate In candidates
                Dim envelopeMap = Await TryLoadEnvelopeAsync(candidate, context, cancellationToken).ConfigureAwait(False)
                If envelopeMap IsNot Nothing Then
                    Return envelopeMap
                End If
            Next

            ' Legacy compatibility: accept pre-envelope snapshots so existing installs keep working.
            For Each candidate In candidates
                Dim legacyMap = Await TryLoadLegacyAsync(candidate, cancellationToken).ConfigureAwait(False)
                If legacyMap IsNot Nothing Then
                    Return legacyMap
                End If
            Next

            Return Nothing
        End Function

        Public Async Function SaveAsync(mapPath As String, map As BadRangeMap, cancellationToken As CancellationToken) As Task
            If String.IsNullOrWhiteSpace(mapPath) Then
                Throw New ArgumentException("Map path is required.", NameOf(mapPath))
            End If

            If map Is Nothing Then
                Throw New ArgumentNullException(NameOf(map))
            End If

            map = NormalizeMap(map)
            map.UpdatedUtc = DateTimeOffset.UtcNow

            Dim context = CreateSecurityContext(mapPath)
            Dim payload = JsonSerializer.SerializeToUtf8Bytes(map, _serializerOptions)
            Dim payloadSha = ComputeSha256Hex(payload)
            Dim savedTicks = DateTimeOffset.UtcNow.UtcTicks
            Dim signature = ComputeEnvelopeSignature(payloadSha, savedTicks, context)

            Dim envelope As New BadRangeMapEnvelope() With {
                .SchemaVersion = CurrentEnvelopeSchemaVersion,
                .SavedUtcTicks = savedTicks,
                .PayloadSha256 = payloadSha,
                .Signature = signature,
                .Payload = map
            }

            Dim envelopeBytes = JsonSerializer.SerializeToUtf8Bytes(envelope, _serializerOptions)
            Dim mirrorPath = BuildMirrorMapPath(mapPath)

            ' Primary + mirror snapshots reduce single-path corruption/loss risk.
            Await SaveSnapshotSetAsync(mapPath, envelopeBytes, cancellationToken).ConfigureAwait(False)
            Await SaveSnapshotSetAsync(mirrorPath, envelopeBytes, cancellationToken).ConfigureAwait(False)
        End Function

        Private Async Function TryLoadEnvelopeAsync(
            snapshotPath As String,
            context As MapSecurityContext,
            cancellationToken As CancellationToken) As Task(Of BadRangeMap)

            If String.IsNullOrWhiteSpace(snapshotPath) OrElse Not File.Exists(snapshotPath) Then
                Return Nothing
            End If

            Try
                Dim bytes = Await File.ReadAllBytesAsync(snapshotPath, cancellationToken).ConfigureAwait(False)
                If bytes.Length = 0 Then
                    Return Nothing
                End If

                Dim envelope = JsonSerializer.Deserialize(Of BadRangeMapEnvelope)(bytes, _serializerOptions)
                If envelope Is Nothing OrElse envelope.Payload Is Nothing Then
                    Return Nothing
                End If

                If envelope.SchemaVersion <= 0 Then
                    Return Nothing
                End If

                Dim payloadBytes = JsonSerializer.SerializeToUtf8Bytes(envelope.Payload, _serializerOptions)
                Dim payloadSha = ComputeSha256Hex(payloadBytes)
                If Not SecureEquals(payloadSha, envelope.PayloadSha256) Then
                    Return Nothing
                End If

                Dim expectedSignature = ComputeEnvelopeSignature(payloadSha, envelope.SavedUtcTicks, context)
                If Not SecureEquals(expectedSignature, envelope.Signature) Then
                    Return Nothing
                End If

                Return NormalizeMap(envelope.Payload)
            Catch ex As OperationCanceledException
                Throw
            Catch
                Return Nothing
            End Try
        End Function

        Private Async Function TryLoadLegacyAsync(snapshotPath As String, cancellationToken As CancellationToken) As Task(Of BadRangeMap)
            If String.IsNullOrWhiteSpace(snapshotPath) OrElse Not File.Exists(snapshotPath) Then
                Return Nothing
            End If

            Try
                Dim bytes = Await File.ReadAllBytesAsync(snapshotPath, cancellationToken).ConfigureAwait(False)
                If bytes.Length = 0 Then
                    Return Nothing
                End If

                Dim map = JsonSerializer.Deserialize(Of BadRangeMap)(bytes, _serializerOptions)
                Return NormalizeMap(map)
            Catch ex As OperationCanceledException
                Throw
            Catch
                Return Nothing
            End Try
        End Function

        Private Async Function SaveSnapshotSetAsync(
            snapshotPath As String,
            payload As Byte(),
            cancellationToken As CancellationToken) As Task

            EnsureDirectoryForPath(snapshotPath)
            ' Keep the last N known-good states before replacing the live snapshot.
            RotateBackups(snapshotPath)
            Await WriteAtomicBytesAsync(snapshotPath, payload, cancellationToken).ConfigureAwait(False)
        End Function

        Private Shared Sub RotateBackups(snapshotPath As String)
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

                ' Single rename is the commit point; partial temp files are never treated as live state.
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

        Private Function CreateSecurityContext(mapPath As String) As MapSecurityContext
            Dim normalizedPath = NormalizePath(mapPath)
            Dim fingerprint = ComputeSha256Hex(Encoding.UTF8.GetBytes(normalizedPath))
            Return New MapSecurityContext() With {
                .PathFingerprint = fingerprint,
                .HmacKey = GetOrCreateHmacKey()
            }
        End Function

        Private Shared Function BuildSnapshotCandidates(mapPath As String) As List(Of String)
            Dim candidates As New List(Of String)()
            If String.IsNullOrWhiteSpace(mapPath) Then
                Return candidates
            End If

            AddUniquePath(candidates, mapPath)
            For generation = 1 To BackupGenerationCount
                AddUniquePath(candidates, GetBackupPath(mapPath, generation))
            Next

            Dim mirrorPath = BuildMirrorMapPath(mapPath)
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

        Private Shared Function BuildMirrorMapPath(mapPath As String) As String
            Dim root = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "XactCopy",
                "badmaps-mirror")
            Directory.CreateDirectory(root)

            Dim normalized = NormalizePath(mapPath)
            Dim pathHash = ComputeSha256Hex(Encoding.UTF8.GetBytes(normalized))
            Dim fileName = Path.GetFileName(mapPath)
            If String.IsNullOrWhiteSpace(fileName) Then
                fileName = "badmap.json"
            End If

            Dim extension = Path.GetExtension(fileName)
            If String.IsNullOrWhiteSpace(extension) Then
                extension = ".json"
            End If

            Dim baseName = Path.GetFileNameWithoutExtension(fileName)
            If String.IsNullOrWhiteSpace(baseName) Then
                baseName = "badmap"
            End If

            Return Path.Combine(root, $"{baseName}-{pathHash.Substring(0, 16)}{extension}")
        End Function

        Private Shared Function GetBackupPath(snapshotPath As String, generation As Integer) As String
            Return $"{snapshotPath}.bak{generation}"
        End Function

        Private Shared Sub EnsureDirectoryForPath(pathValue As String)
            Dim directoryPath = System.IO.Path.GetDirectoryName(pathValue)
            If Not String.IsNullOrWhiteSpace(directoryPath) Then
                Directory.CreateDirectory(directoryPath)
            End If
        End Sub

        Private Shared Function NormalizePath(path As String) As String
            If String.IsNullOrWhiteSpace(path) Then
                Return String.Empty
            End If

            Try
                Return System.IO.Path.GetFullPath(path).Trim().ToUpperInvariant()
            Catch
                Return path.Trim().ToUpperInvariant()
            End Try
        End Function

        Private Shared Function ComputeSha256Hex(bytes As Byte()) As String
            Dim hash = SHA256.HashData(bytes)
            Return Convert.ToHexString(hash).ToLowerInvariant()
        End Function

        Private Shared Function ComputeEnvelopeSignature(
            payloadSha As String,
            savedUtcTicks As Long,
            context As MapSecurityContext) As String

            Dim payload = $"{payloadSha}|{savedUtcTicks}|{context.PathFingerprint}"
            Dim data = Encoding.UTF8.GetBytes(payload)
            Using hmac As New HMACSHA256(context.HmacKey)
                Return Convert.ToBase64String(hmac.ComputeHash(data))
            End Using
        End Function

        Private Shared Function SecureEquals(left As String, right As String) As Boolean
            If left Is Nothing OrElse right Is Nothing Then
                Return False
            End If

            Dim leftBytes = Encoding.UTF8.GetBytes(left)
            Dim rightBytes = Encoding.UTF8.GetBytes(right)
            Return CryptographicOperations.FixedTimeEquals(leftBytes, rightBytes)
        End Function

        Private Shared Function GetOrCreateHmacKey() As Byte()
            SyncLock HmacKeySync
                If CachedHmacKey IsNot Nothing AndAlso CachedHmacKey.Length = HmacKeySizeBytes Then
                    Return DirectCast(CachedHmacKey.Clone(), Byte())
                End If

                Dim securityRoot = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "XactCopy",
                    "security")
                Directory.CreateDirectory(securityRoot)

                Dim keyPath = Path.Combine(securityRoot, SecurityKeyFileName)
                Dim loadedKey As Byte() = Nothing
                If File.Exists(keyPath) Then
                    Try
                        loadedKey = File.ReadAllBytes(keyPath)
                    Catch
                        loadedKey = Nothing
                    End Try
                End If

                If loadedKey Is Nothing OrElse loadedKey.Length <> HmacKeySizeBytes Then
                    ' Regenerate key if missing/invalid and persist atomically.
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

        Private Shared Function NormalizeMap(map As BadRangeMap) As BadRangeMap
            If map Is Nothing Then
                Return Nothing
            End If

            If map.SchemaVersion <= 0 Then
                map.SchemaVersion = 1
            End If

            map.SourceRoot = If(map.SourceRoot, String.Empty).Trim()
            map.SourceIdentity = If(map.SourceIdentity, String.Empty).Trim()
            If map.UpdatedUtc = DateTimeOffset.MinValue Then
                map.UpdatedUtc = DateTimeOffset.UtcNow
            End If

            Dim normalizedFiles As New Dictionary(Of String, BadRangeMapFileEntry)(StringComparer.OrdinalIgnoreCase)
            If map.Files IsNot Nothing Then
                For Each pair In map.Files
                    If String.IsNullOrWhiteSpace(pair.Key) Then
                        Continue For
                    End If

                    Dim entry = If(pair.Value, New BadRangeMapFileEntry() With {.RelativePath = pair.Key})
                    entry.RelativePath = If(entry.RelativePath, pair.Key).Trim()
                    If entry.RelativePath.Length = 0 Then
                        entry.RelativePath = pair.Key
                    End If

                    If entry.LastScanUtc = DateTimeOffset.MinValue Then
                        entry.LastScanUtc = map.UpdatedUtc
                    End If

                    If entry.BadRanges Is Nothing Then
                        entry.BadRanges = New List(Of ByteRange)()
                    End If

                    entry.BadRanges = entry.BadRanges.
                        Where(Function(range) range IsNot Nothing AndAlso range.Length > 0).
                        Select(
                            Function(range)
                                Return New ByteRange() With {
                                    .Offset = Math.Max(0L, range.Offset),
                                    .Length = Math.Max(0, range.Length)
                                }
                            End Function).
                        Where(Function(range) range.Length > 0).
                        OrderBy(Function(range) range.Offset).
                        ToList()

                    normalizedFiles(entry.RelativePath) = entry
                Next
            End If

            map.Files = normalizedFiles
            Return map
        End Function

        Private NotInheritable Class MapSecurityContext
            Public Property PathFingerprint As String = String.Empty
            Public Property HmacKey As Byte() = Array.Empty(Of Byte)()
        End Class

        Private NotInheritable Class BadRangeMapEnvelope
            Public Property SchemaVersion As Integer = CurrentEnvelopeSchemaVersion
            Public Property SavedUtcTicks As Long
            Public Property PayloadSha256 As String = String.Empty
            Public Property Signature As String = String.Empty
            Public Property Payload As BadRangeMap
        End Class
    End Class
End Namespace
