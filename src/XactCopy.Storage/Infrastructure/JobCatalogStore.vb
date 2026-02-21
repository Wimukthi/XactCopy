' -----------------------------------------------------------------------------
' File: src\XactCopy.Storage\Infrastructure\JobCatalogStore.vb
' Purpose: Source file for XactCopy runtime behavior.
' -----------------------------------------------------------------------------

Imports System.IO
Imports System.Linq
Imports System.Text.Json
Imports System.Text.Json.Serialization
Imports XactCopy.Models

Namespace Infrastructure
    ''' <summary>
    ''' Class JobCatalogStore.
    ''' </summary>
    Public Class JobCatalogStore
        Private Const CurrentSchemaVersion As Integer = 2

        Private ReadOnly _catalogPath As String
        Private ReadOnly _serializerOptions As JsonSerializerOptions

        ''' <summary>
        ''' Initializes a new instance.
        ''' </summary>
        Public Sub New(Optional catalogPath As String = Nothing)
            If String.IsNullOrWhiteSpace(catalogPath) Then
                Dim root = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "XactCopy",
                    "jobs")
                _catalogPath = Path.Combine(root, "catalog.json")
            Else
                _catalogPath = catalogPath
            End If

            _serializerOptions = New JsonSerializerOptions() With {
                .WriteIndented = True,
                .PropertyNameCaseInsensitive = True
            }
            _serializerOptions.Converters.Add(New JsonStringEnumConverter())
        End Sub

        ''' <summary>
        ''' Gets or sets CatalogPath.
        ''' </summary>
        Public ReadOnly Property CatalogPath As String
            Get
                Return _catalogPath
            End Get
        End Property

        ''' <summary>
        ''' Computes Load.
        ''' </summary>
        Public Function Load() As JobCatalog
            If Not File.Exists(_catalogPath) Then
                Return New JobCatalog()
            End If

            Try
                Dim json = File.ReadAllText(_catalogPath)
                Dim catalog = JsonSerializer.Deserialize(Of JobCatalog)(json, _serializerOptions)
                Return NormalizeCatalog(catalog)
            Catch
                Return New JobCatalog()
            End Try
        End Function

        ''' <summary>
        ''' Executes Save.
        ''' </summary>
        Public Sub Save(catalog As JobCatalog)
            If catalog Is Nothing Then
                Throw New ArgumentNullException(NameOf(catalog))
            End If

            Dim normalized = NormalizeCatalog(catalog)
            Dim directoryPath = Path.GetDirectoryName(_catalogPath)
            If Not String.IsNullOrWhiteSpace(directoryPath) Then
                Directory.CreateDirectory(directoryPath)
            End If

            Dim tempPath = _catalogPath & ".tmp"
            Dim json = JsonSerializer.Serialize(normalized, _serializerOptions)
            File.WriteAllText(tempPath, json)
            File.Move(tempPath, _catalogPath, overwrite:=True)
        End Sub

        Private Shared Function NormalizeCatalog(catalog As JobCatalog) As JobCatalog
            If catalog Is Nothing Then
                Return New JobCatalog()
            End If

            catalog.SchemaVersion = Math.Max(1, catalog.SchemaVersion)

            If catalog.Jobs Is Nothing Then
                catalog.Jobs = New List(Of ManagedJob)()
            End If

            If catalog.Runs Is Nothing Then
                catalog.Runs = New List(Of ManagedJobRun)()
            End If

            If catalog.QueueEntries Is Nothing Then
                catalog.QueueEntries = New List(Of ManagedJobQueueEntry)()
            End If

            If catalog.QueuedJobIds Is Nothing Then
                catalog.QueuedJobIds = New List(Of String)()
            End If

            For Each job In catalog.Jobs
                If job Is Nothing Then
                    Continue For
                End If

                If String.IsNullOrWhiteSpace(job.JobId) Then
                    job.JobId = Guid.NewGuid().ToString("N")
                End If

                If job.Options Is Nothing Then
                    job.Options = New CopyJobOptions()
                End If
            Next

            For Each run In catalog.Runs
                If run Is Nothing Then
                    Continue For
                End If

                If String.IsNullOrWhiteSpace(run.RunId) Then
                    run.RunId = Guid.NewGuid().ToString("N")
                End If

                If String.IsNullOrWhiteSpace(run.QueueEntryId) Then
                    run.QueueEntryId = String.Empty
                End If

                If run.QueueAttempt < 0 Then
                    run.QueueAttempt = 0
                End If

                If run.Result Is Nothing Then
                    run.Result = Nothing
                End If
            Next

            For Each queueEntry In catalog.QueueEntries
                If queueEntry Is Nothing Then
                    Continue For
                End If

                If String.IsNullOrWhiteSpace(queueEntry.QueueEntryId) Then
                    queueEntry.QueueEntryId = Guid.NewGuid().ToString("N")
                End If

                If queueEntry.EnqueuedUtc = DateTimeOffset.MinValue Then
                    queueEntry.EnqueuedUtc = DateTimeOffset.UtcNow
                End If

                If queueEntry.LastUpdatedUtc = DateTimeOffset.MinValue Then
                    queueEntry.LastUpdatedUtc = queueEntry.EnqueuedUtc
                End If

                If String.IsNullOrWhiteSpace(queueEntry.Trigger) Then
                    queueEntry.Trigger = "queued"
                End If

                If queueEntry.AttemptCount < 0 Then
                    queueEntry.AttemptCount = 0
                End If
            Next

            MigrateLegacyQueue(catalog)
            RemoveDuplicateQueueEntries(catalog)
            catalog.SchemaVersion = CurrentSchemaVersion

            Return catalog
        End Function

        Private Shared Sub MigrateLegacyQueue(catalog As JobCatalog)
            If catalog Is Nothing Then
                Return
            End If

            If catalog.QueueEntries.Count > 0 Then
                Return
            End If

            If catalog.QueuedJobIds Is Nothing OrElse catalog.QueuedJobIds.Count = 0 Then
                Return
            End If

            Dim nowUtc = DateTimeOffset.UtcNow
            For Each queuedJobId In catalog.QueuedJobIds
                If String.IsNullOrWhiteSpace(queuedJobId) Then
                    Continue For
                End If

                catalog.QueueEntries.Add(New ManagedJobQueueEntry() With {
                    .QueueEntryId = Guid.NewGuid().ToString("N"),
                    .JobId = queuedJobId.Trim(),
                    .EnqueuedUtc = nowUtc,
                    .LastUpdatedUtc = nowUtc,
                    .Trigger = "legacy-migration",
                    .EnqueuedBy = "migration"
                })
            Next
        End Sub

        Private Shared Sub RemoveDuplicateQueueEntries(catalog As JobCatalog)
            If catalog Is Nothing OrElse catalog.QueueEntries Is Nothing OrElse catalog.QueueEntries.Count = 0 Then
                Return
            End If

            catalog.QueueEntries = catalog.QueueEntries.
                Where(Function(entry) entry IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(entry.JobId)).
                GroupBy(Function(entry) entry.QueueEntryId, StringComparer.OrdinalIgnoreCase).
                Select(Function(group) group.First()).
                ToList()
        End Sub
    End Class
End Namespace
