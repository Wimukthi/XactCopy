Imports System.IO
Imports System.Text.Json
Imports System.Text.Json.Serialization
Imports XactCopy.Models

Namespace Infrastructure
    Public Class JobCatalogStore
        Private ReadOnly _catalogPath As String
        Private ReadOnly _serializerOptions As JsonSerializerOptions

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

        Public ReadOnly Property CatalogPath As String
            Get
                Return _catalogPath
            End Get
        End Property

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

            If catalog.Jobs Is Nothing Then
                catalog.Jobs = New List(Of ManagedJob)()
            End If

            If catalog.Runs Is Nothing Then
                catalog.Runs = New List(Of ManagedJobRun)()
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

                If run.Result Is Nothing Then
                    run.Result = Nothing
                End If
            Next

            Return catalog
        End Function
    End Class
End Namespace
