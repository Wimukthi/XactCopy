Imports System.IO
Imports System.Security.Cryptography
Imports System.Text
Imports System.Text.Json
Imports System.Text.Json.Serialization
Imports System.Threading
Imports XactCopy.Models

Namespace Infrastructure
    Public Class JobJournalStore
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
            If Not File.Exists(journalPath) Then
                Return Nothing
            End If

            Using stream = New FileStream(journalPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite, 4096, FileOptions.Asynchronous)
                Dim journal = Await JsonSerializer.DeserializeAsync(Of JobJournal)(stream, _serializerOptions, cancellationToken).ConfigureAwait(False)
                Return journal
            End Using
        End Function

        Public Async Function SaveAsync(journalPath As String, journal As JobJournal, cancellationToken As CancellationToken) As Task
            Dim directoryPath = Path.GetDirectoryName(journalPath)
            If Not String.IsNullOrWhiteSpace(directoryPath) Then
                Directory.CreateDirectory(directoryPath)
            End If

            journal.UpdatedUtc = DateTimeOffset.UtcNow
            Dim tempPath = $"{journalPath}.tmp"

            Using stream = New FileStream(tempPath, FileMode.Create, FileAccess.Write, FileShare.None, 4096, FileOptions.Asynchronous)
                Await JsonSerializer.SerializeAsync(stream, journal, _serializerOptions, cancellationToken).ConfigureAwait(False)
                Await stream.FlushAsync(cancellationToken).ConfigureAwait(False)
            End Using

            File.Move(tempPath, journalPath, True)
        End Function
    End Class
End Namespace
