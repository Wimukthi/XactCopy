Namespace Models
    Public Class JobJournal
        Public Property JobId As String = String.Empty
        Public Property SourceRoot As String = String.Empty
        Public Property DestinationRoot As String = String.Empty
        Public Property CreatedUtc As DateTimeOffset = DateTimeOffset.UtcNow
        Public Property UpdatedUtc As DateTimeOffset = DateTimeOffset.UtcNow
        Public Property Files As New Dictionary(Of String, JournalFileEntry)(StringComparer.OrdinalIgnoreCase)
    End Class
End Namespace
