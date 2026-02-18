Namespace Models
    ''' <summary>
    ''' Persistent source-level map of known unreadable byte regions discovered during scan/copy runs.
    ''' </summary>
    Public Class BadRangeMap
        Public Property SchemaVersion As Integer = 1
        Public Property SourceRoot As String = String.Empty
        Public Property SourceIdentity As String = String.Empty
        Public Property UpdatedUtc As DateTimeOffset = DateTimeOffset.UtcNow
        Public Property Files As New Dictionary(Of String, BadRangeMapFileEntry)(StringComparer.OrdinalIgnoreCase)
    End Class
End Namespace
