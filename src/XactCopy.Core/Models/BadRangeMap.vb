Namespace Models
    ''' <summary>
    ''' Persistent source-level map of known unreadable byte regions discovered during scan/copy runs.
    ''' </summary>
    Public Class BadRangeMap
        ''' <summary>
        ''' Gets or sets SchemaVersion.
        ''' </summary>
        Public Property SchemaVersion As Integer = 1
        ''' <summary>
        ''' Gets or sets SourceRoot.
        ''' </summary>
        Public Property SourceRoot As String = String.Empty
        ''' <summary>
        ''' Gets or sets SourceIdentity.
        ''' </summary>
        Public Property SourceIdentity As String = String.Empty
        ''' <summary>
        ''' Gets or sets UpdatedUtc.
        ''' </summary>
        Public Property UpdatedUtc As DateTimeOffset = DateTimeOffset.UtcNow
        ''' <summary>
        ''' Gets or sets Files.
        ''' </summary>
        Public Property Files As New Dictionary(Of String, BadRangeMapFileEntry)(StringComparer.OrdinalIgnoreCase)
    End Class
End Namespace
