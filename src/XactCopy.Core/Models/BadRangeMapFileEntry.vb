Namespace Models
    Public Class BadRangeMapFileEntry
        Public Property RelativePath As String = String.Empty
        Public Property SourceLength As Long
        Public Property LastWriteUtcTicks As Long
        Public Property FileFingerprint As String = String.Empty
        Public Property BadRanges As New List(Of ByteRange)()
        Public Property LastScanUtc As DateTimeOffset = DateTimeOffset.UtcNow
        Public Property LastError As String = String.Empty
    End Class
End Namespace
