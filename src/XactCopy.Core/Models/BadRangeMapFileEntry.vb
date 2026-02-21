Namespace Models
    ''' <summary>
    ''' Per-file bad-range metadata stored inside <see cref="BadRangeMap"/>.
    ''' </summary>
    Public Class BadRangeMapFileEntry
        ''' <summary>
        ''' Gets or sets RelativePath.
        ''' </summary>
        Public Property RelativePath As String = String.Empty
        ''' <summary>
        ''' Gets or sets SourceLength.
        ''' </summary>
        Public Property SourceLength As Long
        ''' <summary>
        ''' Gets or sets LastWriteUtcTicks.
        ''' </summary>
        Public Property LastWriteUtcTicks As Long
        ''' <summary>
        ''' Gets or sets FileFingerprint.
        ''' </summary>
        Public Property FileFingerprint As String = String.Empty
        ''' <summary>
        ''' Gets or sets BadRanges.
        ''' </summary>
        Public Property BadRanges As New List(Of ByteRange)()
        ''' <summary>
        ''' Gets or sets LastScanUtc.
        ''' </summary>
        Public Property LastScanUtc As DateTimeOffset = DateTimeOffset.UtcNow
        ''' <summary>
        ''' Gets or sets LastError.
        ''' </summary>
        Public Property LastError As String = String.Empty
    End Class
End Namespace
