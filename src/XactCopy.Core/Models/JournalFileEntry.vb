Namespace Models
    Public Class JournalFileEntry
        Public Property RelativePath As String = String.Empty
        Public Property SourceLength As Long
        Public Property SourceLastWriteUtcTicks As Long
        Public Property BytesCopied As Long
        Public Property State As FileCopyState = FileCopyState.Pending
        Public Property LastError As String = String.Empty
        Public Property RecoveredRanges As New List(Of ByteRange)()
    End Class
End Namespace
