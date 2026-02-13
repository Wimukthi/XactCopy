Namespace Models
    Public Class CopyJobResult
        Public Property Succeeded As Boolean
        Public Property Cancelled As Boolean
        Public Property TotalFiles As Integer
        Public Property CompletedFiles As Integer
        Public Property FailedFiles As Integer
        Public Property RecoveredFiles As Integer
        Public Property SkippedFiles As Integer
        Public Property TotalBytes As Long
        Public Property CopiedBytes As Long
        Public Property JournalPath As String = String.Empty
        Public Property ErrorMessage As String = String.Empty
    End Class
End Namespace
