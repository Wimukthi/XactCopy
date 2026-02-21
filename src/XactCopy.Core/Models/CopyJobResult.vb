' -----------------------------------------------------------------------------
' File: src\XactCopy.Core\Models\CopyJobResult.vb
' Purpose: Source file for XactCopy runtime behavior.
' -----------------------------------------------------------------------------

Namespace Models
    ''' <summary>
    ''' Class CopyJobResult.
    ''' </summary>
    Public Class CopyJobResult
        ''' <summary>
        ''' Gets or sets Succeeded.
        ''' </summary>
        Public Property Succeeded As Boolean
        ''' <summary>
        ''' Gets or sets Cancelled.
        ''' </summary>
        Public Property Cancelled As Boolean
        ''' <summary>
        ''' Gets or sets TotalFiles.
        ''' </summary>
        Public Property TotalFiles As Integer
        ''' <summary>
        ''' Gets or sets CompletedFiles.
        ''' </summary>
        Public Property CompletedFiles As Integer
        ''' <summary>
        ''' Gets or sets FailedFiles.
        ''' </summary>
        Public Property FailedFiles As Integer
        ''' <summary>
        ''' Gets or sets RecoveredFiles.
        ''' </summary>
        Public Property RecoveredFiles As Integer
        ''' <summary>
        ''' Gets or sets SkippedFiles.
        ''' </summary>
        Public Property SkippedFiles As Integer
        ''' <summary>
        ''' Gets or sets TotalBytes.
        ''' </summary>
        Public Property TotalBytes As Long
        ''' <summary>
        ''' Gets or sets CopiedBytes.
        ''' </summary>
        Public Property CopiedBytes As Long
        ''' <summary>
        ''' Gets or sets JournalPath.
        ''' </summary>
        Public Property JournalPath As String = String.Empty
        ''' <summary>
        ''' Gets or sets ErrorMessage.
        ''' </summary>
        Public Property ErrorMessage As String = String.Empty
    End Class
End Namespace
