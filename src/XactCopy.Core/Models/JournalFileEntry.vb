' -----------------------------------------------------------------------------
' File: src\XactCopy.Core\Models\JournalFileEntry.vb
' Purpose: Source file for XactCopy runtime behavior.
' -----------------------------------------------------------------------------

Namespace Models
    ''' <summary>
    ''' Class JournalFileEntry.
    ''' </summary>
    Public Class JournalFileEntry
        ''' <summary>
        ''' Gets or sets RelativePath.
        ''' </summary>
        Public Property RelativePath As String = String.Empty
        ''' <summary>
        ''' Gets or sets SourceLength.
        ''' </summary>
        Public Property SourceLength As Long
        ''' <summary>
        ''' Gets or sets SourceLastWriteUtcTicks.
        ''' </summary>
        Public Property SourceLastWriteUtcTicks As Long
        ''' <summary>
        ''' Gets or sets BytesCopied.
        ''' </summary>
        Public Property BytesCopied As Long
        ''' <summary>
        ''' Gets or sets State.
        ''' </summary>
        Public Property State As FileCopyState = FileCopyState.Pending
        ''' <summary>
        ''' Gets or sets LastError.
        ''' </summary>
        Public Property LastError As String = String.Empty
        ''' <summary>
        ''' Gets or sets DoNotRetry.
        ''' </summary>
        Public Property DoNotRetry As Boolean
        ''' <summary>
        ''' Gets or sets RecoveredRanges.
        ''' </summary>
        Public Property RecoveredRanges As New List(Of ByteRange)()
        ''' <summary>
        ''' Gets or sets RescueRanges.
        ''' </summary>
        Public Property RescueRanges As New List(Of RescueRange)()
        ''' <summary>
        ''' Gets or sets LastRescuePass.
        ''' </summary>
        Public Property LastRescuePass As String = String.Empty
    End Class
End Namespace
