' -----------------------------------------------------------------------------
' File: src\XactCopy.Core\Models\JobJournal.vb
' Purpose: Source file for XactCopy runtime behavior.
' -----------------------------------------------------------------------------

Namespace Models
    ''' <summary>
    ''' Class JobJournal.
    ''' </summary>
    Public Class JobJournal
        ''' <summary>
        ''' Gets or sets JobId.
        ''' </summary>
        Public Property JobId As String = String.Empty
        ''' <summary>
        ''' Gets or sets SourceRoot.
        ''' </summary>
        Public Property SourceRoot As String = String.Empty
        ''' <summary>
        ''' Gets or sets DestinationRoot.
        ''' </summary>
        Public Property DestinationRoot As String = String.Empty
        ''' <summary>
        ''' Gets or sets CreatedUtc.
        ''' </summary>
        Public Property CreatedUtc As DateTimeOffset = DateTimeOffset.UtcNow
        ''' <summary>
        ''' Gets or sets UpdatedUtc.
        ''' </summary>
        Public Property UpdatedUtc As DateTimeOffset = DateTimeOffset.UtcNow
        ''' <summary>
        ''' Gets or sets Files.
        ''' </summary>
        Public Property Files As New Dictionary(Of String, JournalFileEntry)(StringComparer.OrdinalIgnoreCase)
    End Class
End Namespace
