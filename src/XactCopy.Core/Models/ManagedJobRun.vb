' -----------------------------------------------------------------------------
' File: src\XactCopy.Core\Models\ManagedJobRun.vb
' Purpose: Source file for XactCopy runtime behavior.
' -----------------------------------------------------------------------------

Namespace Models
    ''' <summary>
    ''' Class ManagedJobRun.
    ''' </summary>
    Public Class ManagedJobRun
        ''' <summary>
        ''' Gets or sets RunId.
        ''' </summary>
        Public Property RunId As String = Guid.NewGuid().ToString("N")
        ''' <summary>
        ''' Gets or sets JobId.
        ''' </summary>
        Public Property JobId As String = String.Empty
        ''' <summary>
        ''' Gets or sets DisplayName.
        ''' </summary>
        Public Property DisplayName As String = String.Empty
        ''' <summary>
        ''' Gets or sets SourceRoot.
        ''' </summary>
        Public Property SourceRoot As String = String.Empty
        ''' <summary>
        ''' Gets or sets DestinationRoot.
        ''' </summary>
        Public Property DestinationRoot As String = String.Empty
        ''' <summary>
        ''' Gets or sets Trigger.
        ''' </summary>
        Public Property Trigger As String = "manual"
        ''' <summary>
        ''' Gets or sets QueueEntryId.
        ''' </summary>
        Public Property QueueEntryId As String = String.Empty
        ''' <summary>
        ''' Gets or sets QueueAttempt.
        ''' </summary>
        Public Property QueueAttempt As Integer = 0
        ''' <summary>
        ''' Gets or sets Status.
        ''' </summary>
        Public Property Status As ManagedJobRunStatus = ManagedJobRunStatus.Queued
        ''' <summary>
        ''' Gets or sets StartedUtc.
        ''' </summary>
        Public Property StartedUtc As DateTimeOffset = DateTimeOffset.UtcNow
        ''' <summary>
        ''' Gets or sets LastUpdatedUtc.
        ''' </summary>
        Public Property LastUpdatedUtc As DateTimeOffset = DateTimeOffset.UtcNow
        ''' <summary>
        ''' Gets or sets FinishedUtc.
        ''' </summary>
        Public Property FinishedUtc As DateTimeOffset?
        ''' <summary>
        ''' Gets or sets JournalPath.
        ''' </summary>
        Public Property JournalPath As String = String.Empty
        ''' <summary>
        ''' Gets or sets Summary.
        ''' </summary>
        Public Property Summary As String = String.Empty
        ''' <summary>
        ''' Gets or sets ErrorMessage.
        ''' </summary>
        Public Property ErrorMessage As String = String.Empty
        ''' <summary>
        ''' Gets or sets Result.
        ''' </summary>
        Public Property Result As CopyJobResult
    End Class
End Namespace
