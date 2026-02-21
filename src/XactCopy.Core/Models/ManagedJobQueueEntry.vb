' -----------------------------------------------------------------------------
' File: src\XactCopy.Core\Models\ManagedJobQueueEntry.vb
' Purpose: Source file for XactCopy runtime behavior.
' -----------------------------------------------------------------------------

Namespace Models
    ''' <summary>
    ''' Class ManagedJobQueueEntry.
    ''' </summary>
    Public Class ManagedJobQueueEntry
        ''' <summary>
        ''' Gets or sets QueueEntryId.
        ''' </summary>
        Public Property QueueEntryId As String = Guid.NewGuid().ToString("N")
        ''' <summary>
        ''' Gets or sets JobId.
        ''' </summary>
        Public Property JobId As String = String.Empty
        ''' <summary>
        ''' Gets or sets EnqueuedUtc.
        ''' </summary>
        Public Property EnqueuedUtc As DateTimeOffset = DateTimeOffset.UtcNow
        ''' <summary>
        ''' Gets or sets LastUpdatedUtc.
        ''' </summary>
        Public Property LastUpdatedUtc As DateTimeOffset = DateTimeOffset.UtcNow
        ''' <summary>
        ''' Gets or sets Trigger.
        ''' </summary>
        Public Property Trigger As String = "queued"
        ''' <summary>
        ''' Gets or sets EnqueuedBy.
        ''' </summary>
        Public Property EnqueuedBy As String = String.Empty
        ''' <summary>
        ''' Gets or sets AttemptCount.
        ''' </summary>
        Public Property AttemptCount As Integer = 0
        ''' <summary>
        ''' Gets or sets LastAttemptUtc.
        ''' </summary>
        Public Property LastAttemptUtc As DateTimeOffset?
        ''' <summary>
        ''' Gets or sets LastErrorMessage.
        ''' </summary>
        Public Property LastErrorMessage As String = String.Empty
    End Class
End Namespace
