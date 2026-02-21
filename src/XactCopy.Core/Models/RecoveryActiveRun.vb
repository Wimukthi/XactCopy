' -----------------------------------------------------------------------------
' File: src\XactCopy.Core\Models\RecoveryActiveRun.vb
' Purpose: Source file for XactCopy runtime behavior.
' -----------------------------------------------------------------------------

Namespace Models
    ''' <summary>
    ''' Class RecoveryActiveRun.
    ''' </summary>
    Public Class RecoveryActiveRun
        ''' <summary>
        ''' Gets or sets RunId.
        ''' </summary>
        Public Property RunId As String = String.Empty
        ''' <summary>
        ''' Gets or sets JobId.
        ''' </summary>
        Public Property JobId As String = String.Empty
        ''' <summary>
        ''' Gets or sets JobName.
        ''' </summary>
        Public Property JobName As String = String.Empty
        ''' <summary>
        ''' Gets or sets Trigger.
        ''' </summary>
        Public Property Trigger As String = "manual"
        ''' <summary>
        ''' Gets or sets Options.
        ''' </summary>
        Public Property Options As New CopyJobOptions()
        ''' <summary>
        ''' Gets or sets StartedUtc.
        ''' </summary>
        Public Property StartedUtc As DateTimeOffset = DateTimeOffset.UtcNow
        ''' <summary>
        ''' Gets or sets LastHeartbeatUtc.
        ''' </summary>
        Public Property LastHeartbeatUtc As DateTimeOffset = DateTimeOffset.UtcNow
        ''' <summary>
        ''' Gets or sets LastProgressUtc.
        ''' </summary>
        Public Property LastProgressUtc As DateTimeOffset = DateTimeOffset.UtcNow
        ''' <summary>
        ''' Gets or sets JournalPath.
        ''' </summary>
        Public Property JournalPath As String = String.Empty
    End Class
End Namespace
