' -----------------------------------------------------------------------------
' File: src\XactCopy.Core\Models\RecoveryState.vb
' Purpose: Source file for XactCopy runtime behavior.
' -----------------------------------------------------------------------------

Namespace Models
    ''' <summary>
    ''' Class RecoveryState.
    ''' </summary>
    Public Class RecoveryState
        ''' <summary>
        ''' Gets or sets ProcessSessionId.
        ''' </summary>
        Public Property ProcessSessionId As String = String.Empty
        ''' <summary>
        ''' Gets or sets LastStartUtc.
        ''' </summary>
        Public Property LastStartUtc As DateTimeOffset = DateTimeOffset.MinValue
        ''' <summary>
        ''' Gets or sets LastExitUtc.
        ''' </summary>
        Public Property LastExitUtc As DateTimeOffset = DateTimeOffset.MinValue
        ''' <summary>
        ''' Gets or sets CleanShutdown.
        ''' </summary>
        Public Property CleanShutdown As Boolean = True
        ''' <summary>
        ''' Gets or sets PendingResumePrompt.
        ''' </summary>
        Public Property PendingResumePrompt As Boolean = False
        ''' <summary>
        ''' Gets or sets LastInterruptionReason.
        ''' </summary>
        Public Property LastInterruptionReason As String = String.Empty
        ''' <summary>
        ''' Gets or sets ActiveRun.
        ''' </summary>
        Public Property ActiveRun As RecoveryActiveRun
    End Class
End Namespace
