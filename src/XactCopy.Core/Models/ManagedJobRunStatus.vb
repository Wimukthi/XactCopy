' -----------------------------------------------------------------------------
' File: src\XactCopy.Core\Models\ManagedJobRunStatus.vb
' Purpose: Source file for XactCopy runtime behavior.
' -----------------------------------------------------------------------------

Namespace Models
    ''' <summary>
    ''' Enum ManagedJobRunStatus.
    ''' </summary>
    Public Enum ManagedJobRunStatus
        Queued = 0
        Running = 1
        Paused = 2
        Completed = 3
        Failed = 4
        Cancelled = 5
        Interrupted = 6
    End Enum
End Namespace
