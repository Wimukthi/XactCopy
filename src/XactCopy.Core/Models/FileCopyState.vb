' -----------------------------------------------------------------------------
' File: src\XactCopy.Core\Models\FileCopyState.vb
' Purpose: Source file for XactCopy runtime behavior.
' -----------------------------------------------------------------------------

Namespace Models
    ''' <summary>
    ''' Enum FileCopyState.
    ''' </summary>
    Public Enum FileCopyState
        Pending = 0
        InProgress = 1
        Completed = 2
        CompletedWithRecovery = 3
        Failed = 4
    End Enum
End Namespace
