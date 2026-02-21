' -----------------------------------------------------------------------------
' File: src\XactCopy.Core\Models\SourceMutationPolicy.vb
' Purpose: Source file for XactCopy runtime behavior.
' -----------------------------------------------------------------------------

Namespace Models
    ''' <summary>
    ''' Enum SourceMutationPolicy.
    ''' </summary>
    Public Enum SourceMutationPolicy
        FailFile = 0
        SkipFile = 1
        WaitForReappearance = 2
    End Enum
End Namespace
