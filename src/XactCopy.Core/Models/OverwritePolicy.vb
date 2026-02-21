' -----------------------------------------------------------------------------
' File: src\XactCopy.Core\Models\OverwritePolicy.vb
' Purpose: Source file for XactCopy runtime behavior.
' -----------------------------------------------------------------------------

Namespace Models
    ''' <summary>
    ''' Enum OverwritePolicy.
    ''' </summary>
    Public Enum OverwritePolicy
        Overwrite = 0
        SkipExisting = 1
        OverwriteIfSourceNewer = 2
        Ask = 3
    End Enum
End Namespace
