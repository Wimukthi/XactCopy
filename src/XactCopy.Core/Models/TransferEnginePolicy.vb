' -----------------------------------------------------------------------------
' File: src\XactCopy.Core\Models\TransferEnginePolicy.vb
' Purpose: Source file for XactCopy runtime behavior.
' -----------------------------------------------------------------------------

Namespace Models
    ''' <summary>
    ''' Selects the preferred transfer engine for copy runs.
    ''' </summary>
    Public Enum TransferEnginePolicy
        Auto = 0
        ManagedRescue = 1
        NativeFast = 2
    End Enum
End Namespace
