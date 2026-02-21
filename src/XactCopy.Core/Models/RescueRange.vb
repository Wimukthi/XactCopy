' -----------------------------------------------------------------------------
' File: src\XactCopy.Core\Models\RescueRange.vb
' Purpose: Source file for XactCopy runtime behavior.
' -----------------------------------------------------------------------------

Namespace Models
    ''' <summary>
    ''' Class RescueRange.
    ''' </summary>
    Public Class RescueRange
        ''' <summary>
        ''' Gets or sets Offset.
        ''' </summary>
        Public Property Offset As Long
        ''' <summary>
        ''' Gets or sets Length.
        ''' </summary>
        Public Property Length As Integer
        ''' <summary>
        ''' Gets or sets State.
        ''' </summary>
        Public Property State As RescueRangeState = RescueRangeState.Pending
    End Class
End Namespace
