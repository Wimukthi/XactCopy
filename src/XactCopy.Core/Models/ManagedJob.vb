' -----------------------------------------------------------------------------
' File: src\XactCopy.Core\Models\ManagedJob.vb
' Purpose: Source file for XactCopy runtime behavior.
' -----------------------------------------------------------------------------

Namespace Models
    ''' <summary>
    ''' Class ManagedJob.
    ''' </summary>
    Public Class ManagedJob
        ''' <summary>
        ''' Gets or sets JobId.
        ''' </summary>
        Public Property JobId As String = Guid.NewGuid().ToString("N")
        ''' <summary>
        ''' Gets or sets Name.
        ''' </summary>
        Public Property Name As String = String.Empty
        ''' <summary>
        ''' Gets or sets Options.
        ''' </summary>
        Public Property Options As New CopyJobOptions()
        ''' <summary>
        ''' Gets or sets CreatedUtc.
        ''' </summary>
        Public Property CreatedUtc As DateTimeOffset = DateTimeOffset.UtcNow
        ''' <summary>
        ''' Gets or sets UpdatedUtc.
        ''' </summary>
        Public Property UpdatedUtc As DateTimeOffset = DateTimeOffset.UtcNow
    End Class
End Namespace
