' -----------------------------------------------------------------------------
' File: src\XactCopy.Core\Models\SourceFileDescriptor.vb
' Purpose: Source file for XactCopy runtime behavior.
' -----------------------------------------------------------------------------

Namespace Models
    ''' <summary>
    ''' Class SourceFileDescriptor.
    ''' </summary>
    Public Class SourceFileDescriptor
        ''' <summary>
        ''' Gets or sets RelativePath.
        ''' </summary>
        Public Property RelativePath As String = String.Empty
        ''' <summary>
        ''' Gets or sets FullPath.
        ''' </summary>
        Public Property FullPath As String = String.Empty
        ''' <summary>
        ''' Gets or sets Length.
        ''' </summary>
        Public Property Length As Long
        ''' <summary>
        ''' Gets or sets LastWriteTimeUtc.
        ''' </summary>
        Public Property LastWriteTimeUtc As DateTime
    End Class
End Namespace
