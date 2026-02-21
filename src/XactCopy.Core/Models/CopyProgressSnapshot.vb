' -----------------------------------------------------------------------------
' File: src\XactCopy.Core\Models\CopyProgressSnapshot.vb
' Purpose: Source file for XactCopy runtime behavior.
' -----------------------------------------------------------------------------

Namespace Models
    ''' <summary>
    ''' Class CopyProgressSnapshot.
    ''' </summary>
    Public Class CopyProgressSnapshot
        ''' <summary>
        ''' Gets or sets CurrentFile.
        ''' </summary>
        Public Property CurrentFile As String = String.Empty
        ''' <summary>
        ''' Gets or sets CurrentFileBytesCopied.
        ''' </summary>
        Public Property CurrentFileBytesCopied As Long
        ''' <summary>
        ''' Gets or sets CurrentFileBytesTotal.
        ''' </summary>
        Public Property CurrentFileBytesTotal As Long
        ''' <summary>
        ''' Gets or sets TotalBytesCopied.
        ''' </summary>
        Public Property TotalBytesCopied As Long
        ''' <summary>
        ''' Gets or sets TotalBytes.
        ''' </summary>
        Public Property TotalBytes As Long
        ''' <summary>
        ''' Gets or sets LastChunkBytesTransferred.
        ''' </summary>
        Public Property LastChunkBytesTransferred As Integer
        ''' <summary>
        ''' Gets or sets BufferSizeBytes.
        ''' </summary>
        Public Property BufferSizeBytes As Integer
        ''' <summary>
        ''' Gets or sets CompletedFiles.
        ''' </summary>
        Public Property CompletedFiles As Integer
        ''' <summary>
        ''' Gets or sets FailedFiles.
        ''' </summary>
        Public Property FailedFiles As Integer
        ''' <summary>
        ''' Gets or sets RecoveredFiles.
        ''' </summary>
        Public Property RecoveredFiles As Integer
        ''' <summary>
        ''' Gets or sets SkippedFiles.
        ''' </summary>
        Public Property SkippedFiles As Integer
        ''' <summary>
        ''' Gets or sets TotalFiles.
        ''' </summary>
        Public Property TotalFiles As Integer
        ''' <summary>
        ''' Gets or sets RescuePass.
        ''' </summary>
        Public Property RescuePass As String = String.Empty
        ''' <summary>
        ''' Gets or sets RescueBadRegionCount.
        ''' </summary>
        Public Property RescueBadRegionCount As Integer
        ''' <summary>
        ''' Gets or sets RescueRemainingBytes.
        ''' </summary>
        Public Property RescueRemainingBytes As Long

        ''' <summary>
        ''' Gets or sets OverallProgress.
        ''' </summary>
        Public ReadOnly Property OverallProgress As Double
            Get
                If TotalBytes <= 0 Then
                    If TotalFiles = 0 Then
                        Return 1.0R
                    End If

                    Return CDbl(CompletedFiles + FailedFiles + SkippedFiles) / CDbl(Math.Max(1, TotalFiles))
                End If

                Return Math.Clamp(CDbl(TotalBytesCopied) / CDbl(TotalBytes), 0.0R, 1.0R)
            End Get
        End Property

        ''' <summary>
        ''' Gets or sets CurrentFileProgress.
        ''' </summary>
        Public ReadOnly Property CurrentFileProgress As Double
            Get
                If CurrentFileBytesTotal <= 0 Then
                    Return 0.0R
                End If

                Return Math.Clamp(CDbl(CurrentFileBytesCopied) / CDbl(CurrentFileBytesTotal), 0.0R, 1.0R)
            End Get
        End Property

        ''' <summary>
        ''' Gets or sets LastChunkBufferUtilization.
        ''' </summary>
        Public ReadOnly Property LastChunkBufferUtilization As Double
            Get
                If BufferSizeBytes <= 0 OrElse LastChunkBytesTransferred <= 0 Then
                    Return 0.0R
                End If

                Return Math.Clamp(CDbl(LastChunkBytesTransferred) / CDbl(BufferSizeBytes), 0.0R, 1.0R)
            End Get
        End Property
    End Class
End Namespace
