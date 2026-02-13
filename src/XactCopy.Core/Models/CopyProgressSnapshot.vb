Namespace Models
    Public Class CopyProgressSnapshot
        Public Property CurrentFile As String = String.Empty
        Public Property CurrentFileBytesCopied As Long
        Public Property CurrentFileBytesTotal As Long
        Public Property TotalBytesCopied As Long
        Public Property TotalBytes As Long
        Public Property LastChunkBytesTransferred As Integer
        Public Property BufferSizeBytes As Integer
        Public Property CompletedFiles As Integer
        Public Property FailedFiles As Integer
        Public Property RecoveredFiles As Integer
        Public Property SkippedFiles As Integer
        Public Property TotalFiles As Integer

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

        Public ReadOnly Property CurrentFileProgress As Double
            Get
                If CurrentFileBytesTotal <= 0 Then
                    Return 0.0R
                End If

                Return Math.Clamp(CDbl(CurrentFileBytesCopied) / CDbl(CurrentFileBytesTotal), 0.0R, 1.0R)
            End Get
        End Property

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
