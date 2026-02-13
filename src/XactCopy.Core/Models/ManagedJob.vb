Namespace Models
    Public Class ManagedJob
        Public Property JobId As String = Guid.NewGuid().ToString("N")
        Public Property Name As String = String.Empty
        Public Property Options As New CopyJobOptions()
        Public Property CreatedUtc As DateTimeOffset = DateTimeOffset.UtcNow
        Public Property UpdatedUtc As DateTimeOffset = DateTimeOffset.UtcNow
    End Class
End Namespace
