Namespace Models
    Public Class JobCatalog
        Public Property Jobs As New List(Of ManagedJob)()
        Public Property Runs As New List(Of ManagedJobRun)()
        Public Property QueuedJobIds As New List(Of String)()
    End Class
End Namespace
