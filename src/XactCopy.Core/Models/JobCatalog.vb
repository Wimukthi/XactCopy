Namespace Models
    Public Class JobCatalog
        Public Property SchemaVersion As Integer = 2
        Public Property Jobs As New List(Of ManagedJob)()
        Public Property Runs As New List(Of ManagedJobRun)()
        Public Property QueueEntries As New List(Of ManagedJobQueueEntry)()

        ' Legacy queue representation used by older builds.
        Public Property QueuedJobIds As New List(Of String)()
    End Class
End Namespace
