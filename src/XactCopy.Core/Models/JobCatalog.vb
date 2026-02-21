Namespace Models
    ''' <summary>
    ''' Class JobCatalog.
    ''' </summary>
    Public Class JobCatalog
        ''' <summary>
        ''' Gets or sets SchemaVersion.
        ''' </summary>
        Public Property SchemaVersion As Integer = 2
        ''' <summary>
        ''' Gets or sets Jobs.
        ''' </summary>
        Public Property Jobs As New List(Of ManagedJob)()
        ''' <summary>
        ''' Gets or sets Runs.
        ''' </summary>
        Public Property Runs As New List(Of ManagedJobRun)()
        ''' <summary>
        ''' Gets or sets QueueEntries.
        ''' </summary>
        Public Property QueueEntries As New List(Of ManagedJobQueueEntry)()

        ' Legacy queue representation used by older builds.
        ''' <summary>
        ''' Gets or sets QueuedJobIds.
        ''' </summary>
        Public Property QueuedJobIds As New List(Of String)()
    End Class
End Namespace
