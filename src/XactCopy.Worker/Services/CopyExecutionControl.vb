Imports System.Threading

Namespace Services
    Public NotInheritable Class CopyExecutionControl
        Private _paused As Integer

        Public ReadOnly Property IsPaused As Boolean
            Get
                Return Interlocked.CompareExchange(_paused, 0, 0) = 1
            End Get
        End Property

        Public Sub Pause()
            Interlocked.Exchange(_paused, 1)
        End Sub

        Public Sub [Resume]()
            Interlocked.Exchange(_paused, 0)
        End Sub

        Public Async Function WaitIfPausedAsync(cancellationToken As CancellationToken) As Task
            While IsPaused
                Await Task.Delay(100, cancellationToken).ConfigureAwait(False)
            End While
        End Function
    End Class
End Namespace
