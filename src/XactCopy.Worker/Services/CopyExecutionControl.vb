' -----------------------------------------------------------------------------
' File: src\XactCopy.Worker\Services\CopyExecutionControl.vb
' Purpose: Source file for XactCopy runtime behavior.
' -----------------------------------------------------------------------------

Imports System.Threading

Namespace Services
    ''' <summary>
    ''' Class CopyExecutionControl.
    ''' </summary>
    Public NotInheritable Class CopyExecutionControl
        Private _paused As Integer

        ''' <summary>
        ''' Gets or sets IsPaused.
        ''' </summary>
        Public ReadOnly Property IsPaused As Boolean
            Get
                Return Interlocked.CompareExchange(_paused, 0, 0) = 1
            End Get
        End Property

        ''' <summary>
        ''' Executes Pause.
        ''' </summary>
        Public Sub Pause()
            Interlocked.Exchange(_paused, 1)
        End Sub

        Public Sub [Resume]()
            Interlocked.Exchange(_paused, 0)
        End Sub

        ''' <summary>
        ''' Computes WaitIfPausedAsync.
        ''' </summary>
        Public Async Function WaitIfPausedAsync(cancellationToken As CancellationToken) As Task
            While IsPaused
                Await Task.Delay(100, cancellationToken).ConfigureAwait(False)
            End While
        End Function
    End Class
End Namespace
