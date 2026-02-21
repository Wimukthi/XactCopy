' -----------------------------------------------------------------------------
' File: src\XactCopy.Core\Ipc\JsonMessagePipe.vb
' Purpose: Source file for XactCopy runtime behavior.
' -----------------------------------------------------------------------------

Imports System.IO
Imports System.Text
Imports System.Threading

Namespace Ipc
    ''' <summary>
    ''' Class JsonMessagePipe.
    ''' </summary>
    Public NotInheritable Class JsonMessagePipe
        Private Sub New()
        End Sub

        ''' <summary>
        ''' Computes WriteMessageAsync.
        ''' </summary>
        Public Shared Async Function WriteMessageAsync(
            stream As Stream,
            message As String,
            cancellationToken As CancellationToken) As Task

            Dim messageBytes = Encoding.UTF8.GetBytes(message)
            Dim lengthBytes = BitConverter.GetBytes(messageBytes.Length)

            Await stream.WriteAsync(lengthBytes.AsMemory(0, lengthBytes.Length), cancellationToken).ConfigureAwait(False)
            Await stream.WriteAsync(messageBytes.AsMemory(0, messageBytes.Length), cancellationToken).ConfigureAwait(False)
            Await stream.FlushAsync(cancellationToken).ConfigureAwait(False)
        End Function

        ''' <summary>
        ''' Computes ReadMessageAsync.
        ''' </summary>
        Public Shared Async Function ReadMessageAsync(
            stream As Stream,
            cancellationToken As CancellationToken) As Task(Of String)

            Dim lengthBytes(3) As Byte
            Dim lengthRead = Await ReadExactAsync(stream, lengthBytes, cancellationToken).ConfigureAwait(False)
            If lengthRead = 0 Then
                Return Nothing
            End If

            If lengthRead <> lengthBytes.Length Then
                Throw New EndOfStreamException("IPC stream ended while reading message length.")
            End If

            Dim payloadLength = BitConverter.ToInt32(lengthBytes, 0)
            If payloadLength < 0 Then
                Throw New InvalidDataException($"IPC payload length is invalid: {payloadLength}.")
            End If

            If payloadLength = 0 Then
                Return String.Empty
            End If

            Dim payload(payloadLength - 1) As Byte
            Dim payloadRead = Await ReadExactAsync(stream, payload, cancellationToken).ConfigureAwait(False)
            If payloadRead <> payload.Length Then
                Throw New EndOfStreamException("IPC stream ended while reading message payload.")
            End If

            Return Encoding.UTF8.GetString(payload)
        End Function

        Private Shared Async Function ReadExactAsync(
            stream As Stream,
            buffer As Byte(),
            cancellationToken As CancellationToken) As Task(Of Integer)

            Dim offset = 0
            While offset < buffer.Length
                Dim readNow = Await stream.ReadAsync(buffer.AsMemory(offset, buffer.Length - offset), cancellationToken).ConfigureAwait(False)
                If readNow = 0 Then
                    Return offset
                End If

                offset += readNow
            End While

            Return offset
        End Function
    End Class
End Namespace
