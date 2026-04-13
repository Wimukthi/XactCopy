' -----------------------------------------------------------------------------
' File: src\XactCopy.Core\Ipc\JsonMessagePipe.vb
' Purpose: Source file for XactCopy runtime behavior.
' -----------------------------------------------------------------------------

Imports System.Buffers
Imports System.Buffers.Binary
Imports System.IO
Imports System.Text
Imports System.Threading

Namespace Ipc
    ''' <summary>
    ''' Class JsonMessagePipe.
    ''' </summary>
    Public NotInheritable Class JsonMessagePipe
        ''' <summary>
        ''' Hard ceiling for a single IPC frame payload.
        ''' </summary>
        Public Const MaximumPayloadBytes As Integer = 16 * 1024 * 1024

        Private Sub New()
        End Sub

        ''' <summary>
        ''' Computes WriteMessageAsync.
        ''' </summary>
        Public Shared Async Function WriteMessageAsync(
            stream As Stream,
            message As String,
            cancellationToken As CancellationToken) As Task

            ' Compute byte count without allocating, validate, then rent a pooled buffer
            ' to avoid a fresh heap allocation per IPC frame.
            Dim byteCount = Encoding.UTF8.GetByteCount(message)
            If byteCount > MaximumPayloadBytes Then
                Throw New InvalidDataException(
                    $"IPC payload length {byteCount} exceeds max allowed size {MaximumPayloadBytes}.")
            End If

            Dim messageBuffer = ArrayPool(Of Byte).Shared.Rent(byteCount)
            Try
                Encoding.UTF8.GetBytes(message.AsSpan(), messageBuffer.AsSpan(0, byteCount))

                ' Write 4-byte little-endian length prefix then payload.
                Dim lengthBytes(3) As Byte
                BinaryPrimitives.WriteInt32LittleEndian(lengthBytes.AsSpan(), byteCount)
                Await stream.WriteAsync(lengthBytes.AsMemory(0, 4), cancellationToken).ConfigureAwait(False)
                Await stream.WriteAsync(messageBuffer.AsMemory(0, byteCount), cancellationToken).ConfigureAwait(False)
                Await stream.FlushAsync(cancellationToken).ConfigureAwait(False)
            Finally
                ArrayPool(Of Byte).Shared.Return(messageBuffer)
            End Try
        End Function

        ''' <summary>
        ''' Computes ReadMessageAsync.
        ''' </summary>
        Public Shared Async Function ReadMessageAsync(
            stream As Stream,
            cancellationToken As CancellationToken) As Task(Of String)

            Dim lengthBytes(3) As Byte
            Dim lengthRead = Await ReadExactAsync(stream, lengthBytes.AsMemory(), cancellationToken).ConfigureAwait(False)
            If lengthRead = 0 Then
                Return Nothing
            End If

            If lengthRead <> 4 Then
                Throw New EndOfStreamException("IPC stream ended while reading message length.")
            End If

            Dim payloadLength = BinaryPrimitives.ReadInt32LittleEndian(lengthBytes.AsSpan())
            If payloadLength < 0 OrElse payloadLength > MaximumPayloadBytes Then
                Throw New InvalidDataException(
                    $"IPC payload length is invalid or exceeds limit: {payloadLength} (max {MaximumPayloadBytes}).")
            End If

            If payloadLength = 0 Then
                Return String.Empty
            End If

            ' Rent a pooled buffer for the payload; return it immediately after decoding
            ' to avoid one heap allocation per received IPC frame.
            Dim payload = ArrayPool(Of Byte).Shared.Rent(payloadLength)
            Try
                Dim payloadRead = Await ReadExactAsync(stream, payload.AsMemory(0, payloadLength), cancellationToken).ConfigureAwait(False)
                If payloadRead <> payloadLength Then
                    Throw New EndOfStreamException("IPC stream ended while reading message payload.")
                End If

                Return Encoding.UTF8.GetString(payload, 0, payloadLength)
            Finally
                ArrayPool(Of Byte).Shared.Return(payload)
            End Try
        End Function

        Private Shared Async Function ReadExactAsync(
            stream As Stream,
            buffer As Memory(Of Byte),
            cancellationToken As CancellationToken) As Task(Of Integer)

            Dim offset = 0
            While offset < buffer.Length
                Dim readNow = Await stream.ReadAsync(buffer.Slice(offset), cancellationToken).ConfigureAwait(False)
                If readNow = 0 Then
                    Return offset
                End If

                offset += readNow
            End While

            Return offset
        End Function
    End Class
End Namespace
