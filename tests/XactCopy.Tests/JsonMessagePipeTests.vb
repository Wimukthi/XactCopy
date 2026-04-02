' -----------------------------------------------------------------------------
' File: tests\XactCopy.Tests\JsonMessagePipeTests.vb
' Purpose: Source file for XactCopy runtime behavior.
' -----------------------------------------------------------------------------

Imports System.IO
Imports System.Threading
Imports System.Threading.Tasks
Imports XactCopy.Ipc
Imports Xunit

''' <summary>
''' Class JsonMessagePipeTests.
''' </summary>
Public Class JsonMessagePipeTests
    ''' <summary>
    ''' Computes WriteAndReadMessage_RoundTripsContent.
    ''' </summary>
    <Fact>
    Public Async Function WriteAndReadMessage_RoundTripsContent() As Task
        Using stream As New MemoryStream()
            Await JsonMessagePipe.WriteMessageAsync(stream, "{""Type"":""Ping""}", CancellationToken.None)
            stream.Position = 0

            Dim readBack = Await JsonMessagePipe.ReadMessageAsync(stream, CancellationToken.None)
            Assert.Equal("{""Type"":""Ping""}", readBack)
        End Using
    End Function

    ''' <summary>
    ''' Computes ReadMessage_ReturnsNothingWhenStreamIsAtEnd.
    ''' </summary>
    <Fact>
    Public Async Function ReadMessage_ReturnsNothingWhenStreamIsAtEnd() As Task
        Using stream As New MemoryStream()
            Dim readBack = Await JsonMessagePipe.ReadMessageAsync(stream, CancellationToken.None)
            Assert.Null(readBack)
        End Using
    End Function

    ''' <summary>
    ''' Computes ReadMessage_ThrowsWhenPayloadLengthExceedsMaximum.
    ''' </summary>
    <Fact>
    Public Async Function ReadMessage_ThrowsWhenPayloadLengthExceedsMaximum() As Task
        Using stream As New MemoryStream()
            Dim oversizedLength = JsonMessagePipe.MaximumPayloadBytes + 1
            Dim header = BitConverter.GetBytes(oversizedLength)
            Await stream.WriteAsync(header, 0, header.Length)
            stream.Position = 0

            Await Assert.ThrowsAsync(Of InvalidDataException)(
                Function() JsonMessagePipe.ReadMessageAsync(stream, CancellationToken.None))
        End Using
    End Function

    ''' <summary>
    ''' Computes WriteMessage_ThrowsWhenPayloadLengthExceedsMaximum.
    ''' </summary>
    <Fact>
    Public Async Function WriteMessage_ThrowsWhenPayloadLengthExceedsMaximum() As Task
        Dim oversizedPayload = New String("x"c, JsonMessagePipe.MaximumPayloadBytes + 1)

        Using stream As New MemoryStream()
            Await Assert.ThrowsAsync(Of InvalidDataException)(
                Function() JsonMessagePipe.WriteMessageAsync(stream, oversizedPayload, CancellationToken.None))
        End Using
    End Function
End Class
