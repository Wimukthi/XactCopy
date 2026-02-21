' -----------------------------------------------------------------------------
' File: tests\XactCopy.Tests\BadRangeMapStoreTests.vb
' Purpose: Source file for XactCopy runtime behavior.
' -----------------------------------------------------------------------------

Imports System.Collections.Generic
Imports System.IO
Imports System.Threading
Imports System.Threading.Tasks
Imports XactCopy.Infrastructure
Imports XactCopy.Models
Imports Xunit

''' <summary>
''' Class BadRangeMapStoreTests.
''' </summary>
Public Class BadRangeMapStoreTests
    ''' <summary>
    ''' Computes SaveAndLoadAsync_RoundTripsBadRangeMapData.
    ''' </summary>
    <Fact>
    Public Async Function SaveAndLoadAsync_RoundTripsBadRangeMapData() As Task
        Dim store As New BadRangeMapStore()
        Dim tempDirectory = Path.Combine(Path.GetTempPath(), $"xactcopy-tests-badmap-{Guid.NewGuid():N}")
        Directory.CreateDirectory(tempDirectory)

        Try
            Dim mapPath = Path.Combine(tempDirectory, "badmap.json")
            Dim map As New BadRangeMap() With {
                .SourceRoot = "C:\source",
                .SourceIdentity = "vol:ABCD1234"
            }
            map.Files("video.bin") = New BadRangeMapFileEntry() With {
                .RelativePath = "video.bin",
                .SourceLength = 4096,
                .LastWriteUtcTicks = 123456,
                .FileFingerprint = "fprint",
                .BadRanges = New List(Of ByteRange) From {
                    New ByteRange() With {.Offset = 512, .Length = 128},
                    New ByteRange() With {.Offset = 1024, .Length = 256}
                },
                .LastError = "Read failure"
            }

            Await store.SaveAsync(mapPath, map, CancellationToken.None)
            Dim loaded = Await store.LoadAsync(mapPath, CancellationToken.None)

            Assert.NotNull(loaded)
            Assert.Equal("C:\source", loaded.SourceRoot)
            Assert.True(loaded.Files.ContainsKey("video.bin"))
            Assert.Equal(2, loaded.Files("video.bin").BadRanges.Count)
            Assert.Equal(128, loaded.Files("video.bin").BadRanges(0).Length)
        Finally
            If Directory.Exists(tempDirectory) Then
                Directory.Delete(tempDirectory, recursive:=True)
            End If
        End Try
    End Function

    ''' <summary>
    ''' Computes LoadAsync_UsesBackupWhenPrimarySnapshotIsTampered.
    ''' </summary>
    <Fact>
    Public Async Function LoadAsync_UsesBackupWhenPrimarySnapshotIsTampered() As Task
        Dim store As New BadRangeMapStore()
        Dim tempDirectory = Path.Combine(Path.GetTempPath(), $"xactcopy-tests-badmap-{Guid.NewGuid():N}")
        Directory.CreateDirectory(tempDirectory)

        Try
            Dim mapPath = Path.Combine(tempDirectory, "badmap.json")
            Dim map As New BadRangeMap() With {
                .SourceRoot = "C:\source"
            }
            map.Files("sample.bin") = New BadRangeMapFileEntry() With {
                .RelativePath = "sample.bin",
                .SourceLength = 1024,
                .BadRanges = New List(Of ByteRange) From {
                    New ByteRange() With {.Offset = 0, .Length = 64}
                }
            }

            Await store.SaveAsync(mapPath, map, CancellationToken.None)

            map.Files("sample.bin").BadRanges = New List(Of ByteRange) From {
                New ByteRange() With {.Offset = 0, .Length = 128}
            }
            Await store.SaveAsync(mapPath, map, CancellationToken.None)

            File.WriteAllBytes(mapPath, {CByte(1), CByte(2), CByte(3), CByte(4)})

            Dim loaded = Await store.LoadAsync(mapPath, CancellationToken.None)

            Assert.NotNull(loaded)
            Assert.True(loaded.Files.ContainsKey("sample.bin"))
            Assert.Equal(64, loaded.Files("sample.bin").BadRanges(0).Length)
        Finally
            If Directory.Exists(tempDirectory) Then
                Directory.Delete(tempDirectory, recursive:=True)
            End If
        End Try
    End Function
End Class
