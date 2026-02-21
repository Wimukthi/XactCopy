' -----------------------------------------------------------------------------
' File: src\XactCopy.Worker\Services\DirectoryScanner.vb
' Purpose: Source file for XactCopy runtime behavior.
' -----------------------------------------------------------------------------

Imports System.IO
Imports System.Linq
Imports System.Runtime.InteropServices
Imports Microsoft.Win32.SafeHandles
Imports XactCopy.Models

Namespace Services
    ''' <summary>
    ''' Class SourceScanResult.
    ''' </summary>
    Public NotInheritable Class SourceScanResult
        ''' <summary>
        ''' Gets or sets Files.
        ''' </summary>
        Public Property Files As New List(Of SourceFileDescriptor)()
        ''' <summary>
        ''' Gets or sets Directories.
        ''' </summary>
        Public Property Directories As New List(Of String)()
    End Class

    ''' <summary>
    ''' Class DirectoryScanner.
    ''' </summary>
    Public NotInheritable Class DirectoryScanner
        Private Const FindExInfoBasic As Integer = 1
        Private Const FindExSearchNameMatch As Integer = 0
        Private Const FindFirstExLargeFetch As UInteger = 2UI
        Private Const ErrorFileNotFound As Integer = 2
        Private Const ErrorPathNotFound As Integer = 3
        Private Const ErrorNoMoreFiles As Integer = 18
        Private Const ErrorInvalidParameter As Integer = 87

        Private Sub New()
        End Sub

        ''' <summary>
        ''' Computes ScanSource.
        ''' </summary>
        Public Shared Function ScanSource(
            rootPath As String,
            selectedRelativePaths As IReadOnlyCollection(Of String),
            symlinkHandling As SymlinkHandlingMode,
            copyEmptyDirectories As Boolean,
            log As Action(Of String)) As SourceScanResult

            Dim normalizedRoot = Path.GetFullPath(rootPath)
            Dim discoveredFiles As New List(Of SourceFileDescriptor)()
            Dim discoveredDirectories As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

            Dim filter = SelectionFilter.Create(normalizedRoot, selectedRelativePaths, log)
            Dim visitedDirectoryIdentities As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
            visitedDirectoryIdentities.Add(GetDirectoryIdentity(normalizedRoot, symlinkHandling, log))

            Dim pending As New Stack(Of String)()
            pending.Push(normalizedRoot)

            While pending.Count > 0
                Dim currentDirectory = pending.Pop()
                Dim currentRelativeDirectory = NormalizeRelativePath(Path.GetRelativePath(normalizedRoot, currentDirectory))

                If copyEmptyDirectories AndAlso currentRelativeDirectory.Length > 0 AndAlso filter.ShouldIncludeDirectory(currentRelativeDirectory) Then
                    discoveredDirectories.Add(currentRelativeDirectory)
                End If

                Dim entries As List(Of DirectoryEntryInfo) = Nothing
                Dim enumerationError As Exception = Nothing
                If Not TryEnumerateDirectoryEntries(currentDirectory, entries, enumerationError) Then
                    log?.Invoke($"Directory skipped: {currentDirectory} ({enumerationError.Message})")
                    log?.Invoke($"Files skipped in: {currentDirectory} ({enumerationError.Message})")
                    Continue While
                End If

                For Each entry In entries
                    If entry Is Nothing Then
                        Continue For
                    End If

                    If entry.IsDirectory Then
                        Dim relativeSubDirectory = NormalizeRelativePath(Path.GetRelativePath(normalizedRoot, entry.FullPath))
                        If Not filter.ShouldTraverseDirectory(relativeSubDirectory) Then
                            Continue For
                        End If

                        If entry.IsReparsePoint AndAlso symlinkHandling = SymlinkHandlingMode.Skip Then
                            log?.Invoke($"Skipping symbolic-link directory: {entry.FullPath}")
                            Continue For
                        End If

                        Dim identity = GetDirectoryIdentity(entry.FullPath, symlinkHandling, log)
                        If visitedDirectoryIdentities.Contains(identity) Then
                            log?.Invoke($"Skipping already visited directory target: {entry.FullPath}")
                            Continue For
                        End If

                        visitedDirectoryIdentities.Add(identity)
                        pending.Push(entry.FullPath)
                        Continue For
                    End If

                    Dim relativePath = NormalizeRelativePath(Path.GetRelativePath(normalizedRoot, entry.FullPath))
                    If Not filter.ShouldIncludeFile(relativePath) Then
                        Continue For
                    End If

                    If entry.IsReparsePoint AndAlso symlinkHandling = SymlinkHandlingMode.Skip Then
                        log?.Invoke($"Skipping symbolic-link file: {entry.FullPath}")
                        Continue For
                    End If

                    discoveredFiles.Add(New SourceFileDescriptor() With {
                        .RelativePath = relativePath,
                        .FullPath = entry.FullPath,
                        .Length = entry.Length,
                        .LastWriteTimeUtc = entry.LastWriteTimeUtc
                    })
                Next
            End While

            discoveredFiles.Sort(Function(left, right) StringComparer.OrdinalIgnoreCase.Compare(left.RelativePath, right.RelativePath))

            Dim directories = discoveredDirectories.ToList()
            directories.Sort(StringComparer.OrdinalIgnoreCase)

            Return New SourceScanResult() With {
                .Files = discoveredFiles,
                .Directories = directories
            }
        End Function

        Private Shared Function TryEnumerateDirectoryEntries(
            directoryPath As String,
            ByRef entries As List(Of DirectoryEntryInfo),
            ByRef [error] As Exception) As Boolean

            entries = New List(Of DirectoryEntryInfo)()
            [error] = Nothing

            If String.IsNullOrWhiteSpace(directoryPath) Then
                Return True
            End If

            If TryEnumerateDirectoryEntriesWin32(directoryPath, entries, [error]) Then
                Return True
            End If

            entries.Clear()
            [error] = Nothing
            Try
                For Each subDirectory In Directory.EnumerateDirectories(directoryPath)
                    entries.Add(New DirectoryEntryInfo() With {
                        .FullPath = subDirectory,
                        .IsDirectory = True,
                        .IsReparsePoint = IsReparsePoint(subDirectory),
                        .Length = 0L,
                        .LastWriteTimeUtc = DateTime.MinValue
                    })
                Next

                For Each filePath In Directory.EnumerateFiles(directoryPath)
                    Dim fileInfo As New FileInfo(filePath)
                    entries.Add(New DirectoryEntryInfo() With {
                        .FullPath = filePath,
                        .IsDirectory = False,
                        .IsReparsePoint = IsReparsePoint(filePath),
                        .Length = fileInfo.Length,
                        .LastWriteTimeUtc = fileInfo.LastWriteTimeUtc
                    })
                Next

                Return True
            Catch ex As Exception
                [error] = ex
                Return False
            End Try
        End Function

        Private Shared Function TryEnumerateDirectoryEntriesWin32(
            directoryPath As String,
            entries As List(Of DirectoryEntryInfo),
            ByRef [error] As Exception) As Boolean

            [error] = Nothing
            If entries Is Nothing Then
                entries = New List(Of DirectoryEntryInfo)()
            End If

            Dim searchPath As String
            Try
                searchPath = Path.Combine(directoryPath, "*")
            Catch ex As Exception
                [error] = ex
                Return False
            End Try

            Dim findData As New Win32FindData()
            Dim handle As SafeFindHandle = FindFirstFileEx(
                searchPath,
                FindExInfoBasic,
                findData,
                FindExSearchNameMatch,
                IntPtr.Zero,
                FindFirstExLargeFetch)

            If handle Is Nothing OrElse handle.IsInvalid Then
                Dim win32Error = Marshal.GetLastWin32Error()
                If win32Error = ErrorInvalidParameter Then
                    findData = New Win32FindData()
                    handle = FindFirstFileEx(
                        searchPath,
                        FindExInfoBasic,
                        findData,
                        FindExSearchNameMatch,
                        IntPtr.Zero,
                        0UI)
                    If handle IsNot Nothing AndAlso Not handle.IsInvalid Then
                        win32Error = 0
                    Else
                        win32Error = Marshal.GetLastWin32Error()
                    End If
                End If

                If win32Error = ErrorFileNotFound OrElse
                    win32Error = ErrorPathNotFound OrElse
                    win32Error = ErrorNoMoreFiles Then
                    Return True
                End If

                [error] = New IOException($"Native enumeration failed ({win32Error}).")
                Try
                    handle?.Dispose()
                Catch
                End Try
                Return False
            End If

            Using handle
                Do
                    If Not TryAppendFindDataEntry(directoryPath, findData, entries) Then
                        Continue Do
                    End If
                Loop While FindNextFile(handle, findData)

                Dim finalError = Marshal.GetLastWin32Error()
                If finalError <> 0 AndAlso finalError <> ErrorNoMoreFiles Then
                    [error] = New IOException($"Native enumeration aborted ({finalError}).")
                    Return False
                End If
            End Using

            Return True
        End Function

        Private Shared Function TryAppendFindDataEntry(
            directoryPath As String,
            findData As Win32FindData,
            entries As List(Of DirectoryEntryInfo)) As Boolean

            Dim name = If(findData.FileName, String.Empty)
            If name.Length = 0 OrElse name = "." OrElse name = ".." Then
                Return False
            End If

            Dim attributes = CType(findData.FileAttributes, FileAttributes)
            Dim isDirectory = (attributes And FileAttributes.Directory) = FileAttributes.Directory
            Dim isReparsePoint = (attributes And FileAttributes.ReparsePoint) = FileAttributes.ReparsePoint
            Dim fullPath = Path.Combine(directoryPath, name)
            Dim fileLength = 0L
            If Not isDirectory Then
                Dim rawLength = (CULng(findData.FileSizeHigh) << 32) Or CULng(findData.FileSizeLow)
                fileLength = If(rawLength > CULng(Long.MaxValue), Long.MaxValue, CLng(rawLength))
            End If

            entries.Add(New DirectoryEntryInfo() With {
                .FullPath = fullPath,
                .IsDirectory = isDirectory,
                .IsReparsePoint = isReparsePoint,
                .Length = fileLength,
                .LastWriteTimeUtc = FileTimeToDateTimeUtc(findData.LastWriteTime)
            })
            Return True
        End Function

        Private Shared Function FileTimeToDateTimeUtc(fileTime As FileTime) As DateTime
            Dim ticks = (CULng(fileTime.HighDateTime) << 32) Or CULng(fileTime.LowDateTime)
            If ticks = 0UL Then
                Return DateTime.MinValue
            End If

            Try
                Return DateTime.FromFileTimeUtc(CLng(ticks))
            Catch
                Return DateTime.MinValue
            End Try
        End Function

        Private Shared Function IsReparsePoint(path As String) As Boolean
            Try
                Return (File.GetAttributes(path) And FileAttributes.ReparsePoint) = FileAttributes.ReparsePoint
            Catch
                Return False
            End Try
        End Function

        Private Shared Function GetDirectoryIdentity(path As String, symlinkHandling As SymlinkHandlingMode, log As Action(Of String)) As String
            Dim normalized = NormalizePathKey(path)

            If symlinkHandling <> SymlinkHandlingMode.Follow OrElse Not IsReparsePoint(path) Then
                Return normalized
            End If

            Try
                Dim info As New DirectoryInfo(path)
                Dim target = info.ResolveLinkTarget(returnFinalTarget:=True)
                If target IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(target.FullName) Then
                    Return NormalizePathKey(target.FullName)
                End If
            Catch ex As Exception
                log?.Invoke($"Unable to resolve directory symbolic-link target for {path}: {ex.Message}")
            End Try

            Return normalized
        End Function

        Private Shared Function NormalizePathKey(pathValue As String) As String
            Dim fullPath = Path.GetFullPath(pathValue)
            Return fullPath.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar)
        End Function

        Private Shared Function NormalizeRelativePath(pathValue As String) As String
            Dim text = If(pathValue, String.Empty).Trim()
            If text.Length = 0 OrElse text = "." Then
                Return String.Empty
            End If

            text = text.Replace(Path.AltDirectorySeparatorChar, Path.DirectorySeparatorChar)
            Return text.Trim(Path.DirectorySeparatorChar)
        End Function

        ''' <summary>
        ''' Class SelectionFilter.
        ''' </summary>
        Private NotInheritable Class SelectionFilter
            Private ReadOnly _entries As List(Of String)

            Private Sub New(entries As List(Of String))
                _entries = entries
            End Sub

            ''' <summary>
            ''' Computes Create.
            ''' </summary>
            Public Shared Function Create(rootPath As String, requestedPaths As IReadOnlyCollection(Of String), log As Action(Of String)) As SelectionFilter
                If requestedPaths Is Nothing OrElse requestedPaths.Count = 0 Then
                    Return New SelectionFilter(New List(Of String)())
                End If

                Dim rootFull = Path.GetFullPath(rootPath)
                Dim validEntries As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

                For Each raw In requestedPaths
                    Dim candidate = NormalizeRelativePath(raw)
                    If candidate.Length = 0 Then
                        Return New SelectionFilter(New List(Of String)())
                    End If

                    Dim relativePath = candidate
                    If Path.IsPathRooted(candidate) Then
                        Dim fullCandidate As String
                        Try
                            fullCandidate = Path.GetFullPath(candidate)
                        Catch ex As Exception
                            log?.Invoke($"Ignored invalid selected path '{candidate}' ({ex.Message}).")
                            Continue For
                        End Try

                        If StringComparer.OrdinalIgnoreCase.Equals(fullCandidate.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar), rootFull.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar)) Then
                            Return New SelectionFilter(New List(Of String)())
                        End If

                        Dim rootedPrefix = rootFull.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar) & Path.DirectorySeparatorChar
                        If Not fullCandidate.StartsWith(rootedPrefix, StringComparison.OrdinalIgnoreCase) Then
                            log?.Invoke($"Ignored selected path outside source root: {fullCandidate}")
                            Continue For
                        End If

                        relativePath = NormalizeRelativePath(Path.GetRelativePath(rootFull, fullCandidate))
                    End If

                    If relativePath.Length = 0 Then
                        Return New SelectionFilter(New List(Of String)())
                    End If

                    Dim normalizedFullPath As String
                    Try
                        normalizedFullPath = Path.GetFullPath(Path.Combine(rootFull, relativePath))
                    Catch ex As Exception
                        log?.Invoke($"Ignored invalid selected relative path '{relativePath}' ({ex.Message}).")
                        Continue For
                    End Try

                    Dim sourcePrefix = rootFull.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar) & Path.DirectorySeparatorChar
                    If Not normalizedFullPath.StartsWith(sourcePrefix, StringComparison.OrdinalIgnoreCase) Then
                        log?.Invoke($"Ignored selected relative path outside source root: {relativePath}")
                        Continue For
                    End If

                    validEntries.Add(relativePath)
                Next

                Return New SelectionFilter(validEntries.ToList())
            End Function

            ''' <summary>
            ''' Computes ShouldTraverseDirectory.
            ''' </summary>
            Public Function ShouldTraverseDirectory(relativeDirectory As String) As Boolean
                If _entries.Count = 0 Then
                    Return True
                End If

                Dim normalizedDirectory = NormalizeRelativePath(relativeDirectory)
                If normalizedDirectory.Length = 0 Then
                    Return True
                End If

                Dim directoryPrefix = normalizedDirectory & Path.DirectorySeparatorChar
                For Each entry In _entries
                    If String.Equals(entry, normalizedDirectory, StringComparison.OrdinalIgnoreCase) Then
                        Return True
                    End If

                    If entry.StartsWith(directoryPrefix, StringComparison.OrdinalIgnoreCase) Then
                        Return True
                    End If

                    Dim entryPrefix = entry & Path.DirectorySeparatorChar
                    If normalizedDirectory.StartsWith(entryPrefix, StringComparison.OrdinalIgnoreCase) Then
                        Return True
                    End If
                Next

                Return False
            End Function

            ''' <summary>
            ''' Computes ShouldIncludeDirectory.
            ''' </summary>
            Public Function ShouldIncludeDirectory(relativeDirectory As String) As Boolean
                Return ShouldTraverseDirectory(relativeDirectory)
            End Function

            ''' <summary>
            ''' Computes ShouldIncludeFile.
            ''' </summary>
            Public Function ShouldIncludeFile(relativeFile As String) As Boolean
                If _entries.Count = 0 Then
                    Return True
                End If

                Dim normalizedFile = NormalizeRelativePath(relativeFile)
                If normalizedFile.Length = 0 Then
                    Return False
                End If

                For Each entry In _entries
                    If String.Equals(entry, normalizedFile, StringComparison.OrdinalIgnoreCase) Then
                        Return True
                    End If

                    Dim entryPrefix = entry & Path.DirectorySeparatorChar
                    If normalizedFile.StartsWith(entryPrefix, StringComparison.OrdinalIgnoreCase) Then
                        Return True
                    End If
                Next

                Return False
            End Function
        End Class

        ''' <summary>
        ''' Structure Win32FindData.
        ''' </summary>
        <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Unicode)>
        Private Structure Win32FindData
            Public FileAttributes As UInteger
            Public CreationTime As FileTime
            Public LastAccessTime As FileTime
            Public LastWriteTime As FileTime
            Public FileSizeHigh As UInteger
            Public FileSizeLow As UInteger
            Public Reserved0 As UInteger
            Public Reserved1 As UInteger
            <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=260)>
            Public FileName As String
            <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=14)>
            Public AlternateFileName As String
        End Structure

        ''' <summary>
        ''' Structure FileTime.
        ''' </summary>
        <StructLayout(LayoutKind.Sequential)>
        Private Structure FileTime
            Public LowDateTime As UInteger
            Public HighDateTime As UInteger
        End Structure

        ''' <summary>
        ''' Class DirectoryEntryInfo.
        ''' </summary>
        Private NotInheritable Class DirectoryEntryInfo
            ''' <summary>
            ''' Gets or sets FullPath.
            ''' </summary>
            Public Property FullPath As String = String.Empty
            ''' <summary>
            ''' Gets or sets IsDirectory.
            ''' </summary>
            Public Property IsDirectory As Boolean
            ''' <summary>
            ''' Gets or sets IsReparsePoint.
            ''' </summary>
            Public Property IsReparsePoint As Boolean
            ''' <summary>
            ''' Gets or sets Length.
            ''' </summary>
            Public Property Length As Long
            ''' <summary>
            ''' Gets or sets LastWriteTimeUtc.
            ''' </summary>
            Public Property LastWriteTimeUtc As DateTime = DateTime.MinValue
        End Class

        ''' <summary>
        ''' Class SafeFindHandle.
        ''' </summary>
        Private NotInheritable Class SafeFindHandle
            Inherits SafeHandleZeroOrMinusOneIsInvalid

            ''' <summary>
            ''' Initializes a new instance.
            ''' </summary>
            Public Sub New()
                MyBase.New(ownsHandle:=True)
            End Sub

            Protected Overrides Function ReleaseHandle() As Boolean
                Return FindClose(handle)
            End Function
        End Class

        <DllImport("kernel32.dll", CharSet:=CharSet.Unicode, SetLastError:=True)>
        Private Shared Function FindFirstFileEx(
            lpFileName As String,
            fInfoLevelId As Integer,
            ByRef lpFindFileData As Win32FindData,
            fSearchOp As Integer,
            lpSearchFilter As IntPtr,
            dwAdditionalFlags As UInteger) As SafeFindHandle
        End Function

        <DllImport("kernel32.dll", CharSet:=CharSet.Unicode, SetLastError:=True)>
        Private Shared Function FindNextFile(
            hFindFile As SafeFindHandle,
            ByRef lpFindFileData As Win32FindData) As Boolean
        End Function

        <DllImport("kernel32.dll", SetLastError:=True)>
        Private Shared Function FindClose(hFindFile As IntPtr) As Boolean
        End Function
    End Class
End Namespace
