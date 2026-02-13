Imports System.IO
Imports System.Linq
Imports XactCopy.Models

Namespace Services
    Public NotInheritable Class SourceScanResult
        Public Property Files As New List(Of SourceFileDescriptor)()
        Public Property Directories As New List(Of String)()
    End Class

    Public NotInheritable Class DirectoryScanner
        Private Sub New()
        End Sub

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

                Try
                    For Each subDirectory In Directory.EnumerateDirectories(currentDirectory)
                        Dim relativeSubDirectory = NormalizeRelativePath(Path.GetRelativePath(normalizedRoot, subDirectory))
                        If Not filter.ShouldTraverseDirectory(relativeSubDirectory) Then
                            Continue For
                        End If

                        Dim directoryIsSymlink = IsReparsePoint(subDirectory)
                        If directoryIsSymlink AndAlso symlinkHandling = SymlinkHandlingMode.Skip Then
                            log?.Invoke($"Skipping symbolic-link directory: {subDirectory}")
                            Continue For
                        End If

                        Dim identity = GetDirectoryIdentity(subDirectory, symlinkHandling, log)
                        If visitedDirectoryIdentities.Contains(identity) Then
                            log?.Invoke($"Skipping already visited directory target: {subDirectory}")
                            Continue For
                        End If

                        visitedDirectoryIdentities.Add(identity)
                        pending.Push(subDirectory)
                    Next
                Catch ex As Exception
                    log?.Invoke($"Directory skipped: {currentDirectory} ({ex.Message})")
                End Try

                Try
                    For Each filePath In Directory.EnumerateFiles(currentDirectory)
                        Dim relativePath = NormalizeRelativePath(Path.GetRelativePath(normalizedRoot, filePath))
                        If Not filter.ShouldIncludeFile(relativePath) Then
                            Continue For
                        End If

                        If IsReparsePoint(filePath) AndAlso symlinkHandling = SymlinkHandlingMode.Skip Then
                            log?.Invoke($"Skipping symbolic-link file: {filePath}")
                            Continue For
                        End If

                        Try
                            Dim fileInfo As New FileInfo(filePath)
                            discoveredFiles.Add(New SourceFileDescriptor() With {
                                .RelativePath = relativePath,
                                .FullPath = filePath,
                                .Length = fileInfo.Length,
                                .LastWriteTimeUtc = fileInfo.LastWriteTimeUtc
                            })
                        Catch ex As Exception
                            log?.Invoke($"File metadata skipped: {filePath} ({ex.Message})")
                        End Try
                    Next
                Catch ex As Exception
                    log?.Invoke($"Files skipped in: {currentDirectory} ({ex.Message})")
                End Try
            End While

            discoveredFiles.Sort(Function(left, right) StringComparer.OrdinalIgnoreCase.Compare(left.RelativePath, right.RelativePath))

            Dim directories = discoveredDirectories.ToList()
            directories.Sort(StringComparer.OrdinalIgnoreCase)

            Return New SourceScanResult() With {
                .Files = discoveredFiles,
                .Directories = directories
            }
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

        Private NotInheritable Class SelectionFilter
            Private ReadOnly _entries As List(Of String)

            Private Sub New(entries As List(Of String))
                _entries = entries
            End Sub

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

            Public Function ShouldIncludeDirectory(relativeDirectory As String) As Boolean
                Return ShouldTraverseDirectory(relativeDirectory)
            End Function

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
    End Class
End Namespace
