Imports System.Collections.Generic
Imports System.IO

''' <summary>
''' Class LaunchOptions.
''' </summary>
Public NotInheritable Class LaunchOptions
    ''' <summary>
    ''' Gets or sets IsRecoveryAutostart.
    ''' </summary>
    Public Property IsRecoveryAutostart As Boolean
    ''' <summary>
    ''' Gets or sets ForceResumePrompt.
    ''' </summary>
    Public Property ForceResumePrompt As Boolean

    Private ReadOnly _explorerSourcePaths As New List(Of String)()

    ''' <summary>
    ''' Gets or sets ExplorerFolderPath.
    ''' </summary>
    Public Property ExplorerFolderPath As String = String.Empty

    ''' <summary>
    ''' Gets or sets ExplorerSourcePaths.
    ''' </summary>
    Public ReadOnly Property ExplorerSourcePaths As IReadOnlyList(Of String)
        Get
            Return _explorerSourcePaths
        End Get
    End Property

    ''' <summary>
    ''' Gets or sets ExplorerSourcePath.
    ''' </summary>
    Public ReadOnly Property ExplorerSourcePath As String
        Get
            If _explorerSourcePaths.Count = 0 Then
                Return String.Empty
            End If

            Return _explorerSourcePaths(0)
        End Get
    End Property

    ''' <summary>
    ''' Computes Parse.
    ''' </summary>
    Public Shared Function Parse(args As String()) As LaunchOptions
        Dim options As New LaunchOptions()
        If args Is Nothing Then
            Return options
        End If

        Dim index As Integer = 0
        While index < args.Length
            Dim argumentValue = args(index)
            If String.IsNullOrWhiteSpace(argumentValue) Then
                index += 1
                Continue While
            End If

            If String.Equals(argumentValue, "--recovery-autostart", StringComparison.OrdinalIgnoreCase) Then
                options.IsRecoveryAutostart = True
            ElseIf String.Equals(argumentValue, "--resume-interrupted", StringComparison.OrdinalIgnoreCase) Then
                options.ForceResumePrompt = True
            ElseIf String.Equals(argumentValue, "--from-explorer-folder", StringComparison.OrdinalIgnoreCase) Then
                If index + 1 < args.Length Then
                    index += 1
                    options.ExplorerFolderPath = UnwrapQuotedValue(args(index))
                End If
            ElseIf argumentValue.StartsWith("--from-explorer-folder=", StringComparison.OrdinalIgnoreCase) Then
                options.ExplorerFolderPath = UnwrapQuotedValue(argumentValue.Substring("--from-explorer-folder=".Length))
            ElseIf String.Equals(argumentValue, "--from-explorer", StringComparison.OrdinalIgnoreCase) Then
                index = ParseExplorerSelectionArguments(args, index, options)
            ElseIf argumentValue.StartsWith("--from-explorer=", StringComparison.OrdinalIgnoreCase) Then
                options.AddExplorerSourcePath(argumentValue.Substring("--from-explorer=".Length))
            ElseIf IsLikelyPathArgument(argumentValue) Then
                ' Some shell verb invocations provide paths directly without our explicit flag.
                options.AddExplorerSourcePath(argumentValue)
            End If

            index += 1
        End While

        Return options
    End Function

    Private Shared Function ParseExplorerSelectionArguments(args As String(), flagIndex As Integer, options As LaunchOptions) As Integer
        Dim index = flagIndex
        While index + 1 < args.Length
            Dim candidate = args(index + 1)
            If String.IsNullOrWhiteSpace(candidate) Then
                index += 1
                Continue While
            End If

            If candidate.StartsWith("--", StringComparison.OrdinalIgnoreCase) Then
                Exit While
            End If

            index += 1
            options.AddExplorerSourcePath(candidate)
        End While

        Return index
    End Function

    Private Sub AddExplorerSourcePath(value As String)
        Dim resolved = UnwrapQuotedValue(value)
        If String.IsNullOrWhiteSpace(resolved) Then
            Return
        End If

        If String.Equals(resolved, "%1", StringComparison.OrdinalIgnoreCase) OrElse
           String.Equals(resolved, "%V", StringComparison.OrdinalIgnoreCase) OrElse
           String.Equals(resolved, "%*", StringComparison.OrdinalIgnoreCase) Then
            Return
        End If

        For Each existing In _explorerSourcePaths
            If String.Equals(existing, resolved, StringComparison.OrdinalIgnoreCase) Then
                Return
            End If
        Next

        _explorerSourcePaths.Add(resolved)
    End Sub

    Private Shared Function IsLikelyPathArgument(value As String) As Boolean
        If String.IsNullOrWhiteSpace(value) Then
            Return False
        End If

        Dim text = UnwrapQuotedValue(value)
        If text.Length = 0 Then
            Return False
        End If

        If text.StartsWith("--", StringComparison.OrdinalIgnoreCase) Then
            Return False
        End If

        If text.StartsWith("\\", StringComparison.OrdinalIgnoreCase) Then
            Return True
        End If

        If text.StartsWith(".\", StringComparison.OrdinalIgnoreCase) OrElse
           text.StartsWith("..\", StringComparison.OrdinalIgnoreCase) Then
            Return True
        End If

        Try
            Return Path.IsPathRooted(text)
        Catch
            Return False
        End Try
    End Function

    Private Shared Function UnwrapQuotedValue(value As String) As String
        Dim text = If(value, String.Empty).Trim()
        If text.Length >= 2 AndAlso text.StartsWith("""", StringComparison.Ordinal) AndAlso text.EndsWith("""", StringComparison.Ordinal) Then
            Return text.Substring(1, text.Length - 2)
        End If

        Return text
    End Function
End Class
