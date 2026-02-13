Imports XactCopy.Models

Namespace Configuration
    Friend Enum ExplorerSelectionMode
        SourceFolder = 0
        SelectedItems = 1
    End Enum

    Friend Module SettingsValueConverter
        Public Function NormalizeChoice(value As String, defaultValue As String, ParamArray allowedValues As String()) As String
            Dim candidate = If(value, String.Empty).Trim()
            If candidate.Length = 0 Then
                Return defaultValue
            End If

            For Each allowed In allowedValues
                If String.Equals(candidate, allowed, StringComparison.OrdinalIgnoreCase) Then
                    Return allowed
                End If
            Next

            Return defaultValue
        End Function

        Public Function ClampInteger(value As Integer, minimum As Integer, maximum As Integer, defaultValue As Integer) As Integer
            If value < minimum OrElse value > maximum Then
                Return defaultValue
            End If

            Return value
        End Function

        Public Function ToOverwritePolicy(value As String) As OverwritePolicy
            Select Case NormalizeChoice(value, "overwrite", "overwrite", "skip-existing", "overwrite-if-newer")
                Case "skip-existing"
                    Return OverwritePolicy.SkipExisting
                Case "overwrite-if-newer"
                    Return OverwritePolicy.OverwriteIfSourceNewer
                Case Else
                    Return OverwritePolicy.Overwrite
            End Select
        End Function

        Public Function OverwritePolicyToString(value As OverwritePolicy) As String
            Select Case value
                Case OverwritePolicy.SkipExisting
                    Return "skip-existing"
                Case OverwritePolicy.OverwriteIfSourceNewer
                    Return "overwrite-if-newer"
                Case Else
                    Return "overwrite"
            End Select
        End Function

        Public Function ToSymlinkHandling(value As String) As SymlinkHandlingMode
            Select Case NormalizeChoice(value, "skip", "skip", "follow")
                Case "follow"
                    Return SymlinkHandlingMode.Follow
                Case Else
                    Return SymlinkHandlingMode.Skip
            End Select
        End Function

        Public Function SymlinkHandlingToString(value As SymlinkHandlingMode) As String
            If value = SymlinkHandlingMode.Follow Then
                Return "follow"
            End If

            Return "skip"
        End Function

        Public Function ToVerificationMode(value As String) As VerificationMode
            Select Case NormalizeChoice(value, "none", "none", "sampled", "full")
                Case "sampled"
                    Return VerificationMode.Sampled
                Case "full"
                    Return VerificationMode.Full
                Case Else
                    Return VerificationMode.None
            End Select
        End Function

        Public Function VerificationModeToString(value As VerificationMode) As String
            Select Case value
                Case VerificationMode.Sampled
                    Return "sampled"
                Case VerificationMode.Full
                    Return "full"
                Case Else
                    Return "none"
            End Select
        End Function

        Public Function ToVerificationHashAlgorithm(value As String) As VerificationHashAlgorithm
            Select Case NormalizeChoice(value, "sha256", "sha256", "sha512")
                Case "sha512"
                    Return VerificationHashAlgorithm.Sha512
                Case Else
                    Return VerificationHashAlgorithm.Sha256
            End Select
        End Function

        Public Function VerificationHashAlgorithmToString(value As VerificationHashAlgorithm) As String
            If value = VerificationHashAlgorithm.Sha512 Then
                Return "sha512"
            End If

            Return "sha256"
        End Function

        Public Function ToSalvageFillPattern(value As String) As SalvageFillPattern
            Select Case NormalizeChoice(value, "zero", "zero", "ones", "random")
                Case "ones"
                    Return SalvageFillPattern.Ones
                Case "random"
                    Return SalvageFillPattern.Random
                Case Else
                    Return SalvageFillPattern.Zero
            End Select
        End Function

        Public Function SalvageFillPatternToString(value As SalvageFillPattern) As String
            Select Case value
                Case SalvageFillPattern.Ones
                    Return "ones"
                Case SalvageFillPattern.Random
                    Return "random"
                Case Else
                    Return "zero"
            End Select
        End Function

        Public Function ToExplorerSelectionMode(value As String) As ExplorerSelectionMode
            Select Case NormalizeChoice(value, "selected-items", "source-folder", "selected-items")
                Case "source-folder"
                    Return ExplorerSelectionMode.SourceFolder
                Case Else
                    Return ExplorerSelectionMode.SelectedItems
            End Select
        End Function

        Public Function ExplorerSelectionModeToString(value As ExplorerSelectionMode) As String
            If value = ExplorerSelectionMode.SourceFolder Then
                Return "source-folder"
            End If

            Return "selected-items"
        End Function
    End Module
End Namespace
