' -----------------------------------------------------------------------------
' File: src\XactCopy.UI\Configuration\SettingsValueConverter.vb
' Purpose: Source file for XactCopy runtime behavior.
' -----------------------------------------------------------------------------

Imports System.Drawing
Imports XactCopy.Models

Namespace Configuration
    ''' <summary>
    ''' Enum AccentColorModeChoice.
    ''' </summary>
    Friend Enum AccentColorModeChoice
        Auto = 0
        System = 1
        Custom = 2
    End Enum

    ''' <summary>
    ''' Enum UiDensityChoice.
    ''' </summary>
    Friend Enum UiDensityChoice
        Compact = 0
        Normal = 1
        Comfortable = 2
    End Enum

    ''' <summary>
    ''' Enum GridHeaderStyleChoice.
    ''' </summary>
    Friend Enum GridHeaderStyleChoice
        [Default] = 0
        Minimal = 1
        Prominent = 2
    End Enum

    ''' <summary>
    ''' Enum ProgressBarStyleChoice.
    ''' </summary>
    Friend Enum ProgressBarStyleChoice
        Thin = 0
        Standard = 1
        Thick = 2
    End Enum

    ''' <summary>
    ''' Enum WindowChromeModeChoice.
    ''' </summary>
    Friend Enum WindowChromeModeChoice
        Themed = 0
        Standard = 1
    End Enum

    ''' <summary>
    ''' Enum ExplorerSelectionMode.
    ''' </summary>
    Friend Enum ExplorerSelectionMode
        SourceFolder = 0
        SelectedItems = 1
    End Enum

    ''' <summary>
    ''' Module SettingsValueConverter.
    ''' </summary>
    Friend Module SettingsValueConverter
        ''' <summary>
        ''' Computes NormalizeChoice.
        ''' </summary>
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

        ''' <summary>
        ''' Computes ClampInteger.
        ''' </summary>
        Public Function ClampInteger(value As Integer, minimum As Integer, maximum As Integer, defaultValue As Integer) As Integer
            If value < minimum OrElse value > maximum Then
                Return defaultValue
            End If

            Return value
        End Function

        ''' <summary>
        ''' Computes ToAccentColorModeChoice.
        ''' </summary>
        Public Function ToAccentColorModeChoice(value As String) As AccentColorModeChoice
            Select Case NormalizeChoice(value, "auto", "auto", "system", "custom")
                Case "system"
                    Return AccentColorModeChoice.System
                Case "custom"
                    Return AccentColorModeChoice.Custom
                Case Else
                    Return AccentColorModeChoice.Auto
            End Select
        End Function

        ''' <summary>
        ''' Computes AccentColorModeChoiceToString.
        ''' </summary>
        Public Function AccentColorModeChoiceToString(value As AccentColorModeChoice) As String
            Select Case value
                Case AccentColorModeChoice.System
                    Return "system"
                Case AccentColorModeChoice.Custom
                    Return "custom"
                Case Else
                    Return "auto"
            End Select
        End Function

        ''' <summary>
        ''' Computes ToUiDensityChoice.
        ''' </summary>
        Public Function ToUiDensityChoice(value As String) As UiDensityChoice
            Select Case NormalizeChoice(value, "normal", "compact", "normal", "comfortable")
                Case "compact"
                    Return UiDensityChoice.Compact
                Case "comfortable"
                    Return UiDensityChoice.Comfortable
                Case Else
                    Return UiDensityChoice.Normal
            End Select
        End Function

        ''' <summary>
        ''' Computes UiDensityChoiceToString.
        ''' </summary>
        Public Function UiDensityChoiceToString(value As UiDensityChoice) As String
            Select Case value
                Case UiDensityChoice.Compact
                    Return "compact"
                Case UiDensityChoice.Comfortable
                    Return "comfortable"
                Case Else
                    Return "normal"
            End Select
        End Function

        ''' <summary>
        ''' Computes ToGridHeaderStyleChoice.
        ''' </summary>
        Public Function ToGridHeaderStyleChoice(value As String) As GridHeaderStyleChoice
            Select Case NormalizeChoice(value, "default", "default", "minimal", "prominent")
                Case "minimal"
                    Return GridHeaderStyleChoice.Minimal
                Case "prominent"
                    Return GridHeaderStyleChoice.Prominent
                Case Else
                    Return GridHeaderStyleChoice.Default
            End Select
        End Function

        ''' <summary>
        ''' Computes GridHeaderStyleChoiceToString.
        ''' </summary>
        Public Function GridHeaderStyleChoiceToString(value As GridHeaderStyleChoice) As String
            Select Case value
                Case GridHeaderStyleChoice.Minimal
                    Return "minimal"
                Case GridHeaderStyleChoice.Prominent
                    Return "prominent"
                Case Else
                    Return "default"
            End Select
        End Function

        ''' <summary>
        ''' Computes ToProgressBarStyleChoice.
        ''' </summary>
        Public Function ToProgressBarStyleChoice(value As String) As ProgressBarStyleChoice
            Select Case NormalizeChoice(value, "standard", "thin", "standard", "thick")
                Case "thin"
                    Return ProgressBarStyleChoice.Thin
                Case "thick"
                    Return ProgressBarStyleChoice.Thick
                Case Else
                    Return ProgressBarStyleChoice.Standard
            End Select
        End Function

        ''' <summary>
        ''' Computes ProgressBarStyleChoiceToString.
        ''' </summary>
        Public Function ProgressBarStyleChoiceToString(value As ProgressBarStyleChoice) As String
            Select Case value
                Case ProgressBarStyleChoice.Thin
                    Return "thin"
                Case ProgressBarStyleChoice.Thick
                    Return "thick"
                Case Else
                    Return "standard"
            End Select
        End Function

        ''' <summary>
        ''' Computes ToWindowChromeModeChoice.
        ''' </summary>
        Public Function ToWindowChromeModeChoice(value As String) As WindowChromeModeChoice
            Select Case NormalizeChoice(value, "themed", "themed", "standard")
                Case "standard"
                    Return WindowChromeModeChoice.Standard
                Case Else
                    Return WindowChromeModeChoice.Themed
            End Select
        End Function

        ''' <summary>
        ''' Computes WindowChromeModeChoiceToString.
        ''' </summary>
        Public Function WindowChromeModeChoiceToString(value As WindowChromeModeChoice) As String
            If value = WindowChromeModeChoice.Standard Then
                Return "standard"
            End If

            Return "themed"
        End Function

        ''' <summary>
        ''' Computes NormalizeColorHex.
        ''' </summary>
        Public Function NormalizeColorHex(value As String, defaultValue As String) As String
            Dim candidate = If(value, String.Empty).Trim()
            Dim parsed As Color = Nothing
            If TryParseColorHex(candidate, parsed) Then
                Return ColorToHex(parsed)
            End If

            If TryParseColorHex(defaultValue, parsed) Then
                Return ColorToHex(parsed)
            End If

            Return "#5A78C8"
        End Function

        ''' <summary>
        ''' Computes TryParseColorHex.
        ''' </summary>
        Public Function TryParseColorHex(value As String, ByRef parsedColor As Color) As Boolean
            parsedColor = Color.Empty
            Dim candidate = If(value, String.Empty).Trim()
            If candidate.Length = 0 Then
                Return False
            End If

            If candidate.StartsWith("#", StringComparison.Ordinal) Then
                candidate = candidate.Substring(1)
            End If

            If candidate.Length <> 6 Then
                Return False
            End If

            Dim red As Integer
            Dim green As Integer
            Dim blue As Integer

            If Not Integer.TryParse(candidate.Substring(0, 2), Globalization.NumberStyles.HexNumber, Globalization.CultureInfo.InvariantCulture, red) Then
                Return False
            End If
            If Not Integer.TryParse(candidate.Substring(2, 2), Globalization.NumberStyles.HexNumber, Globalization.CultureInfo.InvariantCulture, green) Then
                Return False
            End If
            If Not Integer.TryParse(candidate.Substring(4, 2), Globalization.NumberStyles.HexNumber, Globalization.CultureInfo.InvariantCulture, blue) Then
                Return False
            End If

            parsedColor = Color.FromArgb(red, green, blue)
            Return True
        End Function

        ''' <summary>
        ''' Computes ColorToHex.
        ''' </summary>
        Public Function ColorToHex(value As Color) As String
            Return $"#{value.R:X2}{value.G:X2}{value.B:X2}"
        End Function

        ''' <summary>
        ''' Computes ToOverwritePolicy.
        ''' </summary>
        Public Function ToOverwritePolicy(value As String) As OverwritePolicy
            Select Case NormalizeChoice(value, "overwrite", "overwrite", "skip-existing", "overwrite-if-newer", "ask")
                Case "skip-existing"
                    Return OverwritePolicy.SkipExisting
                Case "overwrite-if-newer"
                    Return OverwritePolicy.OverwriteIfSourceNewer
                Case "ask"
                    Return OverwritePolicy.Ask
                Case Else
                    Return OverwritePolicy.Overwrite
            End Select
        End Function

        ''' <summary>
        ''' Computes OverwritePolicyToString.
        ''' </summary>
        Public Function OverwritePolicyToString(value As OverwritePolicy) As String
            Select Case value
                Case OverwritePolicy.SkipExisting
                    Return "skip-existing"
                Case OverwritePolicy.OverwriteIfSourceNewer
                    Return "overwrite-if-newer"
                Case OverwritePolicy.Ask
                    Return "ask"
                Case Else
                    Return "overwrite"
            End Select
        End Function

        ''' <summary>
        ''' Computes ToSymlinkHandling.
        ''' </summary>
        Public Function ToSymlinkHandling(value As String) As SymlinkHandlingMode
            Select Case NormalizeChoice(value, "skip", "skip", "follow")
                Case "follow"
                    Return SymlinkHandlingMode.Follow
                Case Else
                    Return SymlinkHandlingMode.Skip
            End Select
        End Function

        ''' <summary>
        ''' Computes SymlinkHandlingToString.
        ''' </summary>
        Public Function SymlinkHandlingToString(value As SymlinkHandlingMode) As String
            If value = SymlinkHandlingMode.Follow Then
                Return "follow"
            End If

            Return "skip"
        End Function

        ''' <summary>
        ''' Computes ToVerificationMode.
        ''' </summary>
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

        ''' <summary>
        ''' Computes VerificationModeToString.
        ''' </summary>
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

        ''' <summary>
        ''' Computes ToVerificationHashAlgorithm.
        ''' </summary>
        Public Function ToVerificationHashAlgorithm(value As String) As VerificationHashAlgorithm
            Select Case NormalizeChoice(value, "sha256", "sha256", "sha512")
                Case "sha512"
                    Return VerificationHashAlgorithm.Sha512
                Case Else
                    Return VerificationHashAlgorithm.Sha256
            End Select
        End Function

        ''' <summary>
        ''' Computes VerificationHashAlgorithmToString.
        ''' </summary>
        Public Function VerificationHashAlgorithmToString(value As VerificationHashAlgorithm) As String
            If value = VerificationHashAlgorithm.Sha512 Then
                Return "sha512"
            End If

            Return "sha256"
        End Function

        ''' <summary>
        ''' Computes ToSalvageFillPattern.
        ''' </summary>
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

        ''' <summary>
        ''' Computes SalvageFillPatternToString.
        ''' </summary>
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

        ''' <summary>
        ''' Computes ToExplorerSelectionMode.
        ''' </summary>
        Public Function ToExplorerSelectionMode(value As String) As ExplorerSelectionMode
            Select Case NormalizeChoice(value, "selected-items", "source-folder", "selected-items")
                Case "source-folder"
                    Return ExplorerSelectionMode.SourceFolder
                Case Else
                    Return ExplorerSelectionMode.SelectedItems
            End Select
        End Function

        ''' <summary>
        ''' Computes ExplorerSelectionModeToString.
        ''' </summary>
        Public Function ExplorerSelectionModeToString(value As ExplorerSelectionMode) As String
            If value = ExplorerSelectionMode.SourceFolder Then
                Return "source-folder"
            End If

            Return "selected-items"
        End Function

        ''' <summary>
        ''' Computes ToWorkerTelemetryProfile.
        ''' </summary>
        Public Function ToWorkerTelemetryProfile(value As String) As WorkerTelemetryProfile
            Select Case NormalizeChoice(value, "normal", "normal", "verbose", "debug")
                Case "verbose"
                    Return WorkerTelemetryProfile.Verbose
                Case "debug"
                    Return WorkerTelemetryProfile.Debug
                Case Else
                    Return WorkerTelemetryProfile.Normal
            End Select
        End Function

        ''' <summary>
        ''' Computes WorkerTelemetryProfileToString.
        ''' </summary>
        Public Function WorkerTelemetryProfileToString(value As WorkerTelemetryProfile) As String
            Select Case value
                Case WorkerTelemetryProfile.Verbose
                    Return "verbose"
                Case WorkerTelemetryProfile.Debug
                    Return "debug"
                Case Else
                    Return "normal"
            End Select
        End Function

        ''' <summary>
        ''' Computes ToSourceMutationPolicy.
        ''' </summary>
        Public Function ToSourceMutationPolicy(value As String) As SourceMutationPolicy
            Select Case NormalizeChoice(value, "fail-file", "fail-file", "skip-file", "wait-for-reappearance")
                Case "skip-file"
                    Return SourceMutationPolicy.SkipFile
                Case "wait-for-reappearance"
                    Return SourceMutationPolicy.WaitForReappearance
                Case Else
                    Return SourceMutationPolicy.FailFile
            End Select
        End Function

        ''' <summary>
        ''' Computes SourceMutationPolicyToString.
        ''' </summary>
        Public Function SourceMutationPolicyToString(value As SourceMutationPolicy) As String
            Select Case value
                Case SourceMutationPolicy.SkipFile
                    Return "skip-file"
                Case SourceMutationPolicy.WaitForReappearance
                    Return "wait-for-reappearance"
                Case Else
                    Return "fail-file"
            End Select
        End Function
    End Module
End Namespace
