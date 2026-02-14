Imports System.IO
Imports System.Text.Json

Namespace Configuration
    Public NotInheritable Class AppSettingsStore
        Private ReadOnly _settingsPath As String

        Public Sub New(Optional settingsPath As String = Nothing)
            If String.IsNullOrWhiteSpace(settingsPath) Then
                Dim root = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "XactCopy")
                _settingsPath = Path.Combine(root, "settings.json")
            Else
                _settingsPath = settingsPath
            End If
        End Sub

        Public ReadOnly Property SettingsPath As String
            Get
                Return _settingsPath
            End Get
        End Property

        Public Function Load() As AppSettings
            If Not File.Exists(_settingsPath) Then
                Return New AppSettings()
            End If

            Try
                Dim json = File.ReadAllText(_settingsPath)
                Dim options As New JsonSerializerOptions() With {
                    .PropertyNameCaseInsensitive = True
                }

                Dim settings = JsonSerializer.Deserialize(Of AppSettings)(json, options)
                If settings Is Nothing Then
                    Return New AppSettings()
                End If

                Return Normalize(settings)
            Catch
                Return New AppSettings()
            End Try
        End Function

        Public Sub Save(settings As AppSettings)
            If settings Is Nothing Then
                Throw New ArgumentNullException(NameOf(settings))
            End If

            Dim settingsDirectory = Path.GetDirectoryName(_settingsPath)
            If Not String.IsNullOrWhiteSpace(settingsDirectory) Then
                Directory.CreateDirectory(settingsDirectory)
            End If

            Dim normalized = Normalize(settings)
            Dim options As New JsonSerializerOptions() With {
                .WriteIndented = True
            }
            Dim json = JsonSerializer.Serialize(normalized, options)
            File.WriteAllText(_settingsPath, json)
        End Sub

        Private Shared Function Normalize(settings As AppSettings) As AppSettings
            If settings Is Nothing Then
                Return New AppSettings()
            End If

            settings.Theme = If(String.IsNullOrWhiteSpace(settings.Theme), "dark", settings.Theme.Trim())
            settings.AccentColorMode =
                SettingsValueConverter.AccentColorModeChoiceToString(
                    SettingsValueConverter.ToAccentColorModeChoice(settings.AccentColorMode))
            settings.AccentColorHex =
                SettingsValueConverter.NormalizeColorHex(settings.AccentColorHex, "#5A78C8")
            settings.WindowChromeMode =
                SettingsValueConverter.WindowChromeModeChoiceToString(
                    SettingsValueConverter.ToWindowChromeModeChoice(settings.WindowChromeMode))
            settings.UiDensity =
                SettingsValueConverter.UiDensityChoiceToString(
                    SettingsValueConverter.ToUiDensityChoice(settings.UiDensity))
            settings.UiScalePercent = NormalizeUiScalePercent(settings.UiScalePercent)
            settings.LogFontFamily = If(settings.LogFontFamily, String.Empty).Trim()
            If String.IsNullOrWhiteSpace(settings.LogFontFamily) Then
                settings.LogFontFamily = "Consolas"
            End If
            settings.LogFontSizePoints =
                SettingsValueConverter.ClampInteger(settings.LogFontSizePoints, 7, 20, 9)
            settings.GridRowHeight =
                SettingsValueConverter.ClampInteger(settings.GridRowHeight, 18, 48, 24)
            settings.GridHeaderStyle =
                SettingsValueConverter.GridHeaderStyleChoiceToString(
                    SettingsValueConverter.ToGridHeaderStyleChoice(settings.GridHeaderStyle))
            settings.ProgressBarStyle =
                SettingsValueConverter.ProgressBarStyleChoiceToString(
                    SettingsValueConverter.ToProgressBarStyleChoice(settings.ProgressBarStyle))
            settings.UserAgent = If(settings.UserAgent, String.Empty).Trim()
            Dim updateReleaseUrl = If(settings.UpdateReleaseUrl, String.Empty).Trim()
            settings.UpdateReleaseUrl = If(String.IsNullOrWhiteSpace(updateReleaseUrl), AppSettings.DefaultUpdateReleaseUrl, updateReleaseUrl)

            settings.RecoveryTouchIntervalSeconds =
                SettingsValueConverter.ClampInteger(settings.RecoveryTouchIntervalSeconds, 1, 60, 2)

            settings.ExplorerSelectionMode =
                SettingsValueConverter.ExplorerSelectionModeToString(
                    SettingsValueConverter.ToExplorerSelectionMode(settings.ExplorerSelectionMode))

            settings.WorkerTelemetryProfile =
                SettingsValueConverter.WorkerTelemetryProfileToString(
                    SettingsValueConverter.ToWorkerTelemetryProfile(settings.WorkerTelemetryProfile))
            settings.WorkerProgressIntervalMs =
                SettingsValueConverter.ClampInteger(settings.WorkerProgressIntervalMs, 20, 1000, 75)
            settings.WorkerMaxLogsPerSecond =
                SettingsValueConverter.ClampInteger(settings.WorkerMaxLogsPerSecond, 0, 5000, 100)
            settings.UiDiagnosticsRefreshMs =
                SettingsValueConverter.ClampInteger(settings.UiDiagnosticsRefreshMs, 100, 5000, 250)
            settings.UiMaxLogLines =
                SettingsValueConverter.ClampInteger(settings.UiMaxLogLines, 1000, 1000000, 50000)

            settings.DefaultBufferSizeMb =
                SettingsValueConverter.ClampInteger(settings.DefaultBufferSizeMb, 1, 256, 4)
            settings.DefaultMaxRetries =
                SettingsValueConverter.ClampInteger(settings.DefaultMaxRetries, 0, 1000, 12)
            settings.DefaultOperationTimeoutSeconds =
                SettingsValueConverter.ClampInteger(settings.DefaultOperationTimeoutSeconds, 1, 3600, 10)
            settings.DefaultPerFileTimeoutSeconds =
                SettingsValueConverter.ClampInteger(settings.DefaultPerFileTimeoutSeconds, 0, 86400, 0)
            settings.DefaultMaxThroughputMbPerSecond =
                SettingsValueConverter.ClampInteger(settings.DefaultMaxThroughputMbPerSecond, 0, 4096, 0)
            settings.DefaultRescueFastScanChunkKb =
                SettingsValueConverter.ClampInteger(settings.DefaultRescueFastScanChunkKb, 0, 262144, 0)
            settings.DefaultRescueTrimChunkKb =
                SettingsValueConverter.ClampInteger(settings.DefaultRescueTrimChunkKb, 0, 262144, 0)
            settings.DefaultRescueScrapeChunkKb =
                SettingsValueConverter.ClampInteger(settings.DefaultRescueScrapeChunkKb, 0, 262144, 0)
            settings.DefaultRescueRetryChunkKb =
                SettingsValueConverter.ClampInteger(settings.DefaultRescueRetryChunkKb, 0, 262144, 0)
            settings.DefaultRescueSplitMinimumKb =
                SettingsValueConverter.ClampInteger(settings.DefaultRescueSplitMinimumKb, 0, 65536, 0)
            settings.DefaultRescueFastScanRetries =
                SettingsValueConverter.ClampInteger(settings.DefaultRescueFastScanRetries, 0, 1000, 0)
            settings.DefaultRescueTrimRetries =
                SettingsValueConverter.ClampInteger(settings.DefaultRescueTrimRetries, 0, 1000, 1)
            settings.DefaultRescueScrapeRetries =
                SettingsValueConverter.ClampInteger(settings.DefaultRescueScrapeRetries, 0, 1000, 2)
            settings.DefaultSampleVerificationChunkKb =
                SettingsValueConverter.ClampInteger(settings.DefaultSampleVerificationChunkKb, 32, 4096, 128)
            settings.DefaultSampleVerificationChunkCount =
                SettingsValueConverter.ClampInteger(settings.DefaultSampleVerificationChunkCount, 1, 64, 3)

            settings.DefaultOverwritePolicy =
                SettingsValueConverter.OverwritePolicyToString(
                    SettingsValueConverter.ToOverwritePolicy(settings.DefaultOverwritePolicy))
            settings.DefaultSymlinkHandling =
                SettingsValueConverter.SymlinkHandlingToString(
                    SettingsValueConverter.ToSymlinkHandling(settings.DefaultSymlinkHandling))
            settings.DefaultVerificationMode =
                SettingsValueConverter.VerificationModeToString(
                    SettingsValueConverter.ToVerificationMode(settings.DefaultVerificationMode))
            settings.DefaultVerificationHashAlgorithm =
                SettingsValueConverter.VerificationHashAlgorithmToString(
                    SettingsValueConverter.ToVerificationHashAlgorithm(settings.DefaultVerificationHashAlgorithm))
            settings.DefaultSalvageFillPattern =
                SettingsValueConverter.SalvageFillPatternToString(
                    SettingsValueConverter.ToSalvageFillPattern(settings.DefaultSalvageFillPattern))

            Return settings
        End Function

        Private Shared Function NormalizeUiScalePercent(value As Integer) As Integer
            Dim allowed = New Integer() {90, 100, 110, 125}
            Dim selected = 100
            Dim minDistance = Integer.MaxValue

            For Each candidate As Integer In allowed
                Dim distance = Math.Abs(candidate - value)
                If distance < minDistance Then
                    minDistance = distance
                    selected = candidate
                End If
            Next

            Return selected
        End Function
    End Class
End Namespace
