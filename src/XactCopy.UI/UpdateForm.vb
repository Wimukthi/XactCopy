Imports System.Diagnostics
Imports System.IO
Imports System.IO.Compression
Imports System.Text
Imports System.Threading
Imports System.Threading.Tasks
Imports XactCopy.Configuration

Friend Class UpdateForm
    Inherits Form

    Private ReadOnly _release As UpdateReleaseInfo
    Private ReadOnly _asset As UpdateAssetInfo
    Private ReadOnly _currentVersion As Version
    Private ReadOnly _settings As AppSettings

    Private ReadOnly _currentValueLabel As New Label()
    Private ReadOnly _latestValueLabel As New Label()
    Private ReadOnly _packageValueLabel As New Label()
    Private ReadOnly _notesTextBox As New TextBox()
    Private ReadOnly _progressBar As New ProgressBar()
    Private ReadOnly _progressLabel As New Label()
    Private ReadOnly _speedLabel As New Label()
    Private ReadOnly _statusLabel As New Label()
    Private ReadOnly _downloadButton As New Button()
    Private ReadOnly _releaseButton As New Button()
    Private ReadOnly _closeButton As New Button()

    Private _cancellation As CancellationTokenSource
    Private _tempRoot As String = String.Empty
    Private _downloadPath As String = String.Empty
    Private _isDownloading As Boolean
    Private _isApplying As Boolean
    Private _allowCloseDuringApply As Boolean

    Public Sub New(release As UpdateReleaseInfo, currentVersion As Version, settings As AppSettings)
        If release Is Nothing Then
            Throw New ArgumentNullException(NameOf(release))
        End If

        If settings Is Nothing Then
            Throw New ArgumentNullException(NameOf(settings))
        End If

        _release = release
        _currentVersion = If(currentVersion, New Version(0, 0))
        _settings = settings
        _asset = UpdateService.SelectBestAsset(release)

        Text = "XactCopy Update"
        StartPosition = FormStartPosition.CenterParent
        MinimumSize = New Size(700, 500)
        Size = New Size(860, 620)
        WindowIconHelper.Apply(Me)

        BuildUi()
        PopulateReleaseInfo()
    End Sub

    Private Sub BuildUi()
        Dim root As New TableLayoutPanel() With {
            .Dock = DockStyle.Fill,
            .ColumnCount = 1,
            .RowCount = 4,
            .Padding = New Padding(12)
        }
        root.RowStyles.Add(New RowStyle(SizeType.AutoSize))
        root.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0F))
        root.RowStyles.Add(New RowStyle(SizeType.AutoSize))
        root.RowStyles.Add(New RowStyle(SizeType.AutoSize))

        root.Controls.Add(BuildSummaryPanel(), 0, 0)

        _notesTextBox.Dock = DockStyle.Fill
        _notesTextBox.Multiline = True
        _notesTextBox.ReadOnly = True
        _notesTextBox.ScrollBars = ScrollBars.Vertical
        _notesTextBox.WordWrap = True
        root.Controls.Add(_notesTextBox, 0, 1)

        root.Controls.Add(BuildProgressPanel(), 0, 2)
        root.Controls.Add(BuildButtonsPanel(), 0, 3)

        Controls.Add(root)
    End Sub

    Private Function BuildSummaryPanel() As Control
        Dim summary As New TableLayoutPanel() With {
            .Dock = DockStyle.Fill,
            .ColumnCount = 2,
            .RowCount = 4,
            .Margin = New Padding(0, 0, 0, 10)
        }
        summary.ColumnStyles.Add(New ColumnStyle(SizeType.AutoSize))
        summary.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0F))
        summary.RowStyles.Add(New RowStyle(SizeType.AutoSize))
        summary.RowStyles.Add(New RowStyle(SizeType.AutoSize))
        summary.RowStyles.Add(New RowStyle(SizeType.AutoSize))
        summary.RowStyles.Add(New RowStyle(SizeType.AutoSize))

        Dim titleLabel As New Label() With {
            .AutoSize = True,
            .Font = New Font(Font, FontStyle.Bold),
            .Text = "Update Available",
            .Margin = New Padding(0, 0, 0, 8)
        }
        summary.Controls.Add(titleLabel, 0, 0)
        summary.SetColumnSpan(titleLabel, 2)

        Dim currentLabel As New Label() With {
            .AutoSize = True,
            .Text = "Current version:",
            .Margin = New Padding(0, 0, 12, 4)
        }
        _currentValueLabel.AutoSize = True
        _currentValueLabel.Margin = New Padding(0, 0, 0, 4)

        Dim latestLabel As New Label() With {
            .AutoSize = True,
            .Text = "Latest version:",
            .Margin = New Padding(0, 0, 12, 4)
        }
        _latestValueLabel.AutoSize = True
        _latestValueLabel.Margin = New Padding(0, 0, 0, 4)

        Dim packageLabel As New Label() With {
            .AutoSize = True,
            .Text = "Package:",
            .Margin = New Padding(0, 0, 12, 0)
        }
        _packageValueLabel.AutoSize = True
        _packageValueLabel.AutoEllipsis = True

        summary.Controls.Add(currentLabel, 0, 1)
        summary.Controls.Add(_currentValueLabel, 1, 1)
        summary.Controls.Add(latestLabel, 0, 2)
        summary.Controls.Add(_latestValueLabel, 1, 2)
        summary.Controls.Add(packageLabel, 0, 3)
        summary.Controls.Add(_packageValueLabel, 1, 3)

        Return summary
    End Function

    Private Function BuildProgressPanel() As Control
        Dim panel As New TableLayoutPanel() With {
            .Dock = DockStyle.Fill,
            .ColumnCount = 1,
            .RowCount = 4,
            .AutoSize = True,
            .Margin = New Padding(0, 10, 0, 10)
        }
        panel.RowStyles.Add(New RowStyle(SizeType.AutoSize))
        panel.RowStyles.Add(New RowStyle(SizeType.AutoSize))
        panel.RowStyles.Add(New RowStyle(SizeType.AutoSize))
        panel.RowStyles.Add(New RowStyle(SizeType.AutoSize))

        _progressBar.Dock = DockStyle.Fill
        _progressBar.Minimum = 0
        _progressBar.Maximum = 100
        _progressBar.Style = ProgressBarStyle.Continuous

        _progressLabel.AutoSize = True
        _progressLabel.Text = "0 B / 0 B"

        _speedLabel.AutoSize = True
        _speedLabel.Text = "Speed: -"

        _statusLabel.AutoSize = True
        _statusLabel.Text = "Ready."

        panel.Controls.Add(_progressBar, 0, 0)
        panel.Controls.Add(_progressLabel, 0, 1)
        panel.Controls.Add(_speedLabel, 0, 2)
        panel.Controls.Add(_statusLabel, 0, 3)

        Return panel
    End Function

    Private Function BuildButtonsPanel() As Control
        Dim panel As New FlowLayoutPanel() With {
            .Dock = DockStyle.Fill,
            .FlowDirection = FlowDirection.RightToLeft,
            .WrapContents = False,
            .AutoSize = True
        }

        _closeButton.Text = "Close"
        _closeButton.AutoSize = True
        AddHandler _closeButton.Click, AddressOf CloseButton_Click

        _releaseButton.Text = "View Release"
        _releaseButton.AutoSize = True
        AddHandler _releaseButton.Click, AddressOf ViewReleaseButton_Click

        _downloadButton.Text = "Download && Install"
        _downloadButton.AutoSize = True
        AddHandler _downloadButton.Click, AddressOf DownloadButton_Click

        panel.Controls.Add(_closeButton)
        panel.Controls.Add(_releaseButton)
        panel.Controls.Add(_downloadButton)

        Return panel
    End Function

    Private Sub PopulateReleaseInfo()
        _currentValueLabel.Text = FormatVersionDisplay(_currentVersion, Nothing)
        _latestValueLabel.Text = FormatVersionDisplay(_release.Version, _release.TagName)
        _packageValueLabel.Text = If(_asset IsNot Nothing, _asset.Name, "No compatible package found in release assets.")
        _notesTextBox.Text = If(String.IsNullOrWhiteSpace(_release.Notes), "No release notes provided.", NormalizeReleaseNotes(_release.Notes))

        If _asset Is Nothing Then
            _downloadButton.Enabled = False
            _statusLabel.Text = "No compatible package is available."
        End If

        If String.IsNullOrWhiteSpace(_release.HtmlUrl) Then
            _releaseButton.Enabled = False
        End If
    End Sub

    Private Shared Function FormatVersionDisplay(version As Version, tagName As String) As String
        If version Is Nothing Then
            Return If(String.IsNullOrWhiteSpace(tagName), "Unknown", tagName)
        End If

        Dim build As Integer = If(version.Build >= 0, version.Build, 0)
        Dim revision As Integer = If(version.Revision >= 0, version.Revision, 0)
        Dim text As String = $"{version.Major}.{version.Minor}.{build}.{revision}"

        If version.Major = 0 AndAlso version.Minor = 0 AndAlso build = 0 AndAlso revision = 0 AndAlso
            Not String.IsNullOrWhiteSpace(tagName) Then
            Return tagName
        End If

        Return text
    End Function

    Private Async Sub DownloadButton_Click(sender As Object, e As EventArgs)
        If _asset Is Nothing OrElse _cancellation IsNot Nothing Then
            Return
        End If

        _isDownloading = True
        _isApplying = False
        _allowCloseDuringApply = False
        _downloadButton.Enabled = False
        _releaseButton.Enabled = False
        _closeButton.Text = "Cancel"
        _closeButton.Enabled = True
        _progressBar.Value = 0
        _statusLabel.Text = "Downloading update package..."
        _cancellation = New CancellationTokenSource()

        Try
            _tempRoot = CreateTempRoot()
            Directory.CreateDirectory(_tempRoot)
            _downloadPath = Path.Combine(_tempRoot, _asset.Name)

            Dim progress = New Progress(Of DownloadProgressInfo)(AddressOf ReportDownloadProgress)
            Await UpdateService.DownloadAssetAsync(_settings, _asset, _downloadPath, progress, _cancellation.Token)

            If _cancellation.IsCancellationRequested Then
                _statusLabel.Text = "Download canceled."
                _closeButton.Text = "Close"
                _downloadButton.Enabled = True
                _releaseButton.Enabled = Not String.IsNullOrWhiteSpace(_release.HtmlUrl)
                Return
            End If

            _isDownloading = False
            _isApplying = True
            _closeButton.Enabled = False
            _statusLabel.Text = "Preparing update..."
            Await ApplyUpdateAsync(_downloadPath)
        Catch ex As OperationCanceledException
            _statusLabel.Text = "Download canceled."
            _closeButton.Text = "Close"
            _downloadButton.Enabled = True
            _releaseButton.Enabled = Not String.IsNullOrWhiteSpace(_release.HtmlUrl)
        Catch ex As Exception
            _statusLabel.Text = "Update failed: " & ex.Message
            _closeButton.Text = "Close"
            _closeButton.Enabled = True
            _downloadButton.Enabled = True
            _releaseButton.Enabled = Not String.IsNullOrWhiteSpace(_release.HtmlUrl)
            _isApplying = False
        Finally
            _isDownloading = False
            _cancellation?.Dispose()
            _cancellation = Nothing
        End Try
    End Sub

    Private Sub CloseButton_Click(sender As Object, e As EventArgs)
        If _isDownloading AndAlso _cancellation IsNot Nothing Then
            _statusLabel.Text = "Canceling..."
            _cancellation.Cancel()
            Return
        End If

        If _isApplying Then
            _statusLabel.Text = "Applying update. XactCopy will close and restart automatically."
            Return
        End If

        Close()
    End Sub

    Private Sub ViewReleaseButton_Click(sender As Object, e As EventArgs)
        If String.IsNullOrWhiteSpace(_release.HtmlUrl) Then
            Return
        End If

        Try
            Process.Start(New ProcessStartInfo(_release.HtmlUrl) With {.UseShellExecute = True})
        Catch
            ' Ignore browser launch failures.
        End Try
    End Sub

    Private Sub ReportDownloadProgress(progress As DownloadProgressInfo)
        _progressBar.Value = Math.Clamp(progress.Percent, 0, 100)
        Dim totalText = If(progress.TotalBytes > 0, UpdateService.FormatBytes(progress.TotalBytes), "Unknown")
        _progressLabel.Text = $"{UpdateService.FormatBytes(progress.BytesReceived)} / {totalText}"
        Dim speedText = If(progress.SpeedBytesPerSec > 0, UpdateService.FormatBytes(CLng(progress.SpeedBytesPerSec)) & "/s", "-")
        _speedLabel.Text = "Speed: " & speedText
    End Sub

    Private Async Function ApplyUpdateAsync(downloadPath As String) As Task
        If String.IsNullOrWhiteSpace(downloadPath) OrElse Not File.Exists(downloadPath) Then
            Throw New FileNotFoundException("Downloaded update package was not found.")
        End If

        Dim targetDirectory = AppContext.BaseDirectory
        Dim executableName = Path.GetFileName(Application.ExecutablePath)
        Dim extension = Path.GetExtension(downloadPath).ToLowerInvariant()

        If Not CanWriteToFolder(targetDirectory) Then
            Throw New InvalidOperationException("The application folder is not writable. Run XactCopy with write access to install updates.")
        End If

        If extension = ".exe" Then
            _allowCloseDuringApply = True
            Process.Start(New ProcessStartInfo(downloadPath) With {.UseShellExecute = True})
            Application.Exit()
            Return
        End If

        If extension <> ".zip" Then
            Throw New InvalidOperationException("Unsupported update package format.")
        End If

        Dim extractRoot = Path.Combine(_tempRoot, "payload")
        If Directory.Exists(extractRoot) Then
            Directory.Delete(extractRoot, recursive:=True)
        End If
        Directory.CreateDirectory(extractRoot)

        Await Task.Run(Sub() ZipFile.ExtractToDirectory(downloadPath, extractRoot, overwriteFiles:=True))

        Dim executableCandidates = Directory.GetFiles(extractRoot, executableName, SearchOption.AllDirectories)
        Dim executablePath = If(executableCandidates.Length > 0, executableCandidates(0), String.Empty)
        Dim sourceDirectory = If(String.IsNullOrWhiteSpace(executablePath), extractRoot, Path.GetDirectoryName(executablePath))
        If String.IsNullOrWhiteSpace(sourceDirectory) OrElse Not Directory.Exists(sourceDirectory) Then
            Throw New InvalidOperationException("Unable to locate update files in downloaded package.")
        End If

        LaunchUpdateScript(sourceDirectory, targetDirectory, executableName, _tempRoot)
        _allowCloseDuringApply = True
        Application.Exit()
    End Function

    Private Shared Function CreateTempRoot() As String
        Dim baseDirectory = Path.Combine(Path.GetTempPath(), "XactCopyUpdate")
        Dim instanceId = Guid.NewGuid().ToString("N")
        Return Path.Combine(baseDirectory, instanceId)
    End Function

    Private Shared Function CanWriteToFolder(folder As String) As Boolean
        Try
            Dim probeFile = Path.Combine(folder, "write-probe.tmp")
            File.WriteAllText(probeFile, "x")
            File.Delete(probeFile)
            Return True
        Catch
            Return False
        End Try
    End Function

    Private Shared Sub LaunchUpdateScript(sourceDirectory As String, targetDirectory As String, executableName As String, tempRoot As String)
        Dim scriptPath = Path.Combine(tempRoot, "apply-update.cmd")
        Dim processId = Process.GetCurrentProcess().Id
        Dim executablePath = Path.Combine(targetDirectory, executableName)
        Dim backupDirectory = Path.Combine(tempRoot, "backup")
        Dim logPath = Path.Combine(tempRoot, "update.log")
        Dim normalizedTarget = targetDirectory
        If normalizedTarget.Length > 3 Then
            normalizedTarget = normalizedTarget.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar)
        End If

        Dim script As New StringBuilder()
        script.AppendLine("@echo off")
        script.AppendLine("setlocal enableextensions")
        script.AppendLine($"set ""PID={processId}""")
        script.AppendLine($"set ""SOURCE={sourceDirectory}""")
        script.AppendLine($"set ""TARGET={normalizedTarget}""")
        script.AppendLine($"set ""BACKUP={backupDirectory}""")
        script.AppendLine($"set ""LOG={logPath}""")
        script.AppendLine(">""%LOG%"" echo XactCopy update log")
        script.AppendLine(":wait")
        script.AppendLine("tasklist /FI ""PID eq %PID%"" | find /I ""%PID%"" >nul")
        script.AppendLine("if not errorlevel 1 (")
        script.AppendLine("  timeout /t 1 /nobreak >nul")
        script.AppendLine("  goto wait")
        script.AppendLine(")")
        script.AppendLine("if exist ""%BACKUP%"" rd /s /q ""%BACKUP%""")
        script.AppendLine("mkdir ""%BACKUP%""")
        script.AppendLine("robocopy ""%TARGET%"" ""%BACKUP%"" /E /COPY:DAT /R:2 /W:1 /NFL /NDL /NJH /NJS /NP >>""%LOG%""")
        script.AppendLine("set ""RC=%ERRORLEVEL%""")
        script.AppendLine("if %RC% GEQ 8 (")
        script.AppendLine("  echo Backup failed with error %RC%. >>""%LOG%""")
        script.AppendLine("  goto launch")
        script.AppendLine(")")
        script.AppendLine("robocopy ""%SOURCE%"" ""%TARGET%"" /E /COPY:DAT /R:2 /W:1 /NFL /NDL /NJH /NJS /NP >>""%LOG%""")
        script.AppendLine("set ""RC=%ERRORLEVEL%""")
        script.AppendLine("if %RC% GEQ 8 (")
        script.AppendLine("  echo Update copy failed with error %RC%. Restoring backup. >>""%LOG%""")
        script.AppendLine("  robocopy ""%BACKUP%"" ""%TARGET%"" /E /COPY:DAT /R:2 /W:1 /NFL /NDL /NJH /NJS /NP >>""%LOG%""")
        script.AppendLine(")")
        script.AppendLine(":launch")
        script.AppendLine($"start """" ""{executablePath}""")
        script.AppendLine("endlocal")

        File.WriteAllText(scriptPath, script.ToString(), Encoding.ASCII)

        Dim startInfo As New ProcessStartInfo(scriptPath) With {
            .WorkingDirectory = tempRoot,
            .UseShellExecute = True,
            .WindowStyle = ProcessWindowStyle.Hidden
        }

        Process.Start(startInfo)
    End Sub

    Protected Overrides Sub OnFormClosing(e As FormClosingEventArgs)
        If _isDownloading AndAlso _cancellation IsNot Nothing Then
            _statusLabel.Text = "Canceling..."
            _cancellation.Cancel()
            e.Cancel = True
            Return
        End If

        If _isApplying AndAlso Not _allowCloseDuringApply Then
            _statusLabel.Text = "Applying update. Please wait..."
            e.Cancel = True
            Return
        End If

        MyBase.OnFormClosing(e)
    End Sub

    Private Shared Function NormalizeReleaseNotes(value As String) As String
        If String.IsNullOrWhiteSpace(value) Then
            Return String.Empty
        End If

        Dim normalized = value.Replace("\r\n", Environment.NewLine).
                                Replace("\n", Environment.NewLine).
                                Replace("\r", Environment.NewLine)
        Return normalized
    End Function
End Class
