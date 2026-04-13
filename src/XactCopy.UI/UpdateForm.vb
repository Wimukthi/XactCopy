Imports System.Diagnostics
Imports System.IO
Imports System.IO.Compression
Imports System.Linq
Imports System.Text
Imports System.Threading
Imports System.Threading.Tasks
Imports XactCopy.Configuration

''' <summary>
''' Class UpdateForm.
''' </summary>
Friend Class UpdateForm
    Inherits Form

    Private ReadOnly _release As UpdateReleaseInfo
    Private ReadOnly _asset As UpdateAssetInfo
    Private ReadOnly _currentVersion As Version
    Private ReadOnly _settings As AppSettings

    Private ReadOnly _currentValueLabel As New Label()
    Private ReadOnly _latestValueLabel As New Label()
    Private ReadOnly _packageValueLabel As New Label()
    Private ReadOnly _notesRichTextBox As New RichTextBox()
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

    ''' <summary>
    ''' Initializes a new instance.
    ''' </summary>
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

        _notesRichTextBox.Dock = DockStyle.Fill
        _notesRichTextBox.ReadOnly = True
        _notesRichTextBox.ScrollBars = RichTextBoxScrollBars.Vertical
        _notesRichTextBox.WordWrap = True
        _notesRichTextBox.BorderStyle = BorderStyle.None
        _notesRichTextBox.DetectUrls = False
        root.Controls.Add(_notesRichTextBox, 0, 1)

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
        RenderMarkdownNotes(_notesRichTextBox, If(String.IsNullOrWhiteSpace(_release.Notes), "No release notes provided.", _release.Notes))

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

            _statusLabel.Text = "Verifying package integrity..."
            Dim expectedSha256 = Await UpdateService.ResolveAssetSha256Async(_settings, _release, _asset, _cancellation.Token)
            If String.IsNullOrWhiteSpace(expectedSha256) Then
                Throw New InvalidOperationException(
                    "Update package checksum is unavailable for this release. Publish a SHA-256 digest (asset digest or .sha256 sidecar) before installing.")
            End If

            Dim actualSha256 = Await UpdateService.ComputeFileSha256Async(_downloadPath, _cancellation.Token)
            If Not String.Equals(expectedSha256, actualSha256, StringComparison.OrdinalIgnoreCase) Then
                Throw New InvalidOperationException(
                    $"Update package integrity check failed (expected {expectedSha256}, got {actualSha256}).")
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

        Dim sourceDirectory = ResolvePayloadSourceDirectory(extractRoot, executableName)
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
        Dim scriptPath = Path.Combine(tempRoot, "apply-update.ps1")
        Dim processId = Process.GetCurrentProcess().Id
        Dim executablePath = Path.Combine(targetDirectory, executableName)
        Dim backupDirectory = Path.Combine(tempRoot, "backup")
        Dim logPath = Path.Combine(tempRoot, "update.log")
        Dim normalizedTarget = targetDirectory
        If normalizedTarget.Length > 3 Then
            normalizedTarget = normalizedTarget.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar)
        End If

        ' Escape single quotes so the embedded literal strings are safe.
        Dim q = Function(s As String) s.Replace("'", "''")
        Dim staleFolder = Path.Combine(normalizedTarget, "win-x64")
        Dim staleDll = Path.Combine(staleFolder, "XactCopy.dll")

        Dim script As New StringBuilder()
        script.AppendLine("$ErrorActionPreference = 'Continue'")
        script.AppendLine($"$TargetPid  = {processId}")
        script.AppendLine($"$Source     = '{q(sourceDirectory)}'")
        script.AppendLine($"$Target     = '{q(normalizedTarget)}'")
        script.AppendLine($"$Backup     = '{q(backupDirectory)}'")
        script.AppendLine($"$Log        = '{q(logPath)}'")
        script.AppendLine($"$TempRoot   = '{q(tempRoot)}'")
        script.AppendLine($"$Executable = '{q(executablePath)}'")
        script.AppendLine($"$StaleDll   = '{q(staleDll)}'")
        script.AppendLine($"$StaleDir   = '{q(staleFolder)}'")
        script.AppendLine()
        script.AppendLine("'XactCopy update log' | Out-File -LiteralPath $Log -Encoding UTF8")
        script.AppendLine()
        script.AppendLine("# Wait for XactCopy UI process to exit.")
        script.AppendLine("while ($null -ne (Get-Process -Id $TargetPid -ErrorAction SilentlyContinue)) {")
        script.AppendLine("    Start-Sleep -Milliseconds 500")
        script.AppendLine("}")
        script.AppendLine()
        script.AppendLine("# Back up current installation.")
        script.AppendLine("if (Test-Path -LiteralPath $Backup) { Remove-Item -LiteralPath $Backup -Recurse -Force }")
        script.AppendLine("$null = New-Item -ItemType Directory -Path $Backup -Force")
        script.AppendLine("robocopy $Target $Backup /E /COPY:DAT /R:2 /W:1 /NFL /NDL /NJH /NJS /NP >> $Log")
        script.AppendLine()
        script.AppendLine("# Copy new files to the install directory.")
        script.AppendLine("robocopy $Source $Target /E /COPY:DAT /R:2 /W:1 /NFL /NDL /NJH /NJS /NP >> $Log")
        script.AppendLine("$copyRc = $LASTEXITCODE")
        script.AppendLine()
        script.AppendLine("if ($copyRc -ge 8) {")
        script.AppendLine("    'Update copy failed. Restoring backup.' | Out-File -LiteralPath $Log -Append -Encoding UTF8")
        script.AppendLine("    robocopy $Backup $Target /E /COPY:DAT /R:2 /W:1 /NFL /NDL /NJH /NJS /NP >> $Log")
        script.AppendLine("} else {")
        script.AppendLine("    # Remove stale win-x64 subfolder from older zip packages.")
        script.AppendLine("    if (Test-Path -LiteralPath $StaleDll) {")
        script.AppendLine("        'Removing stale win-x64 folder.' | Out-File -LiteralPath $Log -Append -Encoding UTF8")
        script.AppendLine("        Remove-Item -LiteralPath $StaleDir -Recurse -Force -ErrorAction SilentlyContinue")
        script.AppendLine("    }")
        script.AppendLine("}")
        script.AppendLine()
        script.AppendLine("# Launch the updated XactCopy.")
        script.AppendLine("Start-Process -FilePath $Executable")
        script.AppendLine()
        script.AppendLine("# Deferred cleanup of the temp directory.")
        script.AppendLine("Start-Sleep -Seconds 10")
        script.AppendLine("Remove-Item -LiteralPath $TempRoot -Recurse -Force -ErrorAction SilentlyContinue")

        File.WriteAllText(scriptPath, script.ToString(), Encoding.UTF8)

        Dim psArgs = $"-NonInteractive -WindowStyle Hidden -ExecutionPolicy Bypass -File ""{scriptPath}"""
        Dim startInfo As New ProcessStartInfo("powershell.exe", psArgs) With {
            .UseShellExecute = False,
            .CreateNoWindow = True
        }
        Process.Start(startInfo)
    End Sub

    Private Shared Function ResolvePayloadSourceDirectory(extractRoot As String, executableName As String) As String
        If String.IsNullOrWhiteSpace(extractRoot) OrElse Not Directory.Exists(extractRoot) Then
            Return String.Empty
        End If

        Dim markerNames As New List(Of String)()
        If Not String.IsNullOrWhiteSpace(executableName) Then
            markerNames.Add(executableName)
        End If
        markerNames.Add("XactCopy.exe")
        markerNames.Add("XactCopy.dll")
        markerNames.Add("XactCopy.runtimeconfig.json")

        Dim candidates As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
        For Each marker In markerNames.Distinct(StringComparer.OrdinalIgnoreCase)
            For Each match In Directory.EnumerateFiles(extractRoot, marker, SearchOption.AllDirectories)
                Dim directoryPath = Path.GetDirectoryName(match)
                If String.IsNullOrWhiteSpace(directoryPath) Then
                    Continue For
                End If

                If IsLikelyPayloadDirectory(directoryPath, executableName) Then
                    candidates.Add(directoryPath)
                End If
            Next
        Next

        If candidates.Count = 0 Then
            Dim childDirectories = Directory.GetDirectories(extractRoot)
            Dim childFiles = Directory.GetFiles(extractRoot)
            If childDirectories.Length = 1 AndAlso childFiles.Length = 0 Then
                Return ResolvePayloadSourceDirectory(childDirectories(0), executableName)
            End If

            Return String.Empty
        End If

        Dim normalizedRoot = EnsureTrailingSeparator(Path.GetFullPath(extractRoot))
        Return candidates.
            OrderBy(Function(pathValue) GetRelativeDepth(normalizedRoot, pathValue)).
            ThenBy(Function(pathValue) pathValue.Length).
            First()
    End Function

    Private Shared Function IsLikelyPayloadDirectory(candidateDirectory As String, executableName As String) As Boolean
        If String.IsNullOrWhiteSpace(candidateDirectory) OrElse Not Directory.Exists(candidateDirectory) Then
            Return False
        End If

        Dim hasManagedHost = File.Exists(Path.Combine(candidateDirectory, "XactCopy.dll")) AndAlso
            File.Exists(Path.Combine(candidateDirectory, "XactCopy.runtimeconfig.json"))
        Dim hasExecutable = False

        If Not String.IsNullOrWhiteSpace(executableName) Then
            hasExecutable = File.Exists(Path.Combine(candidateDirectory, executableName))
        End If

        If Not hasExecutable Then
            hasExecutable = File.Exists(Path.Combine(candidateDirectory, "XactCopy.exe"))
        End If

        Return hasManagedHost OrElse hasExecutable
    End Function

    Private Shared Function EnsureTrailingSeparator(pathValue As String) As String
        If String.IsNullOrWhiteSpace(pathValue) Then
            Return String.Empty
        End If

        If pathValue.EndsWith(Path.DirectorySeparatorChar.ToString(), StringComparison.Ordinal) OrElse
            pathValue.EndsWith(Path.AltDirectorySeparatorChar.ToString(), StringComparison.Ordinal) Then
            Return pathValue
        End If

        Return pathValue & Path.DirectorySeparatorChar
    End Function

    Private Shared Function GetRelativeDepth(rootPath As String, candidatePath As String) As Integer
        Try
            Dim fullCandidate = Path.GetFullPath(candidatePath)
            If fullCandidate.StartsWith(rootPath, StringComparison.OrdinalIgnoreCase) Then
                Dim relative = fullCandidate.Substring(rootPath.Length)
                If relative.Length = 0 Then
                    Return 0
                End If

                Return relative.Split(New Char() {Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar}, StringSplitOptions.RemoveEmptyEntries).Length
            End If
        Catch
        End Try

        Return Integer.MaxValue
    End Function

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

    Protected Overrides Sub OnFormClosed(e As FormClosedEventArgs)
        MyBase.OnFormClosed(e)

        If _isApplying Then
            Return
        End If

        CleanupTempRootNoThrow()
    End Sub

    Private Sub CleanupTempRootNoThrow()
        If String.IsNullOrWhiteSpace(_tempRoot) Then
            Return
        End If

        Try
            If Directory.Exists(_tempRoot) Then
                Directory.Delete(_tempRoot, recursive:=True)
            End If
        Catch
            ' Ignore best-effort cleanup failures.
        End Try
    End Sub

    ''' <summary>
    ''' Renders GitHub-flavoured markdown release notes into a RichTextBox with basic formatting:
    ''' ## headings as bold, ### subheadings as bold, - bullet items with **bold** spans,
    ''' and --- separators as a blank line.
    ''' </summary>
    Private Shared Sub RenderMarkdownNotes(rtb As RichTextBox, notes As String)
        rtb.Clear()
        rtb.SuspendLayout()
        Try
            Dim baseFont = rtb.Font
            Dim boldFont As New Font(baseFont, FontStyle.Bold)
            Try
                Dim lines = notes.Replace(vbCrLf, vbLf).Replace(vbCr, vbLf).
                                  Split({vbLf}, StringSplitOptions.None)

                For Each rawLine In lines
                    Dim line = rawLine.TrimEnd()

                    If line.StartsWith("## ", StringComparison.Ordinal) Then
                        ' H2: bold heading with a blank line above (except at start).
                        If rtb.TextLength > 0 Then AppendRun(rtb, Environment.NewLine, baseFont)
                        AppendRun(rtb, line.Substring(3) & Environment.NewLine, boldFont)

                    ElseIf line.StartsWith("### ", StringComparison.Ordinal) Then
                        ' H3: bold subheading.
                        AppendRun(rtb, line.Substring(4) & Environment.NewLine, boldFont)

                    ElseIf line = "---" Then
                        ' Horizontal rule: blank line separator.
                        AppendRun(rtb, Environment.NewLine, baseFont)

                    ElseIf line.StartsWith("- ", StringComparison.Ordinal) Then
                        ' Bullet point — parse **bold** inline spans.
                        AppendRun(rtb, "  " & ChrW(&H2022) & " ", baseFont)
                        AppendInlineMarkdown(rtb, line.Substring(2), baseFont, boldFont)
                        AppendRun(rtb, Environment.NewLine, baseFont)

                    Else
                        ' Plain text.
                        AppendRun(rtb, line & Environment.NewLine, baseFont)
                    End If
                Next
            Finally
                boldFont.Dispose()
            End Try
        Finally
            rtb.ResumeLayout()
        End Try
        rtb.SelectionStart = 0
        rtb.ScrollToCaret()
    End Sub

    Private Shared Sub AppendRun(rtb As RichTextBox, text As String, font As Font)
        rtb.SelectionStart = rtb.TextLength
        rtb.SelectionLength = 0
        rtb.SelectionFont = font
        rtb.AppendText(text)
    End Sub

    Private Shared Sub AppendInlineMarkdown(rtb As RichTextBox, text As String, baseFont As Font, boldFont As Font)
        Dim remaining = text
        While remaining.Length > 0
            Dim boldOpen = remaining.IndexOf("**", StringComparison.Ordinal)
            If boldOpen < 0 Then
                AppendRun(rtb, remaining, baseFont)
                Return
            End If

            If boldOpen > 0 Then
                AppendRun(rtb, remaining.Substring(0, boldOpen), baseFont)
            End If

            Dim boldClose = remaining.IndexOf("**", boldOpen + 2, StringComparison.Ordinal)
            If boldClose < 0 Then
                ' Unmatched marker — treat the rest as plain text.
                AppendRun(rtb, remaining.Substring(boldOpen), baseFont)
                Return
            End If

            AppendRun(rtb, remaining.Substring(boldOpen + 2, boldClose - boldOpen - 2), boldFont)
            remaining = remaining.Substring(boldClose + 2)
        End While
    End Sub
End Class
