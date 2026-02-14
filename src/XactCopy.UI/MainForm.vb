Imports System.Diagnostics
Imports System.IO
Imports System.Collections.Generic
Imports System.Linq
Imports System.Threading
Imports System.Threading.Tasks

Imports XactCopy.Configuration
Imports XactCopy.Infrastructure
Imports XactCopy.Models
Imports XactCopy.Services

Public Class MainForm
    Inherits Form

    Private Const AppTitle As String = "XactCopy"
    Private Const LayoutWindowKey As String = "MainForm"
    Private Const TaskbarProgressScale As ULong = 1000UL
    Private Const MaxLogCharacters As Integer = 400000
    Private Const LogTrimTargetCharacters As Integer = 280000
    Private Const ProgressRenderIntervalMs As Integer = 50

    Private Shared ReadOnly AppDataRoot As String = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "XactCopy")

    Private Shared ReadOnly JournalsDirectory As String = Path.Combine(AppDataRoot, "journals")
    Private Shared ReadOnly CrashDirectory As String = Path.Combine(AppDataRoot, "crash")

    Private ReadOnly _sourceTextBox As New TextBox()
    Private ReadOnly _destinationTextBox As New TextBox()

    Private ReadOnly _browseSourceButton As New Button()
    Private ReadOnly _browseDestinationButton As New Button()
    Private ReadOnly _startButton As New Button()
    Private ReadOnly _pauseButton As New Button()
    Private ReadOnly _resumeButton As New Button()
    Private ReadOnly _cancelButton As New Button()

    Private ReadOnly _resumeCheckBox As New CheckBox()
    Private ReadOnly _salvageCheckBox As New CheckBox()
    Private ReadOnly _continueCheckBox As New CheckBox()
    Private ReadOnly _verifyCheckBox As New CheckBox()
    Private ReadOnly _adaptiveBufferCheckBox As New CheckBox()
    Private ReadOnly _waitForMediaCheckBox As New CheckBox()

    Private ReadOnly _bufferMbNumeric As New ThemedNumericUpDown()
    Private ReadOnly _retriesNumeric As New ThemedNumericUpDown()
    Private ReadOnly _timeoutSecondsNumeric As New ThemedNumericUpDown()

    Private ReadOnly _overallProgressBar As New ThemedProgressBar()
    Private ReadOnly _currentProgressBar As New ThemedProgressBar()

    Private ReadOnly _overallStatsLabel As New Label()
    Private ReadOnly _currentFileLabel As New Label()
    Private ReadOnly _bytesLabel As New Label()
    Private ReadOnly _throughputLabel As New Label()
    Private ReadOnly _etaLabel As New Label()
    Private ReadOnly _bufferUsageLabel As New Label()
    Private ReadOnly _rescueTelemetryLabel As New Label()
    Private ReadOnly _jobSummaryLabel As New Label()
    Private ReadOnly _journalLabel As New Label()

    Private ReadOnly _logTextBox As New TextBox()

    Private ReadOnly _menuStrip As New MenuStrip()
    Private ReadOnly _fileMenu As New ToolStripMenuItem("&File")
    Private ReadOnly _startMenuItem As New ToolStripMenuItem("&Start Copy")
    Private ReadOnly _pauseMenuItem As New ToolStripMenuItem("&Pause Copy")
    Private ReadOnly _resumeMenuItem As New ToolStripMenuItem("&Resume Copy")
    Private ReadOnly _cancelMenuItem As New ToolStripMenuItem("&Cancel Copy")
    Private ReadOnly _openJournalsMenuItem As New ToolStripMenuItem("Open &Journal Folder")
    Private ReadOnly _openCrashMenuItem As New ToolStripMenuItem("Open &Crash Folder")
    Private ReadOnly _exitMenuItem As New ToolStripMenuItem("E&xit")

    Private ReadOnly _toolsMenu As New ToolStripMenuItem("&Tools")
    Private ReadOnly _settingsMenuItem As New ToolStripMenuItem("&Settings...")
    Private ReadOnly _checkUpdatesMenuItem As New ToolStripMenuItem("Check for &Updates...")

    Private ReadOnly _jobsMenu As New ToolStripMenuItem("&Jobs")
    Private ReadOnly _saveCurrentAsJobMenuItem As New ToolStripMenuItem("&Save Current Options As Job...")
    Private ReadOnly _openJobManagerMenuItem As New ToolStripMenuItem("&Job Manager...")
    Private ReadOnly _resumeInterruptedMenuItem As New ToolStripMenuItem("&Resume Interrupted Job")
    Private ReadOnly _runNextQueuedJobMenuItem As New ToolStripMenuItem("Run &Next Queued Job")

    Private ReadOnly _helpMenu As New ToolStripMenuItem("&Help")
    Private ReadOnly _aboutMenuItem As New ToolStripMenuItem("&About XactCopy")
    Private ReadOnly _toolTip As New ToolTip() With {
        .ShowAlways = True,
        .InitialDelay = 350,
        .ReshowDelay = 120,
        .AutoPopDelay = 24000
    }

    Private ReadOnly _supervisor As WorkerSupervisor
    Private ReadOnly _settingsStore As New AppSettingsStore()
    Private ReadOnly _explorerIntegration As New ExplorerIntegrationService()
    Private ReadOnly _jobManager As JobManagerService
    Private ReadOnly _recoveryService As RecoveryService
    Private ReadOnly _launchOptions As LaunchOptions
    Private ReadOnly _taskbarProgress As New TaskbarProgressController()

    Private _settings As AppSettings
    Private _latestUpdateRelease As UpdateReleaseInfo
    Private _startupRecovery As RecoveryStartupInfo
    Private _interruptedRun As RecoveryActiveRun
    Private _activeRun As ManagedJobRun
    Private _activeRunJournalPath As String = String.Empty
    Private _isCheckingUpdates As Boolean
    Private _isRunning As Boolean
    Private _isPaused As Boolean
    Private _telemetryStartUtc As DateTimeOffset
    Private _telemetryLastSampleUtc As DateTimeOffset
    Private _telemetryLastBytesCopied As Long
    Private _smoothedBytesPerSecond As Double
    Private _bufferSamplesCount As Long
    Private _bufferSampleBytesTotal As Long
    Private _lastChunkUtilization As Double
    Private _lastRecoveryTouchUtc As DateTimeOffset = DateTimeOffset.MinValue
    Private _explorerSelectionSourceRoot As String = String.Empty
    Private ReadOnly _explorerSelectedRelativePaths As New List(Of String)()
    Private _suspendSourceTextChanged As Boolean
    Private _activeSalvageFillPattern As SalvageFillPattern = SalvageFillPattern.Zero
    Private _lastProgressSnapshot As CopyProgressSnapshot
    Private ReadOnly _progressDispatchLock As New Object()
    Private _pendingProgressSnapshot As CopyProgressSnapshot
    Private _progressDispatchQueued As Integer
    Private _lastProgressRenderTick As Long

    Public Sub New(Optional launchOptions As LaunchOptions = Nothing)
        Text = AppTitle
        StartPosition = FormStartPosition.CenterScreen
        MinimumSize = New Size(940, 640)
        Size = New Size(1100, 760)
        WindowIconHelper.Apply(Me)

        _supervisor = New WorkerSupervisor()
        AddHandler _supervisor.LogMessage, AddressOf Supervisor_LogMessage
        AddHandler _supervisor.ProgressChanged, AddressOf Supervisor_ProgressChanged
        AddHandler _supervisor.JobCompleted, AddressOf Supervisor_JobCompleted
        AddHandler _supervisor.WorkerStateChanged, AddressOf Supervisor_WorkerStateChanged
        AddHandler _supervisor.JobPauseStateChanged, AddressOf Supervisor_JobPauseStateChanged

        _launchOptions = If(launchOptions, New LaunchOptions())
        _jobManager = New JobManagerService()
        _recoveryService = New RecoveryService()
        _settings = _settingsStore.Load()
        _startupRecovery = _recoveryService.InitializeSession(_settings, _launchOptions)
        _interruptedRun = _startupRecovery.InterruptedRun

        If _startupRecovery.HasInterruptedRun Then
            _jobManager.MarkRunInterrupted(_startupRecovery.InterruptedRun.RunId, _startupRecovery.InterruptionReason)
            _jobManager.MarkAnyRunningRunsInterrupted(_startupRecovery.InterruptionReason)
        End If

        BuildUi()
        ConfigureToolTips()
        ApplyDefaultSettingsToUi()
        ApplyTheme()
        UpdateCheckMenuCaption()
        ResetCopyTelemetry()
        SetRunningState(isRunning:=False)

        AddHandler Shown, AddressOf MainForm_Shown
    End Sub

    Protected Overrides Sub OnFormClosing(e As FormClosingEventArgs)
        If _isRunning Then
            Dim response = MessageBox.Show(
                "A copy job is still running. Cancel and wait for it to stop before closing?",
                "XactCopy",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Warning)

            If response = DialogResult.Yes Then
                Dim cancelTask = _supervisor.CancelJobAsync("Application is closing.")
                Dim ignoredTask = cancelTask
                AppendLog("Cancellation requested. Close again after the status changes.")
            End If

            e.Cancel = True
            Return
        End If

        Try
            UiLayoutManager.CaptureWindow(Me, LayoutWindowKey)
        Catch
            ' Ignore layout persistence failures on shutdown.
        End Try

        Try
            _settingsStore.Save(_settings)
        Catch
            ' Ignore settings write failures on shutdown.
        End Try

        Try
            _supervisor.Dispose()
        Catch
        End Try

        Try
            _recoveryService.MarkCleanShutdown()
        Catch
        End Try

        MyBase.OnFormClosing(e)
    End Sub

    Private Sub BuildUi()
        ConfigureMenu()

        Dim rootLayout As New TableLayoutPanel() With {
            .Dock = DockStyle.Fill,
            .ColumnCount = 1,
            .RowCount = 6,
            .Padding = New Padding(10)
        }

        rootLayout.RowStyles.Add(New RowStyle(SizeType.AutoSize))
        rootLayout.RowStyles.Add(New RowStyle(SizeType.AutoSize))
        rootLayout.RowStyles.Add(New RowStyle(SizeType.AutoSize))
        rootLayout.RowStyles.Add(New RowStyle(SizeType.AutoSize))
        rootLayout.RowStyles.Add(New RowStyle(SizeType.AutoSize))
        rootLayout.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0F))

        _sourceTextBox.Dock = DockStyle.Fill
        _destinationTextBox.Dock = DockStyle.Fill
        AddHandler _sourceTextBox.TextChanged, AddressOf SourceTextBox_TextChanged

        _browseSourceButton.Text = "Browse..."
        _browseDestinationButton.Text = "Browse..."
        _browseSourceButton.AutoSize = True
        _browseDestinationButton.AutoSize = True

        AddHandler _browseSourceButton.Click, Sub(sender, args) BrowseFolder(_sourceTextBox)
        AddHandler _browseDestinationButton.Click, Sub(sender, args) BrowseFolder(_destinationTextBox)

        rootLayout.Controls.Add(BuildPathRow("Source", _sourceTextBox, _browseSourceButton), 0, 0)
        rootLayout.Controls.Add(BuildPathRow("Destination", _destinationTextBox, _browseDestinationButton), 0, 1)

        _resumeCheckBox.Text = "Resume from journal"
        _resumeCheckBox.AutoSize = True
        _resumeCheckBox.Checked = True

        Dim initialSalvageFillPattern = If(
            _settings Is Nothing,
            SalvageFillPattern.Zero,
            SettingsValueConverter.ToSalvageFillPattern(_settings.DefaultSalvageFillPattern))
        SetActiveSalvageFillPattern(initialSalvageFillPattern)
        _salvageCheckBox.AutoSize = True
        _salvageCheckBox.Checked = True

        _continueCheckBox.Text = "Continue on file errors"
        _continueCheckBox.AutoSize = True
        _continueCheckBox.Checked = True

        _verifyCheckBox.Text = "Verify copied files"
        _verifyCheckBox.AutoSize = True
        _verifyCheckBox.Checked = False

        _adaptiveBufferCheckBox.Text = "Adaptive Buffer (Auto)"
        _adaptiveBufferCheckBox.AutoSize = True
        _adaptiveBufferCheckBox.Checked = False

        _waitForMediaCheckBox.Text = "Wait Forever for Source/Destination"
        _waitForMediaCheckBox.AutoSize = True
        _waitForMediaCheckBox.Checked = False

        _bufferMbNumeric.Minimum = 1D
        _bufferMbNumeric.Maximum = 256D
        _bufferMbNumeric.Value = 4D
        _bufferMbNumeric.Width = 70

        _retriesNumeric.Minimum = 0D
        _retriesNumeric.Maximum = 1000D
        _retriesNumeric.Value = 12D
        _retriesNumeric.Width = 70

        _timeoutSecondsNumeric.Minimum = 1D
        _timeoutSecondsNumeric.Maximum = 3600D
        _timeoutSecondsNumeric.Value = 10D
        _timeoutSecondsNumeric.Width = 70

        rootLayout.Controls.Add(BuildOptionsRow(), 0, 2)

        _startButton.Text = "Start Copy"
        _startButton.AutoSize = True
        AddHandler _startButton.Click, AddressOf StartButton_Click

        _pauseButton.Text = "Pause"
        _pauseButton.AutoSize = True
        AddHandler _pauseButton.Click, AddressOf PauseButton_Click

        _resumeButton.Text = "Resume"
        _resumeButton.AutoSize = True
        AddHandler _resumeButton.Click, AddressOf ResumeButton_Click

        _cancelButton.Text = "Cancel"
        _cancelButton.AutoSize = True
        AddHandler _cancelButton.Click, AddressOf CancelButton_Click

        _jobSummaryLabel.AutoSize = True
        _jobSummaryLabel.Margin = New Padding(16, 8, 0, 0)
        _jobSummaryLabel.Text = "Status: Idle"

        Dim actionsPanel As New FlowLayoutPanel() With {
            .Dock = DockStyle.Fill,
            .AutoSize = True,
            .WrapContents = True,
            .FlowDirection = FlowDirection.LeftToRight
        }
        actionsPanel.Controls.Add(_startButton)
        actionsPanel.Controls.Add(_pauseButton)
        actionsPanel.Controls.Add(_resumeButton)
        actionsPanel.Controls.Add(_cancelButton)
        actionsPanel.Controls.Add(_jobSummaryLabel)
        rootLayout.Controls.Add(actionsPanel, 0, 3)

        _overallProgressBar.Dock = DockStyle.Fill
        _overallProgressBar.Minimum = 0
        _overallProgressBar.Maximum = 1000
        _overallProgressBar.Style = ProgressBarStyle.Continuous
        _overallProgressBar.ShowBorder = False

        _currentProgressBar.Dock = DockStyle.Fill
        _currentProgressBar.Minimum = 0
        _currentProgressBar.Maximum = 1000
        _currentProgressBar.Style = ProgressBarStyle.Continuous
        _currentProgressBar.ShowBorder = False

        _overallStatsLabel.AutoSize = True
        _overallStatsLabel.Text = "Files: 0/0  Failed: 0  Recovered: 0  Skipped: 0"

        _currentFileLabel.AutoSize = True
        _currentFileLabel.Text = "Current: -"
        _currentFileLabel.AutoEllipsis = True

        _bytesLabel.AutoSize = True
        _bytesLabel.Text = "Bytes: 0 B / 0 B"

        _throughputLabel.AutoSize = True
        _throughputLabel.Text = "Speed: 0 B/s (avg 0 B/s)"

        _etaLabel.AutoSize = True
        _etaLabel.Text = "ETA: -"

        _bufferUsageLabel.AutoSize = True
        _bufferUsageLabel.Text = "Buffer Use: last - / -  avg -"

        _rescueTelemetryLabel.AutoSize = True
        _rescueTelemetryLabel.Text = "Rescue: -"

        _journalLabel.AutoSize = True
        _journalLabel.Text = "Journal: -"

        Dim progressLayout As New TableLayoutPanel() With {
            .Dock = DockStyle.Fill,
            .ColumnCount = 1,
            .RowCount = 10,
            .AutoSize = True
        }
        progressLayout.RowStyles.Add(New RowStyle(SizeType.AutoSize))
        progressLayout.RowStyles.Add(New RowStyle(SizeType.AutoSize))
        progressLayout.RowStyles.Add(New RowStyle(SizeType.AutoSize))
        progressLayout.RowStyles.Add(New RowStyle(SizeType.AutoSize))
        progressLayout.RowStyles.Add(New RowStyle(SizeType.AutoSize))
        progressLayout.RowStyles.Add(New RowStyle(SizeType.AutoSize))
        progressLayout.RowStyles.Add(New RowStyle(SizeType.AutoSize))
        progressLayout.RowStyles.Add(New RowStyle(SizeType.AutoSize))
        progressLayout.RowStyles.Add(New RowStyle(SizeType.AutoSize))
        progressLayout.RowStyles.Add(New RowStyle(SizeType.AutoSize))
        progressLayout.Controls.Add(_overallStatsLabel, 0, 0)
        progressLayout.Controls.Add(_overallProgressBar, 0, 1)
        progressLayout.Controls.Add(_currentFileLabel, 0, 2)
        progressLayout.Controls.Add(_currentProgressBar, 0, 3)
        progressLayout.Controls.Add(_bytesLabel, 0, 4)
        progressLayout.Controls.Add(_throughputLabel, 0, 5)
        progressLayout.Controls.Add(_etaLabel, 0, 6)
        progressLayout.Controls.Add(_bufferUsageLabel, 0, 7)
        progressLayout.Controls.Add(_rescueTelemetryLabel, 0, 8)
        progressLayout.Controls.Add(_journalLabel, 0, 9)
        rootLayout.Controls.Add(progressLayout, 0, 4)

        _logTextBox.Dock = DockStyle.Fill
        _logTextBox.Multiline = True
        _logTextBox.ScrollBars = ScrollBars.Both
        _logTextBox.ReadOnly = True
        _logTextBox.WordWrap = False
        _logTextBox.Font = New Font("Consolas", 9.0F, FontStyle.Regular, GraphicsUnit.Point)
        rootLayout.Controls.Add(_logTextBox, 0, 5)

        Dim shellLayout As New TableLayoutPanel() With {
            .Dock = DockStyle.Fill,
            .ColumnCount = 1,
            .RowCount = 2
        }
        shellLayout.RowStyles.Add(New RowStyle(SizeType.AutoSize))
        shellLayout.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0F))
        shellLayout.Controls.Add(_menuStrip, 0, 0)
        shellLayout.Controls.Add(rootLayout, 0, 1)

        Controls.Add(shellLayout)
    End Sub

    Private Sub ConfigureMenu()
        _menuStrip.Dock = DockStyle.Fill
        _menuStrip.GripStyle = ToolStripGripStyle.Hidden

        _startMenuItem.ShortcutKeys = Keys.F5
        _pauseMenuItem.ShortcutKeys = Keys.Control Or Keys.P
        _resumeMenuItem.ShortcutKeys = Keys.Control Or Keys.R
        _cancelMenuItem.ShortcutKeys = Keys.Shift Or Keys.F5
        _exitMenuItem.ShortcutKeys = Keys.Alt Or Keys.F4
        _openJobManagerMenuItem.ShortcutKeys = Keys.Control Or Keys.J
        _startMenuItem.ToolTipText = "Start a copy run with the current options."
        _pauseMenuItem.ToolTipText = "Pause the active copy run."
        _resumeMenuItem.ToolTipText = "Resume a paused copy run."
        _cancelMenuItem.ToolTipText = "Cancel the active copy run."
        _openJournalsMenuItem.ToolTipText = "Open the journal folder used for resume/recovery data."
        _openCrashMenuItem.ToolTipText = "Open captured crash logs."
        _exitMenuItem.ToolTipText = "Close XactCopy."
        _settingsMenuItem.ToolTipText = "Open application settings."
        _checkUpdatesMenuItem.ToolTipText = "Check online for a newer release."
        _saveCurrentAsJobMenuItem.ToolTipText = "Save source, destination, and options as a reusable job."
        _openJobManagerMenuItem.ToolTipText = "Manage saved jobs, queue entries, and run history."
        _resumeInterruptedMenuItem.ToolTipText = "Resume the last interrupted run if recovery data exists."
        _runNextQueuedJobMenuItem.ToolTipText = "Start the next job in queue."
        _aboutMenuItem.ToolTipText = "Show version and license information."

        AddHandler _startMenuItem.Click, AddressOf StartMenuItem_Click
        AddHandler _pauseMenuItem.Click, AddressOf PauseMenuItem_Click
        AddHandler _resumeMenuItem.Click, AddressOf ResumeMenuItem_Click
        AddHandler _cancelMenuItem.Click, AddressOf CancelMenuItem_Click
        AddHandler _openJournalsMenuItem.Click, AddressOf OpenJournalsMenuItem_Click
        AddHandler _openCrashMenuItem.Click, AddressOf OpenCrashMenuItem_Click
        AddHandler _exitMenuItem.Click, Sub(sender, e) Close()

        AddHandler _settingsMenuItem.Click, AddressOf SettingsMenuItem_Click
        AddHandler _checkUpdatesMenuItem.Click, AddressOf CheckUpdatesMenuItem_Click
        AddHandler _saveCurrentAsJobMenuItem.Click, AddressOf SaveCurrentAsJobMenuItem_Click
        AddHandler _openJobManagerMenuItem.Click, AddressOf OpenJobManagerMenuItem_Click
        AddHandler _resumeInterruptedMenuItem.Click, AddressOf ResumeInterruptedMenuItem_Click
        AddHandler _runNextQueuedJobMenuItem.Click, AddressOf RunNextQueuedJobMenuItem_Click

        AddHandler _aboutMenuItem.Click, AddressOf AboutMenuItem_Click

        _fileMenu.DropDownItems.Add(_startMenuItem)
        _fileMenu.DropDownItems.Add(_pauseMenuItem)
        _fileMenu.DropDownItems.Add(_resumeMenuItem)
        _fileMenu.DropDownItems.Add(_cancelMenuItem)
        _fileMenu.DropDownItems.Add(New ToolStripSeparator())
        _fileMenu.DropDownItems.Add(_openJournalsMenuItem)
        _fileMenu.DropDownItems.Add(_openCrashMenuItem)
        _fileMenu.DropDownItems.Add(New ToolStripSeparator())
        _fileMenu.DropDownItems.Add(_exitMenuItem)

        _toolsMenu.DropDownItems.Add(_settingsMenuItem)
        _toolsMenu.DropDownItems.Add(_checkUpdatesMenuItem)

        _jobsMenu.DropDownItems.Add(_saveCurrentAsJobMenuItem)
        _jobsMenu.DropDownItems.Add(_openJobManagerMenuItem)
        _jobsMenu.DropDownItems.Add(New ToolStripSeparator())
        _jobsMenu.DropDownItems.Add(_resumeInterruptedMenuItem)
        _jobsMenu.DropDownItems.Add(_runNextQueuedJobMenuItem)

        _helpMenu.DropDownItems.Add(_aboutMenuItem)

        _menuStrip.Items.Add(_fileMenu)
        _menuStrip.Items.Add(_jobsMenu)
        _menuStrip.Items.Add(_toolsMenu)
        _menuStrip.Items.Add(_helpMenu)

        MainMenuStrip = _menuStrip
    End Sub

    Private Function BuildPathRow(labelText As String, targetTextBox As TextBox, browseButton As Button) As Control
        Dim panel As New TableLayoutPanel() With {
            .Dock = DockStyle.Fill,
            .ColumnCount = 3,
            .AutoSize = True
        }

        panel.ColumnStyles.Add(New ColumnStyle(SizeType.AutoSize))
        panel.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0F))
        panel.ColumnStyles.Add(New ColumnStyle(SizeType.AutoSize))

        Dim labelControl As New Label() With {
            .Text = labelText,
            .AutoSize = True,
            .Anchor = AnchorStyles.Left,
            .Margin = New Padding(0, 8, 10, 0)
        }

        targetTextBox.Margin = New Padding(0, 4, 8, 4)
        browseButton.Margin = New Padding(0, 3, 0, 3)

        panel.Controls.Add(labelControl, 0, 0)
        panel.Controls.Add(targetTextBox, 1, 0)
        panel.Controls.Add(browseButton, 2, 0)

        Return panel
    End Function

    Private Function BuildOptionsRow() As Control
        Dim optionsPanel As New FlowLayoutPanel() With {
            .Dock = DockStyle.Fill,
            .AutoSize = True,
            .AutoSizeMode = AutoSizeMode.GrowAndShrink,
            .WrapContents = True,
            .FlowDirection = FlowDirection.LeftToRight,
            .Padding = New Padding(0),
            .Margin = New Padding(0)
        }

        For Each checkBox In New CheckBox() {_resumeCheckBox, _salvageCheckBox, _continueCheckBox, _verifyCheckBox, _adaptiveBufferCheckBox, _waitForMediaCheckBox}
            checkBox.Anchor = AnchorStyles.Left
            checkBox.Margin = New Padding(0, 4, 12, 0)
        Next

        _bufferMbNumeric.Anchor = AnchorStyles.Left
        _retriesNumeric.Anchor = AnchorStyles.Left
        _timeoutSecondsNumeric.Anchor = AnchorStyles.Left
        _bufferMbNumeric.Margin = New Padding(0, 2, 10, 2)
        _retriesNumeric.Margin = New Padding(0, 2, 10, 2)
        _timeoutSecondsNumeric.Margin = New Padding(0, 2, 0, 2)

        optionsPanel.Controls.Add(_resumeCheckBox)
        optionsPanel.Controls.Add(_salvageCheckBox)
        optionsPanel.Controls.Add(_continueCheckBox)
        optionsPanel.Controls.Add(_verifyCheckBox)
        optionsPanel.Controls.Add(_adaptiveBufferCheckBox)
        optionsPanel.Controls.Add(_waitForMediaCheckBox)

        optionsPanel.Controls.Add(New Label() With {.Text = "Buffer MB (max):", .AutoSize = True, .Margin = New Padding(12, 7, 4, 0)})
        optionsPanel.Controls.Add(_bufferMbNumeric)

        optionsPanel.Controls.Add(New Label() With {.Text = "Retries:", .AutoSize = True, .Margin = New Padding(12, 7, 4, 0)})
        optionsPanel.Controls.Add(_retriesNumeric)

        optionsPanel.Controls.Add(New Label() With {.Text = "Timeout sec:", .AutoSize = True, .Margin = New Padding(12, 7, 4, 0)})
        optionsPanel.Controls.Add(_timeoutSecondsNumeric)

        Return optionsPanel
    End Function

    Private Sub ConfigureToolTips()
        _toolTip.SetToolTip(_sourceTextBox, "Source folder to copy from.")
        _toolTip.SetToolTip(_destinationTextBox, "Destination folder to copy into.")
        _toolTip.SetToolTip(_browseSourceButton, "Browse and select the source folder.")
        _toolTip.SetToolTip(_browseDestinationButton, "Browse and select the destination folder.")

        _toolTip.SetToolTip(_resumeCheckBox, "Use journal data to resume partially copied files.")
        _toolTip.SetToolTip(_salvageCheckBox, "Attempt to recover unreadable source regions using the configured fill pattern.")
        _toolTip.SetToolTip(_continueCheckBox, "Continue with remaining files when a file copy fails.")
        _toolTip.SetToolTip(_verifyCheckBox, "Verify copied output after each file.")
        _toolTip.SetToolTip(_adaptiveBufferCheckBox, "Dynamically adjust transfer buffer size by file profile and I/O behavior.")
        _toolTip.SetToolTip(_waitForMediaCheckBox, "Wait indefinitely if source or destination becomes unavailable.")

        _toolTip.SetToolTip(_bufferMbNumeric, "Maximum transfer buffer in MB (1-256).")
        _toolTip.SetToolTip(_retriesNumeric, "Retry attempts per read/write error (0-1000).")
        _toolTip.SetToolTip(_timeoutSecondsNumeric, "I/O timeout per operation in seconds (1-3600).")

        _toolTip.SetToolTip(_startButton, "Start copying with current options.")
        _toolTip.SetToolTip(_pauseButton, "Pause the active copy run.")
        _toolTip.SetToolTip(_resumeButton, "Resume a paused copy run.")
        _toolTip.SetToolTip(_cancelButton, "Cancel the active copy run.")

        _toolTip.SetToolTip(_overallProgressBar, "Overall run progress.")
        _toolTip.SetToolTip(_currentProgressBar, "Progress of the current file.")
        _toolTip.SetToolTip(_jobSummaryLabel, "Current run state.")
        _toolTip.SetToolTip(_overallStatsLabel, "File counters for completed, failed, recovered, and skipped files.")
        _toolTip.SetToolTip(_currentFileLabel, "Currently processed file path.")
        _toolTip.SetToolTip(_bytesLabel, "Transferred bytes versus total bytes.")
        _toolTip.SetToolTip(_throughputLabel, "Instantaneous and smoothed transfer speed.")
        _toolTip.SetToolTip(_etaLabel, "Estimated remaining time based on current average speed.")
        _toolTip.SetToolTip(_bufferUsageLabel, "Recent and average buffer utilization.")
        _toolTip.SetToolTip(_rescueTelemetryLabel, "AegisRescueCore pass status, unreadable region count, and remaining unrecovered bytes.")
        _toolTip.SetToolTip(_journalLabel, "Active journal file used for resume/recovery.")
        _toolTip.SetToolTip(_logTextBox, "Event-focused operation log (errors, retries, rescue pass summaries). Horizontal and vertical scrolling are enabled.")
    End Sub

    Private Sub SourceTextBox_TextChanged(sender As Object, e As EventArgs)
        If _suspendSourceTextChanged Then
            Return
        End If

        If _explorerSelectedRelativePaths.Count = 0 Then
            Return
        End If

        Dim normalizedSource = NormalizePathKey(_sourceTextBox.Text)
        If String.IsNullOrWhiteSpace(normalizedSource) Then
            ClearExplorerSelectionContext(announce:=False)
            Return
        End If

        If Not String.Equals(normalizedSource, _explorerSelectionSourceRoot, StringComparison.OrdinalIgnoreCase) Then
            ClearExplorerSelectionContext(announce:=False)
        End If
    End Sub

    Private Async Sub StartMenuItem_Click(sender As Object, e As EventArgs)
        If _isRunning Then
            Return
        End If

        Await StartCopyAsync()
    End Sub

    Private Async Sub PauseMenuItem_Click(sender As Object, e As EventArgs)
        Await PauseCopyAsync()
    End Sub

    Private Async Sub ResumeMenuItem_Click(sender As Object, e As EventArgs)
        Await ResumeCopyAsync()
    End Sub

    Private Async Sub CancelMenuItem_Click(sender As Object, e As EventArgs)
        If Not _isRunning Then
            Return
        End If

        Await _supervisor.CancelJobAsync("User requested cancellation.")
    End Sub

    Private Sub OpenJournalsMenuItem_Click(sender As Object, e As EventArgs)
        OpenFolder(JournalsDirectory)
    End Sub

    Private Sub OpenCrashMenuItem_Click(sender As Object, e As EventArgs)
        OpenFolder(CrashDirectory)
    End Sub

    Private Sub SettingsMenuItem_Click(sender As Object, e As EventArgs)
        ShowSettingsDialog()
    End Sub

    Private Async Sub CheckUpdatesMenuItem_Click(sender As Object, e As EventArgs)
        If _latestUpdateRelease IsNot Nothing Then
            Dim currentVersion = UpdateService.ParseVersionSafe(Application.ProductVersion)
            ShowUpdateDialog(_latestUpdateRelease, currentVersion)
            Return
        End If

        Await CheckForUpdatesAsync(showUpToDateDialog:=True)
    End Sub

    Private Sub SaveCurrentAsJobMenuItem_Click(sender As Object, e As EventArgs)
        SaveCurrentOptionsAsJob()
    End Sub

    Private Async Sub OpenJobManagerMenuItem_Click(sender As Object, e As EventArgs)
        Await OpenJobManagerAsync()
    End Sub

    Private Async Sub ResumeInterruptedMenuItem_Click(sender As Object, e As EventArgs)
        Await ResumeInterruptedRunAsync()
    End Sub

    Private Async Sub RunNextQueuedJobMenuItem_Click(sender As Object, e As EventArgs)
        Await RunNextQueuedJobAsync(manualTrigger:=True)
    End Sub

    Private Sub AboutMenuItem_Click(sender As Object, e As EventArgs)
        Dim aboutText As String =
            $"XactCopy{Environment.NewLine}" &
            $"Version: {Application.ProductVersion}{Environment.NewLine}" &
            $"License: GNU GPL v3.0{Environment.NewLine}" &
            $"Resilient copy engine for unstable media with journaled recovery."

        MessageBox.Show(Me, aboutText, "About XactCopy", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub

    Private Sub ShowSettingsDialog()
        Using dialog As New SettingsForm(_settings)
            ThemeManager.ApplyTheme(dialog, ThemeSettings.GetPreferredColorMode(_settings))

            If dialog.ShowDialog(Me) <> DialogResult.OK Then
                Return
            End If

            Dim requiresRestart = dialog.RequiresApplicationRestart
            Dim restartReasonText = dialog.RestartReasonText

            _settings = dialog.ResultSettings
            _settingsStore.Save(_settings)
            SyncExplorerIntegration(announceResult:=True)

            ApplyDefaultSettingsToUi()
            ApplyTheme()
            UpdateCheckMenuCaption()
            UpdateRecoveryMenuState()
            AppendLog("Settings saved.")

            If requiresRestart Then
                OfferRestartAfterSettingsChange(restartReasonText)
            End If
        End Using
    End Sub

    Private Sub OfferRestartAfterSettingsChange(reasonText As String)
        Dim reasonDetails = If(String.IsNullOrWhiteSpace(reasonText), "- Startup behavior settings", reasonText)
        Dim message = "Some settings apply only after restarting XactCopy:" & Environment.NewLine &
            reasonDetails

        If _isRunning Then
            MessageBox.Show(
                Me,
                message & Environment.NewLine & Environment.NewLine &
                "Restart after the current copy run completes.",
                "XactCopy",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information)
            AppendLog("Restart-required settings saved. Restart XactCopy after the current run to apply them fully.")
            Return
        End If

        Dim response = MessageBox.Show(
            Me,
            message & Environment.NewLine & Environment.NewLine & "Restart now?",
            "XactCopy",
            MessageBoxButtons.YesNo,
            MessageBoxIcon.Question)

        If response = DialogResult.Yes Then
            RestartApplication()
        End If
    End Sub

    Private Sub RestartApplication()
        Try
            Dim startInfo As New ProcessStartInfo(Application.ExecutablePath) With {
                .UseShellExecute = True,
                .WorkingDirectory = AppContext.BaseDirectory
            }
            Process.Start(startInfo)
            Close()
        Catch ex As Exception
            AppendLog($"Restart failed: {ex.Message}")
            MessageBox.Show(Me, $"Unable to restart automatically: {ex.Message}", "XactCopy", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End Try
    End Sub

    Private Sub SaveCurrentOptionsAsJob()
        Dim options = BuildOptions()
        If options Is Nothing Then
            Return
        End If

        Dim defaultName = BuildDefaultJobName(options)
        Dim enteredName As String = defaultName
        Dim dialogResult = TextPromptDialog.ShowPrompt(
            Me,
            "Save Job",
            "Enter a name for this saved job.",
            defaultName,
            enteredName)

        If dialogResult <> DialogResult.OK Then
            Return
        End If

        If String.IsNullOrWhiteSpace(enteredName) Then
            MessageBox.Show(Me, "Job name cannot be empty.", "XactCopy", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        Dim savedJob = _jobManager.SaveJob(enteredName, options)
        AppendLog($"Saved job '{savedJob.Name}' ({savedJob.JobId}).")
    End Sub

    Private Async Function OpenJobManagerAsync() As Task
        Using dialog As New JobManagerForm(_jobManager)
            ThemeManager.ApplyTheme(dialog, ThemeSettings.GetPreferredColorMode(_settings))
            Dim result = dialog.ShowDialog(Me)
            If result <> DialogResult.OK Then
                UpdateRecoveryMenuState()
                Return
            End If

            Dim selectedJobId = dialog.RequestedRunJobId
            If String.IsNullOrWhiteSpace(selectedJobId) Then
                UpdateRecoveryMenuState()
                Return
            End If

            Await RunJobByIdAsync(selectedJobId, "job-manager")
        End Using
    End Function

    Private Async Function RunJobByIdAsync(jobId As String, trigger As String) As Task
        If _isRunning Then
            MessageBox.Show(Me, "A copy run is already in progress.", "XactCopy", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return
        End If

        Dim job = _jobManager.GetJobById(jobId)
        If job Is Nothing Then
            MessageBox.Show(Me, "Selected job no longer exists.", "XactCopy", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        Dim normalized = NormalizeAndValidateOptions(job.Options, showDialogs:=True)
        If normalized Is Nothing Then
            Return
        End If

        Dim run = _jobManager.CreateRunForJob(job.JobId, trigger)
        Await StartCopyAsync(normalized, run, clearLog:=True)
    End Function

    Private Async Function RunNextQueuedJobAsync(manualTrigger As Boolean) As Task
        If _isRunning Then
            Return
        End If

        Dim showEmptyQueueDialog = manualTrigger
        Do
            Dim nextJob As ManagedJob = Nothing
            If Not _jobManager.TryDequeueNextJob(nextJob) OrElse nextJob Is Nothing Then
                If showEmptyQueueDialog Then
                    MessageBox.Show(Me, "No queued jobs are waiting to run.", "XactCopy", MessageBoxButtons.OK, MessageBoxIcon.Information)
                End If
                Return
            End If

            Dim normalized = NormalizeAndValidateOptions(nextJob.Options, showDialogs:=manualTrigger)
            Dim queuedRun = _jobManager.CreateRunForJob(nextJob.JobId, If(manualTrigger, "queued-manual", "queued-auto"))

            If normalized Is Nothing Then
                Dim failure As New CopyJobResult() With {
                    .Succeeded = False,
                    .Cancelled = False,
                    .JournalPath = ComputeJournalPath(nextJob.Options),
                    .ErrorMessage = "Queued job skipped due invalid configuration."
                }
                _jobManager.MarkRunCompleted(queuedRun.RunId, failure)
                AppendLog($"Skipped queued job '{nextJob.Name}' because configuration is invalid.")
                showEmptyQueueDialog = False
                Continue Do
            End If

            AppendLog($"Dequeued job '{nextJob.Name}'.")
            Await StartCopyAsync(normalized, queuedRun, clearLog:=True)
            Return
        Loop
    End Function

    Private Async Function ResumeInterruptedRunAsync() As Task
        If _isRunning Then
            MessageBox.Show(Me, "A copy run is already in progress.", "XactCopy", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return
        End If

        If _interruptedRun Is Nothing Then
            _interruptedRun = _recoveryService.GetPendingInterruptedRun()
        End If

        If _interruptedRun Is Nothing Then
            MessageBox.Show(Me, "No interrupted run is currently pending.", "XactCopy", MessageBoxButtons.OK, MessageBoxIcon.Information)
            UpdateRecoveryMenuState()
            Return
        End If

        Dim resumeOptions = NormalizeAndValidateOptions(_interruptedRun.Options, showDialogs:=True)
        If resumeOptions Is Nothing Then
            Return
        End If

        resumeOptions.ResumeFromJournal = True
        Dim runDisplayName = If(String.IsNullOrWhiteSpace(_interruptedRun.JobName), "Recovered Copy Session", _interruptedRun.JobName & " (Recovered)")
        Dim run = _jobManager.CreateAdHocRun(resumeOptions, runDisplayName, "resume-interrupted")
        Await StartCopyAsync(resumeOptions, run, clearLog:=True)
    End Function

    Private Async Function ShowRecoveryPromptAsync() As Task
        If _interruptedRun Is Nothing Then
            _interruptedRun = _recoveryService.GetPendingInterruptedRun()
        End If

        If _interruptedRun Is Nothing Then
            UpdateRecoveryMenuState()
            Return
        End If

        Dim reason = If(_startupRecovery?.InterruptionReason, "The previous copy session ended unexpectedly.")

        Using dialog As New RecoveryPromptForm(_interruptedRun, reason)
            ThemeManager.ApplyTheme(dialog, ThemeSettings.GetPreferredColorMode(_settings))
            Dim result = dialog.ShowDialog(Me)

            If result = DialogResult.OK Then
                Select Case dialog.SelectedAction
                    Case RecoveryPromptAction.ResumeNow
                        Await ResumeInterruptedRunAsync()
                    Case RecoveryPromptAction.OpenJobManager
                        _recoveryService.MarkResumePromptDeferred(_settings, dialog.SuppressPromptForThisRun)
                        Await OpenJobManagerAsync()
                    Case Else
                        _recoveryService.MarkResumePromptDeferred(_settings, dialog.SuppressPromptForThisRun)
                End Select
            Else
                _recoveryService.MarkResumePromptDeferred(_settings, dialog.SuppressPromptForThisRun)
            End If
        End Using

        _interruptedRun = _recoveryService.GetPendingInterruptedRun()
        UpdateRecoveryMenuState()
    End Function

    Private Sub UpdateRecoveryMenuState()
        If _interruptedRun Is Nothing Then
            _interruptedRun = _recoveryService.GetPendingInterruptedRun()
        End If

        _resumeInterruptedMenuItem.Enabled = (Not _isRunning) AndAlso _interruptedRun IsNot Nothing
    End Sub

    Private Shared Function BuildDefaultJobName(options As CopyJobOptions) As String
        Dim sourceName = Path.GetFileName(options.SourceRoot.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar))
        Dim destinationName = Path.GetFileName(options.DestinationRoot.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar))

        If String.IsNullOrWhiteSpace(sourceName) Then
            sourceName = options.SourceRoot
        End If

        If String.IsNullOrWhiteSpace(destinationName) Then
            destinationName = options.DestinationRoot
        End If

        Return $"{sourceName} -> {destinationName}"
    End Function

    Private Shared Function ComputeJournalPath(options As CopyJobOptions) As String
        If options Is Nothing Then
            Return String.Empty
        End If

        If String.IsNullOrWhiteSpace(options.SourceRoot) OrElse String.IsNullOrWhiteSpace(options.DestinationRoot) Then
            Return String.Empty
        End If

        Dim jobId = JobJournalStore.BuildJobId(options.SourceRoot, options.DestinationRoot)
        Return JobJournalStore.GetDefaultJournalPath(jobId)
    End Function

    Private Async Sub MainForm_Shown(sender As Object, e As EventArgs)
        UiLayoutManager.ApplyWindow(Me, LayoutWindowKey)

        If _settings Is Nothing Then
            Return
        End If

        SyncExplorerIntegration(announceResult:=False)
        ApplyExplorerLaunchOptions()

        If _launchOptions IsNot Nothing AndAlso _launchOptions.IsRecoveryAutostart Then
            AppendLog("Launched automatically after logon for recovery.")
        End If

        UpdateRecoveryMenuState()

        If _startupRecovery IsNot Nothing AndAlso _startupRecovery.HasInterruptedRun Then
            AppendLog("Detected an interrupted run from the previous session.")
            If _startupRecovery.ShouldAutoResume Then
                AppendLog("Auto-resume is enabled. Attempting to resume now.")
                Await ResumeInterruptedRunAsync()
            ElseIf _startupRecovery.ShouldPrompt Then
                Await ShowRecoveryPromptAsync()
            Else
                AppendLog("Interrupted run is pending. Use Jobs > Resume Interrupted Job to continue.")
            End If
        End If

        If Not _isRunning Then
            If _settings.CheckUpdatesOnLaunch Then
                Await CheckForUpdatesAsync(showUpToDateDialog:=False)
            Else
                AppendLog("Silent update check skipped (disabled in settings).")
            End If
        End If

        If Not _isRunning AndAlso _settings.AutoRunQueuedJobsOnStartup Then
            AppendLog("Startup queue mode enabled. Looking for queued jobs.")
            Await RunNextQueuedJobAsync(manualTrigger:=False)
        End If
    End Sub

    Private Async Function CheckForUpdatesAsync(showUpToDateDialog As Boolean) As Task
        If _isCheckingUpdates Then
            Return
        End If

        If String.IsNullOrWhiteSpace(_settings.UpdateReleaseUrl) Then
            _latestUpdateRelease = Nothing
            UpdateCheckMenuCaption()
            AppendLog("Update check skipped: release URL is not configured.")
            If showUpToDateDialog Then
                MessageBox.Show(Me,
                                "Open Settings and set an update release URL first.",
                                "XactCopy Updates",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Information)
            End If
            Return
        End If

        _isCheckingUpdates = True
        _checkUpdatesMenuItem.Enabled = False
        UpdateCheckMenuCaption()

        Try
            Dim release As UpdateReleaseInfo = Await UpdateService.GetLatestReleaseAsync(_settings, CancellationToken.None)
            Dim currentVersion As Version = UpdateService.ParseVersionSafe(Application.ProductVersion)

            If Not UpdateService.IsUpdateAvailable(currentVersion, release.Version) Then
                _latestUpdateRelease = Nothing
                AppendLog("Update check complete: already up to date.")
                If showUpToDateDialog Then
                    MessageBox.Show(Me,
                                    "You're already running the latest version.",
                                    "XactCopy Updates",
                                    MessageBoxButtons.OK,
                                    MessageBoxIcon.Information)
                End If
                Return
            End If

            _latestUpdateRelease = release
            Dim latestLabel As String = If(String.IsNullOrWhiteSpace(release.TagName), release.Version.ToString(), release.TagName)
            AppendLog($"Update available: {latestLabel}")

            If showUpToDateDialog Then
                ShowUpdateDialog(release, currentVersion)
            End If
        Catch ex As Exception
            _latestUpdateRelease = Nothing
            AppendLog($"Update check failed: {ex.Message}")
            If showUpToDateDialog Then
                MessageBox.Show(Me,
                                "Update check failed: " & ex.Message,
                                "XactCopy Updates",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Warning)
            End If
        Finally
            _isCheckingUpdates = False
            _checkUpdatesMenuItem.Enabled = True
            UpdateCheckMenuCaption()
        End Try
    End Function

    Private Sub ShowUpdateDialog(release As UpdateReleaseInfo, currentVersion As Version)
        If release Is Nothing Then
            Return
        End If

        If _isRunning Then
            MessageBox.Show(Me,
                            "Finish or cancel the active copy job before installing updates.",
                            "XactCopy Updates",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Information)
            Return
        End If

        Try
            Using dialog As New UpdateForm(release, currentVersion, _settings)
                ThemeManager.ApplyTheme(dialog, ThemeSettings.GetPreferredColorMode(_settings))
                dialog.ShowDialog(Me)
            End Using
        Catch ex As Exception
            AppendLog($"Update dialog failed: {ex.Message}")
            MessageBox.Show(Me,
                            "Unable to open updater: " & ex.Message,
                            "XactCopy Updates",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Warning)
        End Try
    End Sub

    Private Sub UpdateCheckMenuCaption()
        If _isCheckingUpdates Then
            _checkUpdatesMenuItem.Text = "Checking for &Updates..."
            Return
        End If

        _checkUpdatesMenuItem.Text = If(_latestUpdateRelease Is Nothing, "Check for &Updates...", "&Update Available...")
    End Sub

    Private Shared Sub OpenUrl(url As String)
        If String.IsNullOrWhiteSpace(url) Then
            Return
        End If

        Try
            Process.Start(New ProcessStartInfo(url) With {.UseShellExecute = True})
        Catch
            ' Ignore browser launch failures.
        End Try
    End Sub

    Private Shared Sub OpenFolder(path As String)
        Try
            Directory.CreateDirectory(path)
            Process.Start(New ProcessStartInfo(path) With {.UseShellExecute = True})
        Catch
            ' Ignore folder launch failures.
        End Try
    End Sub

    Private Sub SyncExplorerIntegration(announceResult As Boolean)
        Try
            _explorerIntegration.Sync(_settings.EnableExplorerContextMenu, Application.ExecutablePath)
            If announceResult Then
                AppendLog($"Explorer context menu integration {If(_settings.EnableExplorerContextMenu, "enabled", "disabled")}.")
            End If
        Catch ex As Exception
            AppendLog($"Explorer integration update failed: {ex.Message}")
        End Try
    End Sub

    Private Sub ApplyExplorerLaunchOptions()
        If _launchOptions Is Nothing Then
            Return
        End If

        If Not String.IsNullOrWhiteSpace(_launchOptions.ExplorerFolderPath) Then
            Dim resolvedFolder As String = String.Empty
            If Not TryResolveExplorerPath(_launchOptions.ExplorerFolderPath, resolvedFolder) Then
                AppendLog($"Explorer folder path is invalid: {_launchOptions.ExplorerFolderPath}")
                Return
            End If

            If Not Directory.Exists(resolvedFolder) Then
                AppendLog($"Explorer folder path was not found: {resolvedFolder}")
                Return
            End If

            ApplyExplorerSourceFolder(resolvedFolder, "Explorer background folder")
            Return
        End If

        If _launchOptions.ExplorerSourcePaths Is Nothing OrElse _launchOptions.ExplorerSourcePaths.Count = 0 Then
            Return
        End If

        Dim resolvedSelections As New List(Of String)()
        For Each rawPath In _launchOptions.ExplorerSourcePaths
            Dim resolvedPath As String = String.Empty
            If Not TryResolveExplorerPath(rawPath, resolvedPath) Then
                AppendLog($"Explorer selected an invalid path: {rawPath}")
                Continue For
            End If

            If Not File.Exists(resolvedPath) AndAlso Not Directory.Exists(resolvedPath) Then
                AppendLog($"Explorer selection was not found: {resolvedPath}")
                Continue For
            End If

            If resolvedSelections.Any(Function(existing) String.Equals(existing, resolvedPath, StringComparison.OrdinalIgnoreCase)) Then
                Continue For
            End If

            resolvedSelections.Add(resolvedPath)
        Next

        If resolvedSelections.Count = 0 Then
            Return
        End If

        Dim selectionMode = SettingsValueConverter.ToExplorerSelectionMode(_settings.ExplorerSelectionMode)
        If selectionMode = ExplorerSelectionMode.SourceFolder Then
            Dim firstSelection = resolvedSelections(0)
            Dim sourceFolder = If(File.Exists(firstSelection), Path.GetDirectoryName(firstSelection), firstSelection)
            If String.IsNullOrWhiteSpace(sourceFolder) Then
                AppendLog($"Explorer source path has no parent directory: {firstSelection}")
                Return
            End If

            ApplyExplorerSourceFolder(sourceFolder, "Explorer selection")
            Return
        End If

        ApplyExplorerSelectedItems(resolvedSelections)
    End Sub

    Private Sub ApplyExplorerSourceFolder(sourceFolder As String, context As String)
        ClearExplorerSelectionContext(announce:=False)
        If String.IsNullOrWhiteSpace(sourceFolder) Then
            Return
        End If

        _suspendSourceTextChanged = True
        _sourceTextBox.Text = sourceFolder
        _suspendSourceTextChanged = False
        AppendLog($"Source set from {context}: {sourceFolder}")
    End Sub

    Private Sub ApplyExplorerSelectedItems(resolvedSelections As IReadOnlyList(Of String))
        Dim sourceRoot = DetermineExplorerSelectionRoot(resolvedSelections)
        If String.IsNullOrWhiteSpace(sourceRoot) Then
            AppendLog("Explorer selected items do not share a common source root.")
            Return
        End If

        If Not Directory.Exists(sourceRoot) Then
            AppendLog($"Explorer source root was not found: {sourceRoot}")
            Return
        End If

        Dim relativeSelections As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
        For Each itemPath In resolvedSelections
            Dim relativePath = NormalizeRelativePath(Path.GetRelativePath(sourceRoot, itemPath))
            If relativePath.Length = 0 Then
                Continue For
            End If

            relativeSelections.Add(relativePath)
        Next

        If relativeSelections.Count = 0 Then
            ApplyExplorerSourceFolder(sourceRoot, "Explorer selection")
            Return
        End If

        _explorerSelectionSourceRoot = NormalizePathKey(sourceRoot)
        _explorerSelectedRelativePaths.Clear()
        _explorerSelectedRelativePaths.AddRange(relativeSelections)

        _suspendSourceTextChanged = True
        _sourceTextBox.Text = sourceRoot
        _suspendSourceTextChanged = False

        AppendLog($"Source set from Explorer selection: {sourceRoot}")
        AppendLog($"Explorer selected-items mode active ({_explorerSelectedRelativePaths.Count} item(s) queued).")
    End Sub

    Private Shared Function DetermineExplorerSelectionRoot(paths As IReadOnlyList(Of String)) As String
        If paths Is Nothing OrElse paths.Count = 0 Then
            Return String.Empty
        End If

        If paths.Count = 1 AndAlso Directory.Exists(paths(0)) Then
            Dim parentDirectory = Path.GetDirectoryName(paths(0))
            Return If(String.IsNullOrWhiteSpace(parentDirectory), paths(0), parentDirectory)
        End If

        Dim commonPath = DetermineCommonAncestorPath(paths)
        If String.IsNullOrWhiteSpace(commonPath) Then
            Return String.Empty
        End If

        If File.Exists(commonPath) Then
            Dim parentDirectory = Path.GetDirectoryName(commonPath)
            If String.IsNullOrWhiteSpace(parentDirectory) Then
                Return String.Empty
            End If
            Return parentDirectory
        End If

        Return commonPath
    End Function

    Private Shared Function DetermineCommonAncestorPath(paths As IReadOnlyList(Of String)) As String
        Dim commonPath = NormalizeComparablePath(paths(0))

        For index = 1 To paths.Count - 1
            Dim nextPath = NormalizeComparablePath(paths(index))
            commonPath = ReduceToCommonPath(commonPath, nextPath)
            If String.IsNullOrWhiteSpace(commonPath) Then
                Return String.Empty
            End If
        Next

        Return commonPath
    End Function

    Private Shared Function ReduceToCommonPath(leftPath As String, rightPath As String) As String
        Dim candidate = leftPath
        While Not String.IsNullOrWhiteSpace(candidate)
            If String.Equals(candidate, rightPath, StringComparison.OrdinalIgnoreCase) Then
                Return candidate
            End If

            Dim prefix = candidate & Path.DirectorySeparatorChar
            If rightPath.StartsWith(prefix, StringComparison.OrdinalIgnoreCase) Then
                Return candidate
            End If

            candidate = Path.GetDirectoryName(candidate)
            If candidate IsNot Nothing Then
                candidate = NormalizeComparablePath(candidate)
            End If
        End While

        Return String.Empty
    End Function

    Private Shared Function NormalizeComparablePath(pathValue As String) As String
        Dim fullPath = Path.GetFullPath(pathValue)
        If fullPath.Length <= 3 Then
            Return fullPath
        End If

        Return fullPath.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar)
    End Function

    Private Shared Function TryResolveExplorerPath(inputPath As String, ByRef resolvedPath As String) As Boolean
        resolvedPath = String.Empty
        If String.IsNullOrWhiteSpace(inputPath) Then
            Return False
        End If

        Try
            resolvedPath = Path.GetFullPath(inputPath.Trim())
            Return True
        Catch
            Return False
        End Try
    End Function

    Private Shared Function NormalizeRelativePath(value As String) As String
        Dim text = If(value, String.Empty).Trim()
        If text.Length = 0 OrElse text = "." Then
            Return String.Empty
        End If

        text = text.Replace(Path.AltDirectorySeparatorChar, Path.DirectorySeparatorChar)
        Return text.Trim(Path.DirectorySeparatorChar)
    End Function

    Private Shared Function NormalizePathKey(pathValue As String) As String
        If String.IsNullOrWhiteSpace(pathValue) Then
            Return String.Empty
        End If

        Try
            Return Path.GetFullPath(pathValue).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar)
        Catch
            Return String.Empty
        End Try
    End Function

    Private Sub ClearExplorerSelectionContext(announce As Boolean)
        _explorerSelectionSourceRoot = String.Empty
        _explorerSelectedRelativePaths.Clear()
        If announce Then
            AppendLog("Explorer selected-items context cleared.")
        End If
    End Sub

    Private Sub ApplyTheme()
        Dim mode As SystemColorMode = ThemeSettings.GetPreferredColorMode(_settings)
        ThemeManager.ApplyTheme(Me, mode)
    End Sub

    Private Sub BrowseFolder(targetTextBox As TextBox)
        Using dialog As New FolderBrowserDialog()
            dialog.ShowNewFolderButton = True
            dialog.UseDescriptionForTitle = True
            dialog.Description = "Select folder"

            If Directory.Exists(targetTextBox.Text) Then
                dialog.InitialDirectory = targetTextBox.Text
            End If

            If dialog.ShowDialog(Me) = DialogResult.OK Then
                targetTextBox.Text = dialog.SelectedPath
            End If
        End Using
    End Sub

    Private Async Sub StartButton_Click(sender As Object, e As EventArgs)
        If _isRunning Then
            Return
        End If

        Await StartCopyAsync()
    End Sub

    Private Async Sub PauseButton_Click(sender As Object, e As EventArgs)
        Await PauseCopyAsync()
    End Sub

    Private Async Sub ResumeButton_Click(sender As Object, e As EventArgs)
        Await ResumeCopyAsync()
    End Sub

    Private Async Sub CancelButton_Click(sender As Object, e As EventArgs)
        If Not _isRunning Then
            Return
        End If

        AppendLog("Cancellation requested. Waiting for worker acknowledgement.")
        Await _supervisor.CancelJobAsync("User requested cancellation.")
    End Sub

    Private Async Function PauseCopyAsync() As Task
        If Not _isRunning OrElse _isPaused Then
            Return
        End If

        Await _supervisor.PauseJobAsync("User requested pause.")
        _isPaused = True
        _jobSummaryLabel.Text = "Status: Paused"
        If _activeRun IsNot Nothing Then
            _jobManager.MarkRunPaused(_activeRun.RunId)
            _recoveryService.MarkActiveRunPaused(_activeRun.RunId)
        End If
        UpdateShellIndicatorsFromState()
        UpdatePauseResumeUi()
        AppendLog("Pause requested.")
    End Function

    Private Async Function ResumeCopyAsync() As Task
        If Not _isRunning OrElse Not _isPaused Then
            Return
        End If

        Await _supervisor.ResumeJobAsync("User requested resume.")
        _isPaused = False
        _jobSummaryLabel.Text = "Status: Running"
        If _activeRun IsNot Nothing Then
            _jobManager.MarkRunResumed(_activeRun.RunId)
            _recoveryService.TouchActiveRun(_activeRun.RunId)
        End If
        UpdateShellIndicatorsFromState()
        UpdatePauseResumeUi()
        AppendLog("Resume requested.")
    End Function

    Private Async Function StartCopyAsync() As Task
        Dim options As CopyJobOptions = BuildOptions()
        If options Is Nothing Then
            Return
        End If

        Dim run = _jobManager.CreateAdHocRun(options, "Manual Copy", "manual")
        Await StartCopyAsync(options, run, clearLog:=True)
    End Function

    Private Async Function StartCopyAsync(options As CopyJobOptions, run As ManagedJobRun, clearLog As Boolean) As Task
        If options Is Nothing Then
            Return
        End If

        Dim normalizedOptions = NormalizeAndValidateOptions(options, showDialogs:=True)
        If normalizedOptions Is Nothing Then
            Return
        End If

        If run Is Nothing Then
            run = _jobManager.CreateAdHocRun(normalizedOptions, "Manual Copy", "manual")
        End If

        PrepareUiForRun(normalizedOptions, clearLog)

        _isRunning = True
        _isPaused = False
        SetRunningState(isRunning:=True)
        UpdatePauseResumeUi()
        UpdateShellIndicatorsForStarting()

        _activeRun = run
        _activeRunJournalPath = ComputeJournalPath(normalizedOptions)
        _jobManager.MarkRunRunning(_activeRun.RunId, _activeRunJournalPath)
        _recoveryService.MarkJobStarted(_activeRun, normalizedOptions, _activeRunJournalPath, _settings)
        _interruptedRun = Nothing
        UpdateRecoveryMenuState()

        Try
            AppendLog($"Starting run: {_activeRun.DisplayName}")
            AppendLog("Starting copy job through supervisor.")
            AppendLog($"Buffer mode: {If(normalizedOptions.UseAdaptiveBufferSizing, "Adaptive (max " & CInt(_bufferMbNumeric.Value) & " MB)", "Fixed (" & CInt(_bufferMbNumeric.Value) & " MB)")}.")
            AppendLog($"Media wait mode: {If(normalizedOptions.WaitForMediaAvailability, "Enabled (wait forever)", "Disabled")}.")
            If normalizedOptions.SelectedRelativePaths IsNot Nothing AndAlso normalizedOptions.SelectedRelativePaths.Count > 0 Then
                AppendLog($"Selection mode: copying {normalizedOptions.SelectedRelativePaths.Count} selected item(s) from source root.")
            End If
            Await _supervisor.StartJobAsync(normalizedOptions, CancellationToken.None)
            _jobSummaryLabel.Text = "Status: Running"
        Catch ex As Exception
            _isRunning = False
            _isPaused = False
            SetRunningState(isRunning:=False)
            UpdatePauseResumeUi()
            _jobSummaryLabel.Text = "Status: Start failed"
            UpdateShellIndicatorsForFailure("Start failed")
            AppendLog($"Start failed: {ex.Message}")

            Dim failureResult As New CopyJobResult() With {
                .Succeeded = False,
                .Cancelled = False,
                .JournalPath = _activeRunJournalPath,
                .ErrorMessage = ex.Message
            }

            _jobManager.MarkRunCompleted(_activeRun.RunId, failureResult)
            _recoveryService.MarkJobEnded(_activeRun.RunId)
            _activeRun = Nothing
            _activeRunJournalPath = String.Empty
            UpdateRecoveryMenuState()

            MessageBox.Show(Me, ex.ToString(), "XactCopy", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Function

    Private Sub PrepareUiForRun(options As CopyJobOptions, clearLog As Boolean)
        ApplyOptionsToUi(options)

        If clearLog Then
            _logTextBox.Clear()
        End If

        ResetPendingProgressState()
        _overallProgressBar.Value = 0
        _currentProgressBar.Value = 0
        _jobSummaryLabel.Text = "Status: Starting"
        _overallStatsLabel.Text = "Files: 0/0  Failed: 0  Recovered: 0  Skipped: 0"
        _currentFileLabel.Text = "Current: -"
        _bytesLabel.Text = "Bytes: 0 B / 0 B"
        _throughputLabel.Text = "Speed: 0 B/s (avg 0 B/s)"
        _etaLabel.Text = "ETA: -"
        _bufferUsageLabel.Text = $"Buffer Use: last - / {FormatBytes(CLng(_bufferMbNumeric.Value) * 1024L * 1024L)}  avg -"
        _rescueTelemetryLabel.Text = "Rescue: -"
        _journalLabel.Text = "Journal: -"
        ResetCopyTelemetry()
        _lastProgressSnapshot = Nothing
    End Sub

    Private Sub ApplyDefaultSettingsToUi()
        If _settings Is Nothing OrElse _isRunning Then
            Return
        End If

        _resumeCheckBox.Checked = _settings.DefaultResumeFromJournal
        _salvageCheckBox.Checked = _settings.DefaultSalvageUnreadableBlocks
        _continueCheckBox.Checked = _settings.DefaultContinueOnFileError
        _verifyCheckBox.Checked = _settings.DefaultVerifyAfterCopy
        _adaptiveBufferCheckBox.Checked = _settings.DefaultUseAdaptiveBuffer
        _waitForMediaCheckBox.Checked = _settings.DefaultWaitForMediaAvailability
        SetActiveSalvageFillPattern(SettingsValueConverter.ToSalvageFillPattern(_settings.DefaultSalvageFillPattern))

        Dim bufferMb = Math.Max(CInt(_bufferMbNumeric.Minimum), Math.Min(CInt(_bufferMbNumeric.Maximum), _settings.DefaultBufferSizeMb))
        _bufferMbNumeric.Value = bufferMb

        Dim retries = Math.Max(CInt(_retriesNumeric.Minimum), Math.Min(CInt(_retriesNumeric.Maximum), _settings.DefaultMaxRetries))
        _retriesNumeric.Value = retries

        Dim timeoutSeconds = Math.Max(CInt(_timeoutSecondsNumeric.Minimum), Math.Min(CInt(_timeoutSecondsNumeric.Maximum), _settings.DefaultOperationTimeoutSeconds))
        _timeoutSecondsNumeric.Value = timeoutSeconds
    End Sub

    Private Function BuildDefaultCopyOptionsFromSettings() As CopyJobOptions
        Dim safeSettings = If(_settings, New AppSettings())

        Dim defaultMode = SettingsValueConverter.ToVerificationMode(safeSettings.DefaultVerificationMode)
        If defaultMode = VerificationMode.None Then
            defaultMode = VerificationMode.Full
        End If

        Return New CopyJobOptions() With {
            .SelectedRelativePaths = New List(Of String)(),
            .OverwritePolicy = SettingsValueConverter.ToOverwritePolicy(safeSettings.DefaultOverwritePolicy),
            .SymlinkHandling = SettingsValueConverter.ToSymlinkHandling(safeSettings.DefaultSymlinkHandling),
            .CopyEmptyDirectories = safeSettings.DefaultCopyEmptyDirectories,
            .BufferSizeBytes = Math.Max(4096, safeSettings.DefaultBufferSizeMb * 1024 * 1024),
            .UseAdaptiveBufferSizing = safeSettings.DefaultUseAdaptiveBuffer,
            .MaxThroughputBytesPerSecond = Math.Max(0L, CLng(safeSettings.DefaultMaxThroughputMbPerSecond) * 1024L * 1024L),
            .ParallelSmallFileWorkers = 1,
            .SmallFileThresholdBytes = 256 * 1024,
            .WaitForMediaAvailability = safeSettings.DefaultWaitForMediaAvailability,
            .MaxRetries = Math.Max(0, safeSettings.DefaultMaxRetries),
            .OperationTimeout = TimeSpan.FromSeconds(Math.Max(1, safeSettings.DefaultOperationTimeoutSeconds)),
            .PerFileTimeout = TimeSpan.FromSeconds(Math.Max(0, safeSettings.DefaultPerFileTimeoutSeconds)),
            .InitialRetryDelay = TimeSpan.FromMilliseconds(250),
            .MaxRetryDelay = TimeSpan.FromSeconds(8),
            .ResumeFromJournal = safeSettings.DefaultResumeFromJournal,
            .VerifyAfterCopy = safeSettings.DefaultVerifyAfterCopy,
            .VerificationMode = defaultMode,
            .VerificationHashAlgorithm = SettingsValueConverter.ToVerificationHashAlgorithm(safeSettings.DefaultVerificationHashAlgorithm),
            .SampleVerificationChunkBytes = Math.Max(32 * 1024, safeSettings.DefaultSampleVerificationChunkKb * 1024),
            .SampleVerificationChunkCount = Math.Max(1, safeSettings.DefaultSampleVerificationChunkCount),
            .SalvageUnreadableBlocks = safeSettings.DefaultSalvageUnreadableBlocks,
            .SalvageFillPattern = SettingsValueConverter.ToSalvageFillPattern(safeSettings.DefaultSalvageFillPattern),
            .ContinueOnFileError = safeSettings.DefaultContinueOnFileError,
            .PreserveTimestamps = safeSettings.DefaultPreserveTimestamps,
            .WorkerProcessPriorityClass = "Normal",
            .RescueFastScanChunkBytes = Math.Max(0, safeSettings.DefaultRescueFastScanChunkKb) * 1024,
            .RescueTrimChunkBytes = Math.Max(0, safeSettings.DefaultRescueTrimChunkKb) * 1024,
            .RescueScrapeChunkBytes = Math.Max(0, safeSettings.DefaultRescueScrapeChunkKb) * 1024,
            .RescueRetryChunkBytes = Math.Max(0, safeSettings.DefaultRescueRetryChunkKb) * 1024,
            .RescueSplitMinimumBytes = Math.Max(0, safeSettings.DefaultRescueSplitMinimumKb) * 1024,
            .RescueFastScanRetries = Math.Max(0, safeSettings.DefaultRescueFastScanRetries),
            .RescueTrimRetries = Math.Max(0, safeSettings.DefaultRescueTrimRetries),
            .RescueScrapeRetries = Math.Max(0, safeSettings.DefaultRescueScrapeRetries)
        }
    End Function

    Private Sub ApplyOptionsToUi(options As CopyJobOptions)
        _suspendSourceTextChanged = True
        _sourceTextBox.Text = options.SourceRoot
        _suspendSourceTextChanged = False
        _destinationTextBox.Text = options.DestinationRoot
        _resumeCheckBox.Checked = options.ResumeFromJournal
        _salvageCheckBox.Checked = options.SalvageUnreadableBlocks
        _continueCheckBox.Checked = options.ContinueOnFileError
        _verifyCheckBox.Checked = options.VerifyAfterCopy OrElse options.VerificationMode <> VerificationMode.None
        _adaptiveBufferCheckBox.Checked = options.UseAdaptiveBufferSizing
        _waitForMediaCheckBox.Checked = options.WaitForMediaAvailability
        SetActiveSalvageFillPattern(options.SalvageFillPattern)

        Dim bufferMb = Math.Ceiling(CDbl(Math.Max(4096, options.BufferSizeBytes)) / (1024.0R * 1024.0R))
        bufferMb = Math.Max(CDbl(_bufferMbNumeric.Minimum), Math.Min(CDbl(_bufferMbNumeric.Maximum), bufferMb))
        _bufferMbNumeric.Value = CDec(bufferMb)

        Dim retries = Math.Max(CInt(_retriesNumeric.Minimum), Math.Min(CInt(_retriesNumeric.Maximum), options.MaxRetries))
        _retriesNumeric.Value = retries

        Dim timeoutSeconds = Math.Max(1.0R, options.OperationTimeout.TotalSeconds)
        timeoutSeconds = Math.Max(CDbl(_timeoutSecondsNumeric.Minimum), Math.Min(CDbl(_timeoutSecondsNumeric.Maximum), timeoutSeconds))
        _timeoutSecondsNumeric.Value = CDec(timeoutSeconds)

        Dim hasSelectedPaths = options.SelectedRelativePaths IsNot Nothing AndAlso options.SelectedRelativePaths.Count > 0
        If hasSelectedPaths Then
            _explorerSelectionSourceRoot = NormalizePathKey(options.SourceRoot)
            _explorerSelectedRelativePaths.Clear()
            For Each selectedPath In options.SelectedRelativePaths
                Dim normalized = NormalizeRelativePath(selectedPath)
                If normalized.Length > 0 AndAlso Not _explorerSelectedRelativePaths.Contains(normalized, StringComparer.OrdinalIgnoreCase) Then
                    _explorerSelectedRelativePaths.Add(normalized)
                End If
            Next
        Else
            ClearExplorerSelectionContext(announce:=False)
        End If
    End Sub

    Private Function BuildOptions() As CopyJobOptions
        Dim candidate = BuildDefaultCopyOptionsFromSettings()
        candidate.SourceRoot = _sourceTextBox.Text.Trim()
        candidate.DestinationRoot = _destinationTextBox.Text.Trim()
        candidate.BufferSizeBytes = CInt(_bufferMbNumeric.Value) * 1024 * 1024
        candidate.UseAdaptiveBufferSizing = _adaptiveBufferCheckBox.Checked
        candidate.WaitForMediaAvailability = _waitForMediaCheckBox.Checked
        candidate.MaxRetries = CInt(_retriesNumeric.Value)
        candidate.OperationTimeout = TimeSpan.FromSeconds(CDbl(_timeoutSecondsNumeric.Value))
        candidate.ResumeFromJournal = _resumeCheckBox.Checked
        candidate.VerifyAfterCopy = _verifyCheckBox.Checked
        candidate.SalvageUnreadableBlocks = _salvageCheckBox.Checked
        candidate.SalvageFillPattern = _activeSalvageFillPattern
        candidate.ContinueOnFileError = _continueCheckBox.Checked

        If Not candidate.VerifyAfterCopy Then
            candidate.VerificationMode = VerificationMode.None
        ElseIf candidate.VerificationMode = VerificationMode.None Then
            candidate.VerificationMode = VerificationMode.Full
        End If

        Dim sourceKey = NormalizePathKey(candidate.SourceRoot)
        If _explorerSelectedRelativePaths.Count > 0 AndAlso
            sourceKey.Length > 0 AndAlso
            String.Equals(sourceKey, _explorerSelectionSourceRoot, StringComparison.OrdinalIgnoreCase) Then

            candidate.SelectedRelativePaths = New List(Of String)(_explorerSelectedRelativePaths)
        Else
            candidate.SelectedRelativePaths = New List(Of String)()
        End If

        Return NormalizeAndValidateOptions(candidate, showDialogs:=True)
    End Function

    Private Sub SetActiveSalvageFillPattern(pattern As SalvageFillPattern)
        _activeSalvageFillPattern = pattern
        _salvageCheckBox.Text = $"Salvage unreadable blocks ({DescribeSalvageFillPattern(pattern)})"
        _toolTip.SetToolTip(_salvageCheckBox, GetSalvageFillPatternTooltip(pattern))
    End Sub

    Private Shared Function DescribeSalvageFillPattern(pattern As SalvageFillPattern) As String
        Select Case pattern
            Case SalvageFillPattern.Ones
                Return "0xFF-fill"
            Case SalvageFillPattern.Random
                Return "random-fill"
            Case Else
                Return "zero-fill"
        End Select
    End Function

    Private Shared Function GetSalvageFillPatternTooltip(pattern As SalvageFillPattern) As String
        Select Case pattern
            Case SalvageFillPattern.Ones
                Return "Attempt to recover unreadable source regions and fill missing bytes with 0xFF."
            Case SalvageFillPattern.Random
                Return "Attempt to recover unreadable source regions and fill missing bytes with random data."
            Case Else
                Return "Attempt to recover unreadable source regions and fill missing bytes with 0x00."
        End Select
    End Function

    Private Shared Function ClampRescueChunkBytes(value As Integer) As Integer
        Const maxChunkBytes As Integer = 256 * 1024 * 1024
        If value <= 0 Then
            Return 0
        End If

        Return Math.Max(4096, Math.Min(maxChunkBytes, value))
    End Function

    Private Shared Function ClampRescueSplitMinimumBytes(value As Integer) As Integer
        Const maxSplitBytes As Integer = 64 * 1024 * 1024
        If value <= 0 Then
            Return 0
        End If

        Return Math.Max(4096, Math.Min(maxSplitBytes, value))
    End Function

    Private Shared Function ClampRescueRetries(value As Integer) As Integer
        Return Math.Max(0, Math.Min(1000, value))
    End Function

    Private Function NormalizeAndValidateOptions(options As CopyJobOptions, showDialogs As Boolean) As CopyJobOptions
        If options Is Nothing Then
            Return Nothing
        End If

        Dim sourceText = If(options.SourceRoot, String.Empty).Trim()
        Dim destinationText = If(options.DestinationRoot, String.Empty).Trim()

        If String.IsNullOrWhiteSpace(sourceText) Then
            ShowValidation("Select a source folder.", showDialogs)
            Return Nothing
        End If

        If String.IsNullOrWhiteSpace(destinationText) Then
            ShowValidation("Select a destination folder.", showDialogs)
            Return Nothing
        End If

        Dim sourceFull As String
        Dim destinationFull As String
        Try
            sourceFull = Path.GetFullPath(sourceText).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar)
            destinationFull = Path.GetFullPath(destinationText).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar)
        Catch ex As Exception
            ShowValidation("Invalid source or destination path: " & ex.Message, showDialogs)
            Return Nothing
        End Try

        If Not options.WaitForMediaAvailability AndAlso Not Directory.Exists(sourceFull) Then
            ShowValidation($"Source does not exist: {sourceFull}", showDialogs)
            Return Nothing
        End If

        If StringComparer.OrdinalIgnoreCase.Equals(sourceFull, destinationFull) Then
            ShowValidation("Source and destination cannot be the same path.", showDialogs)
            Return Nothing
        End If

        Dim normalizedSelections As New List(Of String)()
        If options.SelectedRelativePaths IsNot Nothing Then
            Dim seen As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
            For Each rawSelection In options.SelectedRelativePaths
                Dim relativeSelection = NormalizeRelativePath(rawSelection)
                If relativeSelection.Length = 0 Then
                    Continue For
                End If

                Dim candidatePath As String
                Try
                    candidatePath = Path.GetFullPath(Path.Combine(sourceFull, relativeSelection))
                Catch
                    Continue For
                End Try

                Dim sourcePrefix = sourceFull & Path.DirectorySeparatorChar
                If Not candidatePath.StartsWith(sourcePrefix, StringComparison.OrdinalIgnoreCase) Then
                    Continue For
                End If

                If Not options.WaitForMediaAvailability AndAlso Not File.Exists(candidatePath) AndAlso Not Directory.Exists(candidatePath) Then
                    Continue For
                End If

                If seen.Add(relativeSelection) Then
                    normalizedSelections.Add(relativeSelection)
                End If
            Next
        End If

        If options.SelectedRelativePaths IsNot Nothing AndAlso options.SelectedRelativePaths.Count > 0 AndAlso
            normalizedSelections.Count = 0 Then
            ShowValidation("Explorer selected-items mode could not find any valid source items to copy.", showDialogs)
            Return Nothing
        End If

        Dim mode = options.VerificationMode
        If options.VerifyAfterCopy AndAlso mode = VerificationMode.None Then
            mode = VerificationMode.Full
        ElseIf Not options.VerifyAfterCopy Then
            mode = VerificationMode.None
        End If

        Return New CopyJobOptions() With {
            .SourceRoot = sourceFull,
            .DestinationRoot = destinationFull,
            .SelectedRelativePaths = normalizedSelections,
            .OverwritePolicy = options.OverwritePolicy,
            .SymlinkHandling = options.SymlinkHandling,
            .CopyEmptyDirectories = options.CopyEmptyDirectories,
            .BufferSizeBytes = Math.Max(4096, options.BufferSizeBytes),
            .UseAdaptiveBufferSizing = options.UseAdaptiveBufferSizing,
            .MaxThroughputBytesPerSecond = Math.Max(0L, options.MaxThroughputBytesPerSecond),
            .ParallelSmallFileWorkers = Math.Max(1, options.ParallelSmallFileWorkers),
            .SmallFileThresholdBytes = Math.Max(4096, options.SmallFileThresholdBytes),
            .WaitForMediaAvailability = options.WaitForMediaAvailability,
            .MaxRetries = Math.Max(0, options.MaxRetries),
            .OperationTimeout = If(options.OperationTimeout <= TimeSpan.Zero, TimeSpan.FromSeconds(10), options.OperationTimeout),
            .PerFileTimeout = If(options.PerFileTimeout < TimeSpan.Zero, TimeSpan.Zero, options.PerFileTimeout),
            .InitialRetryDelay = If(options.InitialRetryDelay <= TimeSpan.Zero, TimeSpan.FromMilliseconds(250), options.InitialRetryDelay),
            .MaxRetryDelay = If(options.MaxRetryDelay <= TimeSpan.Zero, TimeSpan.FromSeconds(8), options.MaxRetryDelay),
            .ResumeFromJournal = options.ResumeFromJournal,
            .VerifyAfterCopy = options.VerifyAfterCopy,
            .VerificationMode = mode,
            .VerificationHashAlgorithm = options.VerificationHashAlgorithm,
            .SampleVerificationChunkBytes = Math.Max(4096, options.SampleVerificationChunkBytes),
            .SampleVerificationChunkCount = Math.Max(1, options.SampleVerificationChunkCount),
            .SalvageUnreadableBlocks = options.SalvageUnreadableBlocks,
            .SalvageFillPattern = options.SalvageFillPattern,
            .ContinueOnFileError = options.ContinueOnFileError,
            .PreserveTimestamps = options.PreserveTimestamps,
            .WorkerProcessPriorityClass = If(options.WorkerProcessPriorityClass, "Normal"),
            .RescueFastScanChunkBytes = ClampRescueChunkBytes(options.RescueFastScanChunkBytes),
            .RescueTrimChunkBytes = ClampRescueChunkBytes(options.RescueTrimChunkBytes),
            .RescueScrapeChunkBytes = ClampRescueChunkBytes(options.RescueScrapeChunkBytes),
            .RescueRetryChunkBytes = ClampRescueChunkBytes(options.RescueRetryChunkBytes),
            .RescueSplitMinimumBytes = ClampRescueSplitMinimumBytes(options.RescueSplitMinimumBytes),
            .RescueFastScanRetries = ClampRescueRetries(options.RescueFastScanRetries),
            .RescueTrimRetries = ClampRescueRetries(options.RescueTrimRetries),
            .RescueScrapeRetries = ClampRescueRetries(options.RescueScrapeRetries)
        }
    End Function

    Private Sub ShowValidation(message As String, showDialog As Boolean)
        AppendLog(message)
        If showDialog Then
            MessageBox.Show(Me, message, "XactCopy", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End If
    End Sub

    Private Sub Supervisor_ProgressChanged(sender As Object, progress As CopyProgressSnapshot)
        If progress Is Nothing Then
            Return
        End If

        SyncLock _progressDispatchLock
            _pendingProgressSnapshot = progress
        End SyncLock

        If Interlocked.CompareExchange(_progressDispatchQueued, 1, 0) <> 0 Then
            Return
        End If

        PostProgressDrain()
    End Sub

    Private Sub PostProgressDrain()
        If IsDisposed OrElse Disposing Then
            Interlocked.Exchange(_progressDispatchQueued, 0)
            Return
        End If

        If InvokeRequired Then
            Try
                BeginInvoke(New MethodInvoker(AddressOf DrainPendingProgressUpdates))
            Catch
                Interlocked.Exchange(_progressDispatchQueued, 0)
            End Try
            Return
        End If

        DrainPendingProgressUpdates()
    End Sub

    Private Sub DrainPendingProgressUpdates()
        If IsDisposed OrElse Disposing Then
            Interlocked.Exchange(_progressDispatchQueued, 0)
            Return
        End If

        Dim deferredDelayMs As Integer = 0
        Dim deferredSnapshot As CopyProgressSnapshot = Nothing

        Do
            Dim snapshot As CopyProgressSnapshot = Nothing
            SyncLock _progressDispatchLock
                snapshot = _pendingProgressSnapshot
                _pendingProgressSnapshot = Nothing
            End SyncLock

            If snapshot Is Nothing Then
                Exit Do
            End If

            Dim nowTick = Environment.TickCount64
            If _lastProgressRenderTick > 0 Then
                Dim elapsed = nowTick - _lastProgressRenderTick
                If elapsed < ProgressRenderIntervalMs Then
                    deferredDelayMs = CInt(Math.Max(1L, ProgressRenderIntervalMs - elapsed))
                    deferredSnapshot = snapshot
                    Exit Do
                End If
            End If

            ApplyProgress(snapshot)
            _lastProgressRenderTick = Environment.TickCount64
        Loop

        If deferredSnapshot IsNot Nothing Then
            SyncLock _progressDispatchLock
                If _pendingProgressSnapshot Is Nothing Then
                    _pendingProgressSnapshot = deferredSnapshot
                End If
            End SyncLock
        End If

        Interlocked.Exchange(_progressDispatchQueued, 0)

        Dim hasPendingSnapshot As Boolean
        SyncLock _progressDispatchLock
            hasPendingSnapshot = _pendingProgressSnapshot IsNot Nothing
        End SyncLock

        If hasPendingSnapshot AndAlso Interlocked.CompareExchange(_progressDispatchQueued, 1, 0) = 0 Then
            If deferredDelayMs > 0 Then
                ScheduleProgressDrain(deferredDelayMs)
            Else
                PostProgressDrain()
            End If
        End If
    End Sub

    Private Sub ScheduleProgressDrain(delayMs As Integer)
        Dim safeDelayMs = Math.Max(1, delayMs)
        Dim drainTask = Task.Run(
            Async Function() As Task
                Try
                    Await Task.Delay(safeDelayMs).ConfigureAwait(False)
                    PostProgressDrain()
                Catch
                    Interlocked.Exchange(_progressDispatchQueued, 0)
                End Try
            End Function)
        Dim ignoredDrainTask = drainTask
    End Sub

    Private Sub ResetPendingProgressState()
        SyncLock _progressDispatchLock
            _pendingProgressSnapshot = Nothing
        End SyncLock
        Interlocked.Exchange(_progressDispatchQueued, 0)
        _lastProgressRenderTick = 0
    End Sub

    Private Sub ApplyProgress(progress As CopyProgressSnapshot)
        If progress Is Nothing OrElse Not _isRunning Then
            Return
        End If

        _lastProgressSnapshot = progress
        _overallProgressBar.Value = ToProgressBarValue(progress.OverallProgress)
        _currentProgressBar.Value = ToProgressBarValue(progress.CurrentFileProgress)

        _overallStatsLabel.Text = $"Files: {progress.CompletedFiles}/{progress.TotalFiles}  Failed: {progress.FailedFiles}  Recovered: {progress.RecoveredFiles}  Skipped: {progress.SkippedFiles}"
        _currentFileLabel.Text = $"Current: {progress.CurrentFile}"
        _bytesLabel.Text = $"Bytes: {FormatBytes(progress.TotalBytesCopied)} / {FormatBytes(progress.TotalBytes)}"
        UpdateTransferTelemetry(progress)
        UpdateRescueTelemetry(progress)
        UpdateShellIndicatorsFromProgress(progress)

        If _activeRun IsNot Nothing Then
            Dim nowUtc = DateTimeOffset.UtcNow
            Dim touchIntervalSeconds = If(_settings Is Nothing, 2, Math.Max(1, _settings.RecoveryTouchIntervalSeconds))
            If _lastRecoveryTouchUtc = DateTimeOffset.MinValue OrElse (nowUtc - _lastRecoveryTouchUtc) >= TimeSpan.FromSeconds(touchIntervalSeconds) Then
                _lastRecoveryTouchUtc = nowUtc
                _recoveryService.TouchActiveRun(_activeRun.RunId)
            End If
        End If
    End Sub

    Private Sub Supervisor_LogMessage(sender As Object, message As String)
        If InvokeRequired Then
            BeginInvoke(New Action(Of String)(AddressOf AppendLog), message)
            Return
        End If

        AppendLog(message)
    End Sub

    Private Sub Supervisor_WorkerStateChanged(sender As Object, stateText As String)
        If InvokeRequired Then
            BeginInvoke(New Action(Of String)(AddressOf ApplyWorkerState), stateText)
            Return
        End If

        ApplyWorkerState(stateText)
    End Sub

    Private Sub Supervisor_JobPauseStateChanged(sender As Object, isPaused As Boolean)
        If InvokeRequired Then
            BeginInvoke(New Action(Of Boolean)(AddressOf ApplyPauseState), isPaused)
            Return
        End If

        ApplyPauseState(isPaused)
    End Sub

    Private Sub ApplyWorkerState(stateText As String)
        If String.IsNullOrWhiteSpace(stateText) Then
            Return
        End If

        AppendLog($"[Supervisor] {stateText}")
    End Sub

    Private Sub ApplyPauseState(isPaused As Boolean)
        _isPaused = isPaused
        If _isRunning Then
            _jobSummaryLabel.Text = If(_isPaused, "Status: Paused", "Status: Running")
            If _activeRun IsNot Nothing Then
                If _isPaused Then
                    _jobManager.MarkRunPaused(_activeRun.RunId)
                    _recoveryService.MarkActiveRunPaused(_activeRun.RunId)
                Else
                    _jobManager.MarkRunResumed(_activeRun.RunId)
                    _recoveryService.TouchActiveRun(_activeRun.RunId)
                End If
            End If
        End If
        UpdateShellIndicatorsFromState()
        UpdatePauseResumeUi()
    End Sub

    Private Sub Supervisor_JobCompleted(sender As Object, result As CopyJobResult)
        If InvokeRequired Then
            BeginInvoke(New Action(Of CopyJobResult)(AddressOf ApplyJobCompleted), result)
            Return
        End If

        ApplyJobCompleted(result)
    End Sub

    Private Sub ApplyJobCompleted(result As CopyJobResult)
        If result Is Nothing Then
            result = New CopyJobResult()
        End If

        Dim completedRun = _activeRun

        _isRunning = False
        _isPaused = False
        ResetPendingProgressState()
        SetRunningState(isRunning:=False)
        UpdatePauseResumeUi()
        ApplyTerminalProgressFromResult(result)
        _journalLabel.Text = $"Journal: {result.JournalPath}"
        _activeRun = Nothing
        _activeRunJournalPath = String.Empty
        _lastRecoveryTouchUtc = DateTimeOffset.MinValue

        If completedRun IsNot Nothing Then
            _jobManager.MarkRunCompleted(completedRun.RunId, result)
            _recoveryService.MarkJobEnded(completedRun.RunId)
        End If
        UpdateRecoveryMenuState()

        If result.Cancelled Then
            _jobSummaryLabel.Text = "Status: Cancelled"
            UpdateShellIndicatorsForCancelled()
            AppendLog("Copy job cancelled.")
            Return
        End If

        If result.Succeeded Then
            _jobSummaryLabel.Text = "Status: Completed successfully"
            UpdateShellIndicatorsForSuccess()
        Else
            _jobSummaryLabel.Text = "Status: Completed with failures"
            UpdateShellIndicatorsForFailure("Completed with failures")
        End If

        AppendLog($"Finished. Completed={result.CompletedFiles}, Failed={result.FailedFiles}, Recovered={result.RecoveredFiles}, Skipped={result.SkippedFiles}")
        If Not String.IsNullOrWhiteSpace(result.ErrorMessage) Then
            AppendLog($"Result message: {result.ErrorMessage}")
        End If

        AppendLog($"Journal saved at: {result.JournalPath}")

        Dim queuedTask = RunNextQueuedJobAsync(manualTrigger:=False)
        Dim ignoredQueuedTask = queuedTask
    End Sub

    Private Sub ApplyTerminalProgressFromResult(result As CopyJobResult)
        Dim previousSnapshot = _lastProgressSnapshot

        Dim totalFiles = Math.Max(0, result.TotalFiles)
        If totalFiles <= 0 AndAlso previousSnapshot IsNot Nothing Then
            totalFiles = Math.Max(0, previousSnapshot.TotalFiles)
        End If

        If totalFiles <= 0 Then
            totalFiles = Math.Max(0, result.CompletedFiles + result.FailedFiles + result.SkippedFiles)
        End If

        Dim completedFiles = Math.Max(0, result.CompletedFiles)
        Dim failedFiles = Math.Max(0, result.FailedFiles)
        Dim recoveredFiles = Math.Max(0, result.RecoveredFiles)
        Dim skippedFiles = Math.Max(0, result.SkippedFiles)

        Dim totalBytes = Math.Max(0L, result.TotalBytes)
        If totalBytes <= 0 AndAlso previousSnapshot IsNot Nothing Then
            totalBytes = Math.Max(0L, previousSnapshot.TotalBytes)
        End If

        Dim copiedBytes = Math.Max(0L, result.CopiedBytes)
        If previousSnapshot IsNot Nothing Then
            copiedBytes = Math.Max(copiedBytes, previousSnapshot.TotalBytesCopied)
        End If

        If result.Succeeded AndAlso Not result.Cancelled AndAlso totalBytes > 0 Then
            copiedBytes = totalBytes
        End If

        Dim overallProgress As Double
        If result.Succeeded AndAlso Not result.Cancelled Then
            overallProgress = 1.0R
        ElseIf totalBytes > 0 Then
            overallProgress = Math.Clamp(CDbl(copiedBytes) / CDbl(totalBytes), 0.0R, 1.0R)
        ElseIf totalFiles > 0 Then
            overallProgress = Math.Clamp(
                CDbl(completedFiles + failedFiles + skippedFiles) / CDbl(totalFiles),
                0.0R,
                1.0R)
        Else
            overallProgress = 0.0R
        End If

        _overallProgressBar.Value = ToProgressBarValue(overallProgress)

        If result.Succeeded AndAlso Not result.Cancelled Then
            _currentProgressBar.Value = 1000
            _currentFileLabel.Text = "Current: -"
            _etaLabel.Text = "ETA: 00:00:00"
        Else
            _currentProgressBar.Value = If(previousSnapshot IsNot Nothing, ToProgressBarValue(previousSnapshot.CurrentFileProgress), 0)
            _currentFileLabel.Text = If(previousSnapshot IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(previousSnapshot.CurrentFile),
                $"Current: {previousSnapshot.CurrentFile}",
                "Current: -")
        End If

        _overallStatsLabel.Text = $"Files: {completedFiles}/{totalFiles}  Failed: {failedFiles}  Recovered: {recoveredFiles}  Skipped: {skippedFiles}"
        _bytesLabel.Text = $"Bytes: {FormatBytes(copiedBytes)} / {FormatBytes(totalBytes)}"

        Dim finalRescuePass = If(result.Succeeded AndAlso Not result.Cancelled, "Complete", If(previousSnapshot IsNot Nothing, previousSnapshot.RescuePass, String.Empty))
        Dim finalRescueBadRegions = If(result.Succeeded AndAlso Not result.Cancelled, 0, If(previousSnapshot IsNot Nothing, previousSnapshot.RescueBadRegionCount, 0))
        Dim finalRescueRemainingBytes = If(result.Succeeded AndAlso Not result.Cancelled, 0L, If(previousSnapshot IsNot Nothing, previousSnapshot.RescueRemainingBytes, 0L))

        _lastProgressSnapshot = New CopyProgressSnapshot() With {
            .CurrentFile = If(result.Succeeded AndAlso Not result.Cancelled, String.Empty, If(previousSnapshot IsNot Nothing, previousSnapshot.CurrentFile, String.Empty)),
            .CurrentFileBytesCopied = If(result.Succeeded AndAlso Not result.Cancelled, 0L, If(previousSnapshot IsNot Nothing, previousSnapshot.CurrentFileBytesCopied, 0L)),
            .CurrentFileBytesTotal = If(result.Succeeded AndAlso Not result.Cancelled, 0L, If(previousSnapshot IsNot Nothing, previousSnapshot.CurrentFileBytesTotal, 0L)),
            .TotalBytesCopied = copiedBytes,
            .TotalBytes = totalBytes,
            .LastChunkBytesTransferred = If(previousSnapshot IsNot Nothing, previousSnapshot.LastChunkBytesTransferred, 0),
            .BufferSizeBytes = If(previousSnapshot IsNot Nothing, previousSnapshot.BufferSizeBytes, 0),
            .CompletedFiles = completedFiles,
            .FailedFiles = failedFiles,
            .RecoveredFiles = recoveredFiles,
            .SkippedFiles = skippedFiles,
            .TotalFiles = totalFiles,
            .RescuePass = If(finalRescuePass, String.Empty),
            .RescueBadRegionCount = Math.Max(0, finalRescueBadRegions),
            .RescueRemainingBytes = Math.Max(0L, finalRescueRemainingBytes)
        }

        UpdateRescueTelemetry(_lastProgressSnapshot)
    End Sub

    Private Sub AppendLog(message As String)
        If String.IsNullOrWhiteSpace(message) Then
            Return
        End If

        _logTextBox.AppendText(message)
        _logTextBox.AppendText(Environment.NewLine)
        TrimLogBufferIfNeeded()
        _logTextBox.SelectionStart = _logTextBox.TextLength
        _logTextBox.ScrollToCaret()
    End Sub

    Private Sub TrimLogBufferIfNeeded()
        If _logTextBox.TextLength <= MaxLogCharacters Then
            Return
        End If

        Dim cutIndex = Math.Max(0, _logTextBox.TextLength - LogTrimTargetCharacters)
        If cutIndex > 0 Then
            Dim newlineIndex = _logTextBox.Text.IndexOf(Environment.NewLine, cutIndex, StringComparison.Ordinal)
            If newlineIndex >= 0 Then
                cutIndex = newlineIndex + Environment.NewLine.Length
            End If
        End If

        If cutIndex <= 0 OrElse cutIndex >= _logTextBox.TextLength Then
            Return
        End If

        _logTextBox.Text = "[Log trimmed to recent events]" & Environment.NewLine & _logTextBox.Text.Substring(cutIndex)
    End Sub

    Private Sub SetRunningState(isRunning As Boolean)
        _sourceTextBox.Enabled = Not isRunning
        _destinationTextBox.Enabled = Not isRunning
        _browseSourceButton.Enabled = Not isRunning
        _browseDestinationButton.Enabled = Not isRunning

        _resumeCheckBox.Enabled = Not isRunning
        _salvageCheckBox.Enabled = Not isRunning
        _continueCheckBox.Enabled = Not isRunning
        _verifyCheckBox.Enabled = Not isRunning
        _adaptiveBufferCheckBox.Enabled = Not isRunning
        _waitForMediaCheckBox.Enabled = Not isRunning
        _bufferMbNumeric.Enabled = Not isRunning
        _retriesNumeric.Enabled = Not isRunning
        _timeoutSecondsNumeric.Enabled = Not isRunning

        _startButton.Enabled = Not isRunning
        _pauseButton.Enabled = isRunning
        _resumeButton.Enabled = isRunning
        _cancelButton.Enabled = isRunning

        _startMenuItem.Enabled = Not isRunning
        _pauseMenuItem.Enabled = isRunning
        _resumeMenuItem.Enabled = isRunning
        _cancelMenuItem.Enabled = isRunning

        _saveCurrentAsJobMenuItem.Enabled = Not isRunning
        _runNextQueuedJobMenuItem.Enabled = Not isRunning
        _openJobManagerMenuItem.Enabled = True

        If Not isRunning Then
            _taskbarProgress.Clear(Handle)
        End If

        UpdatePauseResumeUi()
        UpdateRecoveryMenuState()
    End Sub

    Private Sub UpdatePauseResumeUi()
        If Not _isRunning Then
            _pauseButton.Enabled = False
            _resumeButton.Enabled = False
            _pauseMenuItem.Enabled = False
            _resumeMenuItem.Enabled = False
            Return
        End If

        _pauseButton.Enabled = Not _isPaused
        _resumeButton.Enabled = _isPaused
        _pauseMenuItem.Enabled = Not _isPaused
        _resumeMenuItem.Enabled = _isPaused
    End Sub

    Private Sub ResetCopyTelemetry()
        _telemetryStartUtc = DateTimeOffset.MinValue
        _telemetryLastSampleUtc = DateTimeOffset.MinValue
        _telemetryLastBytesCopied = 0
        _smoothedBytesPerSecond = 0
        _bufferSamplesCount = 0
        _bufferSampleBytesTotal = 0
        _lastChunkUtilization = 0
        _lastProgressSnapshot = Nothing
    End Sub

    Private Sub UpdateTransferTelemetry(progress As CopyProgressSnapshot)
        Dim nowUtc = DateTimeOffset.UtcNow
        If _telemetryStartUtc = DateTimeOffset.MinValue Then
            _telemetryStartUtc = nowUtc
            _telemetryLastSampleUtc = nowUtc
            _telemetryLastBytesCopied = progress.TotalBytesCopied
        End If

        Dim intervalSeconds = (nowUtc - _telemetryLastSampleUtc).TotalSeconds
        Dim deltaBytes = progress.TotalBytesCopied - _telemetryLastBytesCopied
        If intervalSeconds > 0 AndAlso deltaBytes >= 0 Then
            Dim instantBytesPerSecond = CDbl(deltaBytes) / intervalSeconds
            If _smoothedBytesPerSecond <= 0 Then
                _smoothedBytesPerSecond = instantBytesPerSecond
            Else
                _smoothedBytesPerSecond = (_smoothedBytesPerSecond * 0.65R) + (instantBytesPerSecond * 0.35R)
            End If
        End If

        _telemetryLastSampleUtc = nowUtc
        _telemetryLastBytesCopied = progress.TotalBytesCopied

        Dim elapsedSeconds = (nowUtc - _telemetryStartUtc).TotalSeconds
        Dim averageBytesPerSecond As Double = 0
        If elapsedSeconds > 0 Then
            averageBytesPerSecond = CDbl(progress.TotalBytesCopied) / elapsedSeconds
        End If

        _throughputLabel.Text = $"Speed: {FormatBytes(CLng(Math.Round(_smoothedBytesPerSecond)))}" &
            $"/s (avg {FormatBytes(CLng(Math.Round(averageBytesPerSecond)))}/s)"

        Dim remainingBytes = Math.Max(0L, progress.TotalBytes - progress.TotalBytesCopied)
        Dim etaReferenceSpeed = If(_smoothedBytesPerSecond > 1.0R, _smoothedBytesPerSecond, averageBytesPerSecond)
        If remainingBytes <= 0 Then
            _etaLabel.Text = "ETA: 00:00:00"
        ElseIf etaReferenceSpeed > 1.0R Then
            Dim etaSeconds = remainingBytes / etaReferenceSpeed
            _etaLabel.Text = $"ETA: {FormatEta(TimeSpan.FromSeconds(etaSeconds))}"
        Else
            _etaLabel.Text = "ETA: Calculating..."
        End If

        If progress.BufferSizeBytes > 0 AndAlso progress.LastChunkBytesTransferred > 0 Then
            _bufferSamplesCount += 1
            _bufferSampleBytesTotal += progress.LastChunkBytesTransferred
            _lastChunkUtilization = progress.LastChunkBufferUtilization
        End If

        Dim configuredBufferBytes = If(progress.BufferSizeBytes > 0, progress.BufferSizeBytes, CInt(_bufferMbNumeric.Value) * 1024 * 1024)
        Dim averageUtilization As Double = 0
        If configuredBufferBytes > 0 AndAlso _bufferSamplesCount > 0 Then
            averageUtilization = Math.Clamp(
                (CDbl(_bufferSampleBytesTotal) / CDbl(_bufferSamplesCount)) / CDbl(configuredBufferBytes),
                0.0R,
                1.0R)
        End If

        Dim lastChunkText As String
        If progress.LastChunkBytesTransferred > 0 AndAlso configuredBufferBytes > 0 Then
            lastChunkText = $"{FormatBytes(progress.LastChunkBytesTransferred)} / {FormatBytes(configuredBufferBytes)} ({_lastChunkUtilization * 100.0R:0.#}%)"
        Else
            lastChunkText = $"- / {FormatBytes(configuredBufferBytes)}"
        End If

        Dim avgText = If(_bufferSamplesCount > 0, $"{averageUtilization * 100.0R:0.#}%", "-")
        _bufferUsageLabel.Text = $"Buffer Use: last {lastChunkText}  avg {avgText}"
    End Sub

    Private Sub UpdateRescueTelemetry(progress As CopyProgressSnapshot)
        If progress Is Nothing Then
            _rescueTelemetryLabel.Text = "Rescue: -"
            Return
        End If

        Dim passText = If(String.IsNullOrWhiteSpace(progress.RescuePass), "-", progress.RescuePass.Trim())
        Dim badRegions = Math.Max(0, progress.RescueBadRegionCount)
        Dim remainingBytes = Math.Max(0L, progress.RescueRemainingBytes)
        _rescueTelemetryLabel.Text = $"Rescue: {passText}  Bad regions: {badRegions}  Remaining: {FormatBytes(remainingBytes)}"
    End Sub

    Private Sub UpdateShellIndicatorsForStarting()
        Text = $"{AppTitle} - Starting..."
        If IsHandleCreated Then
            _taskbarProgress.SetIndeterminate(Handle)
        End If
    End Sub

    Private Sub UpdateShellIndicatorsFromProgress(progress As CopyProgressSnapshot)
        If progress Is Nothing Then
            Return
        End If

        Text = BuildRunningWindowTitle(progress)

        If Not IsHandleCreated Then
            Return
        End If

        Dim completed = CULng(Math.Round(Math.Clamp(progress.OverallProgress, 0.0R, 1.0R) * CDbl(TaskbarProgressScale)))
        Dim state = If(_isPaused, TaskbarProgressState.Paused, TaskbarProgressState.Normal)
        _taskbarProgress.SetProgress(Handle, completed, TaskbarProgressScale, state)
    End Sub

    Private Sub UpdateShellIndicatorsFromState()
        If Not _isRunning Then
            Text = AppTitle
            Return
        End If

        If _lastProgressSnapshot IsNot Nothing Then
            UpdateShellIndicatorsFromProgress(_lastProgressSnapshot)
            Return
        End If

        If _isPaused Then
            Text = $"{AppTitle} - Paused"
            If IsHandleCreated Then
                _taskbarProgress.SetState(Handle, TaskbarProgressState.Paused)
            End If
            Return
        End If

        Text = $"{AppTitle} - Running"
        If IsHandleCreated Then
            _taskbarProgress.SetIndeterminate(Handle)
        End If
    End Sub

    Private Sub UpdateShellIndicatorsForSuccess()
        Text = $"{AppTitle} - Completed"
        If IsHandleCreated Then
            _taskbarProgress.Clear(Handle)
        End If
    End Sub

    Private Sub UpdateShellIndicatorsForCancelled()
        Text = $"{AppTitle} - Cancelled"
        If IsHandleCreated Then
            _taskbarProgress.Clear(Handle)
        End If
    End Sub

    Private Sub UpdateShellIndicatorsForFailure(reason As String)
        Dim failureReason = If(String.IsNullOrWhiteSpace(reason), "Failed", reason)
        Text = $"{AppTitle} - {failureReason}"

        If Not IsHandleCreated Then
            Return
        End If

        Dim progressValue As ULong = TaskbarProgressScale
        If _lastProgressSnapshot IsNot Nothing Then
            progressValue = CULng(Math.Round(Math.Clamp(_lastProgressSnapshot.OverallProgress, 0.0R, 1.0R) * CDbl(TaskbarProgressScale)))
        End If

        _taskbarProgress.SetProgress(Handle, progressValue, TaskbarProgressScale, TaskbarProgressState.Error)
    End Sub

    Private Function BuildRunningWindowTitle(progress As CopyProgressSnapshot) As String
        Dim statusText = If(_isPaused, "Paused", "Copying")
        Dim percent = Math.Clamp(progress.OverallProgress, 0.0R, 1.0R) * 100.0R
        Dim totalFiles = Math.Max(0, progress.TotalFiles)
        Dim completedFiles = Math.Max(0, progress.CompletedFiles)
        Dim speedBytesPerSecond = Math.Max(0L, CLng(Math.Round(_smoothedBytesPerSecond)))
        Dim speedText = $"{FormatBytes(speedBytesPerSecond)}/s"
        Dim etaText = EstimateEtaText(progress)
        Dim rescueText = String.Empty
        If Not String.IsNullOrWhiteSpace(progress.RescuePass) Then
            Dim badRegions = Math.Max(0, progress.RescueBadRegionCount)
            Dim remainingText = FormatBytes(Math.Max(0L, progress.RescueRemainingBytes))
            rescueText = $" - {progress.RescuePass} [{badRegions} bad, {remainingText} rem]"
        End If

        Return $"{AppTitle} - {statusText} {percent:0.0}% ({completedFiles}/{totalFiles} files) - {speedText} - ETA {etaText}{rescueText}"
    End Function

    Private Function EstimateEtaText(progress As CopyProgressSnapshot) As String
        Dim remainingBytes = Math.Max(0L, progress.TotalBytes - progress.TotalBytesCopied)
        If remainingBytes <= 0 Then
            Return "00:00:00"
        End If

        Dim nowUtc = DateTimeOffset.UtcNow
        Dim elapsedSeconds = If(_telemetryStartUtc = DateTimeOffset.MinValue, 0.0R, (nowUtc - _telemetryStartUtc).TotalSeconds)
        Dim averageBytesPerSecond = If(elapsedSeconds > 0.0R, CDbl(progress.TotalBytesCopied) / elapsedSeconds, 0.0R)
        Dim etaReferenceSpeed = If(_smoothedBytesPerSecond > 1.0R, _smoothedBytesPerSecond, averageBytesPerSecond)
        If etaReferenceSpeed <= 1.0R Then
            Return "-"
        End If

        Return FormatEta(TimeSpan.FromSeconds(remainingBytes / etaReferenceSpeed))
    End Function

    Private Shared Function ToProgressBarValue(progress As Double) As Integer
        Dim bounded = Math.Clamp(progress, 0.0R, 1.0R)
        Return CInt(Math.Round(bounded * 1000.0R))
    End Function

    Private Shared Function FormatEta(value As TimeSpan) As String
        If value < TimeSpan.Zero Then
            value = TimeSpan.Zero
        End If

        If value.TotalHours >= 100 Then
            Return $"{CInt(Math.Floor(value.TotalHours)):000}:{value.Minutes:00}:{value.Seconds:00}"
        End If

        Return $"{CInt(Math.Floor(value.TotalHours)):00}:{value.Minutes:00}:{value.Seconds:00}"
    End Function

    Private Shared Function FormatBytes(value As Long) As String
        If value < 1024 Then
            Return $"{value} B"
        End If

        Dim units = {"KB", "MB", "GB", "TB", "PB"}
        Dim size = CDbl(value)
        Dim unitIndex = -1

        Do
            size /= 1024.0R
            unitIndex += 1
        Loop While size >= 1024.0R AndAlso unitIndex < units.Length - 1

        Return $"{size:0.##} {units(unitIndex)}"
    End Function
End Class


