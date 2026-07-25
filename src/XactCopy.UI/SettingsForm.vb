Imports System.Drawing
Imports System.Linq
Imports System.Reflection
Imports System.Runtime.InteropServices
Imports XactCopy.Configuration
Imports XactCopy.Models

''' <summary>
''' Class SettingsForm.
''' </summary>
Friend Class SettingsForm
    Inherits Form

    Private Const LayoutWindowKey As String = "SettingsForm"
    Private Const LayoutSplitterKey As String = "SettingsForm.MainSplit"
    Private Const WmSetRedraw As Integer = &HB
    Private Shared ReadOnly _fontCacheSyncRoot As New Object()
    Private Shared _cachedLogFontNames As IReadOnlyList(Of String)
    Private Shared ReadOnly _settingsProperties As PropertyInfo() =
        GetType(AppSettings).
            GetProperties(BindingFlags.Public Or BindingFlags.Instance).
            Where(Function(propertyInfo) propertyInfo.CanRead AndAlso propertyInfo.CanWrite).
            ToArray()
    Private Shared ReadOnly _doubleBufferedProperty As PropertyInfo =
        GetType(Control).GetProperty("DoubleBuffered", BindingFlags.Instance Or BindingFlags.NonPublic)

    Private ReadOnly _categoryTreeView As New TreeView()
    Private ReadOnly _pageHostPanel As New Panel()
    Private ReadOnly _pageLookup As New Dictionary(Of String, Control)(StringComparer.OrdinalIgnoreCase)
    Private ReadOnly _pageBuilders As New Dictionary(Of String, Func(Of Control))(StringComparer.OrdinalIgnoreCase)
    Private ReadOnly _pageTitleLabel As New Label()
    Private ReadOnly _mainSplit As New SplitContainer()

    Private ReadOnly _themeComboBox As New ComboBox()
    Private ReadOnly _accentModeComboBox As New ComboBox()
    Private ReadOnly _accentColorTextBox As New TextBox()
    Private ReadOnly _accentColorPickButton As New Button()
    Private ReadOnly _windowChromeModeComboBox As New ComboBox()
    Private ReadOnly _uiDensityComboBox As New ComboBox()
    Private ReadOnly _uiScaleComboBox As New ComboBox()
    Private ReadOnly _logFontFamilyComboBox As New ComboBox()
    Private ReadOnly _logFontSizeNumeric As New ThemedNumericUpDown()
    Private ReadOnly _logColorizeBySeverityCheckBox As New CheckBox()
    Private ReadOnly _gridAlternatingRowsCheckBox As New CheckBox()
    Private ReadOnly _gridRowHeightNumeric As New ThemedNumericUpDown()
    Private ReadOnly _gridHeaderStyleComboBox As New ComboBox()
    Private ReadOnly _showBufferStatusRowCheckBox As New CheckBox()
    Private ReadOnly _showRescueStatusRowCheckBox As New CheckBox()
    Private ReadOnly _showDiagnosticsStatusRowCheckBox As New CheckBox()
    Private ReadOnly _progressBarStyleComboBox As New ComboBox()
    Private ReadOnly _progressBarShowPercentageCheckBox As New CheckBox()

    Private ReadOnly _defaultResumeCheckBox As New CheckBox()
    Private ReadOnly _defaultSalvageCheckBox As New CheckBox()
    Private ReadOnly _defaultContinueOnErrorCheckBox As New CheckBox()
    Private ReadOnly _defaultPreserveTimestampsCheckBox As New CheckBox()
    Private ReadOnly _defaultCopyEmptyDirectoriesCheckBox As New CheckBox()
    Private ReadOnly _defaultWaitForMediaCheckBox As New CheckBox()
    Private ReadOnly _defaultWaitForLockReleaseCheckBox As New CheckBox()
    Private ReadOnly _defaultTreatAccessDeniedContentionCheckBox As New CheckBox()
    Private ReadOnly _defaultUseBadRangeMapCheckBox As New CheckBox()
    Private ReadOnly _defaultSkipKnownBadRangesCheckBox As New CheckBox()
    Private ReadOnly _defaultUpdateBadRangeMapCheckBox As New CheckBox()
    Private ReadOnly _defaultUseRawDiskScanCheckBox As New CheckBox()
    Private ReadOnly _defaultBadRangeMapMaxAgeDaysNumeric As New ThemedNumericUpDown()
    Private ReadOnly _overwritePolicyComboBox As New ComboBox()
    Private ReadOnly _symlinkHandlingComboBox As New ComboBox()
    Private ReadOnly _salvageFillPatternComboBox As New ComboBox()

    Private ReadOnly _defaultAdaptiveBufferCheckBox As New CheckBox()
    Private ReadOnly _defaultFragileModeCheckBox As New CheckBox()
    Private ReadOnly _defaultSkipFileOnFirstReadErrorCheckBox As New CheckBox()
    Private ReadOnly _defaultPersistFragileSkipsCheckBox As New CheckBox()
    Private ReadOnly _defaultBufferMbNumeric As New ThemedNumericUpDown()
    Private ReadOnly _defaultRetriesNumeric As New ThemedNumericUpDown()
    Private ReadOnly _defaultOperationTimeoutNumeric As New ThemedNumericUpDown()
    Private ReadOnly _defaultPerFileTimeoutNumeric As New ThemedNumericUpDown()
    Private ReadOnly _defaultMaxThroughputNumeric As New ThemedNumericUpDown()
    Private ReadOnly _defaultFragileFailureWindowNumeric As New ThemedNumericUpDown()
    Private ReadOnly _defaultFragileFailureThresholdNumeric As New ThemedNumericUpDown()
    Private ReadOnly _defaultFragileCooldownNumeric As New ThemedNumericUpDown()
    Private ReadOnly _defaultLockProbeIntervalNumeric As New ThemedNumericUpDown()
    Private ReadOnly _defaultSourceMutationPolicyComboBox As New ComboBox()
    Private ReadOnly _transferEnginePolicyComboBox As New ComboBox()
    Private ReadOnly _scanPerformanceProfileComboBox As New ComboBox()
    Private ReadOnly _workerProcessPriorityComboBox As New ComboBox()
    Private ReadOnly _parallelSmallFileWorkersNumeric As New ThemedNumericUpDown()
    Private ReadOnly _parallelScanWorkersNumeric As New ThemedNumericUpDown()
    Private ReadOnly _smallFileThresholdKbNumeric As New ThemedNumericUpDown()
    Private ReadOnly _rescueFastChunkKbNumeric As New ThemedNumericUpDown()
    Private ReadOnly _rescueTrimChunkKbNumeric As New ThemedNumericUpDown()
    Private ReadOnly _rescueScrapeChunkKbNumeric As New ThemedNumericUpDown()
    Private ReadOnly _rescueRetryChunkKbNumeric As New ThemedNumericUpDown()
    Private ReadOnly _rescueSplitMinimumKbNumeric As New ThemedNumericUpDown()
    Private ReadOnly _rescueFastRetriesNumeric As New ThemedNumericUpDown()
    Private ReadOnly _rescueTrimRetriesNumeric As New ThemedNumericUpDown()
    Private ReadOnly _rescueScrapeRetriesNumeric As New ThemedNumericUpDown()
    Private ReadOnly _workerTelemetryProfileComboBox As New ComboBox()
    Private ReadOnly _workerProgressIntervalNumeric As New ThemedNumericUpDown()
    Private ReadOnly _workerMaxLogsPerSecondNumeric As New ThemedNumericUpDown()
    Private ReadOnly _uiShowDiagnosticsCheckBox As New CheckBox()
    Private ReadOnly _uiDiagnosticsRefreshNumeric As New ThemedNumericUpDown()
    Private ReadOnly _uiMaxLogLinesNumeric As New ThemedNumericUpDown()

    Private ReadOnly _defaultVerifyCheckBox As New CheckBox()
    Private ReadOnly _defaultVerificationModeComboBox As New ComboBox()
    Private ReadOnly _defaultHashAlgorithmComboBox As New ComboBox()
    Private ReadOnly _sampleChunkKbNumeric As New ThemedNumericUpDown()
    Private ReadOnly _sampleChunkCountNumeric As New ThemedNumericUpDown()

    Private ReadOnly _checkUpdatesOnLaunchCheckBox As New CheckBox()
    Private ReadOnly _updateUrlTextBox As New TextBox()
    Private ReadOnly _userAgentTextBox As New TextBox()

    Private ReadOnly _enableRecoveryAutostartCheckBox As New CheckBox()
    Private ReadOnly _promptResumeAfterCrashCheckBox As New CheckBox()
    Private ReadOnly _autoResumeAfterCrashCheckBox As New CheckBox()
    Private ReadOnly _keepResumePromptCheckBox As New CheckBox()
    Private ReadOnly _autoRunQueuedOnStartupCheckBox As New CheckBox()
    Private ReadOnly _recoveryTouchIntervalNumeric As New ThemedNumericUpDown()

    Private ReadOnly _enableExplorerContextMenuCheckBox As New CheckBox()
    Private ReadOnly _explorerSelectionModeComboBox As New ComboBox()

    Private ReadOnly _okButton As New Button()
    Private ReadOnly _cancelButton As New Button()
    Private ReadOnly _toolTip As New ToolTip() With {
        .ShowAlways = True,
        .InitialDelay = 350,
        .ReshowDelay = 120,
        .AutoPopDelay = 24000
    }

    Private ReadOnly _workingSettings As AppSettings
    Private ReadOnly _originalSettings As AppSettings
    Private ReadOnly _dirtyStateTimer As New Timer() With {.Interval = 120}
    Private _hasChanges As Boolean
    Private _requiresRestart As Boolean
    Private _restartReasonText As String = String.Empty
    Private _suspendDirtyTracking As Boolean
    Private _activePageKey As String = String.Empty
    Private _activePageControl As Control
    Private _initialRevealPending As Boolean = True
    Private _nativeRedrawSuspended As Boolean

    <DllImport("user32.dll", CharSet:=CharSet.Auto)>
    Private Shared Function SendMessage(
        hWnd As IntPtr,
        msg As Integer,
        wParam As IntPtr,
        lParam As IntPtr) As IntPtr
    End Function

    ''' <summary>
    ''' Initializes a new instance.
    ''' </summary>
    Public Sub New(settings As AppSettings)
        If settings Is Nothing Then
            Throw New ArgumentNullException(NameOf(settings))
        End If

        SuspendLayout()
        _workingSettings = CloneSettings(settings)
        _originalSettings = CloneSettings(settings)

        Text = "Settings"
        StartPosition = FormStartPosition.CenterParent
        MinimumSize = New Size(960, 620)
        Size = New Size(1080, 720)
        Opacity = 0.0R
        WindowIconHelper.Apply(Me)
        SetStyle(ControlStyles.AllPaintingInWmPaint Or ControlStyles.OptimizedDoubleBuffer, True)
        UpdateStyles()

        BuildUi()
        EnableDoubleBufferingForContainers(Me)
        ConfigureToolTips()
        WireDirtyTracking(Controls)
        ApplySettingsToControls()

        AddHandler _dirtyStateTimer.Tick, AddressOf DirtyStateTimer_Tick
        AddHandler Load, AddressOf SettingsForm_Load
        AddHandler Shown, AddressOf SettingsForm_Shown
        AddHandler FormClosing, AddressOf SettingsForm_FormClosing
        If _categoryTreeView.Nodes.Count > 0 Then
            _categoryTreeView.SelectedNode = _categoryTreeView.Nodes(0)
            ShowPage(If(_categoryTreeView.Nodes(0).Name, "appearance"))
        End If
        ResumeLayout(performLayout:=True)
    End Sub

    ''' <summary>
    ''' Gets or sets ResultSettings.
    ''' </summary>
    Public ReadOnly Property ResultSettings As AppSettings
        Get
            Return CloneSettings(_workingSettings)
        End Get
    End Property

    ''' <summary>
    ''' Gets or sets HasChanges.
    ''' </summary>
    Public ReadOnly Property HasChanges As Boolean
        Get
            Return _hasChanges
        End Get
    End Property

    ''' <summary>
    ''' Gets or sets RequiresApplicationRestart.
    ''' </summary>
    Public ReadOnly Property RequiresApplicationRestart As Boolean
        Get
            Return _requiresRestart
        End Get
    End Property

    ''' <summary>
    ''' Gets or sets RestartReasonText.
    ''' </summary>
    Public ReadOnly Property RestartReasonText As String
        Get
            Return _restartReasonText
        End Get
    End Property

    Private Sub BuildUi()
        SuspendLayout()
        Dim rootLayout As New TableLayoutPanel() With {
            .Dock = DockStyle.Fill,
            .ColumnCount = 1,
            .RowCount = 2,
            .Padding = New Padding(8)
        }
        rootLayout.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0F))
        rootLayout.RowStyles.Add(New RowStyle(SizeType.AutoSize))

        _mainSplit.Dock = DockStyle.Fill
        _mainSplit.FixedPanel = FixedPanel.Panel1
        _mainSplit.Panel1.Controls.Add(BuildCategoryPane())
        _mainSplit.Panel2.Controls.Add(BuildPagePane())

        Dim buttonPanel As New FlowLayoutPanel() With {
            .Dock = DockStyle.Fill,
            .AutoSize = True,
            .FlowDirection = FlowDirection.RightToLeft,
            .WrapContents = False,
            .Margin = New Padding(0, 6, 0, 0)
        }

        _okButton.Text = "Save"
        _okButton.AutoSize = True
        AddHandler _okButton.Click, AddressOf OkButton_Click

        _cancelButton.Text = "Cancel"
        _cancelButton.AutoSize = True
        AddHandler _cancelButton.Click,
            Sub(sender, e)
                DialogResult = DialogResult.Cancel
                Close()
            End Sub

        buttonPanel.Controls.Add(_cancelButton)
        buttonPanel.Controls.Add(_okButton)

        rootLayout.Controls.Add(_mainSplit, 0, 0)
        rootLayout.Controls.Add(buttonPanel, 0, 1)

        Controls.Add(rootLayout)
        AcceptButton = _okButton
        CancelButton = _cancelButton
        ResumeLayout(performLayout:=True)
    End Sub

    Private Sub WireDirtyTracking(controlCollection As Control.ControlCollection)
        If controlCollection Is Nothing Then
            Return
        End If

        For Each control As Control In controlCollection
            If TypeOf control Is CheckBox Then
                AddHandler DirectCast(control, CheckBox).CheckedChanged, AddressOf InputControl_Changed
            ElseIf TypeOf control Is ComboBox Then
                AddHandler DirectCast(control, ComboBox).SelectedIndexChanged, AddressOf InputControl_Changed
                AddHandler DirectCast(control, ComboBox).TextChanged, AddressOf InputControl_Changed
            ElseIf TypeOf control Is TextBox Then
                AddHandler DirectCast(control, TextBox).TextChanged, AddressOf InputControl_Changed
            ElseIf TypeOf control Is NumericUpDown Then
                AddHandler DirectCast(control, NumericUpDown).ValueChanged, AddressOf InputControl_Changed
            End If

            If control.HasChildren Then
                WireDirtyTracking(control.Controls)
            End If
        Next
    End Sub

    Private Sub InputControl_Changed(sender As Object, e As EventArgs)
        If _suspendDirtyTracking Then
            Return
        End If

        _dirtyStateTimer.Stop()
        _dirtyStateTimer.Start()
    End Sub

    Private Sub DirtyStateTimer_Tick(sender As Object, e As EventArgs)
        _dirtyStateTimer.Stop()
        UpdateDirtyState()
    End Sub

    Private Sub ConfigureToolTips()
        SetDetailedToolTip(
            _categoryTreeView,
            "Pick the settings page to edit.",
            "Changes stay in this dialog until you press Save.")

        SetDetailedToolTip(_themeComboBox, "Changes the color mode for XactCopy windows.", "Use System if you want XactCopy to follow Windows theme changes.")
        SetDetailedToolTip(_accentModeComboBox, "Controls where highlight and progress colors come from.", "Auto is the safest default; Custom uses the color box below.")
        SetDetailedToolTip(_accentColorTextBox, "Enter a custom accent color as #RRGGBB.", "Example: #5A78C8. Used only when Accent source is Custom.")
        SetDetailedToolTip(_accentColorPickButton, "Open a color picker for the custom accent color.", "This updates the #RRGGBB value for you.")
        SetDetailedToolTip(_windowChromeModeComboBox, "Chooses whether XactCopy draws its own title bar styling.", "Use Standard if themed window borders look wrong on this PC.")
        SetDetailedToolTip(_uiDensityComboBox, "Changes spacing between controls.", "Compact fits more on screen; Comfortable is easier to read.")
        SetDetailedToolTip(_uiScaleComboBox, "Scales UI text and controls inside XactCopy.", "Use this when the app feels too small or too large for your display.")
        SetDetailedToolTip(_logFontFamilyComboBox, "Chooses the font used by the operations log.", "Monospace fonts make paths, byte counts, and timings easier to scan.")
        SetDetailedToolTip(_logFontSizeNumeric, "Sets the operations log font size in points.", "Increase this for readability; decrease it to fit more log lines.")
        SetDetailedToolTip(_logColorizeBySeverityCheckBox, "Colors log lines by severity.", "Warnings and errors stand out faster during long runs.")
        SetDetailedToolTip(_gridAlternatingRowsCheckBox, "Adds alternating row shading in grids.", "Useful when scanning job history or queue rows.")
        SetDetailedToolTip(_gridRowHeightNumeric, "Sets the default row height for data grids.", "Raise it if text feels cramped; lower it to fit more rows.")
        SetDetailedToolTip(_gridHeaderStyleComboBox, "Chooses how strongly grid column headers are styled.", "Minimal is quiet; Prominent is easier to see.")
        SetDetailedToolTip(_showBufferStatusRowCheckBox, "Shows the buffer usage row on the main window.", "Useful when tuning transfer speed or checking adaptive buffer behavior.")
        SetDetailedToolTip(_showRescueStatusRowCheckBox, "Shows rescue pass and bad-range status on the main window.", "Keep this on when scanning or copying from questionable media.")
        SetDetailedToolTip(_showDiagnosticsStatusRowCheckBox, "Shows UI event and rendering diagnostics on the main window.", "Useful for troubleshooting UI lag; otherwise it is visual noise.")
        SetDetailedToolTip(_progressBarStyleComboBox, "Changes progress bar thickness.", "Thick is easiest to read; Thin leaves more room for logs.")
        SetDetailedToolTip(_progressBarShowPercentageCheckBox, "Shows percentage text inside progress bars.", "Turn off if you prefer a cleaner progress display.")

        SetDetailedToolTip(_defaultResumeCheckBox, "Uses the journal to continue interrupted runs.", "Recommended for large jobs, removable drives, and any copy that may be stopped.")
        SetDetailedToolTip(_defaultSalvageCheckBox, "Writes placeholder data for unreadable source blocks instead of failing the whole file.", "Use for best-effort recovery from damaged media. Leave off for exact-only copies.")
        SetDetailedToolTip(_defaultContinueOnErrorCheckBox, "Keeps processing other files after one file fails.", "Useful for large batches where one bad file should not stop the run.")
        SetDetailedToolTip(_defaultPreserveTimestampsCheckBox, "Copies source file and folder timestamps to the destination.", "Keep this on for backups and archive migrations.")
        SetDetailedToolTip(_defaultCopyEmptyDirectoriesCheckBox, "Creates empty source folders at the destination.", "Turn on when the folder structure matters even if some folders contain no files.")
        SetDetailedToolTip(_defaultWaitForMediaCheckBox, "Waits when source or destination disappears instead of failing immediately.", "Use for USB disks, network paths, and drives that reconnect.")
        SetDetailedToolTip(_defaultWaitForLockReleaseCheckBox, "Waits for locked files to become readable or writable.", "Useful when antivirus, indexing, or another app temporarily holds a file.")
        SetDetailedToolTip(_defaultTreatAccessDeniedContentionCheckBox, "Treats some Access Denied errors as temporary lock contention.", "Use when security software or sync tools briefly block files.")
        SetDetailedToolTip(_defaultUseBadRangeMapCheckBox, "Loads saved bad-range information for the same source.", "This lets later scans and copies reuse what XactCopy already learned.")
        SetDetailedToolTip(_defaultSkipKnownBadRangesCheckBox, "Avoids rereading ranges already marked unreadable.", "Reduces stress on failing drives, but keeps those ranges as known bad.")
        SetDetailedToolTip(_defaultUpdateBadRangeMapCheckBox, "Saves newly found unreadable ranges after scan or copy runs.", "Keep this on if you want future runs to get faster and gentler.")
        SetDetailedToolTip(_defaultUseRawDiskScanCheckBox, "Uses the experimental raw-disk scan backend for local NTFS files.", "Can scan healthy files faster when running as administrator. Falls back when unsupported.")
        SetDetailedToolTip(_defaultBadRangeMapMaxAgeDaysNumeric, "Controls how old a bad-range map can be before it is ignored.", "0 means never expire. Use a limit if media contents change often.")

        SetDetailedToolTip(_overwritePolicyComboBox, "Chooses what happens when a destination file already exists.", "Ask is safest for manual runs; Skip or Newer is better for repeatable jobs.")
        SetDetailedToolTip(_symlinkHandlingComboBox, "Controls symbolic links in the source tree.", "Skip avoids loops and surprises. Follow copies the linked target content.")
        SetDetailedToolTip(_salvageFillPatternComboBox, "Chooses bytes written where source data cannot be read.", "Zero-fill is predictable. Random-fill avoids repeated blank blocks but is less readable.")

        SetDetailedToolTip(_defaultAdaptiveBufferCheckBox, "Lets XactCopy choose buffer size automatically for each file.", "When on, the manual buffer setting is ignored and live I/O behavior controls sizing.")
        SetDetailedToolTip(_defaultFragileModeCheckBox, "Uses conservative behavior for unstable drives.", "Turn this on for disks that disconnect, stall, click, or get worse under heavy reads.")
        SetDetailedToolTip(_defaultSkipFileOnFirstReadErrorCheckBox, "Stops scanning a file after the first read failure in fragile mode.", "This protects weak media, but may recover less data from that file.")
        SetDetailedToolTip(_defaultPersistFragileSkipsCheckBox, "Remembers fragile-mode skipped files in the journal.", "Prevents resume/recovery from repeatedly hitting the same failing files.")
        SetDetailedToolTip(_defaultBufferMbNumeric, "Sets the buffer size used when adaptive buffering is off.", "Adaptive runs ignore this value and tune themselves automatically.")
        SetDetailedToolTip(_defaultRetriesNumeric, "Sets how many times each failed I/O operation is retried.", "More retries can recover transient errors but can slow or stress bad media.")
        SetDetailedToolTip(_defaultOperationTimeoutNumeric, "Limits how long one read or write operation may stall.", "Shorter detects stuck devices faster. Longer helps slow network or USB storage.")
        SetDetailedToolTip(_defaultPerFileTimeoutNumeric, "Limits total time spent on one file.", "0 disables the limit. Use this to keep a job moving past files that hang.")
        SetDetailedToolTip(_defaultMaxThroughputNumeric, "Caps transfer speed in MB/s.", "0 means unlimited. Use a cap to keep the PC, network, or external disk responsive.")
        SetDetailedToolTip(_transferEnginePolicyComboBox, "Chooses the copy engine used for new copy runs.", "Auto is recommended. Managed rescue favors resilience; Native fast favors speed on clean files.")
        SetDetailedToolTip(_scanPerformanceProfileComboBox, "Chooses how scan-only runs read files.", "Auto fast scan is recommended. Precise scan is slower but maps bad ranges carefully from the start.")
        SetDetailedToolTip(_workerProcessPriorityComboBox, "Sets CPU scheduling priority for the background worker.", "Normal is best for most jobs. High can make the desktop less responsive.")
        SetDetailedToolTip(_parallelSmallFileWorkersNumeric, "Controls how many small files can be copied in parallel.", "0 lets XactCopy choose. 1 disables the parallel small-file phase.")
        SetDetailedToolTip(_parallelScanWorkersNumeric, "Controls how many files fast scan reads at the same time.", "0 lets XactCopy choose. Use 1 for weak drives; higher values suit SSDs.")
        SetDetailedToolTip(_smallFileThresholdKbNumeric, "Files at or below this size can use the parallel small-file phase.", "Raise for many tiny files. Lower if the destination becomes overloaded.")
        SetDetailedToolTip(_defaultFragileFailureWindowNumeric, "Sets the time window used to count repeated fragile-mode failures.", "Shorter reacts to bursts quickly; longer catches slower failure patterns.")
        SetDetailedToolTip(_defaultFragileFailureThresholdNumeric, "Sets how many failed files trigger fragile-mode cooldown.", "Lower protects weak drives sooner. Higher keeps the job moving longer.")
        SetDetailedToolTip(_defaultFragileCooldownNumeric, "Pauses after the fragile failure threshold is reached.", "0 disables cooldown. Use a delay when a drive needs time to recover.")
        SetDetailedToolTip(_defaultLockProbeIntervalNumeric, "Sets how often XactCopy checks whether a locked file is available.", "Lower is more responsive. Higher reduces repeated filesystem probes.")
        SetDetailedToolTip(_defaultSourceMutationPolicyComboBox, "Chooses what to do if a source file disappears mid-run.", "Fail is strict. Skip keeps batches moving. Wait is best for unstable network paths.")
        SetDetailedToolTip(_rescueFastChunkKbNumeric, "Chunk size for the first rescue pass.", "0 = auto. Larger chunks scan faster; smaller chunks localize failures sooner.")
        SetDetailedToolTip(_rescueTrimChunkKbNumeric, "Chunk size used to narrow bad areas after a failure.", "0 = auto. Tune only when you know the media behavior.")
        SetDetailedToolTip(_rescueScrapeChunkKbNumeric, "Chunk size for deeper recovery attempts around bad areas.", "0 = auto. Smaller values can recover more but take longer.")
        SetDetailedToolTip(_rescueRetryChunkKbNumeric, "Chunk size for final retry attempts on bad ranges.", "0 = auto, usually 4 KB. Smaller is slower but more precise.")
        SetDetailedToolTip(_rescueSplitMinimumKbNumeric, "Smallest block size used when splitting failing ranges.", "0 = auto. Lower can isolate damage more closely but increases work.")
        SetDetailedToolTip(_rescueFastRetriesNumeric, "Retries used during the first rescue scan pass.", "Keep low for fragile drives so bad areas are found quickly.")
        SetDetailedToolTip(_rescueTrimRetriesNumeric, "Retries used while trimming bad regions.", "More retries can recover intermittent sectors but increases scan time.")
        SetDetailedToolTip(_rescueScrapeRetriesNumeric, "Retries used during the deeper scrape pass.", "Increase only when recovery quality matters more than drive stress or time.")
        SetDetailedToolTip(_workerTelemetryProfileComboBox, "Controls how much detail the worker reports.", "Normal is best for everyday use. Debug is noisy and intended for troubleshooting.")
        SetDetailedToolTip(_workerProgressIntervalNumeric, "Minimum time between worker progress events.", "Lower feels more live but costs more UI/IPC overhead.")
        SetDetailedToolTip(_workerMaxLogsPerSecondNumeric, "Caps how many log events the worker sends each second.", "0 disables the cap. Use a cap to keep the UI smooth during noisy failures.")
        SetDetailedToolTip(_uiShowDiagnosticsCheckBox, "Shows live UI counters on the main window.", "Useful while tuning performance or investigating UI lag.")
        SetDetailedToolTip(_uiDiagnosticsRefreshNumeric, "Sets how often the UI diagnostics strip refreshes.", "Lower updates more often; higher reduces repaint work.")
        SetDetailedToolTip(_uiMaxLogLinesNumeric, "Limits how many runtime log lines stay in memory.", "Higher preserves more history. Lower uses less memory on very long jobs.")

        SetDetailedToolTip(_defaultVerifyCheckBox, "Verifies destination files after copying.", "Use when integrity matters and extra read time is acceptable.")
        SetDetailedToolTip(_defaultVerificationModeComboBox, "Chooses full-file or sampled verification.", "Full hash is strongest. Sampled hash is faster but cannot prove every byte.")
        SetDetailedToolTip(_defaultHashAlgorithmComboBox, "Chooses the hash used for verification.", "SHA-256 is fast and strong. SHA-512 is available for stricter policies.")
        SetDetailedToolTip(_sampleChunkKbNumeric, "Sets how much data each sampled verification read checks.", "Larger samples improve coverage but take longer.")
        SetDetailedToolTip(_sampleChunkCountNumeric, "Sets how many sample locations are checked per file.", "More samples increase confidence and verification time.")

        SetDetailedToolTip(_checkUpdatesOnLaunchCheckBox, "Checks for new XactCopy releases when the app starts.", "Turn off for offline or locked-down machines.")
        SetDetailedToolTip(_updateUrlTextBox, "Release metadata endpoint used for update checks.", "Use the default unless you host releases from a private endpoint.")
        SetDetailedToolTip(_userAgentTextBox, "HTTP User-Agent sent during update checks.", "Change only if your proxy or release server requires a specific value.")

        SetDetailedToolTip(_enableRecoveryAutostartCheckBox, "Registers startup recovery after an interrupted run.", "Helps resume work after a crash, reboot, or power loss.")
        SetDetailedToolTip(_autoRunQueuedOnStartupCheckBox, "Starts queued jobs automatically when XactCopy launches.", "Use only when the queue is trusted and destinations are expected to be available.")
        SetDetailedToolTip(_promptResumeAfterCrashCheckBox, "Shows a prompt when an interrupted run is found.", "Best default when you want to choose resume, inspect, or dismiss manually.")
        SetDetailedToolTip(_autoResumeAfterCrashCheckBox, "Automatically resumes interrupted runs without asking.", "Use for unattended systems where continuing is always preferred.")
        SetDetailedToolTip(_keepResumePromptCheckBox, "Keeps showing the resume prompt until the interrupted run is handled.", "Prevents forgotten recovery data from being silently ignored.")
        SetDetailedToolTip(_recoveryTouchIntervalNumeric, "Sets how often recovery heartbeat data is written.", "Shorter detects interruptions sooner. Longer reduces small background writes.")

        SetDetailedToolTip(_enableExplorerContextMenuCheckBox, "Adds Copy with XactCopy to Windows Explorer menus.", "Useful when you often start jobs from selected files, folders, or drives.")
        SetDetailedToolTip(_explorerSelectionModeComboBox, "Chooses what Explorer-launched jobs treat as the source.", "Selected items copies only the selection. Full source folder copies the containing folder.")

        SetDetailedToolTip(_okButton, "Saves these settings and closes the dialog.", "Some appearance and integration changes may require restart.")
        SetDetailedToolTip(_cancelButton, "Closes the dialog without saving changes.", "Any edits made since opening Settings are discarded.")
    End Sub

    Private Sub SetDetailedToolTip(control As Control, ParamArray lines() As String)
        _toolTip.SetToolTip(control, TooltipScenarioFormatter.Compose(lines))
    End Sub

    Private Function BuildCategoryPane() As Control
        Dim layout As New TableLayoutPanel() With {
            .Dock = DockStyle.Fill,
            .ColumnCount = 1,
            .RowCount = 2,
            .Padding = New Padding(4)
        }
        layout.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0F))
        layout.RowStyles.Add(New RowStyle(SizeType.AutoSize))
        layout.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0F))

        Dim headerLabel As New Label() With {
            .AutoSize = True,
            .Text = "Categories",
            .Font = New Font(Font, FontStyle.Bold),
            .Margin = New Padding(2, 2, 2, 6)
        }

        _categoryTreeView.Dock = DockStyle.Fill
        _categoryTreeView.HideSelection = False
        _categoryTreeView.FullRowSelect = True
        _categoryTreeView.ShowLines = False
        _categoryTreeView.ShowPlusMinus = False
        _categoryTreeView.ShowRootLines = False
        _categoryTreeView.ItemHeight = 22
        _categoryTreeView.Indent = 16

        _categoryTreeView.BeginUpdate()
        _categoryTreeView.Nodes.Add(New TreeNode("Appearance") With {.Name = "appearance"})
        _categoryTreeView.Nodes.Add(New TreeNode("Copy Defaults") With {.Name = "copy"})
        _categoryTreeView.Nodes.Add(New TreeNode("Performance") With {.Name = "performance"})
        _categoryTreeView.Nodes.Add(New TreeNode("Diagnostics") With {.Name = "diagnostics"})
        _categoryTreeView.Nodes.Add(New TreeNode("Verification") With {.Name = "verification"})
        _categoryTreeView.Nodes.Add(New TreeNode("Updates") With {.Name = "updates"})
        _categoryTreeView.Nodes.Add(New TreeNode("Recovery & Startup") With {.Name = "recovery"})
        _categoryTreeView.Nodes.Add(New TreeNode("Explorer Integration") With {.Name = "explorer"})
        _categoryTreeView.EndUpdate()
        AddHandler _categoryTreeView.AfterSelect, AddressOf CategoryTreeView_AfterSelect

        layout.Controls.Add(headerLabel, 0, 0)
        layout.Controls.Add(_categoryTreeView, 0, 1)
        Return layout
    End Function

    Private Function BuildPagePane() As Control
        Dim layout As New TableLayoutPanel() With {
            .Dock = DockStyle.Fill,
            .ColumnCount = 1,
            .RowCount = 2,
            .Padding = New Padding(6, 4, 4, 4)
        }
        layout.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0F))
        layout.RowStyles.Add(New RowStyle(SizeType.AutoSize))
        layout.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0F))

        _pageTitleLabel.AutoSize = True
        _pageTitleLabel.Font = New Font(Font, FontStyle.Bold)
        _pageTitleLabel.Margin = New Padding(2, 2, 2, 8)

        _pageHostPanel.Dock = DockStyle.Fill
        _pageHostPanel.AutoScroll = True

        ' Only build the default page now; other pages are built on first navigation.
        _pageBuilders("copy") = AddressOf BuildCopyDefaultsPage
        _pageBuilders("performance") = AddressOf BuildPerformancePage
        _pageBuilders("diagnostics") = AddressOf BuildDiagnosticsPage
        _pageBuilders("verification") = AddressOf BuildVerificationPage
        _pageBuilders("updates") = AddressOf BuildUpdatesPage
        _pageBuilders("recovery") = AddressOf BuildRecoveryPage
        _pageBuilders("explorer") = AddressOf BuildExplorerPage

        Dim appearancePage = BuildAppearancePage()
        appearancePage.Dock = DockStyle.Top
        appearancePage.AutoSize = True
        appearancePage.Visible = False
        _pageHostPanel.Controls.Add(appearancePage)
        _pageLookup("appearance") = appearancePage

        layout.Controls.Add(_pageTitleLabel, 0, 0)
        layout.Controls.Add(_pageHostPanel, 0, 1)
        Return layout
    End Function

    Private Function BuildAppearancePage() As Control
        Dim page = CreatePageContainer(5)

        Dim theme = CreateFieldGrid(4)

        _themeComboBox.DropDownStyle = ComboBoxStyle.DropDownList
        _themeComboBox.Items.AddRange(New Object() {"Dark", "System", "Classic"})
        ConfigureComboBoxControl(_themeComboBox, width:=220)

        _accentModeComboBox.DropDownStyle = ComboBoxStyle.DropDownList
        _accentModeComboBox.Items.AddRange(New Object() {"Auto", "System", "Custom"})
        ConfigureComboBoxControl(_accentModeComboBox, width:=220)
        AddHandler _accentModeComboBox.SelectedIndexChanged, AddressOf AppearanceControl_Changed

        ConfigureTextBoxControl(_accentColorTextBox, stretch:=False)
        _accentColorTextBox.Width = 120
        _accentColorTextBox.CharacterCasing = CharacterCasing.Upper

        _accentColorPickButton.Text = "Pick..."
        _accentColorPickButton.AutoSize = True
        AddHandler _accentColorPickButton.Click, AddressOf AccentColorPickButton_Click

        _windowChromeModeComboBox.DropDownStyle = ComboBoxStyle.DropDownList
        _windowChromeModeComboBox.Items.AddRange(New Object() {"Themed title bar (if supported)", "Standard title bar"})
        ConfigureComboBoxControl(_windowChromeModeComboBox, width:=280)

        Dim accentColorPane As New FlowLayoutPanel() With {
            .AutoSize = True,
            .AutoSizeMode = AutoSizeMode.GrowAndShrink,
            .FlowDirection = FlowDirection.LeftToRight,
            .WrapContents = False,
            .Margin = New Padding(0)
        }
        accentColorPane.Controls.Add(_accentColorTextBox)
        accentColorPane.Controls.Add(_accentColorPickButton)

        theme.Controls.Add(CreateFieldLabel("Mode"), 0, 0)
        theme.Controls.Add(_themeComboBox, 1, 0)
        theme.Controls.Add(CreateFieldLabel("Accent source"), 0, 1)
        theme.Controls.Add(_accentModeComboBox, 1, 1)
        theme.Controls.Add(CreateFieldLabel("Custom accent color"), 0, 2)
        theme.Controls.Add(accentColorPane, 1, 2)
        theme.Controls.Add(CreateFieldLabel("Window chrome"), 0, 3)
        theme.Controls.Add(_windowChromeModeComboBox, 1, 3)

        Dim layout = CreateFieldGrid(2)

        _uiDensityComboBox.DropDownStyle = ComboBoxStyle.DropDownList
        _uiDensityComboBox.Items.AddRange(New Object() {"Compact", "Normal", "Comfortable"})
        ConfigureComboBoxControl(_uiDensityComboBox, width:=220)

        _uiScaleComboBox.DropDownStyle = ComboBoxStyle.DropDownList
        _uiScaleComboBox.Items.AddRange(New Object() {"90%", "100%", "110%", "125%"})
        ConfigureComboBoxControl(_uiScaleComboBox, width:=160)

        layout.Controls.Add(CreateFieldLabel("UI density"), 0, 0)
        layout.Controls.Add(_uiDensityComboBox, 1, 0)
        layout.Controls.Add(CreateFieldLabel("UI scale"), 0, 1)
        layout.Controls.Add(_uiScaleComboBox, 1, 1)

        Dim log = CreateFieldGrid(3)

        _logFontFamilyComboBox.DropDownStyle = ComboBoxStyle.DropDownList
        ConfigureComboBoxControl(_logFontFamilyComboBox, width:=320)
        PopulateLogFontFamilyComboBox()

        ConfigureNumeric(_logFontSizeNumeric, 7D, 20D, 9D)
        _logColorizeBySeverityCheckBox.Text = "Color-code log by severity"
        ConfigureCheckBox(_logColorizeBySeverityCheckBox)

        log.Controls.Add(CreateFieldLabel("Operations log font"), 0, 0)
        log.Controls.Add(_logFontFamilyComboBox, 1, 0)
        log.Controls.Add(CreateFieldLabel("Operations log size (pt)"), 0, 1)
        log.Controls.Add(_logFontSizeNumeric, 1, 1)
        log.Controls.Add(_logColorizeBySeverityCheckBox, 0, 2)
        log.SetColumnSpan(_logColorizeBySeverityCheckBox, 3)

        Dim grid = CreateFieldGrid(3)

        _gridAlternatingRowsCheckBox.Text = "Use alternating grid rows"
        ConfigureCheckBox(_gridAlternatingRowsCheckBox)
        ConfigureNumeric(_gridRowHeightNumeric, 18D, 48D, 24D)

        _gridHeaderStyleComboBox.DropDownStyle = ComboBoxStyle.DropDownList
        _gridHeaderStyleComboBox.Items.AddRange(New Object() {"Default", "Minimal", "Prominent"})
        ConfigureComboBoxControl(_gridHeaderStyleComboBox, width:=220)

        grid.Controls.Add(_gridAlternatingRowsCheckBox, 0, 0)
        grid.SetColumnSpan(_gridAlternatingRowsCheckBox, 3)
        grid.Controls.Add(CreateFieldLabel("Grid row height"), 0, 1)
        grid.Controls.Add(_gridRowHeightNumeric, 1, 1)
        grid.Controls.Add(CreateFieldLabel("Grid header style"), 0, 2)
        grid.Controls.Add(_gridHeaderStyleComboBox, 1, 2)

        Dim statusAndProgress = CreateFieldGrid(5)

        _showBufferStatusRowCheckBox.Text = "Show buffer status row"
        _showRescueStatusRowCheckBox.Text = "Show rescue status row"
        _showDiagnosticsStatusRowCheckBox.Text = "Show diagnostics status row"
        _progressBarShowPercentageCheckBox.Text = "Show percentage overlay on progress bars"

        For Each box As CheckBox In New CheckBox() {
            _showBufferStatusRowCheckBox,
            _showRescueStatusRowCheckBox,
            _showDiagnosticsStatusRowCheckBox,
            _progressBarShowPercentageCheckBox
        }
            ConfigureCheckBox(box)
        Next

        AddHandler _showDiagnosticsStatusRowCheckBox.CheckedChanged, AddressOf AppearanceDiagnosticsVisibilityCheckBox_CheckedChanged

        _progressBarStyleComboBox.DropDownStyle = ComboBoxStyle.DropDownList
        _progressBarStyleComboBox.Items.AddRange(New Object() {"Thin", "Standard", "Thick"})
        ConfigureComboBoxControl(_progressBarStyleComboBox, width:=180)

        statusAndProgress.Controls.Add(_showBufferStatusRowCheckBox, 0, 0)
        statusAndProgress.SetColumnSpan(_showBufferStatusRowCheckBox, 3)
        statusAndProgress.Controls.Add(_showRescueStatusRowCheckBox, 0, 1)
        statusAndProgress.SetColumnSpan(_showRescueStatusRowCheckBox, 3)
        statusAndProgress.Controls.Add(_showDiagnosticsStatusRowCheckBox, 0, 2)
        statusAndProgress.SetColumnSpan(_showDiagnosticsStatusRowCheckBox, 3)
        statusAndProgress.Controls.Add(CreateFieldLabel("Progress bar style"), 0, 3)
        statusAndProgress.Controls.Add(_progressBarStyleComboBox, 1, 3)
        statusAndProgress.Controls.Add(_progressBarShowPercentageCheckBox, 0, 4)
        statusAndProgress.SetColumnSpan(_progressBarShowPercentageCheckBox, 3)

        page.Controls.Add(CreateSection("Theme & Accent", theme), 0, 0)
        page.Controls.Add(CreateSection("Layout", layout), 0, 1)
        page.Controls.Add(CreateSection("Operations Log", log), 0, 2)
        page.Controls.Add(CreateSection("Grid Appearance", grid), 0, 3)
        page.Controls.Add(CreateSection("Status & Progress", statusAndProgress), 0, 4)
        Return page
    End Function
    Private Function BuildCopyDefaultsPage() As Control
        Dim page = CreatePageContainer(3)

        Dim behavior = CreateSingleColumnGrid(8)

        _defaultResumeCheckBox.Text = "Resume from journal"
        _defaultSalvageCheckBox.Text = "Salvage unreadable blocks"
        _defaultContinueOnErrorCheckBox.Text = "Continue on file errors"
        _defaultPreserveTimestampsCheckBox.Text = "Preserve source timestamps"
        _defaultCopyEmptyDirectoriesCheckBox.Text = "Copy empty directories"
        _defaultWaitForMediaCheckBox.Text = "Wait forever for source/destination"
        _defaultWaitForLockReleaseCheckBox.Text = "Wait for lock/contention release"
        _defaultTreatAccessDeniedContentionCheckBox.Text = "Treat Access Denied as transient contention"

        For Each box As CheckBox In New CheckBox() {
            _defaultResumeCheckBox,
            _defaultSalvageCheckBox,
            _defaultContinueOnErrorCheckBox,
            _defaultPreserveTimestampsCheckBox,
            _defaultCopyEmptyDirectoriesCheckBox,
            _defaultWaitForMediaCheckBox,
            _defaultWaitForLockReleaseCheckBox,
            _defaultTreatAccessDeniedContentionCheckBox}
            ConfigureCheckBox(box)
            behavior.Controls.Add(box)
        Next

        Dim policy = CreateFieldGrid(3)

        _overwritePolicyComboBox.DropDownStyle = ComboBoxStyle.DropDownList
        _overwritePolicyComboBox.Items.AddRange(New Object() {"Overwrite existing files", "Skip existing files", "Overwrite only if source is newer", "Always ask for each conflict"})
        ConfigureComboBoxControl(_overwritePolicyComboBox, width:=360)

        _symlinkHandlingComboBox.DropDownStyle = ComboBoxStyle.DropDownList
        _symlinkHandlingComboBox.Items.AddRange(New Object() {"Skip symbolic links", "Follow symbolic links"})
        ConfigureComboBoxControl(_symlinkHandlingComboBox, width:=360)

        _salvageFillPatternComboBox.DropDownStyle = ComboBoxStyle.DropDownList
        _salvageFillPatternComboBox.Items.AddRange(New Object() {"Zero-fill unreadable blocks", "0xFF-fill unreadable blocks", "Random-fill unreadable blocks"})
        ConfigureComboBoxControl(_salvageFillPatternComboBox, width:=360)

        policy.Controls.Add(CreateFieldLabel("Overwrite policy"), 0, 0)
        policy.Controls.Add(_overwritePolicyComboBox, 1, 0)
        policy.Controls.Add(CreateFieldLabel("Symlink handling"), 0, 1)
        policy.Controls.Add(_symlinkHandlingComboBox, 1, 1)
        policy.Controls.Add(CreateFieldLabel("Salvage fill pattern"), 0, 2)
        policy.Controls.Add(_salvageFillPatternComboBox, 1, 2)

        Dim badMap = CreateFieldGrid(5)
        _defaultUseBadRangeMapCheckBox.Text = "Use bad-range map when available"
        _defaultSkipKnownBadRangesCheckBox.Text = "Skip known bad ranges during copy"
        _defaultUpdateBadRangeMapCheckBox.Text = "Update bad-range map from scan/copy runs"
        _defaultUseRawDiskScanCheckBox.Text = "Use experimental raw disk scan backend (local NTFS)"

        For Each box As CheckBox In New CheckBox() {
            _defaultUseBadRangeMapCheckBox,
            _defaultSkipKnownBadRangesCheckBox,
            _defaultUpdateBadRangeMapCheckBox,
            _defaultUseRawDiskScanCheckBox}
            ConfigureCheckBox(box)
        Next

        AddHandler _defaultUseBadRangeMapCheckBox.CheckedChanged, AddressOf BadRangeMapControl_CheckedChanged
        ConfigureNumeric(_defaultBadRangeMapMaxAgeDaysNumeric, 0D, 3650D, 30D)
        _defaultBadRangeMapMaxAgeDaysNumeric.Increment = 5D

        badMap.Controls.Add(_defaultUseBadRangeMapCheckBox, 0, 0)
        badMap.SetColumnSpan(_defaultUseBadRangeMapCheckBox, 3)
        badMap.Controls.Add(_defaultSkipKnownBadRangesCheckBox, 0, 1)
        badMap.SetColumnSpan(_defaultSkipKnownBadRangesCheckBox, 3)
        badMap.Controls.Add(_defaultUpdateBadRangeMapCheckBox, 0, 2)
        badMap.SetColumnSpan(_defaultUpdateBadRangeMapCheckBox, 3)
        badMap.Controls.Add(_defaultUseRawDiskScanCheckBox, 0, 3)
        badMap.SetColumnSpan(_defaultUseRawDiskScanCheckBox, 3)
        badMap.Controls.Add(CreateFieldLabel("Map max age (days, 0=never)"), 0, 4)
        badMap.Controls.Add(_defaultBadRangeMapMaxAgeDaysNumeric, 1, 4)

        page.Controls.Add(CreateSection("Default Run Behavior", behavior), 0, 0)
        page.Controls.Add(CreateSection("Policies", policy), 0, 1)
        page.Controls.Add(CreateSection("Bad Range Map", badMap), 0, 2)
        Return page
    End Function

    Private Function BuildPerformancePage() As Control
        Dim page = CreatePageContainer(4)

        Dim body = CreateFieldGrid(12)

        _defaultAdaptiveBufferCheckBox.Text = "Enable adaptive buffer by default"
        ConfigureCheckBox(_defaultAdaptiveBufferCheckBox)
        AddHandler _defaultAdaptiveBufferCheckBox.CheckedChanged, AddressOf DefaultAdaptiveBufferCheckBox_CheckedChanged

        _transferEnginePolicyComboBox.DropDownStyle = ComboBoxStyle.DropDownList
        _transferEnginePolicyComboBox.Items.AddRange(New Object() {"Auto", "Managed rescue", "Native fast"})
        ConfigureComboBoxControl(_transferEnginePolicyComboBox, width:=240)

        _scanPerformanceProfileComboBox.DropDownStyle = ComboBoxStyle.DropDownList
        _scanPerformanceProfileComboBox.Items.AddRange(New Object() {"Auto fast scan", "Fast health scan", "Precise bad-range scan"})
        ConfigureComboBoxControl(_scanPerformanceProfileComboBox, width:=240)

        _workerProcessPriorityComboBox.DropDownStyle = ComboBoxStyle.DropDownList
        _workerProcessPriorityComboBox.Items.AddRange(New Object() {"Idle", "Below normal", "Normal", "Above normal", "High"})
        ConfigureComboBoxControl(_workerProcessPriorityComboBox, width:=240)

        ConfigureNumeric(_defaultBufferMbNumeric, 1D, 256D, 4D)
        ConfigureNumeric(_defaultRetriesNumeric, 0D, 1000D, 12D)
        ConfigureNumeric(_defaultOperationTimeoutNumeric, 1D, 3600D, 10D)
        ConfigureNumeric(_defaultPerFileTimeoutNumeric, 0D, 86400D, 0D)
        ConfigureNumeric(_defaultMaxThroughputNumeric, 0D, 4096D, 0D)
        ConfigureNumeric(_parallelSmallFileWorkersNumeric, 0D, 64D, 0D)
        ConfigureNumeric(_parallelScanWorkersNumeric, 0D, 64D, 0D)
        ConfigureNumeric(_smallFileThresholdKbNumeric, 4D, 1048576D, 256D)

        body.Controls.Add(_defaultAdaptiveBufferCheckBox, 0, 0)
        body.SetColumnSpan(_defaultAdaptiveBufferCheckBox, 3)
        body.Controls.Add(CreateFieldLabel("Transfer engine"), 0, 1)
        body.Controls.Add(_transferEnginePolicyComboBox, 1, 1)
        body.Controls.Add(CreateFieldLabel("Scan profile"), 0, 2)
        body.Controls.Add(_scanPerformanceProfileComboBox, 1, 2)
        body.Controls.Add(CreateFieldLabel("Worker priority"), 0, 3)
        body.Controls.Add(_workerProcessPriorityComboBox, 1, 3)
        body.Controls.Add(CreateFieldLabel("Manual buffer MB"), 0, 4)
        body.Controls.Add(_defaultBufferMbNumeric, 1, 4)
        body.Controls.Add(CreateFieldLabel("Max retries"), 0, 5)
        body.Controls.Add(_defaultRetriesNumeric, 1, 5)
        body.Controls.Add(CreateFieldLabel("Operation timeout (sec)"), 0, 6)
        body.Controls.Add(_defaultOperationTimeoutNumeric, 1, 6)
        body.Controls.Add(CreateFieldLabel("Per-file timeout (sec, 0=disabled)"), 0, 7)
        body.Controls.Add(_defaultPerFileTimeoutNumeric, 1, 7)
        body.Controls.Add(CreateFieldLabel("Max throughput (MB/s, 0=unlimited)"), 0, 8)
        body.Controls.Add(_defaultMaxThroughputNumeric, 1, 8)
        body.Controls.Add(CreateFieldLabel("Small-file workers (0=auto, 1=off)"), 0, 9)
        body.Controls.Add(_parallelSmallFileWorkersNumeric, 1, 9)
        body.Controls.Add(CreateFieldLabel("Scan workers (0=auto, 1=off)"), 0, 10)
        body.Controls.Add(_parallelScanWorkersNumeric, 1, 10)
        body.Controls.Add(CreateFieldLabel("Small-file threshold KB"), 0, 11)
        body.Controls.Add(_smallFileThresholdKbNumeric, 1, 11)

        Dim fragile = CreateFieldGrid(6)
        _defaultFragileModeCheckBox.Text = "Enable fragile media mode by default"
        _defaultSkipFileOnFirstReadErrorCheckBox.Text = "Skip file on first read error (fragile mode)"
        _defaultPersistFragileSkipsCheckBox.Text = "Persist skipped files across resume/recovery"
        ConfigureCheckBox(_defaultFragileModeCheckBox)
        ConfigureCheckBox(_defaultSkipFileOnFirstReadErrorCheckBox)
        ConfigureCheckBox(_defaultPersistFragileSkipsCheckBox)
        AddHandler _defaultFragileModeCheckBox.CheckedChanged, AddressOf FragileControl_CheckedChanged

        ConfigureNumeric(_defaultFragileFailureWindowNumeric, 1D, 3600D, 20D)
        ConfigureNumeric(_defaultFragileFailureThresholdNumeric, 1D, 1000D, 3D)
        ConfigureNumeric(_defaultFragileCooldownNumeric, 0D, 600D, 6D)

        fragile.Controls.Add(_defaultFragileModeCheckBox, 0, 0)
        fragile.SetColumnSpan(_defaultFragileModeCheckBox, 3)
        fragile.Controls.Add(_defaultSkipFileOnFirstReadErrorCheckBox, 0, 1)
        fragile.SetColumnSpan(_defaultSkipFileOnFirstReadErrorCheckBox, 3)
        fragile.Controls.Add(_defaultPersistFragileSkipsCheckBox, 0, 2)
        fragile.SetColumnSpan(_defaultPersistFragileSkipsCheckBox, 3)
        fragile.Controls.Add(CreateFieldLabel("Failure window (sec)"), 0, 3)
        fragile.Controls.Add(_defaultFragileFailureWindowNumeric, 1, 3)
        fragile.Controls.Add(CreateFieldLabel("Failure threshold (files)"), 0, 4)
        fragile.Controls.Add(_defaultFragileFailureThresholdNumeric, 1, 4)
        fragile.Controls.Add(CreateFieldLabel("Cooldown (sec, 0=off)"), 0, 5)
        fragile.Controls.Add(_defaultFragileCooldownNumeric, 1, 5)

        Dim contention = CreateFieldGrid(2)
        ConfigureNumeric(_defaultLockProbeIntervalNumeric, 100D, 10000D, 500D)

        _defaultSourceMutationPolicyComboBox.DropDownStyle = ComboBoxStyle.DropDownList
        _defaultSourceMutationPolicyComboBox.Items.AddRange(New Object() {
            "Fail file when source disappears",
            "Skip missing source file",
            "Wait for source file to reappear"})
        ConfigureComboBoxControl(_defaultSourceMutationPolicyComboBox, width:=360)

        contention.Controls.Add(CreateFieldLabel("Lock probe interval (ms)"), 0, 0)
        contention.Controls.Add(_defaultLockProbeIntervalNumeric, 1, 0)
        contention.Controls.Add(CreateFieldLabel("Source mutation policy"), 0, 1)
        contention.Controls.Add(_defaultSourceMutationPolicyComboBox, 1, 1)

        Dim rescue = CreateFieldGrid(8)

        ConfigureNumeric(_rescueFastChunkKbNumeric, 0D, 262144D, 0D)
        ConfigureNumeric(_rescueTrimChunkKbNumeric, 0D, 262144D, 0D)
        ConfigureNumeric(_rescueScrapeChunkKbNumeric, 0D, 262144D, 0D)
        ConfigureNumeric(_rescueRetryChunkKbNumeric, 0D, 262144D, 0D)
        ConfigureNumeric(_rescueSplitMinimumKbNumeric, 0D, 65536D, 0D)
        ConfigureNumeric(_rescueFastRetriesNumeric, 0D, 1000D, 0D)
        ConfigureNumeric(_rescueTrimRetriesNumeric, 0D, 1000D, 1D)
        ConfigureNumeric(_rescueScrapeRetriesNumeric, 0D, 1000D, 2D)

        rescue.Controls.Add(CreateFieldLabel("FastScan chunk KB (0=auto)"), 0, 0)
        rescue.Controls.Add(_rescueFastChunkKbNumeric, 1, 0)
        rescue.Controls.Add(CreateFieldLabel("TrimSweep chunk KB (0=auto)"), 0, 1)
        rescue.Controls.Add(_rescueTrimChunkKbNumeric, 1, 1)
        rescue.Controls.Add(CreateFieldLabel("Scrape chunk KB (0=auto)"), 0, 2)
        rescue.Controls.Add(_rescueScrapeChunkKbNumeric, 1, 2)
        rescue.Controls.Add(CreateFieldLabel("RetryBad chunk KB (0=auto)"), 0, 3)
        rescue.Controls.Add(_rescueRetryChunkKbNumeric, 1, 3)
        rescue.Controls.Add(CreateFieldLabel("Split minimum KB (0=auto)"), 0, 4)
        rescue.Controls.Add(_rescueSplitMinimumKbNumeric, 1, 4)
        rescue.Controls.Add(CreateFieldLabel("FastScan retries"), 0, 5)
        rescue.Controls.Add(_rescueFastRetriesNumeric, 1, 5)
        rescue.Controls.Add(CreateFieldLabel("TrimSweep retries"), 0, 6)
        rescue.Controls.Add(_rescueTrimRetriesNumeric, 1, 6)
        rescue.Controls.Add(CreateFieldLabel("Scrape retries"), 0, 7)
        rescue.Controls.Add(_rescueScrapeRetriesNumeric, 1, 7)

        page.Controls.Add(CreateSection("Transfer Tuning", body), 0, 0)
        page.Controls.Add(CreateSection("Fragile Media Guard", fragile), 0, 1)
        page.Controls.Add(CreateSection("Contention & Source Mutation", contention), 0, 2)
        page.Controls.Add(CreateSection("Rescue Engine Tuning", rescue), 0, 3)
        Return page
    End Function

    Private Function BuildDiagnosticsPage() As Control
        Dim page = CreatePageContainer(2)

        Dim workerTelemetry = CreateFieldGrid(3)

        _workerTelemetryProfileComboBox.DropDownStyle = ComboBoxStyle.DropDownList
        _workerTelemetryProfileComboBox.Items.AddRange(New Object() {"Normal", "Verbose", "Debug"})
        ConfigureComboBoxControl(_workerTelemetryProfileComboBox, width:=240)

        ConfigureNumeric(_workerProgressIntervalNumeric, 20D, 1000D, 75D)
        ConfigureNumeric(_workerMaxLogsPerSecondNumeric, 0D, 5000D, 100D)

        workerTelemetry.Controls.Add(CreateFieldLabel("Worker telemetry profile"), 0, 0)
        workerTelemetry.Controls.Add(_workerTelemetryProfileComboBox, 1, 0)
        workerTelemetry.Controls.Add(CreateFieldLabel("Progress interval (ms)"), 0, 1)
        workerTelemetry.Controls.Add(_workerProgressIntervalNumeric, 1, 1)
        workerTelemetry.Controls.Add(CreateFieldLabel("Log rate cap (events/sec, 0=unlimited)"), 0, 2)
        workerTelemetry.Controls.Add(_workerMaxLogsPerSecondNumeric, 1, 2)

        Dim uiDiagnostics = CreateFieldGrid(3)

        _uiShowDiagnosticsCheckBox.Text = "Show UI diagnostics strip on main window"
        ConfigureCheckBox(_uiShowDiagnosticsCheckBox)
        AddHandler _uiShowDiagnosticsCheckBox.CheckedChanged, AddressOf DiagnosticsControl_CheckedChanged

        ConfigureNumeric(_uiDiagnosticsRefreshNumeric, 100D, 5000D, 250D)
        ConfigureNumeric(_uiMaxLogLinesNumeric, 1000D, 1000000D, 50000D)
        _uiMaxLogLinesNumeric.Increment = 1000D

        uiDiagnostics.Controls.Add(_uiShowDiagnosticsCheckBox, 0, 0)
        uiDiagnostics.SetColumnSpan(_uiShowDiagnosticsCheckBox, 3)
        uiDiagnostics.Controls.Add(CreateFieldLabel("Diagnostics refresh interval (ms)"), 0, 1)
        uiDiagnostics.Controls.Add(_uiDiagnosticsRefreshNumeric, 1, 1)
        uiDiagnostics.Controls.Add(CreateFieldLabel("Virtual log max lines"), 0, 2)
        uiDiagnostics.Controls.Add(_uiMaxLogLinesNumeric, 1, 2)

        page.Controls.Add(CreateSection("Worker Telemetry", workerTelemetry), 0, 0)
        page.Controls.Add(CreateSection("UI Diagnostics", uiDiagnostics), 0, 1)
        Return page
    End Function

    Private Function BuildVerificationPage() As Control
        Dim page = CreatePageContainer(2)

        Dim defaults = CreateFieldGrid(3)

        _defaultVerifyCheckBox.Text = "Enable verification by default"
        ConfigureCheckBox(_defaultVerifyCheckBox)
        AddHandler _defaultVerifyCheckBox.CheckedChanged, AddressOf DefaultVerifyCheckBox_CheckedChanged

        _defaultVerificationModeComboBox.DropDownStyle = ComboBoxStyle.DropDownList
        _defaultVerificationModeComboBox.Items.AddRange(New Object() {"Full hash", "Sampled hash"})
        ConfigureComboBoxControl(_defaultVerificationModeComboBox, width:=240)
        AddHandler _defaultVerificationModeComboBox.SelectedIndexChanged, AddressOf DefaultVerifyCheckBox_CheckedChanged

        _defaultHashAlgorithmComboBox.DropDownStyle = ComboBoxStyle.DropDownList
        _defaultHashAlgorithmComboBox.Items.AddRange(New Object() {"SHA-256", "SHA-512"})
        ConfigureComboBoxControl(_defaultHashAlgorithmComboBox, width:=240)

        defaults.Controls.Add(_defaultVerifyCheckBox, 0, 0)
        defaults.SetColumnSpan(_defaultVerifyCheckBox, 3)
        defaults.Controls.Add(CreateFieldLabel("Verification mode"), 0, 1)
        defaults.Controls.Add(_defaultVerificationModeComboBox, 1, 1)
        defaults.Controls.Add(CreateFieldLabel("Hash algorithm"), 0, 2)
        defaults.Controls.Add(_defaultHashAlgorithmComboBox, 1, 2)

        Dim sample = CreateFieldGrid(2)

        ConfigureNumeric(_sampleChunkKbNumeric, 32D, 4096D, 128D)
        ConfigureNumeric(_sampleChunkCountNumeric, 1D, 64D, 3D)

        sample.Controls.Add(CreateFieldLabel("Sample chunk size (KB)"), 0, 0)
        sample.Controls.Add(_sampleChunkKbNumeric, 1, 0)
        sample.Controls.Add(CreateFieldLabel("Sample count"), 0, 1)
        sample.Controls.Add(_sampleChunkCountNumeric, 1, 1)

        page.Controls.Add(CreateSection("Verification Defaults", defaults), 0, 0)
        page.Controls.Add(CreateSection("Sampled Mode", sample), 0, 1)
        Return page
    End Function

    Private Function BuildUpdatesPage() As Control
        Dim page = CreatePageContainer(2)

        Dim behavior = CreateSingleColumnGrid(1)
        _checkUpdatesOnLaunchCheckBox.Text = "Check for updates on launch"
        ConfigureCheckBox(_checkUpdatesOnLaunchCheckBox)
        behavior.Controls.Add(_checkUpdatesOnLaunchCheckBox)

        Dim source = CreateFieldGrid(2)

        ConfigureTextBoxControl(_updateUrlTextBox, stretch:=True)
        _updateUrlTextBox.PlaceholderText = AppSettings.DefaultUpdateReleaseUrl
        ConfigureTextBoxControl(_userAgentTextBox, stretch:=True)
        _userAgentTextBox.PlaceholderText = "XactCopy/1.0"

        source.Controls.Add(CreateFieldLabel("Release URL"), 0, 0)
        source.Controls.Add(_updateUrlTextBox, 1, 0)
        source.SetColumnSpan(_updateUrlTextBox, 2)
        source.Controls.Add(CreateFieldLabel("User-Agent"), 0, 1)
        source.Controls.Add(_userAgentTextBox, 1, 1)
        source.SetColumnSpan(_userAgentTextBox, 2)

        page.Controls.Add(CreateSection("Behavior", behavior), 0, 0)
        page.Controls.Add(CreateSection("Release Source", source), 0, 1)
        Return page
    End Function
    Private Function BuildRecoveryPage() As Control
        Dim page = CreatePageContainer(2)

        Dim startup = CreateSingleColumnGrid(2)

        _enableRecoveryAutostartCheckBox.Text = "Auto-start on next logon when a run is interrupted"
        ConfigureCheckBox(_enableRecoveryAutostartCheckBox)
        _autoRunQueuedOnStartupCheckBox.Text = "Automatically run queued jobs on startup"
        ConfigureCheckBox(_autoRunQueuedOnStartupCheckBox)

        startup.Controls.Add(_enableRecoveryAutostartCheckBox, 0, 0)
        startup.Controls.Add(_autoRunQueuedOnStartupCheckBox, 0, 1)

        Dim handling = CreateFieldGrid(4)

        _promptResumeAfterCrashCheckBox.Text = "Prompt to resume interrupted runs"
        ConfigureCheckBox(_promptResumeAfterCrashCheckBox)
        _autoResumeAfterCrashCheckBox.Text = "Auto-resume interrupted runs"
        ConfigureCheckBox(_autoResumeAfterCrashCheckBox)
        _keepResumePromptCheckBox.Text = "Keep prompting until resumed or replaced"
        ConfigureCheckBox(_keepResumePromptCheckBox)

        AddHandler _promptResumeAfterCrashCheckBox.CheckedChanged, AddressOf RecoveryControl_CheckedChanged
        AddHandler _autoResumeAfterCrashCheckBox.CheckedChanged, AddressOf RecoveryControl_CheckedChanged

        ConfigureNumeric(_recoveryTouchIntervalNumeric, 1D, 60D, 2D)

        handling.Controls.Add(_promptResumeAfterCrashCheckBox, 0, 0)
        handling.SetColumnSpan(_promptResumeAfterCrashCheckBox, 3)
        handling.Controls.Add(_autoResumeAfterCrashCheckBox, 0, 1)
        handling.SetColumnSpan(_autoResumeAfterCrashCheckBox, 3)
        handling.Controls.Add(_keepResumePromptCheckBox, 0, 2)
        handling.SetColumnSpan(_keepResumePromptCheckBox, 3)
        handling.Controls.Add(CreateFieldLabel("Recovery heartbeat write interval (sec)"), 0, 3)
        handling.Controls.Add(_recoveryTouchIntervalNumeric, 1, 3)

        page.Controls.Add(CreateSection("Startup", startup), 0, 0)
        page.Controls.Add(CreateSection("Interrupted Run Handling", handling), 0, 1)
        Return page
    End Function

    Private Function BuildExplorerPage() As Control
        Dim page = CreatePageContainer(1)

        Dim integration = CreateFieldGrid(3)

        _enableExplorerContextMenuCheckBox.Text = "Enable Explorer context menu integration"
        ConfigureCheckBox(_enableExplorerContextMenuCheckBox)
        AddHandler _enableExplorerContextMenuCheckBox.CheckedChanged, AddressOf ExplorerControl_CheckedChanged

        _explorerSelectionModeComboBox.DropDownStyle = ComboBoxStyle.DropDownList
        _explorerSelectionModeComboBox.Items.AddRange(New Object() {"Copy selected files/folders", "Copy full source folder"})
        ConfigureComboBoxControl(_explorerSelectionModeComboBox, width:=320)

        Dim detailLabel As New Label() With {
            .Text = "Adds 'Copy with XactCopy' to Explorer for files, folders, drives, and folder background.",
            .AutoSize = False,
            .Dock = DockStyle.Fill,
            .TextAlign = ContentAlignment.MiddleLeft,
            .Margin = New Padding(0, 2, 0, 0)
        }

        integration.Controls.Add(_enableExplorerContextMenuCheckBox, 0, 0)
        integration.SetColumnSpan(_enableExplorerContextMenuCheckBox, 3)
        integration.Controls.Add(CreateFieldLabel("File selection mode"), 0, 1)
        integration.Controls.Add(_explorerSelectionModeComboBox, 1, 1)
        integration.Controls.Add(detailLabel, 0, 2)
        integration.SetColumnSpan(detailLabel, 3)

        page.Controls.Add(CreateSection("Context Menu", integration), 0, 0)
        Return page
    End Function

    Private Shared Sub ConfigureNumeric(control As ThemedNumericUpDown, minimum As Decimal, maximum As Decimal, defaultValue As Decimal)
        control.Minimum = minimum
        control.Maximum = maximum
        control.Value = defaultValue
        control.DecimalPlaces = 0
        control.Increment = 1D
        control.ReadOnly = False
        control.ThousandsSeparator = False
        control.TextAlign = HorizontalAlignment.Left
        control.TabStop = True
        control.Enabled = True
        control.Dock = DockStyle.None
        control.Anchor = AnchorStyles.Left
        control.Width = 120
        control.Margin = New Padding(0, 4, 0, 4)
    End Sub

    Private Shared Sub ConfigureComboBoxControl(control As ComboBox, Optional width As Integer = 320)
        control.Dock = DockStyle.None
        control.Anchor = AnchorStyles.Left
        control.Width = width
        control.Margin = New Padding(0, 4, 0, 4)
        control.IntegralHeight = False
        control.DropDownHeight = 260
    End Sub

    Private Shared Sub ConfigureTextBoxControl(control As TextBox, Optional stretch As Boolean = False)
        control.Dock = If(stretch, DockStyle.Fill, DockStyle.None)
        control.Anchor = If(stretch, AnchorStyles.Left Or AnchorStyles.Right, AnchorStyles.Left)
        If Not stretch Then
            control.Width = 360
        End If
        control.Margin = New Padding(0, 4, 0, 4)
    End Sub

    Private Shared Sub ConfigureCheckBox(control As CheckBox)
        control.AutoSize = True
        control.Dock = DockStyle.None
        control.Anchor = AnchorStyles.Left
        control.Margin = New Padding(0, 3, 0, 3)
        control.Padding = New Padding(0)
    End Sub

    Private Sub PopulateLogFontFamilyComboBox()
        If _logFontFamilyComboBox.Items.Count > 0 Then
            Return
        End If

        Dim names = GetLogFontNameCache()
        Dim items(names.Count - 1) As Object
        For i = 0 To names.Count - 1
            items(i) = names(i)
        Next
        _logFontFamilyComboBox.Items.AddRange(items)
    End Sub

    ''' <summary>
    ''' Pre-populates the system font name cache on a background thread so the
    ''' first Settings dialog open does not block on GDI+ font enumeration.
    ''' </summary>
    Friend Shared Sub WarmFontCache()
        GetLogFontNameCache()
    End Sub

    Private Shared Function GetLogFontNameCache() As IReadOnlyList(Of String)
        Dim cached = _cachedLogFontNames
        If cached IsNot Nothing Then
            Return cached
        End If

        SyncLock _fontCacheSyncRoot
            cached = _cachedLogFontNames
            If cached IsNot Nothing Then
                Return cached
            End If

            Dim preferredFonts = New String() {"Consolas", "Cascadia Mono", "Lucida Console", "Courier New", "Segoe UI"}
            Dim fontNames As New SortedSet(Of String)(StringComparer.OrdinalIgnoreCase)

            For Each fontName As String In preferredFonts
                fontNames.Add(fontName)
            Next

            Try
                For Each family As FontFamily In FontFamily.Families
                    fontNames.Add(family.Name)
                Next
            Catch
                ' Ignore font-enumeration failures and keep fallback list.
            End Try

            _cachedLogFontNames = fontNames.ToList()
            Return _cachedLogFontNames
        End SyncLock
    End Function

    Private Shared Function CreateFieldGrid(rowCount As Integer) As TableLayoutPanel
        Dim grid As New TableLayoutPanel() With {
            .Dock = DockStyle.Top,
            .AutoSize = True,
            .AutoSizeMode = AutoSizeMode.GrowAndShrink,
            .ColumnCount = 3,
            .RowCount = rowCount,
            .Margin = New Padding(0),
            .Padding = New Padding(0)
        }
        grid.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 300.0F))
        grid.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 360.0F))
        grid.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0F))
        For index = 0 To rowCount - 1
            grid.RowStyles.Add(New RowStyle(SizeType.AutoSize))
        Next
        Return grid
    End Function

    Private Shared Function CreateSingleColumnGrid(rowCount As Integer) As TableLayoutPanel
        Dim grid As New TableLayoutPanel() With {
            .Dock = DockStyle.Top,
            .AutoSize = True,
            .AutoSizeMode = AutoSizeMode.GrowAndShrink,
            .ColumnCount = 1,
            .RowCount = rowCount,
            .Margin = New Padding(0),
            .Padding = New Padding(0)
        }
        grid.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0F))
        For index = 0 To rowCount - 1
            grid.RowStyles.Add(New RowStyle(SizeType.AutoSize))
        Next
        Return grid
    End Function

    Private Shared Function CreateFieldLabel(text As String) As Label
        Return New Label() With {
            .Text = text,
            .AutoSize = False,
            .Dock = DockStyle.Fill,
            .TextAlign = ContentAlignment.MiddleLeft,
            .Margin = New Padding(0, 0, 10, 0)
        }
    End Function

    Private Function CreatePageContainer(sectionCount As Integer) As TableLayoutPanel
        Dim page As New TableLayoutPanel() With {
            .Dock = DockStyle.Top,
            .AutoSize = True,
            .AutoSizeMode = AutoSizeMode.GrowAndShrink,
            .ColumnCount = 1,
            .RowCount = sectionCount,
            .Padding = New Padding(2),
            .Margin = New Padding(0)
        }
        page.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0F))
        For index = 0 To sectionCount - 1
            page.RowStyles.Add(New RowStyle(SizeType.AutoSize))
        Next
        Return page
    End Function

    Private Function CreateSection(title As String, body As Control) As Control
        Dim section As New TableLayoutPanel() With {
            .Dock = DockStyle.Top,
            .ColumnCount = 1,
            .RowCount = 2,
            .AutoSize = True,
            .AutoSizeMode = AutoSizeMode.GrowAndShrink,
            .Margin = New Padding(0, 0, 0, 10)
        }
        section.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0F))
        section.RowStyles.Add(New RowStyle(SizeType.AutoSize))
        section.RowStyles.Add(New RowStyle(SizeType.AutoSize))

        Dim titleLabel As New Label() With {.Text = title, .AutoSize = True, .Font = New Font(Font, FontStyle.Bold), .Margin = New Padding(0, 0, 0, 6)}

        Dim bodyHost As New Panel() With {
            .Dock = DockStyle.Top,
            .AutoSize = True,
            .AutoSizeMode = AutoSizeMode.GrowAndShrink,
            .Padding = New Padding(10, 2, 10, 2)
        }
        body.Dock = DockStyle.Top
        bodyHost.Controls.Add(body)

        section.Controls.Add(titleLabel, 0, 0)
        section.Controls.Add(bodyHost, 0, 1)
        Return section
    End Function

    Private Sub SettingsForm_Load(sender As Object, e As EventArgs)
        SuspendNativeRedraw()
        SuspendLayout()
        Try
            UiLayoutManager.ApplyWindow(Me, LayoutWindowKey)
            UiLayoutManager.ApplySplitter(_mainSplit, LayoutSplitterKey, defaultDistance:=210, panel1Min:=170, panel2Min:=500)
            _categoryTreeView.BeginUpdate()
            If _categoryTreeView.Nodes.Count > 0 AndAlso _categoryTreeView.SelectedNode Is Nothing Then
                _categoryTreeView.SelectedNode = _categoryTreeView.Nodes(0)
            End If
            If _activePageControl Is Nothing Then
                ShowPage(If(_categoryTreeView.SelectedNode?.Name, "appearance"))
            End If
            ThemeManager.ApplyTheme(Me, ThemeSettings.GetPreferredColorMode(_workingSettings), _workingSettings)
        Finally
            _categoryTreeView.EndUpdate()
            ResumeLayout(performLayout:=True)
        End Try
    End Sub

    Private Sub SettingsForm_Shown(sender As Object, e As EventArgs)
        If Not _initialRevealPending Then
            Return
        End If

        _initialRevealPending = False
        ResumeNativeRedraw()
        Opacity = 1.0R
        Invalidate(invalidateChildren:=True)
        Update()
    End Sub

    Private Shared Sub EnableDoubleBufferingForContainers(root As Control)
        If root Is Nothing OrElse _doubleBufferedProperty Is Nothing Then
            Return
        End If

        Try
            _doubleBufferedProperty.SetValue(root, True, Nothing)
        Catch
        End Try

        For Each child As Control In root.Controls
            EnableDoubleBufferingForContainers(child)
        Next
    End Sub

    Private Sub SettingsForm_FormClosing(sender As Object, e As FormClosingEventArgs)
        _dirtyStateTimer.Stop()
        ResumeNativeRedraw()
        UiLayoutManager.CaptureSplitter(_mainSplit, LayoutSplitterKey)
        UiLayoutManager.CaptureWindow(Me, LayoutWindowKey)
    End Sub

    Private Sub CategoryTreeView_AfterSelect(sender As Object, e As TreeViewEventArgs)
        ShowPage(If(e.Node?.Name, String.Empty))
    End Sub

    Private Sub ShowPage(pageKey As String)
        Dim normalizedPageKey = NormalizePageKey(pageKey)
        If _activePageControl IsNot Nothing AndAlso
            String.Equals(_activePageKey, normalizedPageKey, StringComparison.OrdinalIgnoreCase) Then
            UpdatePageTitle(normalizedPageKey)
            Return
        End If

        _pageHostPanel.SuspendLayout()
        Try
            Dim selectedPage = EnsurePage(normalizedPageKey)
            If selectedPage Is Nothing Then
                normalizedPageKey = "appearance"
                selectedPage = EnsurePage(normalizedPageKey)
            End If

            If _activePageControl IsNot Nothing AndAlso Not Object.ReferenceEquals(_activePageControl, selectedPage) Then
                _activePageControl.Visible = False
            End If

            If selectedPage IsNot Nothing Then
                selectedPage.Visible = True
                selectedPage.BringToFront()
                _pageHostPanel.AutoScrollPosition = New Point(0, 0)
                _activePageControl = selectedPage
                _activePageKey = normalizedPageKey
            End If
        Finally
            _pageHostPanel.ResumeLayout(performLayout:=True)
        End Try

        UpdatePageTitle(normalizedPageKey)
    End Sub

    Private Function EnsurePage(pageKey As String) As Control
        Dim page As Control = Nothing
        If _pageLookup.TryGetValue(pageKey, page) Then
            Return page
        End If

        Dim builder As Func(Of Control) = Nothing
        If Not _pageBuilders.TryGetValue(pageKey, builder) Then
            Return Nothing
        End If

        page = builder()
        page.Dock = DockStyle.Top
        page.AutoSize = True
        page.Visible = False
        _pageHostPanel.Controls.Add(page)
        _pageLookup(pageKey) = page
        _pageBuilders.Remove(pageKey)
        EnableDoubleBufferingForContainers(page)
        WireDirtyTracking(page.Controls)
        ApplySettingsToControls()
        ThemeManager.ApplyTheme(page, ThemeSettings.GetPreferredColorMode(_workingSettings), _workingSettings)
        Return page
    End Function

    Private Shared Function NormalizePageKey(pageKey As String) As String
        If String.IsNullOrWhiteSpace(pageKey) Then
            Return "appearance"
        End If

        Return pageKey.Trim().ToLowerInvariant()
    End Function

    Private Sub UpdatePageTitle(pageKey As String)
        Dim normalizedPageKey = NormalizePageKey(pageKey)
        Select Case normalizedPageKey
            Case "copy"
                _pageTitleLabel.Text = "Copy Defaults"
            Case "performance"
                _pageTitleLabel.Text = "Performance"
            Case "diagnostics"
                _pageTitleLabel.Text = "Diagnostics"
            Case "verification"
                _pageTitleLabel.Text = "Verification"
            Case "updates"
                _pageTitleLabel.Text = "Updates"
            Case "recovery"
                _pageTitleLabel.Text = "Recovery & Startup"
            Case "explorer"
                _pageTitleLabel.Text = "Explorer Integration"
            Case Else
                _pageTitleLabel.Text = "Appearance"
        End Select
    End Sub

    Private Sub ApplySettingsToControls()
        SuspendLayout()
        _mainSplit.SuspendLayout()
        _pageHostPanel.SuspendLayout()
        _suspendDirtyTracking = True
        Try
            _accentColorTextBox.Text = SettingsValueConverter.NormalizeColorHex(_workingSettings.AccentColorHex, "#5A78C8")
            _windowChromeModeComboBox.SelectedIndex = If(
                SettingsValueConverter.ToWindowChromeModeChoice(_workingSettings.WindowChromeMode) = WindowChromeModeChoice.Standard,
                1,
                0)
            _uiScaleComboBox.SelectedItem = $"{Math.Max(90, Math.Min(125, _workingSettings.UiScalePercent))}%"
            _logFontSizeNumeric.Value = ClampNumeric(_logFontSizeNumeric, _workingSettings.LogFontSizePoints)
            _logColorizeBySeverityCheckBox.Checked = _workingSettings.UiColorizeLogBySeverity
            _gridAlternatingRowsCheckBox.Checked = _workingSettings.GridAlternatingRows
            _gridRowHeightNumeric.Value = ClampNumeric(_gridRowHeightNumeric, _workingSettings.GridRowHeight)
            _showBufferStatusRowCheckBox.Checked = _workingSettings.ShowBufferStatusRow
            _showRescueStatusRowCheckBox.Checked = _workingSettings.ShowRescueStatusRow
            _showDiagnosticsStatusRowCheckBox.Checked = _workingSettings.UiShowDiagnostics
            _progressBarShowPercentageCheckBox.Checked = _workingSettings.ProgressBarShowPercentage

            _checkUpdatesOnLaunchCheckBox.Checked = _workingSettings.CheckUpdatesOnLaunch
            _updateUrlTextBox.Text = If(_workingSettings.UpdateReleaseUrl, String.Empty)
            _userAgentTextBox.Text = If(_workingSettings.UserAgent, String.Empty)

            _enableRecoveryAutostartCheckBox.Checked = _workingSettings.EnableRecoveryAutostart
            _promptResumeAfterCrashCheckBox.Checked = _workingSettings.PromptResumeAfterCrash
            _autoResumeAfterCrashCheckBox.Checked = _workingSettings.AutoResumeAfterCrash
            _keepResumePromptCheckBox.Checked = _workingSettings.KeepResumePromptUntilResolved
            _autoRunQueuedOnStartupCheckBox.Checked = _workingSettings.AutoRunQueuedJobsOnStartup
            _recoveryTouchIntervalNumeric.Value = ClampNumeric(_recoveryTouchIntervalNumeric, _workingSettings.RecoveryTouchIntervalSeconds)

            _enableExplorerContextMenuCheckBox.Checked = _workingSettings.EnableExplorerContextMenu

            _defaultResumeCheckBox.Checked = _workingSettings.DefaultResumeFromJournal
            _defaultSalvageCheckBox.Checked = _workingSettings.DefaultSalvageUnreadableBlocks
            _defaultContinueOnErrorCheckBox.Checked = _workingSettings.DefaultContinueOnFileError
            _defaultPreserveTimestampsCheckBox.Checked = _workingSettings.DefaultPreserveTimestamps
            _defaultCopyEmptyDirectoriesCheckBox.Checked = _workingSettings.DefaultCopyEmptyDirectories
            _defaultWaitForMediaCheckBox.Checked = _workingSettings.DefaultWaitForMediaAvailability
            _defaultWaitForLockReleaseCheckBox.Checked = _workingSettings.DefaultWaitForFileLockRelease
            _defaultTreatAccessDeniedContentionCheckBox.Checked = _workingSettings.DefaultTreatAccessDeniedAsContention
            _defaultUseBadRangeMapCheckBox.Checked = _workingSettings.DefaultUseBadRangeMap
            _defaultSkipKnownBadRangesCheckBox.Checked = _workingSettings.DefaultSkipKnownBadRanges
            _defaultUpdateBadRangeMapCheckBox.Checked = _workingSettings.DefaultUpdateBadRangeMapFromRun
            _defaultUseRawDiskScanCheckBox.Checked = _workingSettings.DefaultUseExperimentalRawDiskScan
            _defaultBadRangeMapMaxAgeDaysNumeric.Value = ClampNumeric(_defaultBadRangeMapMaxAgeDaysNumeric, _workingSettings.DefaultBadRangeMapMaxAgeDays)

            _defaultAdaptiveBufferCheckBox.Checked = _workingSettings.DefaultUseAdaptiveBuffer
            _defaultFragileModeCheckBox.Checked = _workingSettings.DefaultFragileMediaMode
            _defaultSkipFileOnFirstReadErrorCheckBox.Checked = _workingSettings.DefaultSkipFileOnFirstReadError
            _defaultPersistFragileSkipsCheckBox.Checked = _workingSettings.DefaultPersistFragileSkipsAcrossResume
            _defaultBufferMbNumeric.Value = ClampNumeric(_defaultBufferMbNumeric, _workingSettings.DefaultBufferSizeMb)
            _defaultRetriesNumeric.Value = ClampNumeric(_defaultRetriesNumeric, _workingSettings.DefaultMaxRetries)
            _defaultOperationTimeoutNumeric.Value = ClampNumeric(_defaultOperationTimeoutNumeric, _workingSettings.DefaultOperationTimeoutSeconds)
            _defaultPerFileTimeoutNumeric.Value = ClampNumeric(_defaultPerFileTimeoutNumeric, _workingSettings.DefaultPerFileTimeoutSeconds)
            _defaultMaxThroughputNumeric.Value = ClampNumeric(_defaultMaxThroughputNumeric, _workingSettings.DefaultMaxThroughputMbPerSecond)
            _parallelSmallFileWorkersNumeric.Value = ClampNumeric(_parallelSmallFileWorkersNumeric, _workingSettings.DefaultParallelSmallFileWorkers)
            _parallelScanWorkersNumeric.Value = ClampNumeric(_parallelScanWorkersNumeric, _workingSettings.DefaultParallelScanWorkers)
            _smallFileThresholdKbNumeric.Value = ClampNumeric(_smallFileThresholdKbNumeric, _workingSettings.DefaultSmallFileThresholdKb)
            _defaultFragileFailureWindowNumeric.Value = ClampNumeric(_defaultFragileFailureWindowNumeric, _workingSettings.DefaultFragileFailureWindowSeconds)
            _defaultFragileFailureThresholdNumeric.Value = ClampNumeric(_defaultFragileFailureThresholdNumeric, _workingSettings.DefaultFragileFailureThreshold)
            _defaultFragileCooldownNumeric.Value = ClampNumeric(_defaultFragileCooldownNumeric, _workingSettings.DefaultFragileCooldownSeconds)
            _defaultLockProbeIntervalNumeric.Value = ClampNumeric(_defaultLockProbeIntervalNumeric, _workingSettings.DefaultLockContentionProbeIntervalMs)
            _rescueFastChunkKbNumeric.Value = ClampNumeric(_rescueFastChunkKbNumeric, _workingSettings.DefaultRescueFastScanChunkKb)
            _rescueTrimChunkKbNumeric.Value = ClampNumeric(_rescueTrimChunkKbNumeric, _workingSettings.DefaultRescueTrimChunkKb)
            _rescueScrapeChunkKbNumeric.Value = ClampNumeric(_rescueScrapeChunkKbNumeric, _workingSettings.DefaultRescueScrapeChunkKb)
            _rescueRetryChunkKbNumeric.Value = ClampNumeric(_rescueRetryChunkKbNumeric, _workingSettings.DefaultRescueRetryChunkKb)
            _rescueSplitMinimumKbNumeric.Value = ClampNumeric(_rescueSplitMinimumKbNumeric, _workingSettings.DefaultRescueSplitMinimumKb)
            _rescueFastRetriesNumeric.Value = ClampNumeric(_rescueFastRetriesNumeric, _workingSettings.DefaultRescueFastScanRetries)
            _rescueTrimRetriesNumeric.Value = ClampNumeric(_rescueTrimRetriesNumeric, _workingSettings.DefaultRescueTrimRetries)
            _rescueScrapeRetriesNumeric.Value = ClampNumeric(_rescueScrapeRetriesNumeric, _workingSettings.DefaultRescueScrapeRetries)
            _workerProgressIntervalNumeric.Value = ClampNumeric(_workerProgressIntervalNumeric, _workingSettings.WorkerProgressIntervalMs)
            _workerMaxLogsPerSecondNumeric.Value = ClampNumeric(_workerMaxLogsPerSecondNumeric, _workingSettings.WorkerMaxLogsPerSecond)
            _uiDiagnosticsRefreshNumeric.Value = ClampNumeric(_uiDiagnosticsRefreshNumeric, _workingSettings.UiDiagnosticsRefreshMs)
            _uiMaxLogLinesNumeric.Value = ClampNumeric(_uiMaxLogLinesNumeric, _workingSettings.UiMaxLogLines)
            _uiShowDiagnosticsCheckBox.Checked = _workingSettings.UiShowDiagnostics

            _defaultVerifyCheckBox.Checked = _workingSettings.DefaultVerifyAfterCopy
            _sampleChunkKbNumeric.Value = ClampNumeric(_sampleChunkKbNumeric, _workingSettings.DefaultSampleVerificationChunkKb)
            _sampleChunkCountNumeric.Value = ClampNumeric(_sampleChunkCountNumeric, _workingSettings.DefaultSampleVerificationChunkCount)

            Select Case ThemeSettings.GetPreferredColorMode(_workingSettings)
                Case SystemColorMode.Dark
                    _themeComboBox.SelectedItem = "Dark"
                Case SystemColorMode.Classic
                    _themeComboBox.SelectedItem = "Classic"
                Case Else
                    _themeComboBox.SelectedItem = "System"
            End Select

            Select Case SettingsValueConverter.ToAccentColorModeChoice(_workingSettings.AccentColorMode)
                Case AccentColorModeChoice.System
                    _accentModeComboBox.SelectedIndex = 1
                Case AccentColorModeChoice.Custom
                    _accentModeComboBox.SelectedIndex = 2
                Case Else
                    _accentModeComboBox.SelectedIndex = 0
            End Select

            Select Case SettingsValueConverter.ToUiDensityChoice(_workingSettings.UiDensity)
                Case UiDensityChoice.Compact
                    _uiDensityComboBox.SelectedIndex = 0
                Case UiDensityChoice.Comfortable
                    _uiDensityComboBox.SelectedIndex = 2
                Case Else
                    _uiDensityComboBox.SelectedIndex = 1
            End Select

            Select Case SettingsValueConverter.ToGridHeaderStyleChoice(_workingSettings.GridHeaderStyle)
                Case GridHeaderStyleChoice.Minimal
                    _gridHeaderStyleComboBox.SelectedIndex = 1
                Case GridHeaderStyleChoice.Prominent
                    _gridHeaderStyleComboBox.SelectedIndex = 2
                Case Else
                    _gridHeaderStyleComboBox.SelectedIndex = 0
            End Select

            Select Case SettingsValueConverter.ToProgressBarStyleChoice(_workingSettings.ProgressBarStyle)
                Case ProgressBarStyleChoice.Thin
                    _progressBarStyleComboBox.SelectedIndex = 0
                Case ProgressBarStyleChoice.Thick
                    _progressBarStyleComboBox.SelectedIndex = 2
                Case Else
                    _progressBarStyleComboBox.SelectedIndex = 1
            End Select

            Select Case SettingsValueConverter.ToOverwritePolicy(_workingSettings.DefaultOverwritePolicy)
                Case OverwritePolicy.SkipExisting
                    SafeSelectIndex(_overwritePolicyComboBox, 1)
                Case OverwritePolicy.OverwriteIfSourceNewer
                    SafeSelectIndex(_overwritePolicyComboBox, 2)
                Case OverwritePolicy.Ask
                    SafeSelectIndex(_overwritePolicyComboBox, 3)
                Case Else
                    SafeSelectIndex(_overwritePolicyComboBox, 0)
            End Select

            Select Case SettingsValueConverter.ToSymlinkHandling(_workingSettings.DefaultSymlinkHandling)
                Case SymlinkHandlingMode.Follow
                    SafeSelectIndex(_symlinkHandlingComboBox, 1)
                Case Else
                    SafeSelectIndex(_symlinkHandlingComboBox, 0)
            End Select

            Select Case SettingsValueConverter.ToSalvageFillPattern(_workingSettings.DefaultSalvageFillPattern)
                Case SalvageFillPattern.Ones
                    SafeSelectIndex(_salvageFillPatternComboBox, 1)
                Case SalvageFillPattern.Random
                    SafeSelectIndex(_salvageFillPatternComboBox, 2)
                Case Else
                    SafeSelectIndex(_salvageFillPatternComboBox, 0)
            End Select

            Select Case SettingsValueConverter.ToVerificationMode(_workingSettings.DefaultVerificationMode)
                Case VerificationMode.Sampled
                    SafeSelectIndex(_defaultVerificationModeComboBox, 1)
                Case Else
                    SafeSelectIndex(_defaultVerificationModeComboBox, 0)
            End Select

            Select Case SettingsValueConverter.ToSourceMutationPolicy(_workingSettings.DefaultSourceMutationPolicy)
                Case SourceMutationPolicy.SkipFile
                    SafeSelectIndex(_defaultSourceMutationPolicyComboBox, 1)
                Case SourceMutationPolicy.WaitForReappearance
                    SafeSelectIndex(_defaultSourceMutationPolicyComboBox, 2)
                Case Else
                    SafeSelectIndex(_defaultSourceMutationPolicyComboBox, 0)
            End Select

            Select Case SettingsValueConverter.ToTransferEnginePolicy(_workingSettings.DefaultTransferEnginePolicy)
                Case TransferEnginePolicy.ManagedRescue
                    SafeSelectIndex(_transferEnginePolicyComboBox, 1)
                Case TransferEnginePolicy.NativeFast
                    SafeSelectIndex(_transferEnginePolicyComboBox, 2)
                Case Else
                    SafeSelectIndex(_transferEnginePolicyComboBox, 0)
            End Select

            Select Case SettingsValueConverter.ToScanPerformanceProfile(_workingSettings.DefaultScanPerformanceProfile)
                Case ScanPerformanceProfile.Fast
                    SafeSelectIndex(_scanPerformanceProfileComboBox, 1)
                Case ScanPerformanceProfile.Precise
                    SafeSelectIndex(_scanPerformanceProfileComboBox, 2)
                Case Else
                    SafeSelectIndex(_scanPerformanceProfileComboBox, 0)
            End Select

            Select Case SettingsValueConverter.NormalizeWorkerProcessPriorityClass(_workingSettings.DefaultWorkerProcessPriorityClass)
                Case "Idle"
                    SafeSelectIndex(_workerProcessPriorityComboBox, 0)
                Case "BelowNormal"
                    SafeSelectIndex(_workerProcessPriorityComboBox, 1)
                Case "AboveNormal"
                    SafeSelectIndex(_workerProcessPriorityComboBox, 3)
                Case "High"
                    SafeSelectIndex(_workerProcessPriorityComboBox, 4)
                Case Else
                    SafeSelectIndex(_workerProcessPriorityComboBox, 2)
            End Select

            Select Case SettingsValueConverter.ToVerificationHashAlgorithm(_workingSettings.DefaultVerificationHashAlgorithm)
                Case VerificationHashAlgorithm.Sha512
                    SafeSelectIndex(_defaultHashAlgorithmComboBox, 1)
                Case Else
                    SafeSelectIndex(_defaultHashAlgorithmComboBox, 0)
            End Select

            Select Case SettingsValueConverter.ToExplorerSelectionMode(_workingSettings.ExplorerSelectionMode)
                Case ExplorerSelectionMode.SourceFolder
                    SafeSelectIndex(_explorerSelectionModeComboBox, 1)
                Case Else
                    SafeSelectIndex(_explorerSelectionModeComboBox, 0)
            End Select

            Select Case SettingsValueConverter.ToWorkerTelemetryProfile(_workingSettings.WorkerTelemetryProfile)
                Case WorkerTelemetryProfile.Verbose
                    SafeSelectIndex(_workerTelemetryProfileComboBox, 1)
                Case WorkerTelemetryProfile.Debug
                    SafeSelectIndex(_workerTelemetryProfileComboBox, 2)
                Case Else
                    SafeSelectIndex(_workerTelemetryProfileComboBox, 0)
            End Select

            EnsureSelection(_themeComboBox)
            EnsureSelection(_accentModeComboBox)
            EnsureSelection(_windowChromeModeComboBox)
            EnsureSelection(_uiDensityComboBox)
            EnsureSelection(_uiScaleComboBox)
            EnsureSelection(_gridHeaderStyleComboBox)
            EnsureSelection(_progressBarStyleComboBox)
            EnsureSelection(_overwritePolicyComboBox)
            EnsureSelection(_symlinkHandlingComboBox)
            EnsureSelection(_salvageFillPatternComboBox)
            EnsureSelection(_defaultVerificationModeComboBox)
            EnsureSelection(_defaultHashAlgorithmComboBox)
            EnsureSelection(_defaultSourceMutationPolicyComboBox)
            EnsureSelection(_transferEnginePolicyComboBox)
            EnsureSelection(_workerProcessPriorityComboBox)
            EnsureSelection(_explorerSelectionModeComboBox)
            EnsureSelection(_workerTelemetryProfileComboBox)

            Dim requestedLogFont = If(_workingSettings.LogFontFamily, String.Empty).Trim()
            If requestedLogFont.Length = 0 Then
                requestedLogFont = "Consolas"
            End If
            Dim requestedFontIndex = _logFontFamilyComboBox.FindStringExact(requestedLogFont)
            If requestedFontIndex < 0 Then
                _logFontFamilyComboBox.Items.Add(requestedLogFont)
                requestedFontIndex = _logFontFamilyComboBox.Items.Count - 1
            End If
            _logFontFamilyComboBox.SelectedIndex = requestedFontIndex

            UpdateRecoveryControlStates()
            UpdateVerificationControlStates()
            UpdateExplorerControlStates()
            UpdateBadRangeMapControlStates()
            UpdateFragileControlStates()
            UpdatePerformanceControlStates()
            If _showDiagnosticsStatusRowCheckBox.Checked <> _uiShowDiagnosticsCheckBox.Checked Then
                _showDiagnosticsStatusRowCheckBox.Checked = _uiShowDiagnosticsCheckBox.Checked
            End If
            UpdateAppearanceControlStates()
            UpdateDiagnosticsControlStates()
        Finally
            _suspendDirtyTracking = False
            _pageHostPanel.ResumeLayout(performLayout:=True)
            _mainSplit.ResumeLayout(performLayout:=True)
            ResumeLayout(performLayout:=True)
        End Try

        _dirtyStateTimer.Stop()
        UpdateDirtyState()
    End Sub

    Private Sub UpdateDirtyState()
        If _suspendDirtyTracking Then
            Return
        End If

        Dim capturedSettings As AppSettings = Nothing
        If Not TryCaptureSettingsFromControls(capturedSettings, validateInput:=False, showValidationErrors:=False) Then
            _hasChanges = False
            _okButton.Enabled = False
            Return
        End If

        _hasChanges = Not AreSettingsEqual(_originalSettings, capturedSettings)
        _okButton.Enabled = _hasChanges
    End Sub

    Private Function TryCaptureSettingsFromControls(
        ByRef capturedSettings As AppSettings,
        validateInput As Boolean,
        showValidationErrors As Boolean) As Boolean

        Dim updateUrl = _updateUrlTextBox.Text.Trim()
        Dim accentColor = _accentColorTextBox.Text.Trim()
        If validateInput AndAlso updateUrl.Length > 0 AndAlso Not IsValidHttpUrl(updateUrl) Then
            If showValidationErrors Then
                MessageBox.Show(Me, "Release URL must be an absolute http/https URL.", "Settings", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                _updateUrlTextBox.Focus()
            End If
            capturedSettings = Nothing
            Return False
        End If

        If validateInput AndAlso _accentModeComboBox.SelectedIndex = 2 Then
            Dim parsedAccentColor As Color = Nothing
            If Not SettingsValueConverter.TryParseColorHex(accentColor, parsedAccentColor) Then
                If showValidationErrors Then
                    MessageBox.Show(Me, "Custom accent color must use #RRGGBB format.", "Settings", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    _accentColorTextBox.Focus()
                End If
                capturedSettings = Nothing
                Return False
            End If
        End If

        Dim overwritePolicyIndex = If(_overwritePolicyComboBox.SelectedIndex >= 0, _overwritePolicyComboBox.SelectedIndex, 0)
        Dim symlinkHandlingIndex = If(_symlinkHandlingComboBox.SelectedIndex >= 0, _symlinkHandlingComboBox.SelectedIndex, 0)
        Dim salvagePatternIndex = If(_salvageFillPatternComboBox.SelectedIndex >= 0, _salvageFillPatternComboBox.SelectedIndex, 0)
        Dim verificationModeIndex = If(_defaultVerificationModeComboBox.SelectedIndex >= 0, _defaultVerificationModeComboBox.SelectedIndex, 0)
        Dim hashAlgorithmIndex = If(_defaultHashAlgorithmComboBox.SelectedIndex >= 0, _defaultHashAlgorithmComboBox.SelectedIndex, 0)
        Dim sourceMutationPolicyIndex = If(_defaultSourceMutationPolicyComboBox.SelectedIndex >= 0, _defaultSourceMutationPolicyComboBox.SelectedIndex, 0)
        Dim transferEnginePolicyIndex = If(_transferEnginePolicyComboBox.SelectedIndex >= 0, _transferEnginePolicyComboBox.SelectedIndex, 0)
        Dim scanPerformanceProfileIndex = If(_scanPerformanceProfileComboBox.SelectedIndex >= 0, _scanPerformanceProfileComboBox.SelectedIndex, 0)
        Dim workerProcessPriorityIndex = If(_workerProcessPriorityComboBox.SelectedIndex >= 0, _workerProcessPriorityComboBox.SelectedIndex, 2)
        Dim explorerSelectionModeIndex = If(_explorerSelectionModeComboBox.SelectedIndex >= 0, _explorerSelectionModeComboBox.SelectedIndex, 0)
        Dim telemetryProfileIndex = If(_workerTelemetryProfileComboBox.SelectedIndex >= 0, _workerTelemetryProfileComboBox.SelectedIndex, 0)
        Dim accentModeIndex = If(_accentModeComboBox.SelectedIndex >= 0, _accentModeComboBox.SelectedIndex, 0)
        Dim windowChromeIndex = If(_windowChromeModeComboBox.SelectedIndex >= 0, _windowChromeModeComboBox.SelectedIndex, 0)
        Dim uiDensityIndex = If(_uiDensityComboBox.SelectedIndex >= 0, _uiDensityComboBox.SelectedIndex, 1)
        Dim gridHeaderStyleIndex = If(_gridHeaderStyleComboBox.SelectedIndex >= 0, _gridHeaderStyleComboBox.SelectedIndex, 0)
        Dim progressBarStyleIndex = If(_progressBarStyleComboBox.SelectedIndex >= 0, _progressBarStyleComboBox.SelectedIndex, 1)
        Dim uiScalePercent = ParseScalePercent(_uiScaleComboBox)

        capturedSettings = CloneSettings(_workingSettings)
        capturedSettings.Theme = ThemeSettings.ModeToString(GetSelectedMode())
        capturedSettings.AccentColorMode = If(
            accentModeIndex = 1,
            SettingsValueConverter.AccentColorModeChoiceToString(AccentColorModeChoice.System),
            If(accentModeIndex = 2,
               SettingsValueConverter.AccentColorModeChoiceToString(AccentColorModeChoice.Custom),
               SettingsValueConverter.AccentColorModeChoiceToString(AccentColorModeChoice.Auto)))
        capturedSettings.AccentColorHex = SettingsValueConverter.NormalizeColorHex(accentColor, "#5A78C8")
        capturedSettings.WindowChromeMode = If(
            windowChromeIndex = 1,
            SettingsValueConverter.WindowChromeModeChoiceToString(WindowChromeModeChoice.Standard),
            SettingsValueConverter.WindowChromeModeChoiceToString(WindowChromeModeChoice.Themed))
        capturedSettings.UiDensity = If(
            uiDensityIndex = 0,
            SettingsValueConverter.UiDensityChoiceToString(UiDensityChoice.Compact),
            If(uiDensityIndex = 2,
               SettingsValueConverter.UiDensityChoiceToString(UiDensityChoice.Comfortable),
               SettingsValueConverter.UiDensityChoiceToString(UiDensityChoice.Normal)))
        capturedSettings.UiScalePercent = uiScalePercent
        capturedSettings.LogFontFamily = _logFontFamilyComboBox.Text.Trim()
        If capturedSettings.LogFontFamily.Length = 0 Then
            capturedSettings.LogFontFamily = "Consolas"
        End If
        capturedSettings.LogFontSizePoints = CInt(_logFontSizeNumeric.Value)
        capturedSettings.UiColorizeLogBySeverity = _logColorizeBySeverityCheckBox.Checked
        capturedSettings.GridAlternatingRows = _gridAlternatingRowsCheckBox.Checked
        capturedSettings.GridRowHeight = CInt(_gridRowHeightNumeric.Value)
        capturedSettings.GridHeaderStyle = If(
            gridHeaderStyleIndex = 1,
            SettingsValueConverter.GridHeaderStyleChoiceToString(GridHeaderStyleChoice.Minimal),
            If(gridHeaderStyleIndex = 2,
               SettingsValueConverter.GridHeaderStyleChoiceToString(GridHeaderStyleChoice.Prominent),
               SettingsValueConverter.GridHeaderStyleChoiceToString(GridHeaderStyleChoice.Default)))
        capturedSettings.ProgressBarStyle = If(
            progressBarStyleIndex = 0,
            SettingsValueConverter.ProgressBarStyleChoiceToString(ProgressBarStyleChoice.Thin),
            If(progressBarStyleIndex = 2,
               SettingsValueConverter.ProgressBarStyleChoiceToString(ProgressBarStyleChoice.Thick),
               SettingsValueConverter.ProgressBarStyleChoiceToString(ProgressBarStyleChoice.Standard)))
        capturedSettings.ProgressBarShowPercentage = _progressBarShowPercentageCheckBox.Checked
        capturedSettings.ShowBufferStatusRow = _showBufferStatusRowCheckBox.Checked
        capturedSettings.ShowRescueStatusRow = _showRescueStatusRowCheckBox.Checked
        capturedSettings.CheckUpdatesOnLaunch = _checkUpdatesOnLaunchCheckBox.Checked
        capturedSettings.UpdateReleaseUrl = updateUrl
        capturedSettings.UserAgent = _userAgentTextBox.Text.Trim()

        capturedSettings.EnableRecoveryAutostart = _enableRecoveryAutostartCheckBox.Checked
        capturedSettings.AutoResumeAfterCrash = _autoResumeAfterCrashCheckBox.Checked
        capturedSettings.PromptResumeAfterCrash = If(_autoResumeAfterCrashCheckBox.Checked, True, _promptResumeAfterCrashCheckBox.Checked)
        capturedSettings.KeepResumePromptUntilResolved = If(_autoResumeAfterCrashCheckBox.Checked, False, _keepResumePromptCheckBox.Checked)
        capturedSettings.AutoRunQueuedJobsOnStartup = _autoRunQueuedOnStartupCheckBox.Checked
        capturedSettings.RecoveryTouchIntervalSeconds = CInt(_recoveryTouchIntervalNumeric.Value)

        capturedSettings.EnableExplorerContextMenu = _enableExplorerContextMenuCheckBox.Checked
        capturedSettings.ExplorerSelectionMode = If(explorerSelectionModeIndex = 1,
                                                    SettingsValueConverter.ExplorerSelectionModeToString(ExplorerSelectionMode.SourceFolder),
                                                    SettingsValueConverter.ExplorerSelectionModeToString(ExplorerSelectionMode.SelectedItems))

        capturedSettings.DefaultResumeFromJournal = _defaultResumeCheckBox.Checked
        capturedSettings.DefaultSalvageUnreadableBlocks = _defaultSalvageCheckBox.Checked
        capturedSettings.DefaultContinueOnFileError = _defaultContinueOnErrorCheckBox.Checked
        capturedSettings.DefaultPreserveTimestamps = _defaultPreserveTimestampsCheckBox.Checked
        capturedSettings.DefaultCopyEmptyDirectories = _defaultCopyEmptyDirectoriesCheckBox.Checked
        capturedSettings.DefaultWaitForMediaAvailability = _defaultWaitForMediaCheckBox.Checked
        capturedSettings.DefaultWaitForFileLockRelease = _defaultWaitForLockReleaseCheckBox.Checked
        capturedSettings.DefaultTreatAccessDeniedAsContention = _defaultTreatAccessDeniedContentionCheckBox.Checked
        capturedSettings.DefaultUseBadRangeMap = _defaultUseBadRangeMapCheckBox.Checked
        capturedSettings.DefaultSkipKnownBadRanges = _defaultSkipKnownBadRangesCheckBox.Checked
        capturedSettings.DefaultUpdateBadRangeMapFromRun = _defaultUpdateBadRangeMapCheckBox.Checked
        capturedSettings.DefaultUseExperimentalRawDiskScan = _defaultUseRawDiskScanCheckBox.Checked
        capturedSettings.DefaultBadRangeMapMaxAgeDays = CInt(_defaultBadRangeMapMaxAgeDaysNumeric.Value)
        capturedSettings.DefaultUseAdaptiveBuffer = _defaultAdaptiveBufferCheckBox.Checked
        capturedSettings.DefaultFragileMediaMode = _defaultFragileModeCheckBox.Checked
        capturedSettings.DefaultSkipFileOnFirstReadError = _defaultSkipFileOnFirstReadErrorCheckBox.Checked
        capturedSettings.DefaultPersistFragileSkipsAcrossResume = _defaultPersistFragileSkipsCheckBox.Checked
        capturedSettings.DefaultBufferSizeMb = CInt(_defaultBufferMbNumeric.Value)
        capturedSettings.DefaultMaxRetries = CInt(_defaultRetriesNumeric.Value)
        capturedSettings.DefaultOperationTimeoutSeconds = CInt(_defaultOperationTimeoutNumeric.Value)
        capturedSettings.DefaultPerFileTimeoutSeconds = CInt(_defaultPerFileTimeoutNumeric.Value)
        capturedSettings.DefaultMaxThroughputMbPerSecond = CInt(_defaultMaxThroughputNumeric.Value)
        capturedSettings.DefaultParallelSmallFileWorkers = CInt(_parallelSmallFileWorkersNumeric.Value)
        capturedSettings.DefaultParallelScanWorkers = CInt(_parallelScanWorkersNumeric.Value)
        capturedSettings.DefaultSmallFileThresholdKb = CInt(_smallFileThresholdKbNumeric.Value)
        capturedSettings.DefaultFragileFailureWindowSeconds = CInt(_defaultFragileFailureWindowNumeric.Value)
        capturedSettings.DefaultFragileFailureThreshold = CInt(_defaultFragileFailureThresholdNumeric.Value)
        capturedSettings.DefaultFragileCooldownSeconds = CInt(_defaultFragileCooldownNumeric.Value)
        capturedSettings.DefaultLockContentionProbeIntervalMs = CInt(_defaultLockProbeIntervalNumeric.Value)
        capturedSettings.DefaultRescueFastScanChunkKb = CInt(_rescueFastChunkKbNumeric.Value)
        capturedSettings.DefaultRescueTrimChunkKb = CInt(_rescueTrimChunkKbNumeric.Value)
        capturedSettings.DefaultRescueScrapeChunkKb = CInt(_rescueScrapeChunkKbNumeric.Value)
        capturedSettings.DefaultRescueRetryChunkKb = CInt(_rescueRetryChunkKbNumeric.Value)
        capturedSettings.DefaultRescueSplitMinimumKb = CInt(_rescueSplitMinimumKbNumeric.Value)
        capturedSettings.DefaultRescueFastScanRetries = CInt(_rescueFastRetriesNumeric.Value)
        capturedSettings.DefaultRescueTrimRetries = CInt(_rescueTrimRetriesNumeric.Value)
        capturedSettings.DefaultRescueScrapeRetries = CInt(_rescueScrapeRetriesNumeric.Value)
        capturedSettings.WorkerProgressIntervalMs = CInt(_workerProgressIntervalNumeric.Value)
        capturedSettings.WorkerMaxLogsPerSecond = CInt(_workerMaxLogsPerSecondNumeric.Value)
        capturedSettings.UiShowDiagnostics = _uiShowDiagnosticsCheckBox.Checked
        capturedSettings.UiDiagnosticsRefreshMs = CInt(_uiDiagnosticsRefreshNumeric.Value)
        capturedSettings.UiMaxLogLines = CInt(_uiMaxLogLinesNumeric.Value)

        Select Case overwritePolicyIndex
            Case 1
                capturedSettings.DefaultOverwritePolicy = SettingsValueConverter.OverwritePolicyToString(OverwritePolicy.SkipExisting)
            Case 2
                capturedSettings.DefaultOverwritePolicy = SettingsValueConverter.OverwritePolicyToString(OverwritePolicy.OverwriteIfSourceNewer)
            Case 3
                capturedSettings.DefaultOverwritePolicy = SettingsValueConverter.OverwritePolicyToString(OverwritePolicy.Ask)
            Case Else
                capturedSettings.DefaultOverwritePolicy = SettingsValueConverter.OverwritePolicyToString(OverwritePolicy.Overwrite)
        End Select

        Select Case symlinkHandlingIndex
            Case 1
                capturedSettings.DefaultSymlinkHandling = SettingsValueConverter.SymlinkHandlingToString(SymlinkHandlingMode.Follow)
            Case Else
                capturedSettings.DefaultSymlinkHandling = SettingsValueConverter.SymlinkHandlingToString(SymlinkHandlingMode.Skip)
        End Select

        Select Case salvagePatternIndex
            Case 1
                capturedSettings.DefaultSalvageFillPattern = SettingsValueConverter.SalvageFillPatternToString(SalvageFillPattern.Ones)
            Case 2
                capturedSettings.DefaultSalvageFillPattern = SettingsValueConverter.SalvageFillPatternToString(SalvageFillPattern.Random)
            Case Else
                capturedSettings.DefaultSalvageFillPattern = SettingsValueConverter.SalvageFillPatternToString(SalvageFillPattern.Zero)
        End Select

        capturedSettings.DefaultVerifyAfterCopy = _defaultVerifyCheckBox.Checked
        capturedSettings.DefaultVerificationMode = If(verificationModeIndex = 1,
                                                      SettingsValueConverter.VerificationModeToString(VerificationMode.Sampled),
                                                      SettingsValueConverter.VerificationModeToString(VerificationMode.Full))
        capturedSettings.DefaultVerificationHashAlgorithm = If(hashAlgorithmIndex = 1,
                                                               SettingsValueConverter.VerificationHashAlgorithmToString(VerificationHashAlgorithm.Sha512),
                                                               SettingsValueConverter.VerificationHashAlgorithmToString(VerificationHashAlgorithm.Sha256))
        capturedSettings.DefaultSampleVerificationChunkKb = CInt(_sampleChunkKbNumeric.Value)
        capturedSettings.DefaultSampleVerificationChunkCount = CInt(_sampleChunkCountNumeric.Value)
        capturedSettings.DefaultSourceMutationPolicy = If(sourceMutationPolicyIndex = 1,
                                                           SettingsValueConverter.SourceMutationPolicyToString(SourceMutationPolicy.SkipFile),
                                                           If(sourceMutationPolicyIndex = 2,
                                                              SettingsValueConverter.SourceMutationPolicyToString(SourceMutationPolicy.WaitForReappearance),
                                                              SettingsValueConverter.SourceMutationPolicyToString(SourceMutationPolicy.FailFile)))

        capturedSettings.DefaultTransferEnginePolicy = If(transferEnginePolicyIndex = 1,
                                                           SettingsValueConverter.TransferEnginePolicyToString(TransferEnginePolicy.ManagedRescue),
                                                           If(transferEnginePolicyIndex = 2,
                                                              SettingsValueConverter.TransferEnginePolicyToString(TransferEnginePolicy.NativeFast),
                                                              SettingsValueConverter.TransferEnginePolicyToString(TransferEnginePolicy.Auto)))

        capturedSettings.DefaultScanPerformanceProfile = If(scanPerformanceProfileIndex = 1,
                                                             SettingsValueConverter.ScanPerformanceProfileToString(ScanPerformanceProfile.Fast),
                                                             If(scanPerformanceProfileIndex = 2,
                                                                SettingsValueConverter.ScanPerformanceProfileToString(ScanPerformanceProfile.Precise),
                                                                SettingsValueConverter.ScanPerformanceProfileToString(ScanPerformanceProfile.Auto)))

        Select Case workerProcessPriorityIndex
            Case 0
                capturedSettings.DefaultWorkerProcessPriorityClass = "Idle"
            Case 1
                capturedSettings.DefaultWorkerProcessPriorityClass = "BelowNormal"
            Case 3
                capturedSettings.DefaultWorkerProcessPriorityClass = "AboveNormal"
            Case 4
                capturedSettings.DefaultWorkerProcessPriorityClass = "High"
            Case Else
                capturedSettings.DefaultWorkerProcessPriorityClass = "Normal"
        End Select

        Select Case telemetryProfileIndex
            Case 1
                capturedSettings.WorkerTelemetryProfile = SettingsValueConverter.WorkerTelemetryProfileToString(WorkerTelemetryProfile.Verbose)
            Case 2
                capturedSettings.WorkerTelemetryProfile = SettingsValueConverter.WorkerTelemetryProfileToString(WorkerTelemetryProfile.Debug)
            Case Else
                capturedSettings.WorkerTelemetryProfile = SettingsValueConverter.WorkerTelemetryProfileToString(WorkerTelemetryProfile.Normal)
        End Select

        Return True
    End Function

    Private Shared Function BuildRestartRequirementSummary(original As AppSettings, updated As AppSettings) As String
        If original Is Nothing OrElse updated Is Nothing Then
            Return String.Empty
        End If

        Dim reasons As New List(Of String)()

        If Not String.Equals(original.Theme, updated.Theme, StringComparison.OrdinalIgnoreCase) Then
            reasons.Add("Application theme mode")
        End If

        If Not String.Equals(original.WindowChromeMode, updated.WindowChromeMode, StringComparison.OrdinalIgnoreCase) Then
            reasons.Add("Window chrome mode")
        End If

        If original.UiScalePercent <> updated.UiScalePercent Then
            reasons.Add("UI scale")
        End If

        If Not String.Equals(original.UiDensity, updated.UiDensity, StringComparison.OrdinalIgnoreCase) Then
            reasons.Add("UI density")
        End If

        If original.CheckUpdatesOnLaunch <> updated.CheckUpdatesOnLaunch Then
            reasons.Add("Startup update-check behavior")
        End If

        If original.EnableRecoveryAutostart <> updated.EnableRecoveryAutostart Then
            reasons.Add("Recovery auto-start behavior")
        End If

        If original.AutoRunQueuedJobsOnStartup <> updated.AutoRunQueuedJobsOnStartup Then
            reasons.Add("Startup queued-jobs behavior")
        End If

        If original.PromptResumeAfterCrash <> updated.PromptResumeAfterCrash OrElse
            original.AutoResumeAfterCrash <> updated.AutoResumeAfterCrash OrElse
            original.KeepResumePromptUntilResolved <> updated.KeepResumePromptUntilResolved Then
            reasons.Add("Interrupted-run startup handling")
        End If

        If reasons.Count = 0 Then
            Return String.Empty
        End If

        Dim lines As New List(Of String)()
        For Each reason In reasons
            lines.Add($"- {reason}")
        Next

        Return String.Join(Environment.NewLine, lines)
    End Function

    Private Shared Function AreSettingsEqual(left As AppSettings, right As AppSettings) As Boolean
        If left Is right Then
            Return True
        End If

        If left Is Nothing OrElse right Is Nothing Then
            Return False
        End If

        For Each propertyInfo In _settingsProperties
            Dim leftValue = propertyInfo.GetValue(left)
            Dim rightValue = propertyInfo.GetValue(right)
            If Not Object.Equals(leftValue, rightValue) Then
                Return False
            End If
        Next

        Return True
    End Function

    Private Shared Sub CopySettings(source As AppSettings, target As AppSettings)
        If source Is Nothing OrElse target Is Nothing Then
            Return
        End If

        For Each propertyInfo In _settingsProperties
            propertyInfo.SetValue(target, propertyInfo.GetValue(source))
        Next
    End Sub
    Private Shared Function ClampNumeric(control As ThemedNumericUpDown, value As Integer) As Decimal
        Dim candidate As Decimal = value
        If candidate < control.Minimum Then
            Return control.Minimum
        End If
        If candidate > control.Maximum Then
            Return control.Maximum
        End If
        Return candidate
    End Function

    Private Shared Sub EnsureSelection(comboBox As ComboBox)
        If comboBox.SelectedIndex < 0 AndAlso comboBox.Items.Count > 0 Then
            comboBox.SelectedIndex = 0
        End If
    End Sub

    ''' <summary>
    ''' Sets SelectedIndex only when the ComboBox has items, preventing
    ''' ArgumentOutOfRangeException on controls whose page has not been built yet.
    ''' </summary>
    Private Shared Sub SafeSelectIndex(comboBox As ComboBox, index As Integer)
        If comboBox.Items.Count > index Then
            comboBox.SelectedIndex = index
        End If
    End Sub

    Private Sub RecoveryControl_CheckedChanged(sender As Object, e As EventArgs)
        UpdateRecoveryControlStates()
    End Sub

    Private Sub BadRangeMapControl_CheckedChanged(sender As Object, e As EventArgs)
        UpdateBadRangeMapControlStates()
    End Sub

    Private Sub FragileControl_CheckedChanged(sender As Object, e As EventArgs)
        UpdateFragileControlStates()
    End Sub

    Private Sub DefaultAdaptiveBufferCheckBox_CheckedChanged(sender As Object, e As EventArgs)
        UpdatePerformanceControlStates()
    End Sub

    Private Sub DefaultVerifyCheckBox_CheckedChanged(sender As Object, e As EventArgs)
        UpdateVerificationControlStates()
    End Sub

    Private Sub ExplorerControl_CheckedChanged(sender As Object, e As EventArgs)
        UpdateExplorerControlStates()
    End Sub

    Private Sub DiagnosticsControl_CheckedChanged(sender As Object, e As EventArgs)
        If _showDiagnosticsStatusRowCheckBox.Checked <> _uiShowDiagnosticsCheckBox.Checked Then
            _showDiagnosticsStatusRowCheckBox.Checked = _uiShowDiagnosticsCheckBox.Checked
        End If
        UpdateDiagnosticsControlStates()
    End Sub

    Private Sub AppearanceControl_Changed(sender As Object, e As EventArgs)
        UpdateAppearanceControlStates()
    End Sub

    Private Sub AppearanceDiagnosticsVisibilityCheckBox_CheckedChanged(sender As Object, e As EventArgs)
        If _uiShowDiagnosticsCheckBox.Checked <> _showDiagnosticsStatusRowCheckBox.Checked Then
            _uiShowDiagnosticsCheckBox.Checked = _showDiagnosticsStatusRowCheckBox.Checked
        End If
        UpdateDiagnosticsControlStates()
    End Sub

    Private Sub AccentColorPickButton_Click(sender As Object, e As EventArgs)
        Using colorDialog As New ColorDialog()
            colorDialog.FullOpen = True
            colorDialog.AnyColor = True

            Dim parsedColor As Color = Nothing
            If SettingsValueConverter.TryParseColorHex(_accentColorTextBox.Text, parsedColor) Then
                colorDialog.Color = parsedColor
            End If

            If colorDialog.ShowDialog(Me) <> DialogResult.OK Then
                Return
            End If

            _accentColorTextBox.Text = SettingsValueConverter.ColorToHex(colorDialog.Color)
            _accentModeComboBox.SelectedIndex = 2
            UpdateAppearanceControlStates()
            UpdateDirtyState()
        End Using
    End Sub

    Private Sub UpdateRecoveryControlStates()
        Dim autoResumeEnabled = _autoResumeAfterCrashCheckBox.Checked
        _promptResumeAfterCrashCheckBox.Enabled = Not autoResumeEnabled
        If autoResumeEnabled Then
            _keepResumePromptCheckBox.Enabled = False
            Return
        End If
        _keepResumePromptCheckBox.Enabled = _promptResumeAfterCrashCheckBox.Checked
    End Sub

    Private Sub UpdateBadRangeMapControlStates()
        Dim enabled = _defaultUseBadRangeMapCheckBox.Checked
        _defaultSkipKnownBadRangesCheckBox.Enabled = enabled
        _defaultBadRangeMapMaxAgeDaysNumeric.Enabled = enabled
    End Sub

    Private Sub UpdateFragileControlStates()
        Dim enabled = _defaultFragileModeCheckBox.Checked
        _defaultSkipFileOnFirstReadErrorCheckBox.Enabled = enabled
        _defaultPersistFragileSkipsCheckBox.Enabled = enabled
        _defaultFragileFailureWindowNumeric.Enabled = enabled
        _defaultFragileFailureThresholdNumeric.Enabled = enabled
        _defaultFragileCooldownNumeric.Enabled = enabled
    End Sub

    Private Sub UpdatePerformanceControlStates()
        _defaultBufferMbNumeric.Enabled = Not _defaultAdaptiveBufferCheckBox.Checked
    End Sub

    Private Sub UpdateVerificationControlStates()
        ' Keep mode/hash editable even when "verify by default" is unchecked.
        ' This lets users preconfigure sampled settings before enabling verification.
        _defaultVerificationModeComboBox.Enabled = True
        _defaultHashAlgorithmComboBox.Enabled = True

        Dim sampledEnabled = _defaultVerificationModeComboBox.SelectedIndex = 1
        _sampleChunkKbNumeric.Enabled = sampledEnabled
        _sampleChunkCountNumeric.Enabled = sampledEnabled
    End Sub

    Private Sub UpdateExplorerControlStates()
        _explorerSelectionModeComboBox.Enabled = _enableExplorerContextMenuCheckBox.Checked
    End Sub

    Private Sub UpdateDiagnosticsControlStates()
        _uiDiagnosticsRefreshNumeric.Enabled = _uiShowDiagnosticsCheckBox.Checked
    End Sub

    Private Sub UpdateAppearanceControlStates()
        Dim customAccent = _accentModeComboBox.SelectedIndex = 2
        _accentColorTextBox.Enabled = customAccent
        _accentColorPickButton.Enabled = customAccent
    End Sub

    Private Shared Function ParseScalePercent(scaleComboBox As ComboBox) As Integer
        If scaleComboBox Is Nothing Then
            Return 100
        End If

        Dim valueText = If(scaleComboBox.Text, String.Empty).Trim().TrimEnd("%"c)
        Dim parsedValue As Integer
        If Not Integer.TryParse(valueText, parsedValue) Then
            Return 100
        End If

        Return Math.Max(90, Math.Min(125, parsedValue))
    End Function

    Private Sub OkButton_Click(sender As Object, e As EventArgs)
        Dim capturedSettings As AppSettings = Nothing
        If Not TryCaptureSettingsFromControls(capturedSettings, validateInput:=True, showValidationErrors:=True) Then
            Return
        End If

        _hasChanges = Not AreSettingsEqual(_originalSettings, capturedSettings)
        If Not _hasChanges Then
            Return
        End If

        _restartReasonText = BuildRestartRequirementSummary(_originalSettings, capturedSettings)
        _requiresRestart = Not String.IsNullOrWhiteSpace(_restartReasonText)

        CopySettings(capturedSettings, _workingSettings)

        DialogResult = DialogResult.OK
        Close()
    End Sub

    Private Function GetSelectedMode() As SystemColorMode
        Dim selected = If(TryCast(_themeComboBox.SelectedItem, String), String.Empty)
        If String.Equals(selected, "Classic", StringComparison.OrdinalIgnoreCase) Then
            Return SystemColorMode.Classic
        End If
        If String.Equals(selected, "System", StringComparison.OrdinalIgnoreCase) Then
            Return SystemColorMode.System
        End If
        Return SystemColorMode.Dark
    End Function

    Private Shared Function IsValidHttpUrl(text As String) As Boolean
        Dim uri As Uri = Nothing
        If Not Uri.TryCreate(text, UriKind.Absolute, uri) Then
            Return False
        End If
        Return uri.Scheme = Uri.UriSchemeHttp OrElse uri.Scheme = Uri.UriSchemeHttps
    End Function

    Private Shared Function CloneSettings(source As AppSettings) As AppSettings
        Dim clone As New AppSettings()
        If source Is Nothing Then
            Return clone
        End If

        For Each propertyInfo In _settingsProperties
            propertyInfo.SetValue(clone, propertyInfo.GetValue(source))
        Next

        Return clone
    End Function

    Private Sub SuspendNativeRedraw()
        If _nativeRedrawSuspended OrElse Not IsHandleCreated Then
            Return
        End If

        SendMessage(Handle, WmSetRedraw, IntPtr.Zero, IntPtr.Zero)
        _nativeRedrawSuspended = True
    End Sub

    Private Sub ResumeNativeRedraw()
        If Not _nativeRedrawSuspended OrElse Not IsHandleCreated Then
            Return
        End If

        SendMessage(Handle, WmSetRedraw, New IntPtr(1), IntPtr.Zero)
        _nativeRedrawSuspended = False
    End Sub

End Class

