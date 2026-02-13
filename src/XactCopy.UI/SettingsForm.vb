Imports System.Drawing
Imports XactCopy.Configuration
Imports XactCopy.Models

Friend Class SettingsForm
    Inherits Form

    Private Const LayoutWindowKey As String = "SettingsForm"
    Private Const LayoutSplitterKey As String = "SettingsForm.MainSplit"

    Private ReadOnly _categoryTreeView As New TreeView()
    Private ReadOnly _pageHostPanel As New Panel()
    Private ReadOnly _pageLookup As New Dictionary(Of String, Control)(StringComparer.OrdinalIgnoreCase)
    Private ReadOnly _pageTitleLabel As New Label()
    Private ReadOnly _mainSplit As New SplitContainer()

    Private ReadOnly _themeComboBox As New ComboBox()

    Private ReadOnly _defaultResumeCheckBox As New CheckBox()
    Private ReadOnly _defaultSalvageCheckBox As New CheckBox()
    Private ReadOnly _defaultContinueOnErrorCheckBox As New CheckBox()
    Private ReadOnly _defaultPreserveTimestampsCheckBox As New CheckBox()
    Private ReadOnly _defaultCopyEmptyDirectoriesCheckBox As New CheckBox()
    Private ReadOnly _defaultWaitForMediaCheckBox As New CheckBox()
    Private ReadOnly _overwritePolicyComboBox As New ComboBox()
    Private ReadOnly _symlinkHandlingComboBox As New ComboBox()
    Private ReadOnly _salvageFillPatternComboBox As New ComboBox()

    Private ReadOnly _defaultAdaptiveBufferCheckBox As New CheckBox()
    Private ReadOnly _defaultBufferMbNumeric As New ThemedNumericUpDown()
    Private ReadOnly _defaultRetriesNumeric As New ThemedNumericUpDown()
    Private ReadOnly _defaultOperationTimeoutNumeric As New ThemedNumericUpDown()
    Private ReadOnly _defaultPerFileTimeoutNumeric As New ThemedNumericUpDown()
    Private ReadOnly _defaultMaxThroughputNumeric As New ThemedNumericUpDown()

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

    Public Sub New(settings As AppSettings)
        If settings Is Nothing Then
            Throw New ArgumentNullException(NameOf(settings))
        End If

        _workingSettings = CloneSettings(settings)

        Text = "Settings"
        StartPosition = FormStartPosition.CenterParent
        MinimumSize = New Size(960, 620)
        Size = New Size(1080, 720)
        WindowIconHelper.Apply(Me)

        BuildUi()
        ConfigureToolTips()
        ApplySettingsToControls()

        AddHandler Shown, AddressOf SettingsForm_Shown
        AddHandler FormClosing, AddressOf SettingsForm_FormClosing
    End Sub

    Public ReadOnly Property ResultSettings As AppSettings
        Get
            Return CloneSettings(_workingSettings)
        End Get
    End Property

    Private Sub BuildUi()
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
    End Sub

    Private Sub ConfigureToolTips()
        _toolTip.SetToolTip(_categoryTreeView, "Select a category to edit related settings.")
        _toolTip.SetToolTip(_themeComboBox, "Choose the color mode used by XactCopy windows.")

        _toolTip.SetToolTip(_defaultResumeCheckBox, "Enable journal-based resume for new runs by default.")
        _toolTip.SetToolTip(_defaultSalvageCheckBox, "Enable salvage mode for unreadable source regions by default.")
        _toolTip.SetToolTip(_defaultContinueOnErrorCheckBox, "Continue processing remaining files after a file error.")
        _toolTip.SetToolTip(_defaultPreserveTimestampsCheckBox, "Preserve source timestamps on copied files and folders.")
        _toolTip.SetToolTip(_defaultCopyEmptyDirectoriesCheckBox, "Create empty directories in destination even when they contain no files.")
        _toolTip.SetToolTip(_defaultWaitForMediaCheckBox, "Wait indefinitely for source/destination media to become available.")

        _toolTip.SetToolTip(_overwritePolicyComboBox, "Default conflict behavior when destination files already exist.")
        _toolTip.SetToolTip(_symlinkHandlingComboBox, "Choose whether symbolic links are skipped or followed during scans.")
        _toolTip.SetToolTip(_salvageFillPatternComboBox, "Fill pattern used for unreadable blocks when salvage mode is enabled.")

        _toolTip.SetToolTip(_defaultAdaptiveBufferCheckBox, "Enable adaptive transfer buffer sizing by default.")
        _toolTip.SetToolTip(_defaultBufferMbNumeric, "Maximum transfer buffer for new runs (1-256 MB).")
        _toolTip.SetToolTip(_defaultRetriesNumeric, "Default per-operation retry attempts for I/O failures (0-1000).")
        _toolTip.SetToolTip(_defaultOperationTimeoutNumeric, "Default operation timeout in seconds (1-3600).")
        _toolTip.SetToolTip(_defaultPerFileTimeoutNumeric, "Default per-file timeout in seconds (0 disables per-file timeout).")
        _toolTip.SetToolTip(_defaultMaxThroughputNumeric, "Optional throughput cap in MB/s (0 means unlimited).")

        _toolTip.SetToolTip(_defaultVerifyCheckBox, "Enable post-copy verification by default.")
        _toolTip.SetToolTip(_defaultVerificationModeComboBox, "Verification strategy for new runs.")
        _toolTip.SetToolTip(_defaultHashAlgorithmComboBox, "Hash algorithm used by verification.")
        _toolTip.SetToolTip(_sampleChunkKbNumeric, "Chunk size used in sampled verification mode.")
        _toolTip.SetToolTip(_sampleChunkCountNumeric, "Number of chunks to sample per file in sampled mode.")

        _toolTip.SetToolTip(_checkUpdatesOnLaunchCheckBox, "Check for application updates automatically on startup.")
        _toolTip.SetToolTip(_updateUrlTextBox, "Release metadata endpoint (HTTP/HTTPS).")
        _toolTip.SetToolTip(_userAgentTextBox, "HTTP User-Agent header used for update checks.")

        _toolTip.SetToolTip(_enableRecoveryAutostartCheckBox, "Register startup recovery if a run ends unexpectedly.")
        _toolTip.SetToolTip(_autoRunQueuedOnStartupCheckBox, "Automatically start queued jobs after launching XactCopy.")
        _toolTip.SetToolTip(_promptResumeAfterCrashCheckBox, "Show a prompt to resume interrupted runs.")
        _toolTip.SetToolTip(_autoResumeAfterCrashCheckBox, "Automatically resume interrupted runs without prompting.")
        _toolTip.SetToolTip(_keepResumePromptCheckBox, "Keep prompting to resume until the interrupted run is handled.")
        _toolTip.SetToolTip(_recoveryTouchIntervalNumeric, "Seconds between recovery heartbeat writes (1-60).")

        _toolTip.SetToolTip(_enableExplorerContextMenuCheckBox, "Add 'Copy with XactCopy' to Windows Explorer context menus.")
        _toolTip.SetToolTip(_explorerSelectionModeComboBox, "Choose whether Explorer launches copy selected items or whole source folder.")

        _toolTip.SetToolTip(_okButton, "Save all changes and close settings.")
        _toolTip.SetToolTip(_cancelButton, "Discard changes and close settings.")
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

        _categoryTreeView.Nodes.Add(New TreeNode("Appearance") With {.Name = "appearance"})
        _categoryTreeView.Nodes.Add(New TreeNode("Copy Defaults") With {.Name = "copy"})
        _categoryTreeView.Nodes.Add(New TreeNode("Performance") With {.Name = "performance"})
        _categoryTreeView.Nodes.Add(New TreeNode("Verification") With {.Name = "verification"})
        _categoryTreeView.Nodes.Add(New TreeNode("Updates") With {.Name = "updates"})
        _categoryTreeView.Nodes.Add(New TreeNode("Recovery & Startup") With {.Name = "recovery"})
        _categoryTreeView.Nodes.Add(New TreeNode("Explorer Integration") With {.Name = "explorer"})
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

        _pageLookup("appearance") = BuildAppearancePage()
        _pageLookup("copy") = BuildCopyDefaultsPage()
        _pageLookup("performance") = BuildPerformancePage()
        _pageLookup("verification") = BuildVerificationPage()
        _pageLookup("updates") = BuildUpdatesPage()
        _pageLookup("recovery") = BuildRecoveryPage()
        _pageLookup("explorer") = BuildExplorerPage()

        For Each pageControl In _pageLookup.Values
            pageControl.Dock = DockStyle.Fill
            pageControl.Visible = False
            _pageHostPanel.Controls.Add(pageControl)
        Next

        layout.Controls.Add(_pageTitleLabel, 0, 0)
        layout.Controls.Add(_pageHostPanel, 0, 1)
        Return layout
    End Function

    Private Function BuildAppearancePage() As Control
        Dim page = CreatePageContainer(1)

        Dim body = CreateFieldGrid(1)

        _themeComboBox.DropDownStyle = ComboBoxStyle.DropDownList
        _themeComboBox.Items.AddRange(New Object() {"Dark", "System", "Classic"})
        ConfigureComboBoxControl(_themeComboBox, width:=220)

        body.Controls.Add(CreateFieldLabel("Mode"), 0, 0)
        body.Controls.Add(_themeComboBox, 1, 0)

        page.Controls.Add(CreateSection("Theme", body), 0, 0)
        Return page
    End Function
    Private Function BuildCopyDefaultsPage() As Control
        Dim page = CreatePageContainer(2)

        Dim behavior = CreateSingleColumnGrid(6)

        _defaultResumeCheckBox.Text = "Resume from journal"
        _defaultSalvageCheckBox.Text = "Salvage unreadable blocks"
        _defaultContinueOnErrorCheckBox.Text = "Continue on file errors"
        _defaultPreserveTimestampsCheckBox.Text = "Preserve source timestamps"
        _defaultCopyEmptyDirectoriesCheckBox.Text = "Copy empty directories"
        _defaultWaitForMediaCheckBox.Text = "Wait forever for source/destination"

        For Each box As CheckBox In New CheckBox() {_defaultResumeCheckBox, _defaultSalvageCheckBox, _defaultContinueOnErrorCheckBox, _defaultPreserveTimestampsCheckBox, _defaultCopyEmptyDirectoriesCheckBox, _defaultWaitForMediaCheckBox}
            ConfigureCheckBox(box)
            behavior.Controls.Add(box)
        Next

        Dim policy = CreateFieldGrid(3)

        _overwritePolicyComboBox.DropDownStyle = ComboBoxStyle.DropDownList
        _overwritePolicyComboBox.Items.AddRange(New Object() {"Overwrite existing files", "Skip existing files", "Overwrite only if source is newer"})
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

        page.Controls.Add(CreateSection("Default Run Behavior", behavior), 0, 0)
        page.Controls.Add(CreateSection("Policies", policy), 0, 1)
        Return page
    End Function

    Private Function BuildPerformancePage() As Control
        Dim page = CreatePageContainer(1)

        Dim body = CreateFieldGrid(6)

        _defaultAdaptiveBufferCheckBox.Text = "Enable adaptive buffer by default"
        ConfigureCheckBox(_defaultAdaptiveBufferCheckBox)

        ConfigureNumeric(_defaultBufferMbNumeric, 1D, 256D, 4D)
        ConfigureNumeric(_defaultRetriesNumeric, 0D, 1000D, 12D)
        ConfigureNumeric(_defaultOperationTimeoutNumeric, 1D, 3600D, 10D)
        ConfigureNumeric(_defaultPerFileTimeoutNumeric, 0D, 86400D, 0D)
        ConfigureNumeric(_defaultMaxThroughputNumeric, 0D, 4096D, 0D)

        body.Controls.Add(_defaultAdaptiveBufferCheckBox, 0, 0)
        body.SetColumnSpan(_defaultAdaptiveBufferCheckBox, 3)
        body.Controls.Add(CreateFieldLabel("Buffer MB (max)"), 0, 1)
        body.Controls.Add(_defaultBufferMbNumeric, 1, 1)
        body.Controls.Add(CreateFieldLabel("Max retries"), 0, 2)
        body.Controls.Add(_defaultRetriesNumeric, 1, 2)
        body.Controls.Add(CreateFieldLabel("Operation timeout (sec)"), 0, 3)
        body.Controls.Add(_defaultOperationTimeoutNumeric, 1, 3)
        body.Controls.Add(CreateFieldLabel("Per-file timeout (sec, 0=disabled)"), 0, 4)
        body.Controls.Add(_defaultPerFileTimeoutNumeric, 1, 4)
        body.Controls.Add(CreateFieldLabel("Max throughput (MB/s, 0=unlimited)"), 0, 5)
        body.Controls.Add(_defaultMaxThroughputNumeric, 1, 5)

        page.Controls.Add(CreateSection("Transfer Tuning", body), 0, 0)
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
        _updateUrlTextBox.PlaceholderText = "https://api.github.com/repos/owner/repo/releases/latest"
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
            .Dock = DockStyle.Fill,
            .ColumnCount = 1,
            .RowCount = sectionCount + 1,
            .Padding = New Padding(2)
        }
        page.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0F))
        For index = 0 To sectionCount - 1
            page.RowStyles.Add(New RowStyle(SizeType.AutoSize))
        Next
        page.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0F))
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

    Private Sub SettingsForm_Shown(sender As Object, e As EventArgs)
        UiLayoutManager.ApplyWindow(Me, LayoutWindowKey)
        UiLayoutManager.ApplySplitter(_mainSplit, LayoutSplitterKey, defaultDistance:=210, panel1Min:=170, panel2Min:=500)
        If _categoryTreeView.Nodes.Count > 0 Then
            _categoryTreeView.SelectedNode = _categoryTreeView.Nodes(0)
        End If
    End Sub

    Private Sub SettingsForm_FormClosing(sender As Object, e As FormClosingEventArgs)
        UiLayoutManager.CaptureSplitter(_mainSplit, LayoutSplitterKey)
        UiLayoutManager.CaptureWindow(Me, LayoutWindowKey)
    End Sub

    Private Sub CategoryTreeView_AfterSelect(sender As Object, e As TreeViewEventArgs)
        ShowPage(If(e.Node?.Name, String.Empty))
    End Sub

    Private Sub ShowPage(pageKey As String)
        For Each entry In _pageLookup
            entry.Value.Visible = String.Equals(entry.Key, pageKey, StringComparison.OrdinalIgnoreCase)
        Next

        Select Case pageKey
            Case "copy"
                _pageTitleLabel.Text = "Copy Defaults"
            Case "performance"
                _pageTitleLabel.Text = "Performance"
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

        _defaultAdaptiveBufferCheckBox.Checked = _workingSettings.DefaultUseAdaptiveBuffer
        _defaultBufferMbNumeric.Value = ClampNumeric(_defaultBufferMbNumeric, _workingSettings.DefaultBufferSizeMb)
        _defaultRetriesNumeric.Value = ClampNumeric(_defaultRetriesNumeric, _workingSettings.DefaultMaxRetries)
        _defaultOperationTimeoutNumeric.Value = ClampNumeric(_defaultOperationTimeoutNumeric, _workingSettings.DefaultOperationTimeoutSeconds)
        _defaultPerFileTimeoutNumeric.Value = ClampNumeric(_defaultPerFileTimeoutNumeric, _workingSettings.DefaultPerFileTimeoutSeconds)
        _defaultMaxThroughputNumeric.Value = ClampNumeric(_defaultMaxThroughputNumeric, _workingSettings.DefaultMaxThroughputMbPerSecond)

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

        Select Case SettingsValueConverter.ToOverwritePolicy(_workingSettings.DefaultOverwritePolicy)
            Case OverwritePolicy.SkipExisting
                _overwritePolicyComboBox.SelectedIndex = 1
            Case OverwritePolicy.OverwriteIfSourceNewer
                _overwritePolicyComboBox.SelectedIndex = 2
            Case Else
                _overwritePolicyComboBox.SelectedIndex = 0
        End Select

        Select Case SettingsValueConverter.ToSymlinkHandling(_workingSettings.DefaultSymlinkHandling)
            Case SymlinkHandlingMode.Follow
                _symlinkHandlingComboBox.SelectedIndex = 1
            Case Else
                _symlinkHandlingComboBox.SelectedIndex = 0
        End Select

        Select Case SettingsValueConverter.ToSalvageFillPattern(_workingSettings.DefaultSalvageFillPattern)
            Case SalvageFillPattern.Ones
                _salvageFillPatternComboBox.SelectedIndex = 1
            Case SalvageFillPattern.Random
                _salvageFillPatternComboBox.SelectedIndex = 2
            Case Else
                _salvageFillPatternComboBox.SelectedIndex = 0
        End Select

        Select Case SettingsValueConverter.ToVerificationMode(_workingSettings.DefaultVerificationMode)
            Case VerificationMode.Sampled
                _defaultVerificationModeComboBox.SelectedIndex = 1
            Case Else
                _defaultVerificationModeComboBox.SelectedIndex = 0
        End Select

        Select Case SettingsValueConverter.ToVerificationHashAlgorithm(_workingSettings.DefaultVerificationHashAlgorithm)
            Case VerificationHashAlgorithm.Sha512
                _defaultHashAlgorithmComboBox.SelectedIndex = 1
            Case Else
                _defaultHashAlgorithmComboBox.SelectedIndex = 0
        End Select

        Select Case SettingsValueConverter.ToExplorerSelectionMode(_workingSettings.ExplorerSelectionMode)
            Case ExplorerSelectionMode.SourceFolder
                _explorerSelectionModeComboBox.SelectedIndex = 1
            Case Else
                _explorerSelectionModeComboBox.SelectedIndex = 0
        End Select

        EnsureSelection(_themeComboBox)
        EnsureSelection(_overwritePolicyComboBox)
        EnsureSelection(_symlinkHandlingComboBox)
        EnsureSelection(_salvageFillPatternComboBox)
        EnsureSelection(_defaultVerificationModeComboBox)
        EnsureSelection(_defaultHashAlgorithmComboBox)
        EnsureSelection(_explorerSelectionModeComboBox)

        UpdateRecoveryControlStates()
        UpdateVerificationControlStates()
        UpdateExplorerControlStates()
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

    Private Sub RecoveryControl_CheckedChanged(sender As Object, e As EventArgs)
        UpdateRecoveryControlStates()
    End Sub

    Private Sub DefaultVerifyCheckBox_CheckedChanged(sender As Object, e As EventArgs)
        UpdateVerificationControlStates()
    End Sub

    Private Sub ExplorerControl_CheckedChanged(sender As Object, e As EventArgs)
        UpdateExplorerControlStates()
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

    Private Sub OkButton_Click(sender As Object, e As EventArgs)
        Dim updateUrl = _updateUrlTextBox.Text.Trim()
        If updateUrl.Length > 0 AndAlso Not IsValidHttpUrl(updateUrl) Then
            MessageBox.Show(Me, "Release URL must be an absolute http/https URL.", "Settings", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            _updateUrlTextBox.Focus()
            Return
        End If

        _workingSettings.Theme = ThemeSettings.ModeToString(GetSelectedMode())
        _workingSettings.CheckUpdatesOnLaunch = _checkUpdatesOnLaunchCheckBox.Checked
        _workingSettings.UpdateReleaseUrl = updateUrl
        _workingSettings.UserAgent = _userAgentTextBox.Text.Trim()

        _workingSettings.EnableRecoveryAutostart = _enableRecoveryAutostartCheckBox.Checked
        _workingSettings.AutoResumeAfterCrash = _autoResumeAfterCrashCheckBox.Checked
        _workingSettings.PromptResumeAfterCrash = If(_autoResumeAfterCrashCheckBox.Checked, True, _promptResumeAfterCrashCheckBox.Checked)
        _workingSettings.KeepResumePromptUntilResolved = If(_autoResumeAfterCrashCheckBox.Checked, False, _keepResumePromptCheckBox.Checked)
        _workingSettings.AutoRunQueuedJobsOnStartup = _autoRunQueuedOnStartupCheckBox.Checked
        _workingSettings.RecoveryTouchIntervalSeconds = CInt(_recoveryTouchIntervalNumeric.Value)

        _workingSettings.EnableExplorerContextMenu = _enableExplorerContextMenuCheckBox.Checked
        _workingSettings.ExplorerSelectionMode = If(_explorerSelectionModeComboBox.SelectedIndex = 1,
                                                    SettingsValueConverter.ExplorerSelectionModeToString(ExplorerSelectionMode.SourceFolder),
                                                    SettingsValueConverter.ExplorerSelectionModeToString(ExplorerSelectionMode.SelectedItems))

        _workingSettings.DefaultResumeFromJournal = _defaultResumeCheckBox.Checked
        _workingSettings.DefaultSalvageUnreadableBlocks = _defaultSalvageCheckBox.Checked
        _workingSettings.DefaultContinueOnFileError = _defaultContinueOnErrorCheckBox.Checked
        _workingSettings.DefaultPreserveTimestamps = _defaultPreserveTimestampsCheckBox.Checked
        _workingSettings.DefaultCopyEmptyDirectories = _defaultCopyEmptyDirectoriesCheckBox.Checked
        _workingSettings.DefaultWaitForMediaAvailability = _defaultWaitForMediaCheckBox.Checked
        _workingSettings.DefaultUseAdaptiveBuffer = _defaultAdaptiveBufferCheckBox.Checked
        _workingSettings.DefaultBufferSizeMb = CInt(_defaultBufferMbNumeric.Value)
        _workingSettings.DefaultMaxRetries = CInt(_defaultRetriesNumeric.Value)
        _workingSettings.DefaultOperationTimeoutSeconds = CInt(_defaultOperationTimeoutNumeric.Value)
        _workingSettings.DefaultPerFileTimeoutSeconds = CInt(_defaultPerFileTimeoutNumeric.Value)
        _workingSettings.DefaultMaxThroughputMbPerSecond = CInt(_defaultMaxThroughputNumeric.Value)

        Select Case _overwritePolicyComboBox.SelectedIndex
            Case 1
                _workingSettings.DefaultOverwritePolicy = SettingsValueConverter.OverwritePolicyToString(OverwritePolicy.SkipExisting)
            Case 2
                _workingSettings.DefaultOverwritePolicy = SettingsValueConverter.OverwritePolicyToString(OverwritePolicy.OverwriteIfSourceNewer)
            Case Else
                _workingSettings.DefaultOverwritePolicy = SettingsValueConverter.OverwritePolicyToString(OverwritePolicy.Overwrite)
        End Select

        Select Case _symlinkHandlingComboBox.SelectedIndex
            Case 1
                _workingSettings.DefaultSymlinkHandling = SettingsValueConverter.SymlinkHandlingToString(SymlinkHandlingMode.Follow)
            Case Else
                _workingSettings.DefaultSymlinkHandling = SettingsValueConverter.SymlinkHandlingToString(SymlinkHandlingMode.Skip)
        End Select

        Select Case _salvageFillPatternComboBox.SelectedIndex
            Case 1
                _workingSettings.DefaultSalvageFillPattern = SettingsValueConverter.SalvageFillPatternToString(SalvageFillPattern.Ones)
            Case 2
                _workingSettings.DefaultSalvageFillPattern = SettingsValueConverter.SalvageFillPatternToString(SalvageFillPattern.Random)
            Case Else
                _workingSettings.DefaultSalvageFillPattern = SettingsValueConverter.SalvageFillPatternToString(SalvageFillPattern.Zero)
        End Select

        _workingSettings.DefaultVerifyAfterCopy = _defaultVerifyCheckBox.Checked
        _workingSettings.DefaultVerificationMode = If(_defaultVerificationModeComboBox.SelectedIndex = 1,
                                                      SettingsValueConverter.VerificationModeToString(VerificationMode.Sampled),
                                                      SettingsValueConverter.VerificationModeToString(VerificationMode.Full))
        _workingSettings.DefaultVerificationHashAlgorithm = If(_defaultHashAlgorithmComboBox.SelectedIndex = 1,
                                                               SettingsValueConverter.VerificationHashAlgorithmToString(VerificationHashAlgorithm.Sha512),
                                                               SettingsValueConverter.VerificationHashAlgorithmToString(VerificationHashAlgorithm.Sha256))
        _workingSettings.DefaultSampleVerificationChunkKb = CInt(_sampleChunkKbNumeric.Value)
        _workingSettings.DefaultSampleVerificationChunkCount = CInt(_sampleChunkCountNumeric.Value)

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
        Return New AppSettings() With {
            .Theme = source.Theme,
            .CheckUpdatesOnLaunch = source.CheckUpdatesOnLaunch,
            .UpdateReleaseUrl = source.UpdateReleaseUrl,
            .UserAgent = source.UserAgent,
            .EnableRecoveryAutostart = source.EnableRecoveryAutostart,
            .PromptResumeAfterCrash = source.PromptResumeAfterCrash,
            .AutoResumeAfterCrash = source.AutoResumeAfterCrash,
            .KeepResumePromptUntilResolved = source.KeepResumePromptUntilResolved,
            .AutoRunQueuedJobsOnStartup = source.AutoRunQueuedJobsOnStartup,
            .RecoveryTouchIntervalSeconds = source.RecoveryTouchIntervalSeconds,
            .EnableExplorerContextMenu = source.EnableExplorerContextMenu,
            .ExplorerSelectionMode = source.ExplorerSelectionMode,
            .DefaultResumeFromJournal = source.DefaultResumeFromJournal,
            .DefaultSalvageUnreadableBlocks = source.DefaultSalvageUnreadableBlocks,
            .DefaultContinueOnFileError = source.DefaultContinueOnFileError,
            .DefaultVerifyAfterCopy = source.DefaultVerifyAfterCopy,
            .DefaultUseAdaptiveBuffer = source.DefaultUseAdaptiveBuffer,
            .DefaultWaitForMediaAvailability = source.DefaultWaitForMediaAvailability,
            .DefaultBufferSizeMb = source.DefaultBufferSizeMb,
            .DefaultMaxRetries = source.DefaultMaxRetries,
            .DefaultOperationTimeoutSeconds = source.DefaultOperationTimeoutSeconds,
            .DefaultPerFileTimeoutSeconds = source.DefaultPerFileTimeoutSeconds,
            .DefaultMaxThroughputMbPerSecond = source.DefaultMaxThroughputMbPerSecond,
            .DefaultPreserveTimestamps = source.DefaultPreserveTimestamps,
            .DefaultCopyEmptyDirectories = source.DefaultCopyEmptyDirectories,
            .DefaultOverwritePolicy = source.DefaultOverwritePolicy,
            .DefaultSymlinkHandling = source.DefaultSymlinkHandling,
            .DefaultVerificationMode = source.DefaultVerificationMode,
            .DefaultVerificationHashAlgorithm = source.DefaultVerificationHashAlgorithm,
            .DefaultSampleVerificationChunkKb = source.DefaultSampleVerificationChunkKb,
            .DefaultSampleVerificationChunkCount = source.DefaultSampleVerificationChunkCount,
            .DefaultSalvageFillPattern = source.DefaultSalvageFillPattern
        }
    End Function
End Class
