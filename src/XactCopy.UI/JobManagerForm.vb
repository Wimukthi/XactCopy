Imports System.Diagnostics
Imports System.Drawing
Imports System.IO
Imports System.Linq
Imports System.Reflection
Imports System.Runtime.InteropServices
Imports XactCopy.Models
Imports XactCopy.Services

''' <summary>
''' Enum JobManagerRequestAction.
''' </summary>
Friend Enum JobManagerRequestAction
    None = 0
    RunSavedJob = 1
    RunQueuedEntry = 2
End Enum

''' <summary>
''' Class JobManagerForm.
''' </summary>
Friend Class JobManagerForm
    Inherits Form

    Private Const LayoutWindowKey As String = "JobManagerForm.V3"
    Private Const GridLayoutKey As String = "JobManagerForm.V3.Grid"
    Private Const SplitterKey As String = "JobManagerForm.V3.Split"
    Private Const WmSetRedraw As Integer = &HB
    Private Shared ReadOnly _doubleBufferedProperty As PropertyInfo =
        GetType(Control).GetProperty("DoubleBuffered", BindingFlags.Instance Or BindingFlags.NonPublic)

    Private ReadOnly _jobManager As JobManagerService

    Private ReadOnly _summaryLabel As New Label()
    Private ReadOnly _viewLabel As New Label()
    Private ReadOnly _viewComboBox As New ComboBox()
    Private ReadOnly _statusLabel As New Label()
    Private ReadOnly _statusComboBox As New ComboBox()
    Private ReadOnly _searchLabel As New Label()
    Private ReadOnly _searchTextBox As New TextBox()
    Private ReadOnly _refreshButton As New Button()

    Private ReadOnly _mainGrid As New DataGridView()
    Private ReadOnly _detailsTextBox As New TextBox()
    Private ReadOnly _contentSplit As New SplitContainer()

    Private ReadOnly _runNowButton As New Button()
    Private ReadOnly _queueButton As New Button()
    Private ReadOnly _removeQueueButton As New Button()
    Private ReadOnly _moveQueueUpButton As New Button()
    Private ReadOnly _moveQueueDownButton As New Button()
    Private ReadOnly _deleteJobButton As New Button()
    Private ReadOnly _deleteRunButton As New Button()
    Private ReadOnly _renameJobButton As New Button()
    Private ReadOnly _duplicateJobButton As New Button()
    Private ReadOnly _openJournalButton As New Button()
    Private ReadOnly _moveQueueTopButton As New Button()
    Private ReadOnly _moveQueueBottomButton As New Button()
    Private ReadOnly _clearQueueButton As New Button()
    Private ReadOnly _clearHistoryButton As New Button()
    Private ReadOnly _closeButton As New Button()

    Private ReadOnly _gridContextMenu As New ContextMenuStrip()
    Private ReadOnly _autoRefreshTimer As New Timer() With {.Interval = 3000}

    Private ReadOnly _toolTip As New ToolTip() With {
        .ShowAlways = True,
        .InitialDelay = 350,
        .ReshowDelay = 120,
        .AutoPopDelay = 24000
    }

    Private _requestedAction As JobManagerRequestAction = JobManagerRequestAction.None
    Private _requestedRunJobId As String = String.Empty
    Private _requestedQueueEntryId As String = String.Empty

    Private _savedJobs As IReadOnlyList(Of ManagedJob) = Array.Empty(Of ManagedJob)()
    Private _queueEntries As IReadOnlyList(Of JobQueueEntryView) = Array.Empty(Of JobQueueEntryView)()
    Private _runs As IReadOnlyList(Of ManagedJobRun) = Array.Empty(Of ManagedJobRun)()
    Private _savedJobsById As IReadOnlyDictionary(Of String, ManagedJob) =
        New Dictionary(Of String, ManagedJob)(StringComparer.OrdinalIgnoreCase)
    Private _queueEntriesById As IReadOnlyDictionary(Of String, JobQueueEntryView) =
        New Dictionary(Of String, JobQueueEntryView)(StringComparer.OrdinalIgnoreCase)
    Private _runsById As IReadOnlyDictionary(Of String, ManagedJobRun) =
        New Dictionary(Of String, ManagedJobRun)(StringComparer.OrdinalIgnoreCase)
    Private _visibleRows As List(Of GridRowModel) = New List(Of GridRowModel)()
    Private ReadOnly _filterDebounceTimer As New Timer() With {.Interval = 150}
    Private _isPopulatingGrid As Boolean

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
    Public Sub New(jobManager As JobManagerService)
        If jobManager Is Nothing Then
            Throw New ArgumentNullException(NameOf(jobManager))
        End If

        _jobManager = jobManager

        Text = "Job Manager"
        StartPosition = FormStartPosition.CenterParent
        MinimumSize = New Size(1080, 700)
        Size = New Size(1320, 820)
        WindowIconHelper.Apply(Me)
        SetStyle(ControlStyles.AllPaintingInWmPaint Or ControlStyles.OptimizedDoubleBuffer, True)
        UpdateStyles()

        BuildUi()
        EnableDoubleBufferingForContainers(Me)
        ConfigureToolTips()
        ConfigureContextMenu()
        AddHandler _filterDebounceTimer.Tick, AddressOf FilterDebounceTimer_Tick
        AddHandler _autoRefreshTimer.Tick, AddressOf AutoRefreshTimer_Tick

        AddHandler Shown, AddressOf JobManagerForm_Shown
        AddHandler FormClosing, AddressOf JobManagerForm_FormClosing
    End Sub

    ''' <summary>
    ''' Gets or sets RequestedAction.
    ''' </summary>
    Public ReadOnly Property RequestedAction As JobManagerRequestAction
        Get
            Return _requestedAction
        End Get
    End Property

    ''' <summary>
    ''' Gets or sets RequestedRunJobId.
    ''' </summary>
    Public ReadOnly Property RequestedRunJobId As String
        Get
            Return _requestedRunJobId
        End Get
    End Property

    ''' <summary>
    ''' Gets or sets RequestedQueueEntryId.
    ''' </summary>
    Public ReadOnly Property RequestedQueueEntryId As String
        Get
            Return _requestedQueueEntryId
        End Get
    End Property

    Private Sub BuildUi()
        ConfigureGrid()
        ConfigureSplitContainer()

        Dim rootLayout As New TableLayoutPanel() With {
            .Dock = DockStyle.Fill,
            .ColumnCount = 1,
            .RowCount = 4,
            .Padding = New Padding(8)
        }
        rootLayout.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0F))
        rootLayout.RowStyles.Add(New RowStyle(SizeType.AutoSize))
        rootLayout.RowStyles.Add(New RowStyle(SizeType.AutoSize))
        rootLayout.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0F))
        rootLayout.RowStyles.Add(New RowStyle(SizeType.Absolute, 48.0F))

        rootLayout.Controls.Add(BuildHeaderPane(), 0, 0)
        rootLayout.Controls.Add(BuildFilterPane(), 0, 1)
        rootLayout.Controls.Add(_contentSplit, 0, 2)
        rootLayout.Controls.Add(BuildFooterPane(), 0, 3)

        Controls.Add(rootLayout)
    End Sub

    Private Function BuildHeaderPane() As Control
        Dim layout As New TableLayoutPanel() With {
            .Dock = DockStyle.Fill,
            .ColumnCount = 2,
            .RowCount = 1,
            .Margin = New Padding(0, 0, 0, 6),
            .AutoSize = True,
            .AutoSizeMode = AutoSizeMode.GrowAndShrink
        }
        layout.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0F))
        layout.ColumnStyles.Add(New ColumnStyle(SizeType.AutoSize))

        Dim titleLayout As New TableLayoutPanel() With {
            .Dock = DockStyle.Fill,
            .ColumnCount = 1,
            .RowCount = 2,
            .Margin = New Padding(0),
            .AutoSize = True,
            .AutoSizeMode = AutoSizeMode.GrowAndShrink
        }
        titleLayout.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0F))
        titleLayout.RowStyles.Add(New RowStyle(SizeType.AutoSize))
        titleLayout.RowStyles.Add(New RowStyle(SizeType.AutoSize))

        Dim titleLabel As New Label() With {
            .AutoSize = True,
            .Font = New Font(Font, FontStyle.Bold),
            .Text = "Jobs Console",
            .Margin = New Padding(0, 0, 0, 2)
        }

        Dim subtitleLabel As New Label() With {
            .AutoSize = True,
            .Text = "Manage saved jobs, queue order, and execution history from one grid.",
            .Margin = New Padding(0)
        }

        _summaryLabel.AutoSize = True
        _summaryLabel.Anchor = AnchorStyles.Right Or AnchorStyles.Top
        _summaryLabel.Text = "Saved: 0 | Queued: 0 | Runs: 0"

        titleLayout.Controls.Add(titleLabel, 0, 0)
        titleLayout.Controls.Add(subtitleLabel, 0, 1)

        layout.Controls.Add(titleLayout, 0, 0)
        layout.Controls.Add(_summaryLabel, 1, 0)

        Return layout
    End Function

    Private Function BuildFilterPane() As Control
        Dim layout As New TableLayoutPanel() With {
            .Dock = DockStyle.Fill,
            .ColumnCount = 8,
            .RowCount = 1,
            .Margin = New Padding(0, 0, 0, 6),
            .AutoSize = True,
            .AutoSizeMode = AutoSizeMode.GrowAndShrink
        }
        layout.RowStyles.Add(New RowStyle(SizeType.Absolute, 32.0F))
        layout.ColumnStyles.Add(New ColumnStyle(SizeType.AutoSize))
        layout.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 180.0F))
        layout.ColumnStyles.Add(New ColumnStyle(SizeType.AutoSize))
        layout.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 180.0F))
        layout.ColumnStyles.Add(New ColumnStyle(SizeType.AutoSize))
        layout.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0F))
        layout.ColumnStyles.Add(New ColumnStyle(SizeType.AutoSize))
        layout.ColumnStyles.Add(New ColumnStyle(SizeType.AutoSize))

        _viewLabel.Text = "View"
        _viewLabel.AutoSize = True
        _viewLabel.Anchor = AnchorStyles.Left
        _viewLabel.Margin = New Padding(0, 0, 8, 0)

        _viewComboBox.DropDownStyle = ComboBoxStyle.DropDownList
        _viewComboBox.Anchor = AnchorStyles.Left Or AnchorStyles.Right
        _viewComboBox.Margin = New Padding(0, 0, 8, 0)
        _viewComboBox.Items.Add(New EnumChoice(Of JobManagerViewFilter)(JobManagerViewFilter.AllItems, "All Items"))
        _viewComboBox.Items.Add(New EnumChoice(Of JobManagerViewFilter)(JobManagerViewFilter.SavedJobs, "Saved Jobs"))
        _viewComboBox.Items.Add(New EnumChoice(Of JobManagerViewFilter)(JobManagerViewFilter.Queue, "Queue"))
        _viewComboBox.Items.Add(New EnumChoice(Of JobManagerViewFilter)(JobManagerViewFilter.RunHistory, "Run History"))
        AddHandler _viewComboBox.SelectedIndexChanged, AddressOf FilterChanged

        _statusLabel.Text = "Run Status"
        _statusLabel.AutoSize = True
        _statusLabel.Anchor = AnchorStyles.Left
        _statusLabel.Margin = New Padding(0, 0, 8, 0)

        _statusComboBox.DropDownStyle = ComboBoxStyle.DropDownList
        _statusComboBox.Anchor = AnchorStyles.Left Or AnchorStyles.Right
        _statusComboBox.Margin = New Padding(0, 0, 8, 0)
        _statusComboBox.Items.Add(New EnumChoice(Of RunStatusFilter)(RunStatusFilter.Any, "Any"))
        _statusComboBox.Items.Add(New EnumChoice(Of RunStatusFilter)(RunStatusFilter.Queued, "Queued"))
        _statusComboBox.Items.Add(New EnumChoice(Of RunStatusFilter)(RunStatusFilter.Running, "Running"))
        _statusComboBox.Items.Add(New EnumChoice(Of RunStatusFilter)(RunStatusFilter.Paused, "Paused"))
        _statusComboBox.Items.Add(New EnumChoice(Of RunStatusFilter)(RunStatusFilter.Completed, "Completed"))
        _statusComboBox.Items.Add(New EnumChoice(Of RunStatusFilter)(RunStatusFilter.Failed, "Failed"))
        _statusComboBox.Items.Add(New EnumChoice(Of RunStatusFilter)(RunStatusFilter.Cancelled, "Cancelled"))
        _statusComboBox.Items.Add(New EnumChoice(Of RunStatusFilter)(RunStatusFilter.Interrupted, "Interrupted"))
        AddHandler _statusComboBox.SelectedIndexChanged, AddressOf FilterChanged

        _searchLabel.Text = "Search"
        _searchLabel.AutoSize = True
        _searchLabel.Anchor = AnchorStyles.Left
        _searchLabel.Margin = New Padding(0, 0, 8, 0)

        _searchTextBox.Anchor = AnchorStyles.Left Or AnchorStyles.Right
        _searchTextBox.Margin = New Padding(0, 0, 8, 0)
        AddHandler _searchTextBox.TextChanged, AddressOf FilterChanged

        _refreshButton.Text = "Refresh"
        _refreshButton.AutoSize = True
        _refreshButton.Anchor = AnchorStyles.Left Or AnchorStyles.Top
        _refreshButton.Margin = New Padding(0, 1, 0, 0)
        AddHandler _refreshButton.Click,
            Sub(sender, e)
                RefreshData()
            End Sub

        layout.Controls.Add(_viewLabel, 0, 0)
        layout.Controls.Add(_viewComboBox, 1, 0)
        layout.Controls.Add(_statusLabel, 2, 0)
        layout.Controls.Add(_statusComboBox, 3, 0)
        layout.Controls.Add(_searchLabel, 4, 0)
        layout.Controls.Add(_searchTextBox, 5, 0)
        layout.Controls.Add(_refreshButton, 6, 0)

        Return layout
    End Function

    Private Sub ConfigureSplitContainer()
        _contentSplit.Dock = DockStyle.Fill
        _contentSplit.Orientation = Orientation.Horizontal
        _contentSplit.SplitterWidth = 6
        _contentSplit.Panel1.Controls.Add(_mainGrid)

        Dim detailsLayout As New TableLayoutPanel() With {
            .Dock = DockStyle.Fill,
            .ColumnCount = 1,
            .RowCount = 2,
            .Padding = New Padding(0)
        }
        detailsLayout.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0F))
        detailsLayout.RowStyles.Add(New RowStyle(SizeType.AutoSize))
        detailsLayout.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0F))

        Dim detailsTitle As New Label() With {
            .Text = "Selected Item Details",
            .AutoSize = True,
            .Font = New Font(Font, FontStyle.Bold),
            .Margin = New Padding(0, 0, 0, 4)
        }

        _detailsTextBox.Dock = DockStyle.Fill
        _detailsTextBox.Multiline = True
        _detailsTextBox.ReadOnly = True
        _detailsTextBox.ScrollBars = ScrollBars.Both
        _detailsTextBox.WordWrap = False

        detailsLayout.Controls.Add(detailsTitle, 0, 0)
        detailsLayout.Controls.Add(_detailsTextBox, 0, 1)
        _contentSplit.Panel2.Controls.Add(detailsLayout)
    End Sub

    Private Function BuildFooterPane() As Control
        Dim footer As New TableLayoutPanel() With {
            .Dock = DockStyle.Fill,
            .ColumnCount = 2,
            .RowCount = 1,
            .Margin = New Padding(0, 4, 0, 0),
            .AutoSize = False
        }
        footer.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0F))
        footer.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0F))
        footer.ColumnStyles.Add(New ColumnStyle(SizeType.AutoSize))

        Dim actionsPanel As New FlowLayoutPanel() With {
            .Dock = DockStyle.Fill,
            .FlowDirection = FlowDirection.LeftToRight,
            .WrapContents = True,
            .AutoSize = False,
            .Margin = New Padding(0)
        }

        ConfigureActionButton(_runNowButton, "&Run Now", AddressOf RunNowButton_Click)
        ConfigureActionButton(_queueButton, "&Queue", AddressOf QueueButton_Click)
        ConfigureActionButton(_removeQueueButton, "Rem&ove Queue", AddressOf RemoveQueueButton_Click)
        ConfigureActionButton(_moveQueueTopButton, "Top", AddressOf MoveQueueTopButton_Click)
        ConfigureActionButton(_moveQueueUpButton, "Up", AddressOf MoveQueueUpButton_Click)
        ConfigureActionButton(_moveQueueDownButton, "Down", AddressOf MoveQueueDownButton_Click)
        ConfigureActionButton(_moveQueueBottomButton, "Bottom", AddressOf MoveQueueBottomButton_Click)
        ConfigureActionButton(_renameJobButton, "Re&name", AddressOf RenameJobButton_Click)
        ConfigureActionButton(_duplicateJobButton, "&Duplicate", AddressOf DuplicateJobButton_Click)
        ConfigureActionButton(_deleteJobButton, "Delete &Job", AddressOf DeleteJobButton_Click)
        ConfigureActionButton(_deleteRunButton, "Delete Ru&n", AddressOf DeleteRunButton_Click)
        ConfigureActionButton(_openJournalButton, "Open &Journal", AddressOf OpenJournalButton_Click)
        ConfigureActionButton(_clearQueueButton, "Clear Queue", AddressOf ClearQueueButton_Click)
        ConfigureActionButton(_clearHistoryButton, "Clear History", AddressOf ClearHistoryButton_Click)

        actionsPanel.Controls.Add(_runNowButton)
        actionsPanel.Controls.Add(_queueButton)
        actionsPanel.Controls.Add(_removeQueueButton)
        actionsPanel.Controls.Add(_moveQueueTopButton)
        actionsPanel.Controls.Add(_moveQueueUpButton)
        actionsPanel.Controls.Add(_moveQueueDownButton)
        actionsPanel.Controls.Add(_moveQueueBottomButton)
        actionsPanel.Controls.Add(_renameJobButton)
        actionsPanel.Controls.Add(_duplicateJobButton)
        actionsPanel.Controls.Add(_deleteJobButton)
        actionsPanel.Controls.Add(_deleteRunButton)
        actionsPanel.Controls.Add(_openJournalButton)
        actionsPanel.Controls.Add(_clearQueueButton)
        actionsPanel.Controls.Add(_clearHistoryButton)

        _closeButton.Text = "Close"
        _closeButton.AutoSize = True
        _closeButton.Anchor = AnchorStyles.Right Or AnchorStyles.Top
        _closeButton.Margin = New Padding(8, 0, 0, 0)
        AddHandler _closeButton.Click,
            Sub(sender, e)
                DialogResult = DialogResult.Cancel
                Close()
            End Sub

        footer.Controls.Add(actionsPanel, 0, 0)
        footer.Controls.Add(_closeButton, 1, 0)
        Return footer
    End Function

    Private Sub ConfigureActionButton(button As Button, text As String, handler As EventHandler)
        button.Text = text
        button.AutoSize = True
        AddHandler button.Click, handler
    End Sub

    Private Sub ConfigureGrid()
        _mainGrid.Dock = DockStyle.Fill
        _mainGrid.ReadOnly = True
        _mainGrid.AllowUserToAddRows = False
        _mainGrid.AllowUserToDeleteRows = False
        _mainGrid.AllowUserToResizeRows = False
        _mainGrid.AllowUserToOrderColumns = True
        _mainGrid.MultiSelect = False
        _mainGrid.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        _mainGrid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
        _mainGrid.VirtualMode = True
        _mainGrid.EditMode = DataGridViewEditMode.EditProgrammatically
        _mainGrid.AutoGenerateColumns = False
        _mainGrid.RowHeadersVisible = False
        _mainGrid.DefaultCellStyle.WrapMode = DataGridViewTriState.False
        _mainGrid.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        _mainGrid.ColumnHeadersHeight = 28
        _mainGrid.RowTemplate.Height = 24
        _mainGrid.ScrollBars = ScrollBars.Both

        _mainGrid.Columns.Add("Type", "Type")
        _mainGrid.Columns.Add("Name", "Name")
        _mainGrid.Columns.Add("State", "State")
        _mainGrid.Columns.Add("QueuePos", "Queue")
        _mainGrid.Columns.Add("Trigger", "Trigger")
        _mainGrid.Columns.Add("Started", "Started")
        _mainGrid.Columns.Add("Updated", "Updated")
        _mainGrid.Columns.Add("Source", "Source")
        _mainGrid.Columns.Add("Destination", "Destination")
        _mainGrid.Columns.Add("Summary", "Summary")

        _mainGrid.Columns("Type").Width = 85
        _mainGrid.Columns("Name").Width = 220
        _mainGrid.Columns("State").Width = 120
        _mainGrid.Columns("QueuePos").Width = 75
        _mainGrid.Columns("Trigger").Width = 130
        _mainGrid.Columns("Started").Width = 160
        _mainGrid.Columns("Updated").Width = 160
        _mainGrid.Columns("Source").Width = 320
        _mainGrid.Columns("Destination").Width = 320
        _mainGrid.Columns("Summary").Width = 380

        AddHandler _mainGrid.CellValueNeeded, AddressOf MainGrid_CellValueNeeded
        AddHandler _mainGrid.CellFormatting, AddressOf MainGrid_CellFormatting
        AddHandler _mainGrid.SelectionChanged, AddressOf MainGrid_SelectionChanged
        AddHandler _mainGrid.CellDoubleClick, AddressOf MainGrid_CellDoubleClick
        AddHandler _mainGrid.KeyDown, AddressOf MainGrid_KeyDown
    End Sub

    Private Sub ConfigureToolTips()
        SetDetailedToolTip(_summaryLabel, "Shows counts for saved jobs, queued jobs, and run history.", "Use it as a quick sanity check after filtering.")
        SetDetailedToolTip(_viewComboBox, "Chooses whether the grid shows saved jobs, queue entries, or run history.", "The action buttons change based on this view.")
        SetDetailedToolTip(_statusComboBox, "Filters run history by final status.", "Only affects the Run History view.")
        SetDetailedToolTip(_searchTextBox, "Filters rows by job name, paths, trigger, or summary text.", "Useful when history gets large.")
        SetDetailedToolTip(_refreshButton, "Reloads jobs, queue, and history from disk.", "Use after another window or process changes job data.")

        SetDetailedToolTip(_runNowButton, "Runs the selected saved job or queue entry immediately.", "Shortcut: Enter.")
        SetDetailedToolTip(_queueButton, "Adds the selected saved job to the end of the queue.", "Queued jobs can run later or at startup if that setting is enabled.")
        SetDetailedToolTip(_removeQueueButton, "Removes the selected entry from the queue.", "The saved job itself is not deleted.")
        SetDetailedToolTip(_moveQueueTopButton, "Moves the selected queued job to the front.", "Use when a job should run next.")
        SetDetailedToolTip(_moveQueueUpButton, "Moves the selected queued job one slot earlier.")
        SetDetailedToolTip(_moveQueueDownButton, "Moves the selected queued job one slot later.")
        SetDetailedToolTip(_moveQueueBottomButton, "Moves the selected queued job to the end.", "Use when lower priority work can wait.")
        SetDetailedToolTip(_renameJobButton, "Renames the selected saved job.", "Shortcut: F2. Paths and options are unchanged.")
        SetDetailedToolTip(_duplicateJobButton, "Creates a new saved job from the selected job.", "Useful for making a variant without rebuilding options.")
        SetDetailedToolTip(_deleteJobButton, "Deletes the selected saved job.", "Shortcut: Delete. Run history is not removed.")
        SetDetailedToolTip(_deleteRunButton, "Deletes the selected run history entry.", "Shortcut: Delete. Saved jobs and journals are not changed.")
        SetDetailedToolTip(_openJournalButton, "Opens the journal file for the selected run.", "Useful for troubleshooting resume, recovery, and file status.")
        SetDetailedToolTip(_clearQueueButton, "Removes every queued entry.", "Saved jobs remain available.")
        SetDetailedToolTip(_clearHistoryButton, "Deletes all run history entries.", "This does not delete saved jobs.")
        SetDetailedToolTip(_detailsTextBox, "Shows details for the selected row.", "Includes paths, options, run result, and journal data when available.")
    End Sub

    Private Sub SetDetailedToolTip(control As Control, ParamArray lines() As String)
        _toolTip.SetToolTip(control, TooltipScenarioFormatter.Compose(lines))
    End Sub

    Private Sub JobManagerForm_Shown(sender As Object, e As EventArgs)
        UiLayoutManager.ApplyWindow(Me, LayoutWindowKey)
        UiLayoutManager.ApplyGridLayout(_mainGrid, GridLayoutKey)

        _viewComboBox.SelectedIndex = 0
        _statusComboBox.SelectedIndex = 0
        _filterDebounceTimer.Stop()

        RefreshData()
        _autoRefreshTimer.Start()

        ' Apply splitter after initial layout pass so persisted distance/min sizes are honored.
        BeginInvoke(New Action(
            Sub()
                UiLayoutManager.ApplySplitter(
                    _contentSplit,
                    SplitterKey,
                    defaultDistance:=CInt(Math.Max(260, _contentSplit.Height * 0.68R)),
                    panel1Min:=320,
                    panel2Min:=160)
            End Sub))
    End Sub

    Private Sub JobManagerForm_FormClosing(sender As Object, e As FormClosingEventArgs)
        _filterDebounceTimer.Stop()
        _autoRefreshTimer.Stop()
        UiLayoutManager.CaptureGridLayout(_mainGrid, GridLayoutKey)
        UiLayoutManager.CaptureSplitter(_contentSplit, SplitterKey)
        UiLayoutManager.CaptureWindow(Me, LayoutWindowKey)
    End Sub

    Private Sub FilterChanged(sender As Object, e As EventArgs)
        _filterDebounceTimer.Stop()
        _filterDebounceTimer.Start()
    End Sub

    Private Sub FilterDebounceTimer_Tick(sender As Object, e As EventArgs)
        _filterDebounceTimer.Stop()
        PopulateGrid(keepSelection:=True)
    End Sub

    Private Sub RefreshData()
        _filterDebounceTimer.Stop()
        _savedJobs = _jobManager.GetJobs()
        _queueEntries = _jobManager.GetQueueEntries()
        _runs = _jobManager.GetRecentRuns(500)
        _savedJobsById = BuildDictionary(_savedJobs, Function(job) job?.JobId)
        _queueEntriesById = BuildDictionary(_queueEntries, Function(entry) entry?.QueueEntryId)
        _runsById = BuildDictionary(_runs, Function(run) run?.RunId)

        _summaryLabel.Text = $"Saved: {_savedJobs.Count} | Queued: {_queueEntries.Count} | Runs: {_runs.Count}"

        PopulateGrid(keepSelection:=True)
    End Sub

    Private Sub PopulateGrid(keepSelection As Boolean)
        Dim previousTag = If(keepSelection, GetSelectedTag(), Nothing)
        Dim view = GetSelectedViewFilter()
        Dim runStatusFilter = GetSelectedRunStatusFilter()
        Dim searchText = If(_searchTextBox.Text, String.Empty).Trim()
        Dim newRows As New List(Of GridRowModel)(Math.Max(64, _savedJobs.Count + _queueEntries.Count + _runs.Count))

        If view = JobManagerViewFilter.AllItems OrElse view = JobManagerViewFilter.SavedJobs Then
            For Each job In _savedJobs
                Dim summary = BuildSavedJobSummary(job)
                If Not MatchesSearch(searchText, "saved", job.Name, "ready", job.Options.SourceRoot, job.Options.DestinationRoot, summary) Then
                    Continue For
                End If

                newRows.Add(New GridRowModel(
                    "Saved",
                    job.Name,
                    "Ready",
                    "-",
                    "manual",
                    job.CreatedUtc.ToLocalTime().ToString("yyyy-MM-dd HH:mm:ss"),
                    job.UpdatedUtc.ToLocalTime().ToString("yyyy-MM-dd HH:mm:ss"),
                    job.Options.SourceRoot,
                    job.Options.DestinationRoot,
                    summary,
                    New GridRowTag(RowKind.SavedJob, job.JobId, job.JobId)))
            Next
        End If

        If view = JobManagerViewFilter.AllItems OrElse view = JobManagerViewFilter.Queue Then
            For Each entry In _queueEntries
                Dim state = If(entry.AttemptCount > 0, "Retry queued", "Queued")
                Dim summary = BuildQueueSummary(entry)
                If Not MatchesSearch(searchText, "queue", entry.JobName, state, entry.SourceRoot, entry.DestinationRoot, summary, entry.Trigger) Then
                    Continue For
                End If

                newRows.Add(New GridRowModel(
                    "Queue",
                    entry.JobName,
                    state,
                    entry.Position.ToString(),
                    entry.Trigger,
                    entry.EnqueuedUtc.ToLocalTime().ToString("yyyy-MM-dd HH:mm:ss"),
                    entry.LastUpdatedUtc.ToLocalTime().ToString("yyyy-MM-dd HH:mm:ss"),
                    entry.SourceRoot,
                    entry.DestinationRoot,
                    summary,
                    New GridRowTag(RowKind.QueueEntry, entry.QueueEntryId, entry.JobId)))
            Next
        End If

        If view = JobManagerViewFilter.AllItems OrElse view = JobManagerViewFilter.RunHistory Then
            For Each run In _runs
                If runStatusFilter.HasValue AndAlso run.Status <> runStatusFilter.Value Then
                    Continue For
                End If

                Dim summary = If(String.IsNullOrWhiteSpace(run.Summary), "-", run.Summary)
                If Not MatchesSearch(searchText, "run", run.DisplayName, run.Status.ToString(), run.SourceRoot, run.DestinationRoot, summary, run.Trigger) Then
                    Continue For
                End If

                Dim queueDisplay = If(run.QueueAttempt > 0, run.QueueAttempt.ToString(), "-")
                newRows.Add(New GridRowModel(
                    "Run",
                    run.DisplayName,
                    run.Status.ToString(),
                    queueDisplay,
                    run.Trigger,
                    run.StartedUtc.ToLocalTime().ToString("yyyy-MM-dd HH:mm:ss"),
                    run.LastUpdatedUtc.ToLocalTime().ToString("yyyy-MM-dd HH:mm:ss"),
                    run.SourceRoot,
                    run.DestinationRoot,
                    summary,
                    New GridRowTag(RowKind.RunHistory, run.RunId, run.JobId)))
            Next
        End If

        _isPopulatingGrid = True
        SuspendDrawing(_mainGrid)
        _mainGrid.SuspendLayout()
        Try
            _visibleRows = newRows
            _mainGrid.RowCount = _visibleRows.Count
            RestoreSelection(previousTag)
        Finally
            _mainGrid.ResumeLayout(performLayout:=False)
            ResumeDrawing(_mainGrid)
            _isPopulatingGrid = False
        End Try

        UpdateActionStates()
        RenderSelectedItemDetails()
    End Sub

    Private Function BuildSavedJobSummary(job As ManagedJob) As String
        If job Is Nothing OrElse job.Options Is Nothing Then
            Return "Saved job definition"
        End If

        Dim mode = If(job.Options.UseAdaptiveBufferSizing, "Adaptive auto buffer", "Manual buffer")
        Dim verify = If(job.Options.VerifyAfterCopy, "Verify", "No verify")
        Return $"{mode}; retries {Math.Max(0, job.Options.MaxRetries)}; {verify}."
    End Function

    Private Function BuildQueueSummary(entry As JobQueueEntryView) As String
        If entry Is Nothing Then
            Return "Queued"
        End If

        Dim baseText = If(String.IsNullOrWhiteSpace(entry.EnqueuedBy), "Queued", $"Queued by {entry.EnqueuedBy}")
        If String.IsNullOrWhiteSpace(entry.LastErrorMessage) Then
            Return baseText
        End If

        Return $"{baseText}; last error: {entry.LastErrorMessage}"
    End Function

    Private Function MatchesSearch(searchText As String, ParamArray fields() As String) As Boolean
        If String.IsNullOrWhiteSpace(searchText) Then
            Return True
        End If

        For Each field In fields
            If field Is Nothing Then
                Continue For
            End If

            If field.IndexOf(searchText, StringComparison.OrdinalIgnoreCase) >= 0 Then
                Return True
            End If
        Next

        Return False
    End Function

    Private Sub RestoreSelection(previousTag As GridRowTag)
        Dim targetIndex As Integer = -1

        If previousTag IsNot Nothing Then
            For index = 0 To _visibleRows.Count - 1
                Dim tag = _visibleRows(index).Tag
                If tag Is Nothing Then
                    Continue For
                End If

                If tag.Kind = previousTag.Kind AndAlso
                    String.Equals(tag.PrimaryId, previousTag.PrimaryId, StringComparison.OrdinalIgnoreCase) Then
                    targetIndex = index
                    Exit For
                End If
            Next
        End If

        If targetIndex < 0 AndAlso _visibleRows.Count > 0 Then
            targetIndex = 0
        End If

        _mainGrid.ClearSelection()
        If targetIndex < 0 OrElse targetIndex >= _visibleRows.Count Then
            _mainGrid.CurrentCell = Nothing
            Return
        End If

        _mainGrid.CurrentCell = _mainGrid.Rows(targetIndex).Cells(0)
        _mainGrid.Rows(targetIndex).Selected = True
    End Sub

    Private Sub MainGrid_CellValueNeeded(sender As Object, e As DataGridViewCellValueEventArgs)
        If e.RowIndex < 0 OrElse e.RowIndex >= _visibleRows.Count Then
            Return
        End If

        Dim row = _visibleRows(e.RowIndex)
        Select Case e.ColumnIndex
            Case 0
                e.Value = row.TypeText
            Case 1
                e.Value = row.NameText
            Case 2
                e.Value = row.StateText
            Case 3
                e.Value = row.QueueText
            Case 4
                e.Value = row.TriggerText
            Case 5
                e.Value = row.StartedText
            Case 6
                e.Value = row.UpdatedText
            Case 7
                e.Value = row.SourceText
            Case 8
                e.Value = row.DestinationText
            Case 9
                e.Value = row.SummaryText
            Case Else
                e.Value = String.Empty
        End Select
    End Sub

    Private Sub MainGrid_SelectionChanged(sender As Object, e As EventArgs)
        If _isPopulatingGrid Then
            Return
        End If

        UpdateActionStates()
        RenderSelectedItemDetails()
    End Sub

    Private Sub MainGrid_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs)
        If e.RowIndex < 0 Then
            Return
        End If

        Dim selectedTag = GetSelectedTag()
        If selectedTag Is Nothing Then
            Return
        End If

        Select Case selectedTag.Kind
            Case RowKind.SavedJob, RowKind.QueueEntry
                RunNowSelected()
            Case RowKind.RunHistory
                OpenJournalForSelected()
        End Select
    End Sub

    Private Sub MainGrid_CellFormatting(sender As Object, e As DataGridViewCellFormattingEventArgs)
        If e.RowIndex < 0 OrElse e.RowIndex >= _visibleRows.Count Then
            Return
        End If

        Dim row = _visibleRows(e.RowIndex)
        Dim state = row.StateText

        Select Case state
            Case "Failed"
                e.CellStyle.ForeColor = Color.FromArgb(200, 40, 40)
            Case "Running"
                e.CellStyle.ForeColor = Color.FromArgb(20, 140, 60)
            Case "Completed"
                e.CellStyle.ForeColor = Color.FromArgb(30, 100, 180)
            Case "Cancelled", "Interrupted"
                e.CellStyle.ForeColor = Color.FromArgb(160, 120, 20)
            Case "Paused"
                e.CellStyle.ForeColor = Color.FromArgb(140, 80, 180)
            Case "Queued", "Retry queued"
                e.CellStyle.ForeColor = Color.FromArgb(100, 100, 100)
        End Select
    End Sub

    Private Sub MainGrid_KeyDown(sender As Object, e As KeyEventArgs)
        Select Case e.KeyCode
            Case Keys.Enter
                e.Handled = True
                e.SuppressKeyPress = True
                Dim selectedTag = GetSelectedTag()
                If selectedTag IsNot Nothing Then
                    Select Case selectedTag.Kind
                        Case RowKind.SavedJob, RowKind.QueueEntry
                            RunNowSelected()
                        Case RowKind.RunHistory
                            OpenJournalForSelected()
                    End Select
                End If

            Case Keys.Delete
                e.Handled = True
                e.SuppressKeyPress = True
                Dim selectedTag = GetSelectedTag()
                If selectedTag IsNot Nothing Then
                    Select Case selectedTag.Kind
                        Case RowKind.SavedJob
                            DeleteJobButton_Click(sender, EventArgs.Empty)
                        Case RowKind.QueueEntry
                            RemoveQueueButton_Click(sender, EventArgs.Empty)
                        Case RowKind.RunHistory
                            DeleteRunButton_Click(sender, EventArgs.Empty)
                    End Select
                End If

            Case Keys.F5
                e.Handled = True
                e.SuppressKeyPress = True
                RefreshData()

            Case Keys.F2
                e.Handled = True
                e.SuppressKeyPress = True
                RenameJobButton_Click(sender, EventArgs.Empty)
        End Select
    End Sub

    Private Sub AutoRefreshTimer_Tick(sender As Object, e As EventArgs)
        Dim hasActiveRuns = False
        For Each run In _runs
            If run IsNot Nothing AndAlso (run.Status = ManagedJobRunStatus.Running OrElse run.Status = ManagedJobRunStatus.Paused OrElse run.Status = ManagedJobRunStatus.Queued) Then
                hasActiveRuns = True
                Exit For
            End If
        Next

        If hasActiveRuns Then
            RefreshData()
        End If
    End Sub

    Private Sub RunNowButton_Click(sender As Object, e As EventArgs)
        RunNowSelected()
    End Sub

    Private Sub RunNowSelected()
        Dim selectedTag = GetSelectedTag()
        If selectedTag Is Nothing Then
            Return
        End If

        Select Case selectedTag.Kind
            Case RowKind.SavedJob
                _requestedAction = JobManagerRequestAction.RunSavedJob
                _requestedRunJobId = selectedTag.JobId
                _requestedQueueEntryId = String.Empty
                DialogResult = DialogResult.OK
                Close()

            Case RowKind.QueueEntry
                _requestedAction = JobManagerRequestAction.RunQueuedEntry
                _requestedQueueEntryId = selectedTag.PrimaryId
                _requestedRunJobId = selectedTag.JobId
                DialogResult = DialogResult.OK
                Close()

            Case Else
                MessageBox.Show(Me,
                                "Select a saved job or queue entry to run.",
                                "Job Manager",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Information)
        End Select
    End Sub

    Private Sub QueueButton_Click(sender As Object, e As EventArgs)
        Dim selectedTag = GetSelectedTag()
        If selectedTag Is Nothing Then
            Return
        End If

        Dim queued As Boolean
        Select Case selectedTag.Kind
            Case RowKind.SavedJob
                queued = _jobManager.QueueJob(selectedTag.JobId, trigger:="job-manager")
            Case RowKind.RunHistory
                If String.IsNullOrWhiteSpace(selectedTag.JobId) Then
                    MessageBox.Show(Me,
                                    "This run is not linked to a saved job, so it cannot be queued.",
                                    "Job Manager",
                                    MessageBoxButtons.OK,
                                    MessageBoxIcon.Information)
                    Return
                End If
                queued = _jobManager.QueueJob(selectedTag.JobId, trigger:="rerun-history")
            Case Else
                MessageBox.Show(Me,
                                "Queue action applies to saved jobs and saved-job run entries.",
                                "Job Manager",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Information)
                Return
        End Select

        If Not queued Then
            MessageBox.Show(Me,
                            "Unable to queue this item. It may already be queued.",
                            "Job Manager",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Information)
        End If

        RefreshData()
    End Sub

    Private Sub RemoveQueueButton_Click(sender As Object, e As EventArgs)
        Dim selectedTag = GetSelectedTag()
        If selectedTag Is Nothing OrElse selectedTag.Kind <> RowKind.QueueEntry Then
            MessageBox.Show(Me,
                            "Select a queue entry first.",
                            "Job Manager",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Information)
            Return
        End If

        _jobManager.RemoveQueuedJob(selectedTag.PrimaryId)
        RefreshData()
    End Sub

    Private Sub MoveQueueUpButton_Click(sender As Object, e As EventArgs)
        MoveQueueEntry(QueueMoveDirection.Up)
    End Sub

    Private Sub MoveQueueDownButton_Click(sender As Object, e As EventArgs)
        MoveQueueEntry(QueueMoveDirection.Down)
    End Sub

    Private Sub MoveQueueEntry(direction As QueueMoveDirection)
        Dim selectedTag = GetSelectedTag()
        If selectedTag Is Nothing OrElse selectedTag.Kind <> RowKind.QueueEntry Then
            MessageBox.Show(Me,
                            "Select a queue entry first.",
                            "Job Manager",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Information)
            Return
        End If

        _jobManager.MoveQueueEntry(selectedTag.PrimaryId, direction)
        RefreshData()
    End Sub

    Private Sub DeleteJobButton_Click(sender As Object, e As EventArgs)
        Dim selectedTag = GetSelectedTag()
        If selectedTag Is Nothing OrElse selectedTag.Kind <> RowKind.SavedJob Then
            MessageBox.Show(Me,
                            "Select a saved job first.",
                            "Job Manager",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Information)
            Return
        End If

        Dim response = MessageBox.Show(Me,
                                       "Delete selected saved job?",
                                       "Job Manager",
                                       MessageBoxButtons.YesNo,
                                       MessageBoxIcon.Warning)
        If response <> DialogResult.Yes Then
            Return
        End If

        _jobManager.DeleteJob(selectedTag.JobId)
        RefreshData()
    End Sub

    Private Sub DeleteRunButton_Click(sender As Object, e As EventArgs)
        Dim selectedTag = GetSelectedTag()
        If selectedTag Is Nothing OrElse selectedTag.Kind <> RowKind.RunHistory Then
            MessageBox.Show(Me,
                            "Select a run history row first.",
                            "Job Manager",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Information)
            Return
        End If

        Dim response = MessageBox.Show(Me,
                                       "Delete selected run history entry?",
                                       "Job Manager",
                                       MessageBoxButtons.YesNo,
                                       MessageBoxIcon.Warning)
        If response <> DialogResult.Yes Then
            Return
        End If

        _jobManager.DeleteRun(selectedTag.PrimaryId)
        RefreshData()
    End Sub

    Private Sub OpenJournalButton_Click(sender As Object, e As EventArgs)
        OpenJournalForSelected()
    End Sub

    Private Sub ClearQueueButton_Click(sender As Object, e As EventArgs)
        If _queueEntries.Count = 0 Then
            Return
        End If

        Dim response = MessageBox.Show(Me,
                                       "Clear all queued jobs?",
                                       "Job Manager",
                                       MessageBoxButtons.YesNo,
                                       MessageBoxIcon.Warning)
        If response <> DialogResult.Yes Then
            Return
        End If

        _jobManager.ClearQueue()
        RefreshData()
    End Sub

    Private Sub ClearHistoryButton_Click(sender As Object, e As EventArgs)
        If _runs.Count = 0 Then
            Return
        End If

        Dim response = MessageBox.Show(Me,
                                       "Clear all run history entries?",
                                       "Job Manager",
                                       MessageBoxButtons.YesNo,
                                       MessageBoxIcon.Warning)
        If response <> DialogResult.Yes Then
            Return
        End If

        _jobManager.ClearRunHistory(keepLatest:=0)
        RefreshData()
    End Sub

    Private Sub MoveQueueTopButton_Click(sender As Object, e As EventArgs)
        MoveQueueEntry(QueueMoveDirection.Top)
    End Sub

    Private Sub MoveQueueBottomButton_Click(sender As Object, e As EventArgs)
        MoveQueueEntry(QueueMoveDirection.Bottom)
    End Sub

    Private Sub RenameJobButton_Click(sender As Object, e As EventArgs)
        Dim selectedTag = GetSelectedTag()
        If selectedTag Is Nothing OrElse selectedTag.Kind <> RowKind.SavedJob Then
            Return
        End If

        Dim currentJob As ManagedJob = Nothing
        _savedJobsById.TryGetValue(selectedTag.JobId, currentJob)
        If currentJob Is Nothing Then
            Return
        End If

        Dim newName = PromptForInput("Rename Job", "Enter new name:", currentJob.Name)
        If newName Is Nothing OrElse newName.Trim() = currentJob.Name Then
            Return
        End If

        _jobManager.RenameJob(selectedTag.JobId, newName)
        RefreshData()
    End Sub

    Private Sub DuplicateJobButton_Click(sender As Object, e As EventArgs)
        Dim selectedTag = GetSelectedTag()
        If selectedTag Is Nothing OrElse selectedTag.Kind <> RowKind.SavedJob Then
            Return
        End If

        Dim currentJob As ManagedJob = Nothing
        _savedJobsById.TryGetValue(selectedTag.JobId, currentJob)
        If currentJob Is Nothing Then
            Return
        End If

        Dim suggestedName = currentJob.Name & " (Copy)"
        Dim newName = PromptForInput("Duplicate Job", "Enter name for the copy:", suggestedName)
        If newName Is Nothing Then
            Return
        End If

        Dim duplicated = _jobManager.DuplicateJob(selectedTag.JobId, newName)
        If duplicated Is Nothing Then
            MessageBox.Show(Me,
                            "Failed to duplicate job.",
                            "Job Manager",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Error)
            Return
        End If

        RefreshData()
    End Sub

    Private Sub OpenJournalForSelected()
        Dim selectedTag = GetSelectedTag()
        If selectedTag Is Nothing OrElse selectedTag.Kind <> RowKind.RunHistory Then
            Return
        End If

        Dim runEntry As ManagedJobRun = Nothing
        _runsById.TryGetValue(selectedTag.PrimaryId, runEntry)
        If runEntry Is Nothing OrElse String.IsNullOrWhiteSpace(runEntry.JournalPath) Then
            MessageBox.Show(Me,
                            "Selected run does not have a journal path.",
                            "Job Manager",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Information)
            Return
        End If

        Dim journalPath = runEntry.JournalPath
        If Not File.Exists(journalPath) Then
            MessageBox.Show(Me,
                            "Journal file no longer exists.",
                            "Job Manager",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Information)
            Return
        End If

        Try
            Process.Start(New ProcessStartInfo() With {
                .FileName = journalPath,
                .UseShellExecute = True
            })
        Catch ex As Exception
            MessageBox.Show(Me,
                            ex.Message,
                            "Job Manager",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub ConfigureContextMenu()
        _gridContextMenu.ShowImageMargin = False
        AddHandler _gridContextMenu.Opening, AddressOf GridContextMenu_Opening
        _mainGrid.ContextMenuStrip = _gridContextMenu
    End Sub

    Private Sub GridContextMenu_Opening(sender As Object, e As System.ComponentModel.CancelEventArgs)
        _gridContextMenu.Items.Clear()

        Dim hitTest = _mainGrid.HitTest(
            _mainGrid.PointToClient(Cursor.Position).X,
            _mainGrid.PointToClient(Cursor.Position).Y)
        If hitTest.RowIndex < 0 Then
            e.Cancel = True
            Return
        End If

        _mainGrid.ClearSelection()
        _mainGrid.CurrentCell = _mainGrid.Rows(hitTest.RowIndex).Cells(0)
        _mainGrid.Rows(hitTest.RowIndex).Selected = True

        Dim selectedTag = GetSelectedTag()
        If selectedTag Is Nothing Then
            e.Cancel = True
            Return
        End If

        Select Case selectedTag.Kind
            Case RowKind.SavedJob
                _gridContextMenu.Items.Add("Run Now", Nothing, AddressOf RunNowButton_Click)
                _gridContextMenu.Items.Add("Queue", Nothing, AddressOf QueueButton_Click)
                _gridContextMenu.Items.Add(New ToolStripSeparator())
                _gridContextMenu.Items.Add("Rename", Nothing, AddressOf RenameJobButton_Click)
                _gridContextMenu.Items.Add("Duplicate", Nothing, AddressOf DuplicateJobButton_Click)
                _gridContextMenu.Items.Add(New ToolStripSeparator())
                _gridContextMenu.Items.Add("Delete Job", Nothing, AddressOf DeleteJobButton_Click)

            Case RowKind.QueueEntry
                _gridContextMenu.Items.Add("Run Now", Nothing, AddressOf RunNowButton_Click)
                _gridContextMenu.Items.Add(New ToolStripSeparator())
                _gridContextMenu.Items.Add("Move to Top", Nothing, AddressOf MoveQueueTopButton_Click)
                _gridContextMenu.Items.Add("Move Up", Nothing, AddressOf MoveQueueUpButton_Click)
                _gridContextMenu.Items.Add("Move Down", Nothing, AddressOf MoveQueueDownButton_Click)
                _gridContextMenu.Items.Add("Move to Bottom", Nothing, AddressOf MoveQueueBottomButton_Click)
                _gridContextMenu.Items.Add(New ToolStripSeparator())
                _gridContextMenu.Items.Add("Remove from Queue", Nothing, AddressOf RemoveQueueButton_Click)

            Case RowKind.RunHistory
                _gridContextMenu.Items.Add("Open Journal", Nothing, AddressOf OpenJournalButton_Click)
                If Not String.IsNullOrWhiteSpace(selectedTag.JobId) Then
                    _gridContextMenu.Items.Add("Re-queue Job", Nothing, AddressOf QueueButton_Click)
                End If
                _gridContextMenu.Items.Add(New ToolStripSeparator())
                _gridContextMenu.Items.Add("Delete Run", Nothing, AddressOf DeleteRunButton_Click)
        End Select
    End Sub

    Private Function PromptForInput(title As String, prompt As String, defaultValue As String) As String
        Using dialog As New Form()
            dialog.Text = title
            dialog.StartPosition = FormStartPosition.CenterParent
            dialog.FormBorderStyle = FormBorderStyle.FixedDialog
            dialog.MinimizeBox = False
            dialog.MaximizeBox = False
            dialog.ShowInTaskbar = False
            dialog.Size = New Size(420, 160)
            WindowIconHelper.Apply(dialog)

            Dim promptLabel As New Label() With {
                .AutoSize = True,
                .Location = New Point(12, 16),
                .Text = prompt
            }

            Dim inputBox As New TextBox() With {
                .Location = New Point(12, 40),
                .Width = 380,
                .Text = If(defaultValue, String.Empty)
            }

            Dim okButton As New Button() With {
                .Text = "OK",
                .DialogResult = DialogResult.OK,
                .Location = New Point(236, 80),
                .Width = 75
            }

            Dim cancelButton As New Button() With {
                .Text = "Cancel",
                .DialogResult = DialogResult.Cancel,
                .Location = New Point(317, 80),
                .Width = 75
            }

            dialog.AcceptButton = okButton
            dialog.CancelButton = cancelButton
            dialog.Controls.AddRange({promptLabel, inputBox, okButton, cancelButton})

            inputBox.SelectAll()

            If dialog.ShowDialog(Me) = DialogResult.OK AndAlso Not String.IsNullOrWhiteSpace(inputBox.Text) Then
                Return inputBox.Text.Trim()
            End If

            Return Nothing
        End Using
    End Function

    Private Sub UpdateActionStates()
        Dim selectedTag = GetSelectedTag()
        Dim hasSelection = selectedTag IsNot Nothing

        Dim isSaved = hasSelection AndAlso selectedTag.Kind = RowKind.SavedJob
        Dim isQueue = hasSelection AndAlso selectedTag.Kind = RowKind.QueueEntry
        Dim isRun = hasSelection AndAlso selectedTag.Kind = RowKind.RunHistory
        Dim isRunWithSavedJob = isRun AndAlso Not String.IsNullOrWhiteSpace(selectedTag.JobId)

        _runNowButton.Enabled = isSaved OrElse isQueue
        _queueButton.Enabled = isSaved OrElse isRunWithSavedJob
        _removeQueueButton.Enabled = isQueue
        _moveQueueTopButton.Enabled = isQueue
        _moveQueueUpButton.Enabled = isQueue
        _moveQueueDownButton.Enabled = isQueue
        _moveQueueBottomButton.Enabled = isQueue
        _renameJobButton.Enabled = isSaved
        _duplicateJobButton.Enabled = isSaved
        _deleteJobButton.Enabled = isSaved
        _deleteRunButton.Enabled = isRun
        _openJournalButton.Enabled = isRun
        _clearQueueButton.Enabled = _queueEntries.Count > 0
        _clearHistoryButton.Enabled = _runs.Count > 0

        Dim view = GetSelectedViewFilter()
        Dim supportsRunFilter = (view = JobManagerViewFilter.AllItems OrElse view = JobManagerViewFilter.RunHistory)
        _statusLabel.Enabled = supportsRunFilter
        _statusComboBox.Enabled = supportsRunFilter
    End Sub

    Private Sub RenderSelectedItemDetails()
        Dim selectedTag = GetSelectedTag()
        If selectedTag Is Nothing Then
            _detailsTextBox.Text = "Select an item to view details."
            Return
        End If

        Select Case selectedTag.Kind
            Case RowKind.SavedJob
                Dim job As ManagedJob = Nothing
                _savedJobsById.TryGetValue(selectedTag.JobId, job)
                _detailsTextBox.Text = BuildSavedJobDetails(job)

            Case RowKind.QueueEntry
                Dim entry As JobQueueEntryView = Nothing
                _queueEntriesById.TryGetValue(selectedTag.PrimaryId, entry)
                _detailsTextBox.Text = BuildQueueEntryDetails(entry)

            Case RowKind.RunHistory
                Dim run As ManagedJobRun = Nothing
                _runsById.TryGetValue(selectedTag.PrimaryId, run)
                _detailsTextBox.Text = BuildRunDetails(run)

            Case Else
                _detailsTextBox.Text = String.Empty
        End Select
    End Sub

    Private Function BuildSavedJobDetails(job As ManagedJob) As String
        If job Is Nothing Then
            Return "Saved job details unavailable."
        End If

        Dim options = If(job.Options, New CopyJobOptions())
        Dim lines As New List(Of String)()
        lines.Add("Type: Saved Job")
        lines.Add($"Name: {job.Name}")
        lines.Add($"Job ID: {job.JobId}")
        lines.Add($"Created: {job.CreatedUtc.ToLocalTime():yyyy-MM-dd HH:mm:ss}")
        lines.Add($"Updated: {job.UpdatedUtc.ToLocalTime():yyyy-MM-dd HH:mm:ss}")
        lines.Add($"Source: {options.SourceRoot}")
        lines.Add($"Destination: {options.DestinationRoot}")
        lines.Add($"Overwrite: {options.OverwritePolicy}")
        lines.Add($"Verify: {options.VerifyAfterCopy}")
        lines.Add($"Adaptive Buffer: {options.UseAdaptiveBufferSizing}")
        lines.Add($"Manual Buffer MB: {Math.Max(1, CInt(Math.Round(options.BufferSizeBytes / 1024.0R / 1024.0R)))}")
        lines.Add($"Retries: {Math.Max(0, options.MaxRetries)}")
        lines.Add($"Timeout: {Math.Max(0, CInt(options.OperationTimeout.TotalSeconds))} sec")
        Return String.Join(Environment.NewLine, lines)
    End Function

    Private Function BuildQueueEntryDetails(entry As JobQueueEntryView) As String
        If entry Is Nothing Then
            Return "Queue entry details unavailable."
        End If

        Dim lines As New List(Of String)()
        lines.Add("Type: Queue Entry")
        lines.Add($"Queue Entry ID: {entry.QueueEntryId}")
        lines.Add($"Position: {entry.Position}")
        lines.Add($"Job: {entry.JobName}")
        lines.Add($"Job ID: {entry.JobId}")
        lines.Add($"Trigger: {entry.Trigger}")
        lines.Add($"Enqueued: {entry.EnqueuedUtc.ToLocalTime():yyyy-MM-dd HH:mm:ss}")
        lines.Add($"Updated: {entry.LastUpdatedUtc.ToLocalTime():yyyy-MM-dd HH:mm:ss}")
        lines.Add($"Attempts: {entry.AttemptCount}")
        lines.Add($"Source: {entry.SourceRoot}")
        lines.Add($"Destination: {entry.DestinationRoot}")
        If Not String.IsNullOrWhiteSpace(entry.LastErrorMessage) Then
            lines.Add($"Last Error: {entry.LastErrorMessage}")
        End If
        Return String.Join(Environment.NewLine, lines)
    End Function

    Private Function BuildRunDetails(run As ManagedJobRun) As String
        If run Is Nothing Then
            Return "Run details unavailable."
        End If

        Dim lines As New List(Of String)()
        lines.Add("Type: Run")
        lines.Add($"Run ID: {run.RunId}")
        lines.Add($"Display Name: {run.DisplayName}")
        lines.Add($"Status: {run.Status}")
        lines.Add($"Trigger: {run.Trigger}")
        lines.Add($"Started: {run.StartedUtc.ToLocalTime():yyyy-MM-dd HH:mm:ss}")
        lines.Add($"Updated: {run.LastUpdatedUtc.ToLocalTime():yyyy-MM-dd HH:mm:ss}")
        If run.FinishedUtc.HasValue Then
            lines.Add($"Finished: {run.FinishedUtc.Value.ToLocalTime():yyyy-MM-dd HH:mm:ss}")
        End If
        lines.Add($"Source: {run.SourceRoot}")
        lines.Add($"Destination: {run.DestinationRoot}")
        lines.Add($"Queue Entry ID: {If(String.IsNullOrWhiteSpace(run.QueueEntryId), "-", run.QueueEntryId)}")
        lines.Add($"Queue Attempt: {run.QueueAttempt}")
        lines.Add($"Journal: {If(String.IsNullOrWhiteSpace(run.JournalPath), "-", run.JournalPath)}")
        lines.Add($"Summary: {If(String.IsNullOrWhiteSpace(run.Summary), "-", run.Summary)}")
        If Not String.IsNullOrWhiteSpace(run.ErrorMessage) Then
            lines.Add($"Error: {run.ErrorMessage}")
        End If
        Return String.Join(Environment.NewLine, lines)
    End Function

    Private Function GetSelectedTag() As GridRowTag
        Dim rowIndex = GetSelectedRowIndex()
        If rowIndex < 0 OrElse rowIndex >= _visibleRows.Count Then
            Return Nothing
        End If

        Return _visibleRows(rowIndex).Tag
    End Function

    Private Function GetSelectedRowIndex() As Integer
        If _mainGrid.CurrentCell IsNot Nothing Then
            Return _mainGrid.CurrentCell.RowIndex
        End If

        If _mainGrid.SelectedCells.Count > 0 Then
            Return _mainGrid.SelectedCells(0).RowIndex
        End If

        If _mainGrid.SelectedRows.Count > 0 Then
            Return _mainGrid.SelectedRows(0).Index
        End If

        Return -1
    End Function

    Private Shared Function BuildDictionary(Of TItem)(
        items As IEnumerable(Of TItem),
        keySelector As Func(Of TItem, String)) As IReadOnlyDictionary(Of String, TItem)

        Dim dictionary As New Dictionary(Of String, TItem)(StringComparer.OrdinalIgnoreCase)
        If items Is Nothing OrElse keySelector Is Nothing Then
            Return dictionary
        End If

        For Each item In items
            If item Is Nothing Then
                Continue For
            End If

            Dim key = keySelector(item)
            If String.IsNullOrWhiteSpace(key) Then
                Continue For
            End If

            dictionary(key) = item
        Next

        Return dictionary
    End Function

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

    Private Shared Sub SuspendDrawing(target As Control)
        If target Is Nothing OrElse Not target.IsHandleCreated Then
            Return
        End If

        SendMessage(target.Handle, WmSetRedraw, IntPtr.Zero, IntPtr.Zero)
    End Sub

    Private Shared Sub ResumeDrawing(target As Control)
        If target Is Nothing OrElse Not target.IsHandleCreated Then
            Return
        End If

        SendMessage(target.Handle, WmSetRedraw, New IntPtr(1), IntPtr.Zero)
        target.Invalidate(invalidateChildren:=True)
        target.Update()
    End Sub

    Private Function GetSelectedViewFilter() As JobManagerViewFilter
        Dim selected = TryCast(_viewComboBox.SelectedItem, EnumChoice(Of JobManagerViewFilter))
        If selected Is Nothing Then
            Return JobManagerViewFilter.AllItems
        End If

        Return selected.Value
    End Function

    Private Function GetSelectedRunStatusFilter() As ManagedJobRunStatus?
        Dim selected = TryCast(_statusComboBox.SelectedItem, EnumChoice(Of RunStatusFilter))
        If selected Is Nothing OrElse selected.Value = RunStatusFilter.Any Then
            Return Nothing
        End If

        Select Case selected.Value
            Case RunStatusFilter.Queued
                Return ManagedJobRunStatus.Queued
            Case RunStatusFilter.Running
                Return ManagedJobRunStatus.Running
            Case RunStatusFilter.Paused
                Return ManagedJobRunStatus.Paused
            Case RunStatusFilter.Completed
                Return ManagedJobRunStatus.Completed
            Case RunStatusFilter.Failed
                Return ManagedJobRunStatus.Failed
            Case RunStatusFilter.Cancelled
                Return ManagedJobRunStatus.Cancelled
            Case RunStatusFilter.Interrupted
                Return ManagedJobRunStatus.Interrupted
            Case Else
                Return Nothing
        End Select
    End Function

    ''' <summary>
    ''' Enum JobManagerViewFilter.
    ''' </summary>
    Private Enum JobManagerViewFilter
        AllItems = 0
        SavedJobs = 1
        Queue = 2
        RunHistory = 3
    End Enum

    ''' <summary>
    ''' Enum RunStatusFilter.
    ''' </summary>
    Private Enum RunStatusFilter
        Any = 0
        Queued = 1
        Running = 2
        Paused = 3
        Completed = 4
        Failed = 5
        Cancelled = 6
        Interrupted = 7
    End Enum

    ''' <summary>
    ''' Enum RowKind.
    ''' </summary>
    Private Enum RowKind
        SavedJob = 0
        QueueEntry = 1
        RunHistory = 2
    End Enum

    ''' <summary>
    ''' Class GridRowModel.
    ''' </summary>
    Private NotInheritable Class GridRowModel
        ''' <summary>
        ''' Initializes a new instance.
        ''' </summary>
        Public Sub New(
            typeText As String,
            nameText As String,
            stateText As String,
            queueText As String,
            triggerText As String,
            startedText As String,
            updatedText As String,
            sourceText As String,
            destinationText As String,
            summaryText As String,
            tag As GridRowTag)

            Me.TypeText = If(typeText, String.Empty)
            Me.NameText = If(nameText, String.Empty)
            Me.StateText = If(stateText, String.Empty)
            Me.QueueText = If(queueText, String.Empty)
            Me.TriggerText = If(triggerText, String.Empty)
            Me.StartedText = If(startedText, String.Empty)
            Me.UpdatedText = If(updatedText, String.Empty)
            Me.SourceText = If(sourceText, String.Empty)
            Me.DestinationText = If(destinationText, String.Empty)
            Me.SummaryText = If(summaryText, String.Empty)
            Me.Tag = tag
        End Sub

        ''' <summary>
        ''' Gets or sets TypeText.
        ''' </summary>
        Public ReadOnly Property TypeText As String
        ''' <summary>
        ''' Gets or sets NameText.
        ''' </summary>
        Public ReadOnly Property NameText As String
        ''' <summary>
        ''' Gets or sets StateText.
        ''' </summary>
        Public ReadOnly Property StateText As String
        ''' <summary>
        ''' Gets or sets QueueText.
        ''' </summary>
        Public ReadOnly Property QueueText As String
        ''' <summary>
        ''' Gets or sets TriggerText.
        ''' </summary>
        Public ReadOnly Property TriggerText As String
        ''' <summary>
        ''' Gets or sets StartedText.
        ''' </summary>
        Public ReadOnly Property StartedText As String
        ''' <summary>
        ''' Gets or sets UpdatedText.
        ''' </summary>
        Public ReadOnly Property UpdatedText As String
        ''' <summary>
        ''' Gets or sets SourceText.
        ''' </summary>
        Public ReadOnly Property SourceText As String
        ''' <summary>
        ''' Gets or sets DestinationText.
        ''' </summary>
        Public ReadOnly Property DestinationText As String
        ''' <summary>
        ''' Gets or sets SummaryText.
        ''' </summary>
        Public ReadOnly Property SummaryText As String
        ''' <summary>
        ''' Gets or sets Tag.
        ''' </summary>
        Public ReadOnly Property Tag As GridRowTag
    End Class

    ''' <summary>
    ''' Class GridRowTag.
    ''' </summary>
    Private NotInheritable Class GridRowTag
        ''' <summary>
        ''' Initializes a new instance.
        ''' </summary>
        Public Sub New(kind As RowKind, primaryId As String, jobId As String)
            Me.Kind = kind
            Me.PrimaryId = If(primaryId, String.Empty)
            Me.JobId = If(jobId, String.Empty)
        End Sub

        ''' <summary>
        ''' Gets or sets Kind.
        ''' </summary>
        Public ReadOnly Property Kind As RowKind
        ''' <summary>
        ''' Gets or sets PrimaryId.
        ''' </summary>
        Public ReadOnly Property PrimaryId As String
        ''' <summary>
        ''' Gets or sets JobId.
        ''' </summary>
        Public ReadOnly Property JobId As String
    End Class

    ''' <summary>
    ''' Class EnumChoice.
    ''' </summary>
    Private NotInheritable Class EnumChoice(Of T)
        ''' <summary>
        ''' Initializes a new instance.
        ''' </summary>
        Public Sub New(value As T, caption As String)
            Me.Value = value
            Me.Caption = If(caption, String.Empty)
        End Sub

        ''' <summary>
        ''' Gets or sets Value.
        ''' </summary>
        Public ReadOnly Property Value As T
        ''' <summary>
        ''' Gets or sets Caption.
        ''' </summary>
        Public ReadOnly Property Caption As String

        ''' <summary>
        ''' Computes ToString.
        ''' </summary>
        Public Overrides Function ToString() As String
            Return Caption
        End Function
    End Class
End Class

