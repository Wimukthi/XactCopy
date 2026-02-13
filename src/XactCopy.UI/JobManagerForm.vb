Imports System.Drawing
Imports XactCopy.Models
Imports XactCopy.Services

Friend Class JobManagerForm
    Inherits Form

    Private Const LayoutWindowKey As String = "JobManagerForm"
    Private Const GridLayoutSavedKey As String = "JobManagerForm.SingleGrid.SavedJobs"
    Private Const GridLayoutQueueKey As String = "JobManagerForm.SingleGrid.Queue"
    Private Const GridLayoutHistoryKey As String = "JobManagerForm.SingleGrid.RunHistory"

    Private ReadOnly _jobManager As JobManagerService

    Private ReadOnly _summaryLabel As New Label()
    Private ReadOnly _hintLabel As New Label()

    Private ReadOnly _viewLabel As New Label()
    Private ReadOnly _viewComboBox As New ComboBox()
    Private ReadOnly _runSelectedButton As New Button()
    Private ReadOnly _queueSelectedButton As New Button()
    Private ReadOnly _removeFromQueueButton As New Button()
    Private ReadOnly _deleteSelectedButton As New Button()
    Private ReadOnly _refreshButton As New Button()
    Private ReadOnly _closeButton As New Button()

    Private ReadOnly _mainGrid As New DataGridView()
    Private ReadOnly _toolTip As New ToolTip() With {
        .ShowAlways = True,
        .InitialDelay = 350,
        .ReshowDelay = 120,
        .AutoPopDelay = 24000
    }

    Private _requestedRunJobId As String = String.Empty
    Private _currentView As JobManagerView = JobManagerView.SavedJobs

    Public Sub New(jobManager As JobManagerService)
        If jobManager Is Nothing Then
            Throw New ArgumentNullException(NameOf(jobManager))
        End If

        _jobManager = jobManager

        Text = "Job Manager"
        StartPosition = FormStartPosition.CenterParent
        MinimumSize = New Size(980, 640)
        Size = New Size(1180, 760)
        WindowIconHelper.Apply(Me)

        BuildUi()
        ConfigureToolTips()

        AddHandler Shown, AddressOf JobManagerForm_Shown
        AddHandler FormClosing, AddressOf JobManagerForm_FormClosing
    End Sub

    Public ReadOnly Property RequestedRunJobId As String
        Get
            Return _requestedRunJobId
        End Get
    End Property

    Private Sub BuildUi()
        ConfigureGrid()

        Dim rootLayout As New TableLayoutPanel() With {
            .Dock = DockStyle.Fill,
            .ColumnCount = 1,
            .RowCount = 4,
            .Padding = New Padding(10)
        }
        rootLayout.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0F))
        rootLayout.RowStyles.Add(New RowStyle(SizeType.AutoSize))
        rootLayout.RowStyles.Add(New RowStyle(SizeType.AutoSize))
        rootLayout.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0F))
        rootLayout.RowStyles.Add(New RowStyle(SizeType.AutoSize))

        rootLayout.Controls.Add(BuildHeaderPane(), 0, 0)
        rootLayout.Controls.Add(BuildCommandPane(), 0, 1)
        rootLayout.Controls.Add(_mainGrid, 0, 2)
        rootLayout.Controls.Add(BuildFooterPane(), 0, 3)

        Controls.Add(rootLayout)
    End Sub

    Private Sub ConfigureToolTips()
        _toolTip.SetToolTip(_viewComboBox, "Choose which dataset to display in the grid: Saved Jobs, Queue, or Run History.")
        _toolTip.SetToolTip(_runSelectedButton, "Run the selected saved job immediately.")
        _toolTip.SetToolTip(_queueSelectedButton, "Queue the selected saved job for later execution.")
        _toolTip.SetToolTip(_removeFromQueueButton, "Remove the selected queue entry.")
        _toolTip.SetToolTip(_deleteSelectedButton, "Delete the selected saved job definition.")
        _toolTip.SetToolTip(_refreshButton, "Reload jobs, queue, and run history.")
        _toolTip.SetToolTip(_mainGrid, "Main list view. Double-click a saved job row to run it.")
        _toolTip.SetToolTip(_summaryLabel, "Counts of saved jobs, queued jobs, and recorded runs.")
        _toolTip.SetToolTip(_hintLabel, "Quick usage hint for this dialog.")
        _toolTip.SetToolTip(_closeButton, "Close Job Manager.")
    End Sub

    Private Function BuildHeaderPane() As Control
        Dim layout As New TableLayoutPanel() With {
            .Dock = DockStyle.Fill,
            .ColumnCount = 2,
            .RowCount = 1,
            .Margin = New Padding(0, 0, 0, 8)
        }
        layout.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0F))
        layout.ColumnStyles.Add(New ColumnStyle(SizeType.AutoSize))

        Dim titleLayout As New TableLayoutPanel() With {
            .Dock = DockStyle.Fill,
            .ColumnCount = 1,
            .RowCount = 2
        }
        titleLayout.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0F))
        titleLayout.RowStyles.Add(New RowStyle(SizeType.AutoSize))
        titleLayout.RowStyles.Add(New RowStyle(SizeType.AutoSize))

        Dim titleLabel As New Label() With {
            .AutoSize = True,
            .Font = New Font(Font, FontStyle.Bold),
            .Text = "Job Manager",
            .Margin = New Padding(0, 0, 0, 2)
        }

        Dim subtitleLabel As New Label() With {
            .AutoSize = True,
            .Text = "One-grid view for saved jobs, queue, and run history.",
            .Margin = New Padding(0)
        }

        _summaryLabel.AutoSize = True
        _summaryLabel.Anchor = AnchorStyles.Right Or AnchorStyles.Top
        _summaryLabel.Margin = New Padding(10, 2, 0, 0)
        _summaryLabel.Text = "Saved: 0 | Queued: 0 | Runs: 0"

        titleLayout.Controls.Add(titleLabel, 0, 0)
        titleLayout.Controls.Add(subtitleLabel, 0, 1)

        layout.Controls.Add(titleLayout, 0, 0)
        layout.Controls.Add(_summaryLabel, 1, 0)

        Return layout
    End Function

    Private Function BuildCommandPane() As Control
        Dim pane As New FlowLayoutPanel() With {
            .Dock = DockStyle.Fill,
            .FlowDirection = FlowDirection.LeftToRight,
            .WrapContents = True,
            .AutoSize = True,
            .Margin = New Padding(0, 0, 0, 8)
        }

        _viewLabel.Text = "View:"
        _viewLabel.AutoSize = True
        _viewLabel.Margin = New Padding(0, 8, 6, 0)

        _viewComboBox.DropDownStyle = ComboBoxStyle.DropDownList
        _viewComboBox.Width = 200
        _viewComboBox.Items.Add(New ViewItem(JobManagerView.SavedJobs, "Saved Jobs"))
        _viewComboBox.Items.Add(New ViewItem(JobManagerView.Queue, "Queue"))
        _viewComboBox.Items.Add(New ViewItem(JobManagerView.RunHistory, "Run History"))
        AddHandler _viewComboBox.SelectedIndexChanged, AddressOf ViewComboBox_SelectedIndexChanged

        _runSelectedButton.Text = "Run Selected"
        _runSelectedButton.AutoSize = True
        AddHandler _runSelectedButton.Click, AddressOf RunSelectedButton_Click

        _queueSelectedButton.Text = "Queue Selected"
        _queueSelectedButton.AutoSize = True
        AddHandler _queueSelectedButton.Click, AddressOf QueueSelectedButton_Click

        _removeFromQueueButton.Text = "Remove From Queue"
        _removeFromQueueButton.AutoSize = True
        AddHandler _removeFromQueueButton.Click, AddressOf RemoveFromQueueButton_Click

        _deleteSelectedButton.Text = "Delete Selected Job"
        _deleteSelectedButton.AutoSize = True
        AddHandler _deleteSelectedButton.Click, AddressOf DeleteSelectedButton_Click

        _refreshButton.Text = "Refresh"
        _refreshButton.AutoSize = True
        AddHandler _refreshButton.Click, AddressOf RefreshButton_Click

        pane.Controls.Add(_viewLabel)
        pane.Controls.Add(_viewComboBox)
        pane.Controls.Add(_runSelectedButton)
        pane.Controls.Add(_queueSelectedButton)
        pane.Controls.Add(_removeFromQueueButton)
        pane.Controls.Add(_deleteSelectedButton)
        pane.Controls.Add(_refreshButton)

        Return pane
    End Function

    Private Function BuildFooterPane() As Control
        Dim footer As New TableLayoutPanel() With {
            .Dock = DockStyle.Fill,
            .ColumnCount = 2,
            .RowCount = 1,
            .Margin = New Padding(0, 8, 0, 0)
        }
        footer.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0F))
        footer.ColumnStyles.Add(New ColumnStyle(SizeType.AutoSize))

        _hintLabel.AutoSize = True
        _hintLabel.Text = "Tip: double-click a saved job row to run immediately."
        _hintLabel.Anchor = AnchorStyles.Left Or AnchorStyles.Top

        _closeButton.Text = "Close"
        _closeButton.AutoSize = True
        AddHandler _closeButton.Click,
            Sub(sender, e)
                DialogResult = DialogResult.Cancel
                Close()
            End Sub

        footer.Controls.Add(_hintLabel, 0, 0)
        footer.Controls.Add(_closeButton, 1, 0)

        Return footer
    End Function

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
        _mainGrid.RowHeadersVisible = False
        _mainGrid.DefaultCellStyle.WrapMode = DataGridViewTriState.False
        _mainGrid.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        _mainGrid.ColumnHeadersHeight = 28
        _mainGrid.RowTemplate.Height = 24
        _mainGrid.ScrollBars = ScrollBars.Both

        AddHandler _mainGrid.SelectionChanged, AddressOf MainGrid_SelectionChanged
        AddHandler _mainGrid.CellDoubleClick, AddressOf MainGrid_CellDoubleClick
    End Sub

    Private Sub JobManagerForm_Shown(sender As Object, e As EventArgs)
        UiLayoutManager.ApplyWindow(Me, LayoutWindowKey)

        _viewComboBox.SelectedIndex = 0
        RefreshData()
    End Sub

    Private Sub JobManagerForm_FormClosing(sender As Object, e As FormClosingEventArgs)
        CaptureGridLayoutForCurrentView()
        UiLayoutManager.CaptureWindow(Me, LayoutWindowKey)
    End Sub

    Private Sub ViewComboBox_SelectedIndexChanged(sender As Object, e As EventArgs)
        Dim selectedView = GetSelectedView()
        If selectedView = _currentView AndAlso _mainGrid.Columns.Count > 0 Then
            Return
        End If

        SwitchView(selectedView)
        RefreshData()
    End Sub

    Private Sub RefreshButton_Click(sender As Object, e As EventArgs)
        RefreshData()
    End Sub

    Private Sub RefreshData()
        Dim savedJobs = _jobManager.GetJobs()
        Dim queuedJobs = _jobManager.GetQueuedJobs()
        Dim runs = _jobManager.GetRecentRuns(250)

        _summaryLabel.Text = $"Saved: {savedJobs.Count} | Queued: {queuedJobs.Count} | Runs: {runs.Count}"

        Select Case _currentView
            Case JobManagerView.SavedJobs
                PopulateSavedJobs(savedJobs)
            Case JobManagerView.Queue
                PopulateQueue(queuedJobs)
            Case Else
                PopulateRunHistory(runs)
        End Select

        UpdateActionStates()
    End Sub

    Private Sub SwitchView(view As JobManagerView)
        CaptureGridLayoutForCurrentView()
        _currentView = view

        ConfigureColumnsForCurrentView()
        UiLayoutManager.ApplyGridLayout(_mainGrid, GetLayoutKeyForCurrentView())
        UpdateActionStates()
    End Sub

    Private Sub ConfigureColumnsForCurrentView()
        _mainGrid.Columns.Clear()
        _mainGrid.Rows.Clear()

        Select Case _currentView
            Case JobManagerView.SavedJobs
                Dim nameColumn = _mainGrid.Columns.Add("Name", "Name")
                Dim sourceColumn = _mainGrid.Columns.Add("Source", "Source")
                Dim destinationColumn = _mainGrid.Columns.Add("Destination", "Destination")
                Dim updatedColumn = _mainGrid.Columns.Add("Updated", "Updated")

                _mainGrid.Columns(nameColumn).Width = 240
                _mainGrid.Columns(sourceColumn).Width = 380
                _mainGrid.Columns(destinationColumn).Width = 380
                _mainGrid.Columns(updatedColumn).Width = 170

            Case JobManagerView.Queue
                Dim positionColumn = _mainGrid.Columns.Add("Position", "#")
                Dim nameColumn = _mainGrid.Columns.Add("Name", "Job")
                Dim sourceColumn = _mainGrid.Columns.Add("Source", "Source")
                Dim destinationColumn = _mainGrid.Columns.Add("Destination", "Destination")

                _mainGrid.Columns(positionColumn).Width = 50
                _mainGrid.Columns(nameColumn).Width = 260
                _mainGrid.Columns(sourceColumn).Width = 410
                _mainGrid.Columns(destinationColumn).Width = 410

            Case Else
                Dim startedColumn = _mainGrid.Columns.Add("Started", "Started")
                Dim runColumn = _mainGrid.Columns.Add("DisplayName", "Run")
                Dim statusColumn = _mainGrid.Columns.Add("Status", "Status")
                Dim sourceColumn = _mainGrid.Columns.Add("Source", "Source")
                Dim destinationColumn = _mainGrid.Columns.Add("Destination", "Destination")
                Dim summaryColumn = _mainGrid.Columns.Add("Summary", "Summary")

                _mainGrid.Columns(startedColumn).Width = 160
                _mainGrid.Columns(runColumn).Width = 240
                _mainGrid.Columns(statusColumn).Width = 120
                _mainGrid.Columns(sourceColumn).Width = 330
                _mainGrid.Columns(destinationColumn).Width = 330
                _mainGrid.Columns(summaryColumn).Width = 360
        End Select
    End Sub

    Private Sub PopulateSavedJobs(jobs As IReadOnlyList(Of ManagedJob))
        _mainGrid.Rows.Clear()

        For Each job In jobs
            Dim rowIndex = _mainGrid.Rows.Add(
                job.Name,
                job.Options.SourceRoot,
                job.Options.DestinationRoot,
                job.UpdatedUtc.ToLocalTime().ToString("yyyy-MM-dd HH:mm:ss"))

            _mainGrid.Rows(rowIndex).Tag = New GridRowTag(RowKind.SavedJob, job.JobId)
        Next
    End Sub

    Private Sub PopulateQueue(queue As IReadOnlyList(Of ManagedJob))
        _mainGrid.Rows.Clear()

        Dim position = 1
        For Each job In queue
            Dim rowIndex = _mainGrid.Rows.Add(
                position.ToString(),
                job.Name,
                job.Options.SourceRoot,
                job.Options.DestinationRoot)

            _mainGrid.Rows(rowIndex).Tag = New GridRowTag(RowKind.QueuedJob, job.JobId)
            position += 1
        Next
    End Sub

    Private Sub PopulateRunHistory(runs As IReadOnlyList(Of ManagedJobRun))
        _mainGrid.Rows.Clear()

        For Each run In runs
            Dim summaryText = If(String.IsNullOrWhiteSpace(run.Summary), "-", run.Summary)
            Dim rowIndex = _mainGrid.Rows.Add(
                run.StartedUtc.ToLocalTime().ToString("yyyy-MM-dd HH:mm:ss"),
                run.DisplayName,
                run.Status.ToString(),
                run.SourceRoot,
                run.DestinationRoot,
                summaryText)

            _mainGrid.Rows(rowIndex).Tag = New GridRowTag(RowKind.RunHistoryEntry, run.RunId)
        Next
    End Sub

    Private Sub RunSelectedButton_Click(sender As Object, e As EventArgs)
        RunSelectedSavedJob()
    End Sub

    Private Sub QueueSelectedButton_Click(sender As Object, e As EventArgs)
        Dim selectedTag = GetSelectedTag()
        If selectedTag Is Nothing OrElse selectedTag.Kind <> RowKind.SavedJob Then
            MessageBox.Show(Me, "Select a saved job row first.", "Job Manager", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return
        End If

        If Not _jobManager.QueueJob(selectedTag.Id) Then
            MessageBox.Show(Me,
                            "Unable to queue that job. It may already be queued.",
                            "Job Manager",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Information)
            Return
        End If

        RefreshData()
    End Sub

    Private Sub RemoveFromQueueButton_Click(sender As Object, e As EventArgs)
        Dim selectedTag = GetSelectedTag()
        If selectedTag Is Nothing OrElse selectedTag.Kind <> RowKind.QueuedJob Then
            MessageBox.Show(Me, "Select a queued job row first.", "Job Manager", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return
        End If

        _jobManager.RemoveQueuedJob(selectedTag.Id)
        RefreshData()
    End Sub

    Private Sub DeleteSelectedButton_Click(sender As Object, e As EventArgs)
        Dim selectedTag = GetSelectedTag()
        If selectedTag Is Nothing OrElse selectedTag.Kind <> RowKind.SavedJob Then
            MessageBox.Show(Me, "Select a saved job row first.", "Job Manager", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return
        End If

        Dim response = MessageBox.Show(Me,
                                       "Delete selected job?",
                                       "Job Manager",
                                       MessageBoxButtons.YesNo,
                                       MessageBoxIcon.Warning)
        If response <> DialogResult.Yes Then
            Return
        End If

        _jobManager.DeleteJob(selectedTag.Id)
        RefreshData()
    End Sub

    Private Sub MainGrid_SelectionChanged(sender As Object, e As EventArgs)
        UpdateActionStates()
    End Sub

    Private Sub MainGrid_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs)
        If e.RowIndex < 0 Then
            Return
        End If

        If _currentView = JobManagerView.SavedJobs Then
            RunSelectedSavedJob()
        End If
    End Sub

    Private Sub RunSelectedSavedJob()
        Dim selectedTag = GetSelectedTag()
        If selectedTag Is Nothing OrElse selectedTag.Kind <> RowKind.SavedJob Then
            MessageBox.Show(Me, "Select a saved job row first.", "Job Manager", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return
        End If

        _requestedRunJobId = selectedTag.Id
        DialogResult = DialogResult.OK
        Close()
    End Sub

    Private Sub UpdateActionStates()
        Dim hasSelection = _mainGrid.SelectedRows.Count > 0
        Dim inSavedView = _currentView = JobManagerView.SavedJobs
        Dim inQueueView = _currentView = JobManagerView.Queue

        _runSelectedButton.Visible = inSavedView
        _queueSelectedButton.Visible = inSavedView
        _deleteSelectedButton.Visible = inSavedView
        _removeFromQueueButton.Visible = inQueueView

        _runSelectedButton.Enabled = inSavedView AndAlso hasSelection
        _queueSelectedButton.Enabled = inSavedView AndAlso hasSelection
        _deleteSelectedButton.Enabled = inSavedView AndAlso hasSelection
        _removeFromQueueButton.Enabled = inQueueView AndAlso hasSelection

        Select Case _currentView
            Case JobManagerView.SavedJobs
                _hintLabel.Text = "Tip: double-click a saved job row to run immediately."
            Case JobManagerView.Queue
                _hintLabel.Text = "Queue order is top-to-bottom and runs in that sequence."
            Case Else
                _hintLabel.Text = "Run history is read-only."
        End Select
    End Sub

    Private Function GetSelectedTag() As GridRowTag
        If _mainGrid.SelectedRows.Count = 0 Then
            Return Nothing
        End If

        Return TryCast(_mainGrid.SelectedRows(0).Tag, GridRowTag)
    End Function

    Private Function GetSelectedView() As JobManagerView
        Dim selectedItem = TryCast(_viewComboBox.SelectedItem, ViewItem)
        If selectedItem Is Nothing Then
            Return JobManagerView.SavedJobs
        End If

        Return selectedItem.View
    End Function

    Private Function GetLayoutKeyForCurrentView() As String
        Select Case _currentView
            Case JobManagerView.Queue
                Return GridLayoutQueueKey
            Case JobManagerView.RunHistory
                Return GridLayoutHistoryKey
            Case Else
                Return GridLayoutSavedKey
        End Select
    End Function

    Private Sub CaptureGridLayoutForCurrentView()
        If _mainGrid.Columns.Count = 0 Then
            Return
        End If

        UiLayoutManager.CaptureGridLayout(_mainGrid, GetLayoutKeyForCurrentView())
    End Sub

    Private Enum JobManagerView
        SavedJobs = 0
        Queue = 1
        RunHistory = 2
    End Enum

    Private Enum RowKind
        SavedJob = 0
        QueuedJob = 1
        RunHistoryEntry = 2
    End Enum

    Private NotInheritable Class GridRowTag
        Public Sub New(kind As RowKind, id As String)
            Me.Kind = kind
            Me.Id = If(id, String.Empty)
        End Sub

        Public ReadOnly Property Kind As RowKind
        Public ReadOnly Property Id As String
    End Class

    Private NotInheritable Class ViewItem
        Public Sub New(view As JobManagerView, caption As String)
            Me.View = view
            Me.Caption = If(caption, String.Empty)
        End Sub

        Public ReadOnly Property View As JobManagerView
        Public ReadOnly Property Caption As String

        Public Overrides Function ToString() As String
            Return Caption
        End Function
    End Class
End Class
