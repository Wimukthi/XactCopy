Imports XactCopy.Models

Friend Enum RecoveryPromptAction
    ResumeNow = 0
    OpenJobManager = 1
    Dismiss = 2
End Enum

Friend Class RecoveryPromptForm
    Inherits Form

    Private Const LayoutWindowKey As String = "RecoveryPromptForm"

    Private ReadOnly _resumeButton As New Button()
    Private ReadOnly _manageButton As New Button()
    Private ReadOnly _dismissButton As New Button()
    Private ReadOnly _suppressPromptCheckBox As New CheckBox()
    Private ReadOnly _detailsBox As New TextBox()
    Private ReadOnly _toolTip As New ToolTip() With {
        .ShowAlways = True,
        .InitialDelay = 350,
        .ReshowDelay = 120,
        .AutoPopDelay = 24000
    }

    Private _selectedAction As RecoveryPromptAction = RecoveryPromptAction.Dismiss

    Public Sub New(interruptedRun As RecoveryActiveRun, interruptionReason As String)
        If interruptedRun Is Nothing Then
            Throw New ArgumentNullException(NameOf(interruptedRun))
        End If

        Text = "Resume Interrupted Job"
        StartPosition = FormStartPosition.CenterParent
        FormBorderStyle = FormBorderStyle.FixedDialog
        ShowInTaskbar = False
        MinimizeBox = False
        MaximizeBox = False
        ClientSize = New Size(720, 320)
        WindowIconHelper.Apply(Me)

        BuildUi(interruptedRun, interruptionReason)

        AddHandler Shown, AddressOf RecoveryPromptForm_Shown
        AddHandler FormClosing, AddressOf RecoveryPromptForm_FormClosing
    End Sub

    Public ReadOnly Property SelectedAction As RecoveryPromptAction
        Get
            Return _selectedAction
        End Get
    End Property

    Public ReadOnly Property SuppressPromptForThisRun As Boolean
        Get
            Return _suppressPromptCheckBox.Checked
        End Get
    End Property

    Private Sub BuildUi(interruptedRun As RecoveryActiveRun, interruptionReason As String)
        Dim runName = If(String.IsNullOrWhiteSpace(interruptedRun.JobName), "Interrupted Copy Session", interruptedRun.JobName)
        Dim startedText = interruptedRun.StartedUtc.ToLocalTime().ToString("yyyy-MM-dd HH:mm:ss")

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

        Dim titleLabel As New Label() With {
            .AutoSize = True,
            .Font = New Font(Font, FontStyle.Bold),
            .Text = "XactCopy detected an interrupted copy session."
        }

        _detailsBox.Dock = DockStyle.Fill
        _detailsBox.Multiline = True
        _detailsBox.ReadOnly = True
        _detailsBox.WordWrap = True
        _detailsBox.ScrollBars = ScrollBars.Vertical
        _detailsBox.Text =
            $"Job: {runName}{Environment.NewLine}" &
            $"Started: {startedText}{Environment.NewLine}" &
            $"Source: {interruptedRun.Options.SourceRoot}{Environment.NewLine}" &
            $"Destination: {interruptedRun.Options.DestinationRoot}{Environment.NewLine}" &
            $"Journal: {interruptedRun.JournalPath}{Environment.NewLine}{Environment.NewLine}" &
            $"Reason: {If(String.IsNullOrWhiteSpace(interruptionReason), "Unexpected exit", interruptionReason)}{Environment.NewLine}{Environment.NewLine}" &
            "Choose Resume Now to continue from the journal."

        _toolTip.SetToolTip(_detailsBox, "Details of the interrupted run and recovery journal path.")
        _toolTip.SetToolTip(_suppressPromptCheckBox, "Hide this prompt for this specific interrupted run.")
        _toolTip.SetToolTip(_resumeButton, "Resume the interrupted run now.")
        _toolTip.SetToolTip(_manageButton, "Open Job Manager to inspect or run jobs manually.")
        _toolTip.SetToolTip(_dismissButton, "Close this prompt and decide later.")

        _suppressPromptCheckBox.AutoSize = True
        _suppressPromptCheckBox.Text = "Do not prompt again for this interrupted run"

        _resumeButton.AutoSize = True
        _resumeButton.Text = "Resume Now"
        AddHandler _resumeButton.Click,
            Sub(sender, e)
                _selectedAction = RecoveryPromptAction.ResumeNow
                DialogResult = DialogResult.OK
                Close()
            End Sub

        _manageButton.AutoSize = True
        _manageButton.Text = "Open Job Manager"
        AddHandler _manageButton.Click,
            Sub(sender, e)
                _selectedAction = RecoveryPromptAction.OpenJobManager
                DialogResult = DialogResult.OK
                Close()
            End Sub

        _dismissButton.AutoSize = True
        _dismissButton.Text = "Not Now"
        AddHandler _dismissButton.Click,
            Sub(sender, e)
                _selectedAction = RecoveryPromptAction.Dismiss
                DialogResult = DialogResult.Cancel
                Close()
            End Sub

        Dim buttonPanel As New FlowLayoutPanel() With {
            .Dock = DockStyle.Fill,
            .AutoSize = True,
            .FlowDirection = FlowDirection.RightToLeft,
            .WrapContents = False
        }
        buttonPanel.Controls.Add(_dismissButton)
        buttonPanel.Controls.Add(_manageButton)
        buttonPanel.Controls.Add(_resumeButton)

        root.Controls.Add(titleLabel, 0, 0)
        root.Controls.Add(_detailsBox, 0, 1)
        root.Controls.Add(_suppressPromptCheckBox, 0, 2)
        root.Controls.Add(buttonPanel, 0, 3)

        Controls.Add(root)
        AcceptButton = _resumeButton
        CancelButton = _dismissButton
    End Sub

    Private Sub RecoveryPromptForm_Shown(sender As Object, e As EventArgs)
        UiLayoutManager.ApplyWindow(Me, LayoutWindowKey)
    End Sub

    Private Sub RecoveryPromptForm_FormClosing(sender As Object, e As FormClosingEventArgs)
        UiLayoutManager.CaptureWindow(Me, LayoutWindowKey)
    End Sub
End Class
