Imports System.Drawing
Imports System.IO
Imports System.Runtime.InteropServices

''' <summary>
''' Class AboutForm.
''' </summary>
Friend Class AboutForm
    Inherits Form

    Private Const AuthorName As String = "Wimukthi Bandara"
    Private Const ProjectUrl As String = "https://github.com/Wimukthi/XactCopy"

    Private Shared ReadOnly LogoFilePath As String = Path.Combine(
        AppDomain.CurrentDomain.BaseDirectory,
        "Assets",
        "logo.png")

    Private ReadOnly _bannerPanel As New Panel()
    Private ReadOnly _logoPictureBox As New PictureBox()
    Private ReadOnly _versionOverlayLabel As New Label()
    Private ReadOnly _licenseRichTextBox As New RichTextBox()
    Private ReadOnly _detailsTextBox As New TextBox()
    Private ReadOnly _closeButton As New Button()
    Private ReadOnly _toolTip As New ToolTip() With {
        .ShowAlways = True,
        .InitialDelay = 350,
        .ReshowDelay = 120,
        .AutoPopDelay = 24000
    }

    ''' <summary>
    ''' Initializes a new instance.
    ''' </summary>
    Public Sub New()
        Text = "About XactCopy"
        StartPosition = FormStartPosition.CenterParent
        FormBorderStyle = FormBorderStyle.FixedSingle
        MaximizeBox = False
        MinimizeBox = False
        ShowInTaskbar = False
        MinimumSize = New Size(620, 560)
        Size = New Size(620, 560)
        WindowIconHelper.Apply(Me)

        BuildUi()
        ConfigureToolTips()
        PopulateContent()
    End Sub

    Protected Overrides Sub OnFormClosed(e As FormClosedEventArgs)
        Try
            If _logoPictureBox.Image IsNot Nothing Then
                _logoPictureBox.Image.Dispose()
                _logoPictureBox.Image = Nothing
            End If
        Catch
            ' Ignore cleanup failures for best-effort resource release.
        End Try

        MyBase.OnFormClosed(e)
    End Sub

    Private Sub BuildUi()
        Dim rootLayout As New TableLayoutPanel() With {
            .Dock = DockStyle.Fill,
            .ColumnCount = 1,
            .RowCount = 4,
            .Padding = New Padding(10)
        }
        rootLayout.RowStyles.Add(New RowStyle(SizeType.Absolute, 96.0F))
        rootLayout.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0F))
        rootLayout.RowStyles.Add(New RowStyle(SizeType.AutoSize))
        rootLayout.RowStyles.Add(New RowStyle(SizeType.AutoSize))

        _bannerPanel.Dock = DockStyle.Fill
        _bannerPanel.Margin = New Padding(0, 0, 0, 8)

        _logoPictureBox.Dock = DockStyle.Fill
        _logoPictureBox.SizeMode = PictureBoxSizeMode.Zoom
        _logoPictureBox.TabStop = False

        _versionOverlayLabel.Dock = DockStyle.Bottom
        _versionOverlayLabel.Height = 24
        _versionOverlayLabel.TextAlign = ContentAlignment.MiddleCenter
        _versionOverlayLabel.Padding = New Padding(0)
        _versionOverlayLabel.Font = New Font("Lucida Console", 9.0F, FontStyle.Regular, GraphicsUnit.Point)

        _bannerPanel.Controls.Add(_logoPictureBox)
        _bannerPanel.Controls.Add(_versionOverlayLabel)

        _licenseRichTextBox.Dock = DockStyle.Fill
        _licenseRichTextBox.ReadOnly = True
        _licenseRichTextBox.BorderStyle = BorderStyle.None
        _licenseRichTextBox.ScrollBars = RichTextBoxScrollBars.Vertical
        _licenseRichTextBox.DetectUrls = False
        _licenseRichTextBox.Margin = New Padding(0, 0, 0, 8)

        _detailsTextBox.Dock = DockStyle.Fill
        _detailsTextBox.Multiline = True
        _detailsTextBox.ReadOnly = True
        _detailsTextBox.BorderStyle = BorderStyle.None
        _detailsTextBox.TabStop = False
        _detailsTextBox.Height = 80
        _detailsTextBox.Margin = New Padding(0, 0, 0, 8)

        Dim buttonPanel As New FlowLayoutPanel() With {
            .Dock = DockStyle.Fill,
            .AutoSize = True,
            .FlowDirection = FlowDirection.RightToLeft,
            .WrapContents = False,
            .Margin = New Padding(0)
        }

        _closeButton.Text = "&Close"
        _closeButton.AutoSize = True
        AddHandler _closeButton.Click, Sub(sender, e) Close()

        buttonPanel.Controls.Add(_closeButton)

        rootLayout.Controls.Add(_bannerPanel, 0, 0)
        rootLayout.Controls.Add(_licenseRichTextBox, 0, 1)
        rootLayout.Controls.Add(_detailsTextBox, 0, 2)
        rootLayout.Controls.Add(buttonPanel, 0, 3)

        Controls.Add(rootLayout)
        AcceptButton = _closeButton
        CancelButton = _closeButton
    End Sub

    Private Sub ConfigureToolTips()
        _toolTip.SetToolTip(_closeButton, "Close the About window.")
        _toolTip.SetToolTip(_licenseRichTextBox, "License and system snapshot information.")
        _toolTip.SetToolTip(_detailsTextBox, "Author and project details.")
        _toolTip.SetToolTip(_versionOverlayLabel, "Current application version.")
        _toolTip.SetToolTip(_logoPictureBox, "XactCopy logo.")
    End Sub

    Private Sub PopulateContent()
        Dim capturedAt = DateTime.Now
        Dim totalRam = GetTotalPhysicalMemoryDisplay()

        _versionOverlayLabel.Text = $"Version {Application.ProductVersion}"

        _licenseRichTextBox.Text =
            "XactCopy is a resilient file copier designed for unstable storage media." & Environment.NewLine & Environment.NewLine &
            "System Info Snapshot:" & Environment.NewLine &
            $"- Captured: {capturedAt:yyyy-MM-dd HH:mm:ss}" & Environment.NewLine &
            $"- Machine: {Environment.MachineName}" & Environment.NewLine &
            $"- OS: {RuntimeInformation.OSDescription}" & Environment.NewLine &
            $"- OS Architecture: {RuntimeInformation.OSArchitecture}" & Environment.NewLine &
            $"- Process Architecture: {RuntimeInformation.ProcessArchitecture}" & Environment.NewLine &
            $"- Runtime: {RuntimeInformation.FrameworkDescription}" & Environment.NewLine &
            $"- Process Type: {If(Environment.Is64BitProcess, "64-bit", "32-bit")}" & Environment.NewLine &
            $"- Logical Processors: {Environment.ProcessorCount}" & Environment.NewLine &
            $"- Total RAM: {totalRam}"

        _detailsTextBox.Text =
            "Program: XactCopy" & Environment.NewLine &
            $"Author: {AuthorName}" & Environment.NewLine &
            "License: GNU GPL v3.0" & Environment.NewLine &
            $"Repository: {ProjectUrl}"

        LoadLogoImage()
    End Sub

    Private Shared Function GetTotalPhysicalMemoryDisplay() As String
        Try
            Dim computerInfo As New Microsoft.VisualBasic.Devices.ComputerInfo()
            Return FormatBytes(computerInfo.TotalPhysicalMemory)
        Catch
            Return "Unavailable"
        End Try
    End Function

    Private Shared Function FormatBytes(valueInBytes As ULong) As String
        Dim units As String() = {"B", "KB", "MB", "GB", "TB", "PB"}
        Dim value As Double = valueInBytes
        Dim unitIndex As Integer = 0

        While value >= 1024.0R AndAlso unitIndex < units.Length - 1
            value /= 1024.0R
            unitIndex += 1
        End While

        Return $"{value:0.##} {units(unitIndex)}"
    End Function

    Private Sub LoadLogoImage()
        If Not File.Exists(LogoFilePath) Then
            _logoPictureBox.Image = Nothing
            Return
        End If

        Try
            Using loadedImage = Image.FromFile(LogoFilePath)
                _logoPictureBox.Image = CType(loadedImage.Clone(), Image)
            End Using
        Catch
            _logoPictureBox.Image = Nothing
        End Try
    End Sub
End Class
