Imports System.Drawing
Imports System.Runtime.CompilerServices
Imports System.Windows.Forms
Imports XactCopy.Configuration

Friend Module UiAppearanceManager
    Private NotInheritable Class FontBaseline
        Public Sub New(font As Font)
            FamilyName = If(font?.FontFamily?.Name, SystemFonts.MessageBoxFont.FontFamily.Name)
            SizeInPoints = If(font Is Nothing, SystemFonts.MessageBoxFont.SizeInPoints, font.SizeInPoints)
            Style = If(font Is Nothing, FontStyle.Regular, font.Style)
        End Sub

        Public ReadOnly Property FamilyName As String
        Public ReadOnly Property SizeInPoints As Single
        Public ReadOnly Property Style As FontStyle
    End Class

    Private ReadOnly _fontBaselines As New ConditionalWeakTable(Of Control, FontBaseline)()

    Public Sub Apply(root As Control, settings As AppSettings)
        If root Is Nothing Then
            Return
        End If

        Dim safeSettings = If(settings, New AppSettings())
        Dim scaleFactor = GetEffectiveFontScale(safeSettings)
        Dim buttonHeight = GetButtonHeight(safeSettings)
        Dim compactPadding = GetButtonVerticalPadding(safeSettings)
        Dim treeItemHeight = GetTreeItemHeight(safeSettings)

        ApplyToControl(root, scaleFactor, buttonHeight, compactPadding, treeItemHeight)
    End Sub

    Private Sub ApplyToControl(control As Control, scaleFactor As Single, buttonHeight As Integer, buttonPadding As Integer, treeItemHeight As Integer)
        If control Is Nothing Then
            Return
        End If

        ApplyScaledFont(control, scaleFactor)

        If TypeOf control Is Button Then
            Dim button = DirectCast(control, Button)
            button.MinimumSize = New Size(button.MinimumSize.Width, buttonHeight)
            button.Padding = New Padding(10, buttonPadding, 10, buttonPadding)
        ElseIf TypeOf control Is TreeView Then
            Dim tree = DirectCast(control, TreeView)
            tree.ItemHeight = treeItemHeight
        End If

        For Each child As Control In control.Controls
            ApplyToControl(child, scaleFactor, buttonHeight, buttonPadding, treeItemHeight)
        Next
    End Sub

    Private Sub ApplyScaledFont(control As Control, scaleFactor As Single)
        If control Is Nothing Then
            Return
        End If

        Dim baseline = _fontBaselines.GetValue(control, Function(key) New FontBaseline(key.Font))
        Dim targetSize = Math.Max(7.0F, Math.Min(20.0F, baseline.SizeInPoints * scaleFactor))
        Dim currentFont = control.Font

        If currentFont IsNot Nothing AndAlso
            String.Equals(currentFont.FontFamily.Name, baseline.FamilyName, StringComparison.OrdinalIgnoreCase) AndAlso
            currentFont.Style = baseline.Style AndAlso
            Math.Abs(currentFont.SizeInPoints - targetSize) < 0.05F Then
            Return
        End If

        Try
            control.Font = New Font(baseline.FamilyName, targetSize, baseline.Style, GraphicsUnit.Point)
        Catch
            control.Font = New Font(SystemFonts.MessageBoxFont.FontFamily, targetSize, baseline.Style, GraphicsUnit.Point)
        End Try
    End Sub

    Public Function GetEffectiveFontScale(settings As AppSettings) As Single
        Dim safeSettings = If(settings, New AppSettings())
        Dim uiScale = Math.Max(0.9F, Math.Min(1.25F, safeSettings.UiScalePercent / 100.0F))
        Dim densityFactor As Single = 1.0F

        Select Case SettingsValueConverter.ToUiDensityChoice(safeSettings.UiDensity)
            Case UiDensityChoice.Compact
                densityFactor = 0.92F
            Case UiDensityChoice.Comfortable
                densityFactor = 1.08F
            Case Else
                densityFactor = 1.0F
        End Select

        Return uiScale * densityFactor
    End Function

    Private Function GetButtonHeight(settings As AppSettings) As Integer
        Select Case SettingsValueConverter.ToUiDensityChoice(settings.UiDensity)
            Case UiDensityChoice.Compact
                Return 24
            Case UiDensityChoice.Comfortable
                Return 34
            Case Else
                Return 28
        End Select
    End Function

    Private Function GetButtonVerticalPadding(settings As AppSettings) As Integer
        Select Case SettingsValueConverter.ToUiDensityChoice(settings.UiDensity)
            Case UiDensityChoice.Compact
                Return 2
            Case UiDensityChoice.Comfortable
                Return 5
            Case Else
                Return 3
        End Select
    End Function

    Private Function GetTreeItemHeight(settings As AppSettings) As Integer
        Select Case SettingsValueConverter.ToUiDensityChoice(settings.UiDensity)
            Case UiDensityChoice.Compact
                Return 20
            Case UiDensityChoice.Comfortable
                Return 26
            Case Else
                Return 22
        End Select
    End Function
End Module
