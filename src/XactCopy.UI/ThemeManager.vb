Imports System.Drawing
Imports System.Windows.Forms

Friend Module ThemeManager
    Public Sub ApplyTheme(root As Control, mode As SystemColorMode)
        Dim palette = ThemePalette.FromMode(mode)
        ApplyToControl(root, palette)
    End Sub

    Private Sub ApplyToControl(control As Control, palette As ThemePalette)
        If control Is Nothing Then
            Return
        End If

        If TypeOf control Is ThemedProgressBar Then
            Dim themedProgressBar = DirectCast(control, ThemedProgressBar)
            themedProgressBar.TrackColor = If(palette.IsDark, Color.FromArgb(18, 18, 18), SystemColors.ControlLight)
            themedProgressBar.FillColor = If(palette.IsDark, palette.FocusBorder, SystemColors.Highlight)
            themedProgressBar.BorderColor = If(palette.IsDark, palette.Border, SystemColors.ControlDark)
            themedProgressBar.ShowBorder = False
            themedProgressBar.BackColor = themedProgressBar.TrackColor
            themedProgressBar.ForeColor = themedProgressBar.FillColor
        ElseIf TypeOf control Is TextBox Then
            Dim textBox = DirectCast(control, TextBox)
            EnsureBorderWrapper(textBox, palette)
            textBox.BackColor = palette.Field
            textBox.ForeColor = palette.Text
            textBox.BorderStyle = BorderStyle.None
        ElseIf TypeOf control Is ComboBox Then
            Dim combo = DirectCast(control, ComboBox)
            EnsureBorderWrapper(combo, palette)
            combo.BackColor = palette.Field
            combo.ForeColor = palette.Text
            combo.FlatStyle = If(palette.IsDark, FlatStyle.Flat, FlatStyle.Standard)
        ElseIf TypeOf control Is ThemedNumericUpDown Then
            Dim numeric = DirectCast(control, ThemedNumericUpDown)
            numeric.ApplyColors(palette.Field, palette.Text, palette.Border)
        ElseIf TypeOf control Is Button Then
            Dim button = DirectCast(control, Button)
            If palette.IsDark Then
                button.FlatStyle = FlatStyle.Flat
                button.FlatAppearance.BorderColor = palette.Border
                button.FlatAppearance.BorderSize = 1
                button.FlatAppearance.MouseOverBackColor = Color.FromArgb(45, 45, 45)
                button.FlatAppearance.MouseDownBackColor = Color.FromArgb(55, 55, 55)
                button.BackColor = palette.Surface
                button.ForeColor = palette.Text
                button.UseVisualStyleBackColor = False
            Else
                button.FlatStyle = FlatStyle.Standard
                button.UseVisualStyleBackColor = True
                button.ForeColor = palette.Text
            End If
        ElseIf TypeOf control Is Label Then
            control.ForeColor = If(control.Enabled, palette.Text, palette.MutedText)
        ElseIf TypeOf control Is CheckBox OrElse TypeOf control Is RadioButton Then
            control.ForeColor = If(control.Enabled, palette.Text, palette.MutedText)
        ElseIf TypeOf control Is TreeView Then
            Dim tree = DirectCast(control, TreeView)
            EnsureBorderWrapper(tree, palette)
            tree.BackColor = palette.Field
            tree.ForeColor = palette.Text
            tree.BorderStyle = BorderStyle.None
        ElseIf TypeOf control Is ListBox Then
            Dim listBox = DirectCast(control, ListBox)
            EnsureBorderWrapper(listBox, palette)
            listBox.BackColor = palette.Field
            listBox.ForeColor = palette.Text
            listBox.BorderStyle = BorderStyle.None
        ElseIf TypeOf control Is ThemedGroupBox Then
            Dim themedGroup = DirectCast(control, ThemedGroupBox)
            themedGroup.BackColor = palette.Surface
            themedGroup.ForeColor = palette.Text
            themedGroup.BorderColor = palette.Border
            themedGroup.TitleColor = palette.Text
        ElseIf TypeOf control Is GroupBox Then
            control.BackColor = palette.Surface
            control.ForeColor = palette.Text
        ElseIf TypeOf control Is TabControl OrElse TypeOf control Is TabPage OrElse TypeOf control Is Panel OrElse
            TypeOf control Is TableLayoutPanel OrElse TypeOf control Is FlowLayoutPanel OrElse TypeOf control Is SplitContainer Then
            control.BackColor = palette.Surface
            control.ForeColor = palette.Text
        ElseIf TypeOf control Is DataGridView Then
            Dim grid = DirectCast(control, DataGridView)
            EnsureBorderWrapper(grid, palette)
            grid.EnableHeadersVisualStyles = False
            grid.BackgroundColor = palette.Surface
            grid.GridColor = palette.Border
            grid.BorderStyle = BorderStyle.None

            Dim headerStyle As New DataGridViewCellStyle() With {
                .BackColor = palette.Surface,
                .ForeColor = palette.Text,
                .SelectionBackColor = palette.Surface,
                .SelectionForeColor = palette.Text
            }
            grid.ColumnHeadersDefaultCellStyle = headerStyle
            grid.RowHeadersDefaultCellStyle = headerStyle

            Dim cellStyle As New DataGridViewCellStyle() With {
                .BackColor = palette.Field,
                .ForeColor = palette.Text,
                .SelectionBackColor = If(palette.IsDark, palette.FocusBorder, SystemColors.Highlight),
                .SelectionForeColor = If(palette.IsDark, palette.Text, SystemColors.HighlightText)
            }
            grid.DefaultCellStyle = cellStyle
            grid.AlternatingRowsDefaultCellStyle = New DataGridViewCellStyle(cellStyle) With {
                .BackColor = If(palette.IsDark, Color.FromArgb(40, 40, 40), palette.Field)
            }
        ElseIf TypeOf control Is MenuStrip Then
            Dim menu = DirectCast(control, MenuStrip)
            menu.BackColor = palette.Surface
            menu.ForeColor = palette.Text
            For Each item As ToolStripItem In menu.Items
                ApplyToolStripItemTheme(item, palette)
            Next
        Else
            If TypeOf control Is Form Then
                control.BackColor = palette.Base
                control.ForeColor = palette.Text
            End If
        End If

        For Each child As Control In control.Controls
            ApplyToControl(child, palette)
        Next
    End Sub

    Private Sub ApplyToolStripItemTheme(item As ToolStripItem, palette As ThemePalette)
        If item Is Nothing Then
            Return
        End If

        item.BackColor = palette.Surface
        item.ForeColor = palette.Text

        Dim dropDownItem = TryCast(item, ToolStripDropDownItem)
        If dropDownItem Is Nothing Then
            Return
        End If

        dropDownItem.DropDown.BackColor = palette.Surface
        dropDownItem.DropDown.ForeColor = palette.Text
        For Each child As ToolStripItem In dropDownItem.DropDownItems
            ApplyToolStripItemTheme(child, palette)
        Next
    End Sub

    Private Sub EnsureBorderWrapper(control As Control, palette As ThemePalette)
        If control Is Nothing OrElse control.Parent Is Nothing Then
            Return
        End If

        Dim parentPanel = TryCast(control.Parent, Panel)
        If parentPanel IsNot Nothing AndAlso Object.Equals(parentPanel.Tag, "ThemedBorder") Then
            parentPanel.BackColor = palette.Border
            parentPanel.Padding = New Padding(1)
            Return
        End If

        Dim parent = control.Parent
        Dim wrapper As New Panel() With {
            .Tag = "ThemedBorder",
            .BackColor = palette.Border,
            .Padding = New Padding(1),
            .Margin = control.Margin,
            .Dock = control.Dock,
            .Anchor = control.Anchor,
            .Location = control.Location,
            .Size = control.Size,
            .Name = control.Name & "_border",
            .TabStop = False,
            .TabIndex = control.TabIndex
        }

        control.Margin = New Padding(0)
        control.Dock = DockStyle.Fill

        Dim table = TryCast(parent, TableLayoutPanel)
        If table IsNot Nothing Then
            Dim row = table.GetRow(control)
            Dim col = table.GetColumn(control)
            Dim rowSpan = table.GetRowSpan(control)
            Dim colSpan = table.GetColumnSpan(control)
            table.Controls.Remove(control)
            wrapper.Controls.Add(control)
            table.Controls.Add(wrapper, col, row)
            If rowSpan > 1 Then
                table.SetRowSpan(wrapper, rowSpan)
            End If
            If colSpan > 1 Then
                table.SetColumnSpan(wrapper, colSpan)
            End If
        Else
            Dim index = parent.Controls.GetChildIndex(control)
            parent.Controls.Remove(control)
            wrapper.Controls.Add(control)
            parent.Controls.Add(wrapper)
            parent.Controls.SetChildIndex(wrapper, index)
        End If
    End Sub

    Private Structure ThemePalette
        Public Property IsDark As Boolean
        Public Property Base As Color
        Public Property Surface As Color
        Public Property Field As Color
        Public Property Border As Color
        Public Property FocusBorder As Color
        Public Property Text As Color
        Public Property MutedText As Color

        Public Shared Function FromMode(mode As SystemColorMode) As ThemePalette
            If mode = SystemColorMode.Dark Then
                Return New ThemePalette() With {
                    .IsDark = True,
                    .Base = Color.FromArgb(24, 24, 24),
                    .Surface = Color.FromArgb(32, 32, 32),
                    .Field = Color.FromArgb(45, 45, 45),
                    .Border = Color.FromArgb(70, 70, 70),
                    .FocusBorder = Color.FromArgb(90, 120, 200),
                    .Text = Color.Gainsboro,
                    .MutedText = Color.FromArgb(140, 140, 140)
                }
            End If

            Return New ThemePalette() With {
                .IsDark = False,
                .Base = SystemColors.Control,
                .Surface = SystemColors.Control,
                .Field = SystemColors.Window,
                .Border = SystemColors.ControlDark,
                .FocusBorder = SystemColors.Highlight,
                .Text = SystemColors.ControlText,
                .MutedText = SystemColors.GrayText
            }
        End Function
    End Structure
End Module
