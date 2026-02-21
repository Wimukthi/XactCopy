' -----------------------------------------------------------------------------
' File: src\XactCopy.UI\TextPromptDialog.vb
' Purpose: Source file for XactCopy runtime behavior.
' -----------------------------------------------------------------------------

''' <summary>
''' Class TextPromptDialog.
''' </summary>
Friend NotInheritable Class TextPromptDialog
    Private Sub New()
    End Sub

    ''' <summary>
    ''' Computes ShowPrompt.
    ''' </summary>
    Public Shared Function ShowPrompt(owner As IWin32Window, title As String, prompt As String, defaultValue As String, ByRef value As String) As DialogResult
        Using dialog As New Form()
            dialog.Text = If(String.IsNullOrWhiteSpace(title), "Input", title)
            dialog.StartPosition = FormStartPosition.CenterParent
            dialog.FormBorderStyle = FormBorderStyle.FixedDialog
            dialog.MinimizeBox = False
            dialog.MaximizeBox = False
            dialog.ShowInTaskbar = False
            dialog.ClientSize = New Size(520, 160)
            WindowIconHelper.Apply(dialog)

            Dim root As New TableLayoutPanel() With {
                .Dock = DockStyle.Fill,
                .ColumnCount = 1,
                .RowCount = 3,
                .Padding = New Padding(12)
            }
            root.RowStyles.Add(New RowStyle(SizeType.AutoSize))
            root.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0F))
            root.RowStyles.Add(New RowStyle(SizeType.AutoSize))

            Dim promptLabel As New Label() With {
                .AutoSize = True,
                .Text = If(String.IsNullOrWhiteSpace(prompt), "Value", prompt)
            }

            Dim valueTextBox As New TextBox() With {
                .Dock = DockStyle.Fill,
                .Text = If(defaultValue, String.Empty)
            }
            Dim resolvedValue As String = If(defaultValue, String.Empty)

            Dim okButton As New Button() With {
                .AutoSize = True,
                .Text = "OK"
            }

            Dim cancelButton As New Button() With {
                .AutoSize = True,
                .Text = "Cancel"
            }

            AddHandler okButton.Click,
                Sub(sender, e)
                    resolvedValue = valueTextBox.Text.Trim()
                    dialog.DialogResult = DialogResult.OK
                    dialog.Close()
                End Sub

            AddHandler cancelButton.Click,
                Sub(sender, e)
                    dialog.DialogResult = DialogResult.Cancel
                    dialog.Close()
                End Sub

            Dim buttonPanel As New FlowLayoutPanel() With {
                .Dock = DockStyle.Fill,
                .AutoSize = True,
                .FlowDirection = FlowDirection.RightToLeft,
                .WrapContents = False
            }
            buttonPanel.Controls.Add(cancelButton)
            buttonPanel.Controls.Add(okButton)

            root.Controls.Add(promptLabel, 0, 0)
            root.Controls.Add(valueTextBox, 0, 1)
            root.Controls.Add(buttonPanel, 0, 2)
            dialog.Controls.Add(root)

            dialog.AcceptButton = okButton
            dialog.CancelButton = cancelButton

            valueTextBox.SelectionStart = 0
            valueTextBox.SelectionLength = valueTextBox.TextLength
            valueTextBox.Focus()

            Dim result = dialog.ShowDialog(owner)
            If result <> DialogResult.OK Then
                value = defaultValue
            Else
                value = resolvedValue
            End If

            Return result
        End Using
    End Function
End Class
