' -----------------------------------------------------------------------------
' File: src\XactCopy.UI\ThemedNumericUpDown.vb
' Purpose: Source file for XactCopy runtime behavior.
' -----------------------------------------------------------------------------

Imports System.ComponentModel
Imports System.Drawing
Imports System.Windows.Forms

''' <summary>
''' Class ThemedNumericUpDown.
''' </summary>
Friend Class ThemedNumericUpDown
    Inherits NumericUpDown

    Private _borderColor As Color = SystemColors.ControlDark

    ''' <summary>
    ''' Initializes a new instance.
    ''' </summary>
    Public Sub New()
        BorderStyle = BorderStyle.None
    End Sub

    ''' <summary>
    ''' Gets or sets BorderColor.
    ''' </summary>
    <Browsable(False), DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)>
    Public Property BorderColor As Color
        Get
            Return _borderColor
        End Get
        Set(value As Color)
            _borderColor = value
            Invalidate()
        End Set
    End Property

    ''' <summary>
    ''' Executes ApplyColors.
    ''' </summary>
    Public Sub ApplyColors(backColor As Color, foreColor As Color, borderColor As Color)
        Me.BackColor = backColor
        Me.ForeColor = foreColor
        Me.BorderColor = borderColor
        UpdateChildColors()
    End Sub

    Protected Overrides Sub OnHandleCreated(e As EventArgs)
        MyBase.OnHandleCreated(e)
        BorderStyle = BorderStyle.None
        UpdateChildColors()
    End Sub

    Protected Overrides Sub OnBackColorChanged(e As EventArgs)
        MyBase.OnBackColorChanged(e)
        UpdateChildColors()
    End Sub

    Protected Overrides Sub OnForeColorChanged(e As EventArgs)
        MyBase.OnForeColorChanged(e)
        UpdateChildColors()
    End Sub

    Private Sub UpdateChildColors()
        For Each child As Control In Controls
            child.BackColor = BackColor
            child.ForeColor = ForeColor
        Next
    End Sub

    Protected Overrides Sub WndProc(ByRef m As Message)
        MyBase.WndProc(m)

        Const WM_PAINT As Integer = &HF
        Const WM_NCPAINT As Integer = &H85

        If m.Msg = WM_PAINT OrElse m.Msg = WM_NCPAINT Then
            Using g As Graphics = Graphics.FromHwnd(Handle)
                Dim rect As New Rectangle(0, 0, Width - 1, Height - 1)
                Using pen As New Pen(_borderColor)
                    g.DrawRectangle(pen, rect)
                End Using
            End Using
        End If
    End Sub
End Class
