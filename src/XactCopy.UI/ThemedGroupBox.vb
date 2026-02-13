Imports System.ComponentModel
Imports System.Drawing
Imports System.Windows.Forms

Friend Class ThemedGroupBox
    Inherits GroupBox

    Private _borderColor As Color = Color.FromArgb(70, 70, 70)
    Private _titleColor As Color = Color.Gainsboro

    Public Sub New()
        SetStyle(ControlStyles.UserPaint Or
                 ControlStyles.AllPaintingInWmPaint Or
                 ControlStyles.OptimizedDoubleBuffer Or
                 ControlStyles.ResizeRedraw, True)

        Padding = New Padding(8, 22, 8, 8)
    End Sub

    <Browsable(False), DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)>
    Public Property BorderColor As Color
        Get
            Return _borderColor
        End Get
        Set(value As Color)
            If _borderColor = value Then
                Return
            End If

            _borderColor = value
            Invalidate()
        End Set
    End Property

    <Browsable(False), DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)>
    Public Property TitleColor As Color
        Get
            Return _titleColor
        End Get
        Set(value As Color)
            If _titleColor = value Then
                Return
            End If

            _titleColor = value
            Invalidate()
        End Set
    End Property

    Protected Overrides Sub OnPaintBackground(e As PaintEventArgs)
        e.Graphics.Clear(BackColor)
    End Sub

    Protected Overrides Sub OnPaint(e As PaintEventArgs)
        Dim textValue = If(Text, String.Empty)
        Dim hasText = textValue.Length > 0

        Dim textFlags = TextFormatFlags.Left Or TextFormatFlags.NoPrefix Or TextFormatFlags.EndEllipsis
        Dim measuredText = If(hasText, TextRenderer.MeasureText(e.Graphics, textValue, Font, New Size(Integer.MaxValue, Integer.MaxValue), textFlags), Size.Empty)

        Dim titleHeight = If(hasText, Math.Max(Font.Height, measuredText.Height), Font.Height)
        Dim topBorderY = Math.Max(6, titleHeight \ 2)

        Dim borderRect As New Rectangle(0, topBorderY, Math.Max(0, Width - 1), Math.Max(0, Height - topBorderY - 1))
        If borderRect.Width <= 0 OrElse borderRect.Height <= 0 Then
            Return
        End If

        Using pen As New Pen(_borderColor)
            If hasText Then
                Dim textX = 10
                Dim textRect As New Rectangle(textX, 0, measuredText.Width + 6, titleHeight)

                Dim leftSegmentEnd = Math.Max(borderRect.Left, textRect.Left - 4)
                If leftSegmentEnd > borderRect.Left Then
                    e.Graphics.DrawLine(pen, borderRect.Left, borderRect.Top, leftSegmentEnd, borderRect.Top)
                End If

                Dim rightSegmentStart = Math.Min(borderRect.Right, textRect.Right + 2)
                If rightSegmentStart < borderRect.Right Then
                    e.Graphics.DrawLine(pen, rightSegmentStart, borderRect.Top, borderRect.Right, borderRect.Top)
                End If

                e.Graphics.DrawLine(pen, borderRect.Left, borderRect.Top, borderRect.Left, borderRect.Bottom)
                e.Graphics.DrawLine(pen, borderRect.Right, borderRect.Top, borderRect.Right, borderRect.Bottom)
                e.Graphics.DrawLine(pen, borderRect.Left, borderRect.Bottom, borderRect.Right, borderRect.Bottom)

                Using textBackground As New SolidBrush(BackColor)
                    e.Graphics.FillRectangle(textBackground, textRect)
                End Using

                TextRenderer.DrawText(e.Graphics, textValue, Font, New Point(textRect.Left + 2, 0), _titleColor, textFlags)
            Else
                e.Graphics.DrawRectangle(pen, borderRect)
            End If
        End Using
    End Sub
End Class
