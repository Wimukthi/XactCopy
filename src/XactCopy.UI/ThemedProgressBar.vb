' -----------------------------------------------------------------------------
' File: src\XactCopy.UI\ThemedProgressBar.vb
' Purpose: Source file for XactCopy runtime behavior.
' -----------------------------------------------------------------------------

Option Strict On
Option Explicit On

Imports System.ComponentModel
Imports System.Drawing
Imports System.Windows.Forms

''' <summary>
''' Class ThemedProgressBar.
''' </summary>
Friend Class ThemedProgressBar
    Inherits Control

    Private _minimum As Integer = 0
    Private _maximum As Integer = 100
    Private _value As Integer = 0
    Private _style As ProgressBarStyle = ProgressBarStyle.Continuous
    Private _trackColor As Color = Color.FromArgb(24, 24, 24)
    Private _fillColor As Color = Color.FromArgb(90, 120, 200)
    Private _borderColor As Color = Color.FromArgb(70, 70, 70)
    Private _showBorder As Boolean = False
    Private _showPercentageText As Boolean = False
    Private _percentageTextColor As Color = Color.Gainsboro

    ''' <summary>
    ''' Initializes a new instance.
    ''' </summary>
    Public Sub New()
        SetStyle(ControlStyles.UserPaint Or
                 ControlStyles.AllPaintingInWmPaint Or
                 ControlStyles.OptimizedDoubleBuffer Or
                 ControlStyles.ResizeRedraw Or
                 ControlStyles.Opaque, True)
        Size = New Size(100, 17)
        TabStop = False
        BackColor = _trackColor
        ForeColor = _fillColor
    End Sub

    ''' <summary>
    ''' Gets or sets Minimum.
    ''' </summary>
    <DefaultValue(0)>
    Public Property Minimum As Integer
        Get
            Return _minimum
        End Get
        Set(value As Integer)
            If _minimum = value Then
                Return
            End If

            _minimum = value
            If _maximum < _minimum Then
                _maximum = _minimum
            End If
            If _value < _minimum Then
                _value = _minimum
            End If

            Invalidate()
        End Set
    End Property

    ''' <summary>
    ''' Gets or sets Maximum.
    ''' </summary>
    <DefaultValue(100)>
    Public Property Maximum As Integer
        Get
            Return _maximum
        End Get
        Set(value As Integer)
            If _maximum = value Then
                Return
            End If

            _maximum = Math.Max(value, _minimum)
            If _value > _maximum Then
                _value = _maximum
            End If

            Invalidate()
        End Set
    End Property

    ''' <summary>
    ''' Gets or sets Value.
    ''' </summary>
    <DefaultValue(0)>
    Public Property Value As Integer
        Get
            Return _value
        End Get
        Set(value As Integer)
            Dim clamped = Math.Max(_minimum, Math.Min(_maximum, value))
            If _value = clamped Then
                Return
            End If

            _value = clamped
            Invalidate()
        End Set
    End Property

    ''' <summary>
    ''' Gets or sets Style.
    ''' </summary>
    <DefaultValue(GetType(ProgressBarStyle), "Continuous")>
    Public Property Style As ProgressBarStyle
        Get
            Return _style
        End Get
        Set(value As ProgressBarStyle)
            If _style = value Then
                Return
            End If

            _style = value
            Invalidate()
        End Set
    End Property

    ''' <summary>
    ''' Gets or sets TrackColor.
    ''' </summary>
    <DefaultValue(GetType(Color), "24, 24, 24")>
    Public Property TrackColor As Color
        Get
            Return _trackColor
        End Get
        Set(value As Color)
            If _trackColor = value Then
                Return
            End If

            _trackColor = value
            Invalidate()
        End Set
    End Property

    ''' <summary>
    ''' Gets or sets FillColor.
    ''' </summary>
    <DefaultValue(GetType(Color), "90, 120, 200")>
    Public Property FillColor As Color
        Get
            Return _fillColor
        End Get
        Set(value As Color)
            If _fillColor = value Then
                Return
            End If

            _fillColor = value
            Invalidate()
        End Set
    End Property

    ''' <summary>
    ''' Gets or sets BorderColor.
    ''' </summary>
    <DefaultValue(GetType(Color), "70, 70, 70")>
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

    ''' <summary>
    ''' Gets or sets ShowBorder.
    ''' </summary>
    <DefaultValue(False)>
    Public Property ShowBorder As Boolean
        Get
            Return _showBorder
        End Get
        Set(value As Boolean)
            If _showBorder = value Then
                Return
            End If

            _showBorder = value
            Invalidate()
        End Set
    End Property

    ''' <summary>
    ''' Gets or sets ShowPercentageText.
    ''' </summary>
    <DefaultValue(False)>
    Public Property ShowPercentageText As Boolean
        Get
            Return _showPercentageText
        End Get
        Set(value As Boolean)
            If _showPercentageText = value Then
                Return
            End If

            _showPercentageText = value
            Invalidate()
        End Set
    End Property

    ''' <summary>
    ''' Gets or sets PercentageTextColor.
    ''' </summary>
    <DefaultValue(GetType(Color), "220, 220, 220")>
    Public Property PercentageTextColor As Color
        Get
            Return _percentageTextColor
        End Get
        Set(value As Color)
            If _percentageTextColor = value Then
                Return
            End If

            _percentageTextColor = value
            Invalidate()
        End Set
    End Property

    Protected Overrides Sub OnPaint(e As PaintEventArgs)
        Dim rect As Rectangle = ClientRectangle
        If rect.Width <= 0 OrElse rect.Height <= 0 Then
            Return
        End If

        e.Graphics.Clear(_trackColor)

        Dim contentRect As Rectangle = rect
        If _showBorder AndAlso rect.Width > 2 AndAlso rect.Height > 2 Then
            contentRect = New Rectangle(1, 1, rect.Width - 2, rect.Height - 2)
        End If

        If _style <> ProgressBarStyle.Marquee Then
            Dim range As Integer = Math.Max(1, _maximum - _minimum)
            Dim progress As Double = CDbl(_value - _minimum) / range
            Dim fillWidth As Integer = CInt(Math.Round(contentRect.Width * progress))

            If fillWidth > 0 AndAlso contentRect.Height > 0 Then
                Using fillBrush As New SolidBrush(_fillColor)
                    e.Graphics.FillRectangle(fillBrush, New Rectangle(contentRect.X, contentRect.Y, fillWidth, contentRect.Height))
                End Using
            End If
        End If

        If _showBorder AndAlso rect.Width > 1 AndAlso rect.Height > 1 Then
            Using borderPen As New Pen(_borderColor)
                e.Graphics.DrawRectangle(borderPen, New Rectangle(0, 0, rect.Width - 1, rect.Height - 1))
            End Using
        End If

        If _showPercentageText AndAlso _style <> ProgressBarStyle.Marquee Then
            Dim range As Integer = Math.Max(1, _maximum - _minimum)
            Dim progress As Double = CDbl(_value - _minimum) / range
            progress = Math.Max(0.0R, Math.Min(1.0R, progress))
            Dim text = $"{progress * 100.0R:0.0}%"
            TextRenderer.DrawText(
                e.Graphics,
                text,
                Font,
                contentRect,
                _percentageTextColor,
                TextFormatFlags.HorizontalCenter Or TextFormatFlags.VerticalCenter Or TextFormatFlags.NoPrefix)
        End If
    End Sub

    Protected Overrides Sub OnPaintBackground(pevent As PaintEventArgs)
        If pevent Is Nothing Then
            Return
        End If

        pevent.Graphics.Clear(_trackColor)
    End Sub
End Class
