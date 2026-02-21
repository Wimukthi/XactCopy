Option Strict On
Option Explicit On

Imports System.Drawing
Imports System.IO
Imports System.Linq
Imports System.Text.Json
Imports System.Windows.Forms

''' <summary>
''' Class UiWindowLayout.
''' </summary>
Friend NotInheritable Class UiWindowLayout
    ''' <summary>
    ''' Gets or sets X.
    ''' </summary>
    Public Property X As Integer
    ''' <summary>
    ''' Gets or sets Y.
    ''' </summary>
    Public Property Y As Integer
    ''' <summary>
    ''' Gets or sets Width.
    ''' </summary>
    Public Property Width As Integer
    ''' <summary>
    ''' Gets or sets Height.
    ''' </summary>
    Public Property Height As Integer
    ''' <summary>
    ''' Gets or sets WindowState.
    ''' </summary>
    Public Property WindowState As FormWindowState = FormWindowState.Normal
    ''' <summary>
    ''' Gets or sets Dpi.
    ''' </summary>
    Public Property Dpi As Integer = 96
End Class

''' <summary>
''' Class UiGridColumnLayout.
''' </summary>
Friend NotInheritable Class UiGridColumnLayout
    ''' <summary>
    ''' Gets or sets Width.
    ''' </summary>
    Public Property Width As Integer
    ''' <summary>
    ''' Gets or sets DisplayIndex.
    ''' </summary>
    Public Property DisplayIndex As Integer
End Class

''' <summary>
''' Class UiGridLayout.
''' </summary>
Friend NotInheritable Class UiGridLayout
    ''' <summary>
    ''' Gets or sets Columns.
    ''' </summary>
    Public Property Columns As Dictionary(Of String, UiGridColumnLayout) =
        New Dictionary(Of String, UiGridColumnLayout)(StringComparer.OrdinalIgnoreCase)
End Class

''' <summary>
''' Class UiLayoutState.
''' </summary>
Friend NotInheritable Class UiLayoutState
    ''' <summary>
    ''' Gets or sets Windows.
    ''' </summary>
    Public Property Windows As Dictionary(Of String, UiWindowLayout) =
        New Dictionary(Of String, UiWindowLayout)(StringComparer.OrdinalIgnoreCase)

    ''' <summary>
    ''' Gets or sets Splitters.
    ''' </summary>
    Public Property Splitters As Dictionary(Of String, Integer) =
        New Dictionary(Of String, Integer)(StringComparer.OrdinalIgnoreCase)

    ''' <summary>
    ''' Gets or sets Grids.
    ''' </summary>
    Public Property Grids As Dictionary(Of String, UiGridLayout) =
        New Dictionary(Of String, UiGridLayout)(StringComparer.OrdinalIgnoreCase)
End Class

''' <summary>
''' Class UiLayoutStore.
''' </summary>
Friend NotInheritable Class UiLayoutStore
    Private ReadOnly _layoutPath As String

    ''' <summary>
    ''' Initializes a new instance.
    ''' </summary>
    Public Sub New(Optional layoutPath As String = Nothing)
        If String.IsNullOrWhiteSpace(layoutPath) Then
            Dim root = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "XactCopy")
            _layoutPath = Path.Combine(root, "ui-layout.json")
        Else
            _layoutPath = layoutPath
        End If
    End Sub

    ''' <summary>
    ''' Computes Load.
    ''' </summary>
    Public Function Load() As UiLayoutState
        If Not File.Exists(_layoutPath) Then
            Return New UiLayoutState()
        End If

        Try
            Dim json = File.ReadAllText(_layoutPath)
            Dim options As New JsonSerializerOptions() With {
                .PropertyNameCaseInsensitive = True
            }
            Dim state = JsonSerializer.Deserialize(Of UiLayoutState)(json, options)
            If state Is Nothing Then
                Return New UiLayoutState()
            End If

            NormalizeState(state)
            Return state
        Catch
            Return New UiLayoutState()
        End Try
    End Function

    ''' <summary>
    ''' Executes Save.
    ''' </summary>
    Public Sub Save(state As UiLayoutState)
        If state Is Nothing Then
            Return
        End If

        Try
            Dim directoryPath = Path.GetDirectoryName(_layoutPath)
            If Not String.IsNullOrWhiteSpace(directoryPath) Then
                Directory.CreateDirectory(directoryPath)
            End If

            Dim options As New JsonSerializerOptions() With {
                .WriteIndented = True
            }
            Dim json = JsonSerializer.Serialize(state, options)
            File.WriteAllText(_layoutPath, json)
        Catch
            ' Ignore layout persistence failures.
        End Try
    End Sub

    Private Shared Sub NormalizeState(state As UiLayoutState)
        If state.Windows Is Nothing Then
            state.Windows = New Dictionary(Of String, UiWindowLayout)(StringComparer.OrdinalIgnoreCase)
        Else
            state.Windows = New Dictionary(Of String, UiWindowLayout)(state.Windows, StringComparer.OrdinalIgnoreCase)
        End If

        If state.Splitters Is Nothing Then
            state.Splitters = New Dictionary(Of String, Integer)(StringComparer.OrdinalIgnoreCase)
        Else
            state.Splitters = New Dictionary(Of String, Integer)(state.Splitters, StringComparer.OrdinalIgnoreCase)
        End If

        If state.Grids Is Nothing Then
            state.Grids = New Dictionary(Of String, UiGridLayout)(StringComparer.OrdinalIgnoreCase)
        Else
            state.Grids = New Dictionary(Of String, UiGridLayout)(state.Grids, StringComparer.OrdinalIgnoreCase)
        End If

        Dim gridKeys = state.Grids.Keys.ToArray()
        For Each key In gridKeys
            Dim layout = state.Grids(key)
            If layout Is Nothing Then
                state.Grids(key) = New UiGridLayout()
                Continue For
            End If

            If layout.Columns Is Nothing Then
                layout.Columns = New Dictionary(Of String, UiGridColumnLayout)(StringComparer.OrdinalIgnoreCase)
            Else
                layout.Columns = New Dictionary(Of String, UiGridColumnLayout)(layout.Columns, StringComparer.OrdinalIgnoreCase)
            End If
        Next
    End Sub
End Class

''' <summary>
''' Module UiLayoutManager.
''' </summary>
Friend Module UiLayoutManager
    Private ReadOnly _syncRoot As New Object()
    Private ReadOnly _store As New UiLayoutStore()
    Private _state As UiLayoutState

    ''' <summary>
    ''' Executes ApplyWindow.
    ''' </summary>
    Public Sub ApplyWindow(target As Form, windowKey As String)
        If target Is Nothing OrElse String.IsNullOrWhiteSpace(windowKey) Then
            Return
        End If

        Dim savedLayout As UiWindowLayout = Nothing

        SyncLock _syncRoot
            Dim state = GetStateLocked()
            If Not state.Windows.TryGetValue(windowKey, savedLayout) OrElse savedLayout Is Nothing Then
                Return
            End If
        End SyncLock

        If savedLayout.Width <= 0 OrElse savedLayout.Height <= 0 Then
            Return
        End If

        Dim scale = GetScaleFactor(target.DeviceDpi, savedLayout.Dpi)

        Dim desiredWidth = ScaleValue(savedLayout.Width, scale)
        Dim desiredHeight = ScaleValue(savedLayout.Height, scale)
        Dim desiredX = ScaleValue(savedLayout.X, scale)
        Dim desiredY = ScaleValue(savedLayout.Y, scale)

        If target.MinimumSize.Width > 0 Then
            desiredWidth = Math.Max(target.MinimumSize.Width, desiredWidth)
        End If

        If target.MinimumSize.Height > 0 Then
            desiredHeight = Math.Max(target.MinimumSize.Height, desiredHeight)
        End If

        Dim desiredBounds As New Rectangle(desiredX, desiredY, desiredWidth, desiredHeight)
        Dim visibleBounds = EnsureVisibleBounds(desiredBounds)

        target.StartPosition = FormStartPosition.Manual
        target.Bounds = visibleBounds

        If savedLayout.WindowState = FormWindowState.Maximized Then
            target.WindowState = FormWindowState.Maximized
        Else
            target.WindowState = FormWindowState.Normal
        End If
    End Sub

    ''' <summary>
    ''' Executes CaptureWindow.
    ''' </summary>
    Public Sub CaptureWindow(target As Form, windowKey As String)
        If target Is Nothing OrElse String.IsNullOrWhiteSpace(windowKey) Then
            Return
        End If

        Dim state = target.WindowState
        If state = FormWindowState.Minimized Then
            state = FormWindowState.Normal
        End If

        Dim bounds = If(target.WindowState = FormWindowState.Normal, target.Bounds, target.RestoreBounds)
        If bounds.Width <= 0 OrElse bounds.Height <= 0 Then
            Return
        End If

        Dim layout As New UiWindowLayout() With {
            .X = bounds.X,
            .Y = bounds.Y,
            .Width = bounds.Width,
            .Height = bounds.Height,
            .WindowState = state,
            .Dpi = If(target.DeviceDpi > 0, target.DeviceDpi, 96)
        }

        SyncLock _syncRoot
            Dim current = GetStateLocked()
            current.Windows(windowKey) = layout
            _store.Save(current)
        End SyncLock
    End Sub

    ''' <summary>
    ''' Executes ApplySplitter.
    ''' </summary>
    Public Sub ApplySplitter(split As SplitContainer, splitterKey As String, defaultDistance As Integer, panel1Min As Integer, panel2Min As Integer)
        If split Is Nothing Then
            Return
        End If

        Dim desiredDistance = defaultDistance

        If Not String.IsNullOrWhiteSpace(splitterKey) Then
            SyncLock _syncRoot
                Dim state = GetStateLocked()
                Dim savedDistance As Integer = 0
                If state.Splitters.TryGetValue(splitterKey, savedDistance) AndAlso savedDistance > 0 Then
                    desiredDistance = savedDistance
                End If
            End SyncLock
        End If

        ConfigureSplitterSafe(split, desiredDistance, panel1Min, panel2Min)
    End Sub

    ''' <summary>
    ''' Executes CaptureSplitter.
    ''' </summary>
    Public Sub CaptureSplitter(split As SplitContainer, splitterKey As String)
        If split Is Nothing OrElse String.IsNullOrWhiteSpace(splitterKey) Then
            Return
        End If

        Dim availableLength = If(split.Orientation = Orientation.Vertical, split.Width, split.Height)
        If availableLength <= 0 Then
            Return
        End If

        Dim distance = split.SplitterDistance
        If distance <= 0 Then
            Return
        End If

        SyncLock _syncRoot
            Dim state = GetStateLocked()
            state.Splitters(splitterKey) = distance
            _store.Save(state)
        End SyncLock
    End Sub

    ''' <summary>
    ''' Executes ApplyGridLayout.
    ''' </summary>
    Public Sub ApplyGridLayout(grid As DataGridView, gridKey As String)
        If grid Is Nothing OrElse String.IsNullOrWhiteSpace(gridKey) Then
            Return
        End If

        Dim saved As UiGridLayout = Nothing

        SyncLock _syncRoot
            Dim state = GetStateLocked()
            If Not state.Grids.TryGetValue(gridKey, saved) OrElse saved Is Nothing OrElse saved.Columns Is Nothing Then
                Return
            End If
        End SyncLock

        Dim orderedColumns As New List(Of KeyValuePair(Of DataGridViewColumn, UiGridColumnLayout))()
        For Each column As DataGridViewColumn In grid.Columns
            If String.IsNullOrWhiteSpace(column.Name) Then
                Continue For
            End If

            Dim columnLayout As UiGridColumnLayout = Nothing
            If Not saved.Columns.TryGetValue(column.Name, columnLayout) OrElse columnLayout Is Nothing Then
                Continue For
            End If

            If columnLayout.Width > 0 Then
                Dim minWidth = Math.Max(10, column.MinimumWidth)
                column.Width = Math.Max(minWidth, columnLayout.Width)
            End If

            orderedColumns.Add(New KeyValuePair(Of DataGridViewColumn, UiGridColumnLayout)(column, columnLayout))
        Next

        Dim sortedColumns = orderedColumns.
            OrderBy(Function(entry) entry.Value.DisplayIndex).
            ToList()

        For index = 0 To sortedColumns.Count - 1
            Try
                sortedColumns(index).Key.DisplayIndex = index
            Catch
                ' Ignore invalid display index transitions.
            End Try
        Next
    End Sub

    ''' <summary>
    ''' Executes CaptureGridLayout.
    ''' </summary>
    Public Sub CaptureGridLayout(grid As DataGridView, gridKey As String)
        If grid Is Nothing OrElse String.IsNullOrWhiteSpace(gridKey) Then
            Return
        End If

        Dim layout As New UiGridLayout()
        For Each column As DataGridViewColumn In grid.Columns
            If String.IsNullOrWhiteSpace(column.Name) Then
                Continue For
            End If

            layout.Columns(column.Name) = New UiGridColumnLayout() With {
                .Width = Math.Max(10, column.Width),
                .DisplayIndex = column.DisplayIndex
            }
        Next

        SyncLock _syncRoot
            Dim state = GetStateLocked()
            state.Grids(gridKey) = layout
            _store.Save(state)
        End SyncLock
    End Sub

    ''' <summary>
    ''' Executes ConfigureSplitterSafe.
    ''' </summary>
    Public Sub ConfigureSplitterSafe(split As SplitContainer, desiredDistance As Integer, panel1Min As Integer, panel2Min As Integer)
        If split Is Nothing Then
            Return
        End If

        Dim availableLength = If(split.Orientation = Orientation.Vertical, split.Width, split.Height)
        If availableLength <= 80 Then
            Return
        End If

        Dim safePanel1Min = Math.Max(0, Math.Min(panel1Min, availableLength - 40))
        Dim safePanel2Min = Math.Max(0, Math.Min(panel2Min, availableLength - safePanel1Min - 20))

        split.Panel1MinSize = safePanel1Min
        split.Panel2MinSize = safePanel2Min

        Dim minimumDistance = split.Panel1MinSize
        Dim maximumDistance = Math.Max(minimumDistance, availableLength - split.Panel2MinSize)
        Dim clampedDistance = Math.Max(minimumDistance, Math.Min(maximumDistance, desiredDistance))

        Try
            split.SplitterDistance = clampedDistance
        Catch
            ' Ignore transient layout timing issues.
        End Try
    End Sub

    Private Function GetStateLocked() As UiLayoutState
        If _state Is Nothing Then
            _state = _store.Load()
        End If

        Return _state
    End Function

    Private Function GetScaleFactor(currentDpi As Integer, savedDpi As Integer) As Single
        Dim effectiveCurrentDpi = If(currentDpi > 0, currentDpi, 96)
        Dim effectiveSavedDpi = If(savedDpi > 0, savedDpi, effectiveCurrentDpi)

        If effectiveCurrentDpi <= 0 OrElse effectiveSavedDpi <= 0 Then
            Return 1.0F
        End If

        Return CSng(effectiveCurrentDpi) / CSng(effectiveSavedDpi)
    End Function

    Private Function ScaleValue(value As Integer, scale As Single) As Integer
        If scale <= 0 Then
            Return value
        End If

        Return CInt(Math.Round(value * scale))
    End Function

    Private Function EnsureVisibleBounds(bounds As Rectangle) As Rectangle
        Dim targetScreen As Screen = Nothing
        For Each candidate In Screen.AllScreens
            If candidate.WorkingArea.IntersectsWith(bounds) Then
                targetScreen = candidate
                Exit For
            End If
        Next

        If targetScreen Is Nothing Then
            targetScreen = Screen.PrimaryScreen
        End If

        Dim workingArea = targetScreen.WorkingArea

        Dim width = Math.Min(bounds.Width, workingArea.Width)
        Dim height = Math.Min(bounds.Height, workingArea.Height)
        Dim x = Math.Max(workingArea.Left, Math.Min(bounds.X, workingArea.Right - width))
        Dim y = Math.Max(workingArea.Top, Math.Min(bounds.Y, workingArea.Bottom - height))

        Return New Rectangle(x, y, width, height)
    End Function
End Module
