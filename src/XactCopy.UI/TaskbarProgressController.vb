Imports System.Runtime.InteropServices

Friend Enum TaskbarProgressState
    NoProgress = 0
    Indeterminate = 1
    Normal = 2
    [Error] = 4
    Paused = 8
End Enum

<StructLayout(LayoutKind.Sequential)>
Friend Structure TaskbarRect
    Public Left As Integer
    Public Top As Integer
    Public Right As Integer
    Public Bottom As Integer
End Structure

<ComImport>
<Guid("56FDF342-FD6D-11d0-958A-006097C9A090")>
<InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
Friend Interface ITaskbarList
    Sub HrInit()
    Sub AddTab(hwnd As IntPtr)
    Sub DeleteTab(hwnd As IntPtr)
    Sub ActivateTab(hwnd As IntPtr)
    Sub SetActiveAlt(hwnd As IntPtr)
End Interface

<ComImport>
<Guid("602D4995-B13A-429b-A66E-1935E44F4317")>
<InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
Friend Interface ITaskbarList2
    Inherits ITaskbarList
    Shadows Sub HrInit()
    Shadows Sub AddTab(hwnd As IntPtr)
    Shadows Sub DeleteTab(hwnd As IntPtr)
    Shadows Sub ActivateTab(hwnd As IntPtr)
    Shadows Sub SetActiveAlt(hwnd As IntPtr)
    Sub MarkFullscreenWindow(hwnd As IntPtr, <MarshalAs(UnmanagedType.Bool)> fFullscreen As Boolean)
End Interface

<ComImport>
<Guid("EA1AFB91-9E28-4B86-90E9-9E9F8A5EEA84")>
<InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
Friend Interface ITaskbarList3
    Inherits ITaskbarList2
    Shadows Sub HrInit()
    Shadows Sub AddTab(hwnd As IntPtr)
    Shadows Sub DeleteTab(hwnd As IntPtr)
    Shadows Sub ActivateTab(hwnd As IntPtr)
    Shadows Sub SetActiveAlt(hwnd As IntPtr)
    Shadows Sub MarkFullscreenWindow(hwnd As IntPtr, <MarshalAs(UnmanagedType.Bool)> fFullscreen As Boolean)
    Sub SetProgressValue(hwnd As IntPtr, ullCompleted As ULong, ullTotal As ULong)
    Sub SetProgressState(hwnd As IntPtr, tbpFlags As TaskbarProgressState)
    Sub RegisterTab(hwndTab As IntPtr, hwndMDI As IntPtr)
    Sub UnregisterTab(hwndTab As IntPtr)
    Sub SetTabOrder(hwndTab As IntPtr, hwndInsertBefore As IntPtr)
    Sub SetTabActive(hwndTab As IntPtr, hwndMDI As IntPtr, dwReserved As UInteger)
    Sub ThumbBarAddButtons(hwnd As IntPtr, cButtons As UInteger, pButtons As IntPtr)
    Sub ThumbBarUpdateButtons(hwnd As IntPtr, cButtons As UInteger, pButtons As IntPtr)
    Sub ThumbBarSetImageList(hwnd As IntPtr, himl As IntPtr)
    Sub SetOverlayIcon(hwnd As IntPtr, hIcon As IntPtr, <MarshalAs(UnmanagedType.LPWStr)> pszDescription As String)
    Sub SetThumbnailTooltip(hwnd As IntPtr, <MarshalAs(UnmanagedType.LPWStr)> pszTip As String)
    Sub SetThumbnailClip(hwnd As IntPtr, ByRef clip As TaskbarRect)
End Interface

Friend NotInheritable Class TaskbarProgressController
    Private ReadOnly _taskbar As ITaskbarList3
    Private ReadOnly _supported As Boolean

    Public Sub New()
        Try
            Dim version = Environment.OSVersion.Version
            If version.Major < 6 OrElse (version.Major = 6 AndAlso version.Minor < 1) Then
                _supported = False
                Return
            End If

            Dim taskbarType = Type.GetTypeFromCLSID(New Guid("56FDF344-FD6D-11d0-958A-006097C9A090"), throwOnError:=False)
            If taskbarType Is Nothing Then
                _supported = False
                Return
            End If

            _taskbar = TryCast(Activator.CreateInstance(taskbarType), ITaskbarList3)
            If _taskbar Is Nothing Then
                _supported = False
                Return
            End If

            _taskbar.HrInit()
            _supported = True
        Catch
            _supported = False
        End Try
    End Sub

    Public Sub SetState(hwnd As IntPtr, state As TaskbarProgressState)
        If Not CanUse(hwnd) Then
            Return
        End If

        Try
            _taskbar.SetProgressState(hwnd, state)
        Catch
        End Try
    End Sub

    Public Sub SetProgress(hwnd As IntPtr, completed As ULong, total As ULong, state As TaskbarProgressState)
        If Not CanUse(hwnd) Then
            Return
        End If

        Dim safeTotal = Math.Max(1UL, total)
        Dim safeCompleted = Math.Min(completed, safeTotal)

        Try
            _taskbar.SetProgressState(hwnd, state)
            _taskbar.SetProgressValue(hwnd, safeCompleted, safeTotal)
        Catch
        End Try
    End Sub

    Public Sub Clear(hwnd As IntPtr)
        SetState(hwnd, TaskbarProgressState.NoProgress)
    End Sub

    Public Sub SetIndeterminate(hwnd As IntPtr)
        SetState(hwnd, TaskbarProgressState.Indeterminate)
    End Sub

    Private Function CanUse(hwnd As IntPtr) As Boolean
        Return _supported AndAlso _taskbar IsNot Nothing AndAlso hwnd <> IntPtr.Zero
    End Function
End Class

