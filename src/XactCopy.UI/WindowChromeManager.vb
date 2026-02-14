Imports System.Runtime.InteropServices
Imports Microsoft.Win32
Imports System.Windows.Forms
Imports XactCopy.Configuration

Friend Module WindowChromeManager
    Private Const DwmUseImmersiveDarkMode As Integer = 20
    Private Const DwmUseImmersiveDarkModeLegacy As Integer = 19

    <DllImport("dwmapi.dll")>
    Private Function DwmSetWindowAttribute(hwnd As IntPtr, dwAttribute As Integer, ByRef pvAttribute As Integer, cbAttribute As Integer) As Integer
    End Function

    Public Sub Apply(form As Form, mode As SystemColorMode, settings As AppSettings)
        If form Is Nothing Then
            Return
        End If

        Dim chromeMode = SettingsValueConverter.ToWindowChromeModeChoice(If(settings?.WindowChromeMode, String.Empty))
        Dim useThemedChrome = chromeMode = WindowChromeModeChoice.Themed
        Dim enableDarkTitleBar = useThemedChrome AndAlso ResolveShouldUseDarkTitleBar(mode)

        If form.IsHandleCreated Then
            ApplyDarkTitleBarFlag(form.Handle, enableDarkTitleBar)
            Return
        End If

        Dim createdHandler As EventHandler = Nothing
        createdHandler =
            Sub(sender, e)
                Try
                    ApplyDarkTitleBarFlag(form.Handle, enableDarkTitleBar)
                Finally
                    RemoveHandler form.HandleCreated, createdHandler
                End Try
            End Sub

        AddHandler form.HandleCreated, createdHandler
    End Sub

    Private Function ResolveShouldUseDarkTitleBar(mode As SystemColorMode) As Boolean
        Select Case mode
            Case SystemColorMode.Dark
                Return True
            Case SystemColorMode.System
                Return IsWindowsAppsDarkModeEnabled()
            Case Else
                Return False
        End Select
    End Function

    Private Function IsWindowsAppsDarkModeEnabled() As Boolean
        Const personalizeKey As String = "Software\Microsoft\Windows\CurrentVersion\Themes\Personalize"
        Const valueName As String = "AppsUseLightTheme"

        Try
            Using key = Registry.CurrentUser.OpenSubKey(personalizeKey, writable:=False)
                Dim rawValue = key?.GetValue(valueName)
                If rawValue Is Nothing Then
                    Return False
                End If

                Dim value = Convert.ToInt32(rawValue, Globalization.CultureInfo.InvariantCulture)
                Return value = 0
            End Using
        Catch
            Return False
        End Try
    End Function

    Private Sub ApplyDarkTitleBarFlag(windowHandle As IntPtr, enabled As Boolean)
        If windowHandle = IntPtr.Zero Then
            Return
        End If

        Dim flag = If(enabled, 1, 0)
        Dim size = Marshal.SizeOf(GetType(Integer))

        Dim result = DwmSetWindowAttribute(windowHandle, DwmUseImmersiveDarkMode, flag, size)
        If result <> 0 Then
            DwmSetWindowAttribute(windowHandle, DwmUseImmersiveDarkModeLegacy, flag, size)
        End If
    End Sub
End Module
