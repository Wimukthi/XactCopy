Imports System.IO
Imports System.Text
Imports System.Threading

Friend Module Program

    Private ReadOnly CrashLogDirectory As String = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "XactCopy",
        "crash")

    <STAThread()>
    Friend Sub Main(args As String())
        AddHandler Application.ThreadException, AddressOf HandleThreadException
        AddHandler AppDomain.CurrentDomain.UnhandledException, AddressOf HandleUnhandledException
        Application.SetUnhandledExceptionMode(UnhandledExceptionMode.CatchException)

        Dim launchOptions As LaunchOptions = LaunchOptions.Parse(args)
        Application.SetHighDpiMode(HighDpiMode.SystemAware)
        Application.EnableVisualStyles()
        ConfigureCompatibleTextRenderingDefault()
        Application.SetColorMode(ThemeSettings.GetPreferredColorMode())
        Application.Run(New MainForm(launchOptions))
    End Sub

    Private Sub ConfigureCompatibleTextRenderingDefault()
        Try
            ' Keep WinForms startup resilient even if a window/control handle was created
            ' earlier by framework code or another startup path.
            Application.SetCompatibleTextRenderingDefault(False)
        Catch ex As InvalidOperationException
            LogFatal("Startup warning: compatible text rendering default could not be set.", ex)
        End Try
    End Sub

    Private Sub HandleThreadException(sender As Object, e As ThreadExceptionEventArgs)
        LogFatal("UI thread exception", e.Exception)
        MessageBox.Show(
            "XactCopy captured an unexpected UI error. A crash report was written to LocalAppData\\XactCopy\\crash.",
            "XactCopy",
            MessageBoxButtons.OK,
            MessageBoxIcon.Error)
    End Sub

    Private Sub HandleUnhandledException(sender As Object, e As UnhandledExceptionEventArgs)
        Dim exceptionObject = TryCast(e.ExceptionObject, Exception)
        If exceptionObject Is Nothing Then
            exceptionObject = New Exception("Unknown unhandled exception.")
        End If

        LogFatal("Unhandled exception", exceptionObject)
        MessageBox.Show(
            "XactCopy captured an unexpected fatal error. A crash report was written to LocalAppData\\XactCopy\\crash.",
            "XactCopy",
            MessageBoxButtons.OK,
            MessageBoxIcon.Error)
    End Sub

    Private Sub LogFatal(context As String, ex As Exception)
        Try
            Directory.CreateDirectory(CrashLogDirectory)
            Dim fileName = $"crash-{DateTimeOffset.UtcNow:yyyyMMdd-HHmmss-fff}.log"
            Dim fullPath = Path.Combine(CrashLogDirectory, fileName)

            Dim builder As New StringBuilder()
            builder.AppendLine($"TimestampUtc: {DateTimeOffset.UtcNow:O}")
            builder.AppendLine($"Context: {context}")
            builder.AppendLine("Exception:")
            builder.AppendLine(ex.ToString())

            File.WriteAllText(fullPath, builder.ToString(), Encoding.UTF8)
        Catch
            ' Ignore logging failures in fatal-path handlers.
        End Try
    End Sub

End Module
