Imports System.IO
Imports System.IO.Pipes
Imports System.Collections.Generic
Imports System.Security.Cryptography
Imports System.Security.Principal
Imports System.Text
Imports System.Text.Json
Imports System.Threading
Imports System.Threading.Tasks

''' <summary>
''' Module Program.
''' </summary>
Friend Module Program
    Private Const SingleInstanceMutexName As String = "Local\XactCopy.UI.SingleInstance"
    Private Const ActivationPipeNameBase As String = "xactcopy-ui-activation-v1"
    Private Const ActivationPipeConnectTimeoutMs As Integer = 1500
    Private Const ActivationPipeReadBufferBytes As Integer = 4096
    Private Const ActivationPipeMaxPayloadBytes As Integer = 64 * 1024
    Private Const ActivationPipeMaxArgs As Integer = 128
    Private Const ActivationPipeMaxArgChars As Integer = 8192

    Private ReadOnly CrashLogDirectory As String = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "XactCopy",
        "crash")
    Private ReadOnly ActivationPipeName As String = BuildActivationPipeName()

    <STAThread()>
    Friend Sub Main(args As String())
        AddHandler Application.ThreadException, AddressOf HandleThreadException
        AddHandler AppDomain.CurrentDomain.UnhandledException, AddressOf HandleUnhandledException
        Application.SetUnhandledExceptionMode(UnhandledExceptionMode.CatchException)

        Dim launchOptions As LaunchOptions = LaunchOptions.Parse(args)
        Dim createdNew As Boolean = False
        Using instanceMutex As New Mutex(initiallyOwned:=True, name:=SingleInstanceMutexName, createdNew:=createdNew)
            If Not createdNew Then
                If Not ForwardLaunchToPrimaryInstance(args) Then
                    MessageBox.Show(
                        "XactCopy is already running and could not be activated automatically.",
                        "XactCopy",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information)
                End If
                Return
            End If

            Application.SetHighDpiMode(HighDpiMode.SystemAware)
            Application.EnableVisualStyles()
            ConfigureCompatibleTextRenderingDefault()
            Application.SetColorMode(ThemeSettings.GetPreferredColorMode())

            Dim mainForm As New MainForm(launchOptions)
            Dim activationCts As New CancellationTokenSource()
            Dim activationServerTask = StartActivationServerAsync(mainForm, activationCts.Token)

            Try
                Application.Run(mainForm)
            Finally
                activationCts.Cancel()
                Try
                    activationServerTask.Wait(TimeSpan.FromSeconds(2))
                Catch
                    ' Best-effort shutdown for background activation server.
                End Try
                activationCts.Dispose()
            End Try
        End Using
    End Sub

    Private Function ForwardLaunchToPrimaryInstance(args As String()) As Boolean
        Try
            Dim payload = JsonSerializer.Serialize(If(args, Array.Empty(Of String)()))
            Dim payloadBytes = Encoding.UTF8.GetByteCount(payload)
            If payloadBytes > ActivationPipeMaxPayloadBytes Then
                Return False
            End If

            Using client As New NamedPipeClientStream(
                ".",
                ActivationPipeName,
                PipeDirection.Out,
                PipeOptions.Asynchronous)

                client.Connect(ActivationPipeConnectTimeoutMs)

                Using writer As New StreamWriter(client, New UTF8Encoding(encoderShouldEmitUTF8Identifier:=False), bufferSize:=4096, leaveOpen:=True)
                    writer.AutoFlush = True
                    writer.WriteLine(payload)
                End Using
            End Using

            Return True
        Catch
            Return False
        End Try
    End Function

    Private Function StartActivationServerAsync(mainForm As MainForm, cancellationToken As CancellationToken) As Task
        If mainForm Is Nothing Then
            Return Task.CompletedTask
        End If

        Return Task.Run(
            Async Function()
                While Not cancellationToken.IsCancellationRequested
                    Using server As New NamedPipeServerStream(
                        ActivationPipeName,
                        PipeDirection.In,
                        maxNumberOfServerInstances:=1,
                        PipeTransmissionMode.Byte,
                        PipeOptions.Asynchronous Or PipeOptions.CurrentUserOnly)

                        Try
                            Await server.WaitForConnectionAsync(cancellationToken).ConfigureAwait(False)
                        Catch ex As OperationCanceledException
                            Exit While
                        Catch ex As Exception
                            LogWarning("Activation server wait failed.", ex)
                            Continue While
                        End Try

                        Try
                            Dim payload = Await ReadActivationPayloadAsync(server, cancellationToken).ConfigureAwait(False)
                            If String.IsNullOrWhiteSpace(payload) Then
                                Continue While
                            End If

                            Dim forwardedArgs As String() = Nothing
                            Try
                                forwardedArgs = JsonSerializer.Deserialize(Of String())(payload)
                            Catch
                                forwardedArgs = Array.Empty(Of String)()
                            End Try

                            Dim sanitizedArgs = SanitizeForwardedArgs(forwardedArgs)
                            Dim forwardedOptions = LaunchOptions.Parse(sanitizedArgs)
                            If mainForm.IsDisposed Then
                                Continue While
                            End If

                            Await WaitForMainFormHandleAsync(mainForm, cancellationToken).ConfigureAwait(False)
                            If mainForm.IsDisposed OrElse Not mainForm.IsHandleCreated Then
                                Continue While
                            End If

                            mainForm.BeginInvoke(New Action(Of LaunchOptions)(AddressOf mainForm.HandleForwardedLaunch), forwardedOptions)
                        Catch ex As OperationCanceledException
                            Exit While
                        Catch ex As InvalidDataException
                            LogWarning("Activation server rejected oversized or malformed payload.", ex)
                        Catch ex As Exception
                            LogWarning("Activation server receive failed.", ex)
                        End Try
                    End Using
                End While
            End Function,
            cancellationToken)
    End Function

    Private Async Function WaitForMainFormHandleAsync(mainForm As MainForm, cancellationToken As CancellationToken) As Task
        If mainForm Is Nothing Then
            Return
        End If

        Dim spinCount As Integer = 0
        While Not mainForm.IsDisposed AndAlso Not mainForm.IsHandleCreated AndAlso spinCount < 40
            cancellationToken.ThrowIfCancellationRequested()
            Await Task.Delay(50, cancellationToken).ConfigureAwait(False)
            spinCount += 1
        End While
    End Function

    Private Sub ConfigureCompatibleTextRenderingDefault()
        Try
            ' Keep WinForms startup resilient even if a window/control handle was created
            ' earlier by framework code or another startup path.
            Application.SetCompatibleTextRenderingDefault(False)
        Catch ex As InvalidOperationException
            LogWarning("Startup warning: compatible text rendering default could not be set.", ex)
        End Try
    End Sub

    Private Function BuildActivationPipeName() As String
        Dim identityToken = Environment.UserName
        Try
            Using identity = WindowsIdentity.GetCurrent()
                If identity IsNot Nothing AndAlso identity.User IsNot Nothing Then
                    identityToken = identity.User.Value
                End If
            End Using
        Catch
        End Try

        If String.IsNullOrWhiteSpace(identityToken) Then
            identityToken = "default"
        End If

        Dim hash = SHA256.HashData(Encoding.UTF8.GetBytes(identityToken))
        Dim suffix = Convert.ToHexString(hash).Substring(0, 16).ToLowerInvariant()
        Return $"{ActivationPipeNameBase}-{suffix}"
    End Function

    Private Async Function ReadActivationPayloadAsync(stream As Stream, cancellationToken As CancellationToken) As Task(Of String)
        If stream Is Nothing Then
            Return String.Empty
        End If

        Dim buffer(ActivationPipeReadBufferBytes - 1) As Byte
        Using payloadStream As New MemoryStream()
            While payloadStream.Length <= ActivationPipeMaxPayloadBytes
                Dim bytesRead = Await stream.ReadAsync(buffer.AsMemory(0, buffer.Length), cancellationToken).ConfigureAwait(False)
                If bytesRead <= 0 Then
                    Exit While
                End If

                Dim newlineIndex = Array.IndexOf(buffer, CByte(10), 0, bytesRead)
                If newlineIndex >= 0 Then
                    Dim bytesToWrite = newlineIndex
                    If bytesToWrite > 0 AndAlso buffer(bytesToWrite - 1) = CByte(13) Then
                        bytesToWrite -= 1
                    End If

                    If bytesToWrite > 0 Then
                        payloadStream.Write(buffer, 0, bytesToWrite)
                    End If
                    Exit While
                End If

                payloadStream.Write(buffer, 0, bytesRead)
                If payloadStream.Length > ActivationPipeMaxPayloadBytes Then
                    Throw New InvalidDataException("Activation payload exceeded the maximum allowed size.")
                End If
            End While

            If payloadStream.Length = 0 Then
                Return String.Empty
            End If

            Return Encoding.UTF8.GetString(payloadStream.ToArray())
        End Using
    End Function

    Private Function SanitizeForwardedArgs(args As String()) As String()
        If args Is Nothing OrElse args.Length = 0 Then
            Return Array.Empty(Of String)()
        End If

        Dim sanitized As New List(Of String)(Math.Min(args.Length, ActivationPipeMaxArgs))
        For Each rawValue In args
            If sanitized.Count >= ActivationPipeMaxArgs Then
                Exit For
            End If

            Dim text = If(rawValue, String.Empty)
            If text.Length > ActivationPipeMaxArgChars Then
                text = text.Substring(0, ActivationPipeMaxArgChars)
            End If
            sanitized.Add(text)
        Next

        Return sanitized.ToArray()
    End Function

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

    Private Sub LogWarning(context As String, ex As Exception)
        WriteDiagnosticLog("warn", context, ex)
    End Sub

    Private Sub LogFatal(context As String, ex As Exception)
        WriteDiagnosticLog("crash", context, ex)
    End Sub

    Private Sub WriteDiagnosticLog(filePrefix As String, context As String, ex As Exception)
        Try
            Directory.CreateDirectory(CrashLogDirectory)
            Dim safePrefix = If(String.IsNullOrWhiteSpace(filePrefix), "diag", filePrefix.Trim().ToLowerInvariant())
            Dim fileName = $"{safePrefix}-{DateTimeOffset.UtcNow:yyyyMMdd-HHmmss-fff}.log"
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
