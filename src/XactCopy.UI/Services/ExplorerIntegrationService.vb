Imports System.IO
Imports Microsoft.Win32

Namespace Services
    ''' <summary>
    ''' Class ExplorerIntegrationService.
    ''' </summary>
    Friend NotInheritable Class ExplorerIntegrationService
        Private Const ShellClassesRoot As String = "Software\Classes"
        Private Const MenuKeyName As String = "XactCopy"

        ' Use "%1" as the reliable primary selection token and keep "%*" as an
        ' optional tail for shells that still provide expanded multi-select args.
        Private Shared ReadOnly _targets As RegistrationTarget() = {
            New RegistrationTarget("Directory", "--from-explorer ""%1"" %*", supportsMultiSelect:=True),
            New RegistrationTarget("Drive", "--from-explorer ""%1"" %*", supportsMultiSelect:=True),
            New RegistrationTarget("*", "--from-explorer ""%1"" %*", supportsMultiSelect:=True),
            New RegistrationTarget("Directory\Background", "--from-explorer-folder ""%V""", supportsMultiSelect:=False)
        }

        ''' <summary>
        ''' Executes Sync.
        ''' </summary>
        Public Sub Sync(enabled As Boolean, executablePath As String)
            Dim resolvedPath = If(executablePath, String.Empty).Trim()
            If enabled Then
                If String.IsNullOrWhiteSpace(resolvedPath) OrElse Not File.Exists(resolvedPath) Then
                    Throw New FileNotFoundException("Cannot register Explorer integration because the executable path is invalid.", resolvedPath)
                End If

                Register(resolvedPath)
            Else
                Unregister()
            End If
        End Sub

        ''' <summary>
        ''' Computes IsRegistered.
        ''' </summary>
        Public Function IsRegistered() As Boolean
            For Each target In _targets
                Dim keyPath = BuildMenuPath(target.ShellPath)
                Using key = Registry.CurrentUser.OpenSubKey(keyPath, writable:=False)
                    If key Is Nothing Then
                        Return False
                    End If
                End Using
            Next

            Return True
        End Function

        Private Sub Register(executablePath As String)
            For Each target In _targets
                Dim menuPath = BuildMenuPath(target.ShellPath)
                Using menuKey = Registry.CurrentUser.CreateSubKey(menuPath, writable:=True)
                    If menuKey Is Nothing Then
                        Throw New InvalidOperationException($"Unable to create Explorer integration key: {menuPath}")
                    End If

                    menuKey.SetValue(String.Empty, "Copy with XactCopy", RegistryValueKind.String)
                    menuKey.SetValue("Icon", executablePath, RegistryValueKind.String)
                    If target.SupportsMultiSelect Then
                        menuKey.SetValue("MultiSelectModel", "Player", RegistryValueKind.String)
                    Else
                        menuKey.DeleteValue("MultiSelectModel", throwOnMissingValue:=False)
                    End If

                    Using commandKey = menuKey.CreateSubKey("command", writable:=True)
                        If commandKey Is Nothing Then
                            Throw New InvalidOperationException($"Unable to create Explorer integration command key: {menuPath}\command")
                        End If

                        Dim command = $"""{executablePath}"" {target.CommandArguments}"
                        commandKey.SetValue(String.Empty, command, RegistryValueKind.String)
                    End Using
                End Using
            Next
        End Sub

        Private Sub Unregister()
            For Each target In _targets
                Try
                    Registry.CurrentUser.DeleteSubKeyTree(BuildMenuPath(target.ShellPath), throwOnMissingSubKey:=False)
                Catch
                    ' Ignore partial cleanup failures.
                End Try
            Next
        End Sub

        Private Shared Function BuildMenuPath(shellPath As String) As String
            Return $"{ShellClassesRoot}\{shellPath}\shell\{MenuKeyName}"
        End Function

        ''' <summary>
        ''' Class RegistrationTarget.
        ''' </summary>
        Private NotInheritable Class RegistrationTarget
            ''' <summary>
            ''' Initializes a new instance.
            ''' </summary>
            Public Sub New(shellPath As String, commandArguments As String, supportsMultiSelect As Boolean)
                Me.ShellPath = If(shellPath, String.Empty)
                Me.CommandArguments = If(commandArguments, String.Empty)
                Me.SupportsMultiSelect = supportsMultiSelect
            End Sub

            ''' <summary>
            ''' Gets or sets ShellPath.
            ''' </summary>
            Public ReadOnly Property ShellPath As String
            ''' <summary>
            ''' Gets or sets CommandArguments.
            ''' </summary>
            Public ReadOnly Property CommandArguments As String
            ''' <summary>
            ''' Gets or sets SupportsMultiSelect.
            ''' </summary>
            Public ReadOnly Property SupportsMultiSelect As Boolean
        End Class
    End Class
End Namespace
