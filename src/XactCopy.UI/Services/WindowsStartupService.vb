' -----------------------------------------------------------------------------
' File: src\XactCopy.UI\Services\WindowsStartupService.vb
' Purpose: Source file for XactCopy runtime behavior.
' -----------------------------------------------------------------------------

Imports Microsoft.Win32

Namespace Services
    ''' <summary>
    ''' Class WindowsStartupService.
    ''' </summary>
    Friend NotInheritable Class WindowsStartupService
        Private Const RunOnceKeyPath As String = "Software\Microsoft\Windows\CurrentVersion\RunOnce"
        Private Const RecoveryEntryName As String = "XactCopyRecoveryResume"

        ''' <summary>
        ''' Computes RegisterRecoveryRunOnce.
        ''' </summary>
        Public Function RegisterRecoveryRunOnce() As Boolean
            Try
                Using key = Registry.CurrentUser.CreateSubKey(RunOnceKeyPath, writable:=True)
                    If key Is Nothing Then
                        Return False
                    End If

                    Dim executablePath = Application.ExecutablePath
                    Dim command = $"""{executablePath}"" --recovery-autostart --resume-interrupted"
                    key.SetValue(RecoveryEntryName, command, RegistryValueKind.String)
                    Return True
                End Using
            Catch
                Return False
            End Try
        End Function

        ''' <summary>
        ''' Computes ClearRecoveryRunOnce.
        ''' </summary>
        Public Function ClearRecoveryRunOnce() As Boolean
            Try
                Using key = Registry.CurrentUser.OpenSubKey(RunOnceKeyPath, writable:=True)
                    If key Is Nothing Then
                        Return True
                    End If

                    If key.GetValue(RecoveryEntryName) IsNot Nothing Then
                        key.DeleteValue(RecoveryEntryName, throwOnMissingValue:=False)
                    End If
                End Using

                Return True
            Catch
                Return False
            End Try
        End Function
    End Class
End Namespace
