Imports System.Drawing

''' <summary>
''' Class WindowIconHelper.
''' </summary>
Friend NotInheritable Class WindowIconHelper
    Private Shared ReadOnly SyncRoot As New Object()
    Private Shared _cachedIcon As Icon

    Private Sub New()
    End Sub

    ''' <summary>
    ''' Executes Apply.
    ''' </summary>
    Public Shared Sub Apply(form As Form)
        If form Is Nothing Then
            Return
        End If

        Try
            Dim icon = GetIcon()
            If icon IsNot Nothing Then
                form.Icon = CType(icon.Clone(), Icon)
            End If
        Catch
            ' Never block UI flow because of icon loading issues.
        End Try
    End Sub

    Private Shared Function GetIcon() As Icon
        If _cachedIcon IsNot Nothing Then
            Return _cachedIcon
        End If

        SyncLock SyncRoot
            If _cachedIcon Is Nothing Then
                Try
                    _cachedIcon = Icon.ExtractAssociatedIcon(Application.ExecutablePath)
                Catch
                    _cachedIcon = Nothing
                End Try

                If _cachedIcon Is Nothing Then
                    _cachedIcon = SystemIcons.Application
                End If
            End If
        End SyncLock

        Return _cachedIcon
    End Function
End Class
