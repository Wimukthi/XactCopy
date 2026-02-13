Imports XactCopy.Configuration

Friend Module ThemeSettings
    Public Function GetPreferredColorMode() As SystemColorMode
        Dim store As New AppSettingsStore()
        Dim settings = store.Load()
        Return GetPreferredColorMode(settings)
    End Function

    Public Function GetPreferredColorMode(settings As AppSettings) As SystemColorMode
        Dim value As String = If(settings?.Theme, "dark")
        If String.Equals(value, "dark", StringComparison.OrdinalIgnoreCase) Then
            Return SystemColorMode.Dark
        End If

        If String.Equals(value, "classic", StringComparison.OrdinalIgnoreCase) Then
            Return SystemColorMode.Classic
        End If

        Return SystemColorMode.System
    End Function

    Public Sub SavePreferredColorMode(mode As SystemColorMode)
        Dim store As New AppSettingsStore()
        Dim settings = store.Load()
        settings.Theme = ModeToString(mode)
        store.Save(settings)
    End Sub

    Public Function ModeToString(mode As SystemColorMode) As String
        Select Case mode
            Case SystemColorMode.Dark
                Return "dark"
            Case SystemColorMode.Classic
                Return "classic"
            Case Else
                Return "system"
        End Select
    End Function
End Module
