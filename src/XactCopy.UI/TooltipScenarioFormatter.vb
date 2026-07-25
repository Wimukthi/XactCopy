' -----------------------------------------------------------------------------
' File: src\XactCopy.UI\TooltipScenarioFormatter.vb
' Purpose: Builds normalized multiline tooltip text.
' -----------------------------------------------------------------------------

Imports System
Imports System.Collections.Generic

Friend Module TooltipScenarioFormatter
    Public Function Compose(description As String) As String
        Return Compose(New String() {description})
    End Function

    Public Function Compose(ParamArray lines() As String) As String
        If lines Is Nothing OrElse lines.Length = 0 Then
            Return String.Empty
        End If

        Dim normalized As New List(Of String)()
        For Each rawLine In lines
            If String.IsNullOrWhiteSpace(rawLine) Then
                Continue For
            End If

            Dim splitLines = rawLine.
                Replace(ControlChars.CrLf, ControlChars.Lf).
                Replace(ControlChars.Cr, ControlChars.Lf).
                Split(ControlChars.Lf)

            For Each splitLine In splitLines
                Dim trimmed = If(splitLine, String.Empty).Trim()
                If trimmed.Length > 0 Then
                    normalized.Add(trimmed)
                End If
            Next
        Next

        Return String.Join(Environment.NewLine, normalized)
    End Function
End Module
