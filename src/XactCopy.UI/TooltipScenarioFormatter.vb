' -----------------------------------------------------------------------------
' File: src\XactCopy.UI\TooltipScenarioFormatter.vb
' Purpose: Builds detailed tooltip text with scenario-oriented guidance.
' -----------------------------------------------------------------------------

Imports System

Friend Module TooltipScenarioFormatter
    Public Function Compose(description As String) As String
        Dim summary = If(description, String.Empty).Trim()
        If String.IsNullOrWhiteSpace(summary) Then
            Return String.Empty
        End If

        If summary.IndexOf("Useful when:", StringComparison.OrdinalIgnoreCase) >= 0 Then
            Return summary
        End If

        Return summary & Environment.NewLine & "Useful when: " & InferScenario(summary)
    End Function

    Private Function InferScenario(description As String) As String
        Dim text = description.ToLowerInvariant()

        If text.Contains("resume") OrElse text.Contains("journal") OrElse text.Contains("recovery") Then
            Return "a copy/scan was interrupted by power loss, app crash, or media disconnect and you want to continue safely."
        End If

        If text.Contains("salvage") OrElse text.Contains("unreadable") Then
            Return "the source media has bad sectors and you prefer a best-effort copy over a hard failure."
        End If

        If text.Contains("continue") AndAlso text.Contains("error") Then
            Return "you are copying large batches and want one bad file to not stop the entire run."
        End If

        If text.Contains("verify") OrElse text.Contains("hash") Then
            Return "data integrity matters (archives, backups, or transfer validation) and extra verification time is acceptable."
        End If

        If text.Contains("adaptive buffer") OrElse text.Contains("buffer") Then
            Return "workloads include mixed small/large files and you want throughput tuned automatically per file behavior."
        End If

        If text.Contains("retry") Then
            Return "I/O errors are intermittent and temporary retries can recover without manual intervention."
        End If

        If text.Contains("timeout") Then
            Return "a device occasionally stalls and you need bounded waits to avoid indefinitely blocked operations."
        End If

        If text.Contains("wait") AndAlso (text.Contains("media") OrElse text.Contains("source") OrElse text.Contains("destination")) Then
            Return "source/destination paths are on removable or network media that may disappear and return later."
        End If

        If text.Contains("fragile") OrElse text.Contains("first read error") OrElse text.Contains("circuit breaker") Then
            Return "the drive is highly unstable and minimizing repeated hammering is more important than maximum recovery."
        End If

        If text.Contains("bad-range") OrElse text.Contains("bad block") OrElse text.Contains("scan") Then
            Return "you are working with suspect media and want to reuse prior unreadable-range knowledge to reduce stress and time."
        End If

        If text.Contains("overwrite") OrElse text.Contains("conflict") Then
            Return "destination already has files and you need deterministic behavior (replace, skip, or prompt)."
        End If

        If text.Contains("symlink") Then
            Return "the source tree contains symbolic links and you need explicit control over link traversal."
        End If

        If text.Contains("throughput") Then
            Return "you need to cap transfer speed to reduce system impact on interactive or shared workloads."
        End If

        If text.Contains("theme") OrElse text.Contains("accent") OrElse text.Contains("chrome") OrElse text.Contains("density") OrElse text.Contains("scale") Then
            Return "you want better readability/ergonomics on your display or a UI look that matches your desktop style."
        End If

        If text.Contains("log") OrElse text.Contains("diagnostic") OrElse text.Contains("telemetry") Then
            Return "you are troubleshooting behavior and need deeper runtime visibility."
        End If

        If text.Contains("update") OrElse text.Contains("release") OrElse text.Contains("user-agent") Then
            Return "you want controlled update behavior for online, offline, or enterprise network environments."
        End If

        If text.Contains("explorer") OrElse text.Contains("context menu") Then
            Return "you launch copy jobs from File Explorer and want faster right-click workflow integration."
        End If

        If text.Contains("queue") OrElse text.Contains("job") OrElse text.Contains("history") Then
            Return "you run recurring tasks and need repeatable execution, scheduling order, and auditability."
        End If

        If text.Contains("progress") OrElse text.Contains("eta") OrElse text.Contains("speed") Then
            Return "you want quick visibility into transfer pace and completion expectations during long runs."
        End If

        If text.Contains("close") OrElse text.Contains("save") OrElse text.Contains("cancel") Then
            Return "you need to commit or discard configuration changes without ambiguity."
        End If

        Return "you need predictable behavior and clearer control for the current copy/scan workload."
    End Function
End Module
