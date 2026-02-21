' -----------------------------------------------------------------------------
' File: src\XactCopy.UI\UpdateService.vb
' Purpose: Source file for XactCopy runtime behavior.
' -----------------------------------------------------------------------------

Imports System.Collections.Generic
Imports System.Net.Http
Imports System.Text.Json
Imports System.Threading
Imports System.Threading.Tasks
Imports System.Diagnostics
Imports System.IO

Imports XactCopy.Configuration

''' <summary>
''' Class UpdateAssetInfo.
''' </summary>
Friend Class UpdateAssetInfo
    ''' <summary>
    ''' Gets or sets Name.
    ''' </summary>
    Public Property Name As String = String.Empty
    ''' <summary>
    ''' Gets or sets DownloadUrl.
    ''' </summary>
    Public Property DownloadUrl As String = String.Empty
    ''' <summary>
    ''' Gets or sets Size.
    ''' </summary>
    Public Property Size As Long
End Class

''' <summary>
''' Class UpdateReleaseInfo.
''' </summary>
Friend Class UpdateReleaseInfo
    ''' <summary>
    ''' Gets or sets TagName.
    ''' </summary>
    Public Property TagName As String = String.Empty
    ''' <summary>
    ''' Gets or sets Version.
    ''' </summary>
    Public Property Version As Version = New Version(0, 0)
    ''' <summary>
    ''' Gets or sets Title.
    ''' </summary>
    Public Property Title As String = String.Empty
    ''' <summary>
    ''' Gets or sets Notes.
    ''' </summary>
    Public Property Notes As String = String.Empty
    ''' <summary>
    ''' Gets or sets HtmlUrl.
    ''' </summary>
    Public Property HtmlUrl As String = String.Empty
    ''' <summary>
    ''' Gets or sets Assets.
    ''' </summary>
    Public Property Assets As List(Of UpdateAssetInfo) = New List(Of UpdateAssetInfo)()
End Class

''' <summary>
''' Structure DownloadProgressInfo.
''' </summary>
Friend Structure DownloadProgressInfo
    Public ReadOnly BytesReceived As Long
    Public ReadOnly TotalBytes As Long
    Public ReadOnly SpeedBytesPerSec As Double
    Public ReadOnly Elapsed As TimeSpan

    ''' <summary>
    ''' Initializes a new instance.
    ''' </summary>
    Public Sub New(bytesReceived As Long, totalBytes As Long, speedBytesPerSec As Double, elapsed As TimeSpan)
        Me.BytesReceived = bytesReceived
        Me.TotalBytes = totalBytes
        Me.SpeedBytesPerSec = speedBytesPerSec
        Me.Elapsed = elapsed
    End Sub

    ''' <summary>
    ''' Gets or sets Percent.
    ''' </summary>
    Public ReadOnly Property Percent As Integer
        Get
            If TotalBytes <= 0 Then
                Return 0
            End If

            Dim percentValue As Double = (CDbl(BytesReceived) / CDbl(TotalBytes)) * 100.0R
            Return CInt(Math.Clamp(Math.Round(percentValue), 0.0R, 100.0R))
        End Get
    End Property
End Structure

''' <summary>
''' Module UpdateService.
''' </summary>
Friend Module UpdateService
    Private Const DefaultTimeoutSeconds As Integer = 20

    ''' <summary>
    ''' Computes GetLatestReleaseAsync.
    ''' </summary>
    Public Async Function GetLatestReleaseAsync(settings As AppSettings, cancellationToken As CancellationToken) As Task(Of UpdateReleaseInfo)
        If settings Is Nothing Then
            Throw New ArgumentNullException(NameOf(settings))
        End If

        Dim releaseUrl As String = settings.UpdateReleaseUrl
        If String.IsNullOrWhiteSpace(releaseUrl) Then
            Throw New InvalidOperationException("Update release URL is not configured.")
        End If

        Using client As New HttpClient()
            client.Timeout = TimeSpan.FromSeconds(DefaultTimeoutSeconds)

            Dim userAgent As String = ResolveUserAgent(settings)
            If Not String.IsNullOrWhiteSpace(userAgent) Then
                client.DefaultRequestHeaders.UserAgent.ParseAdd(userAgent)
            End If
            client.DefaultRequestHeaders.Accept.ParseAdd("application/vnd.github+json")

            Using response As HttpResponseMessage = Await client.GetAsync(releaseUrl, HttpCompletionOption.ResponseHeadersRead, cancellationToken)
                response.EnsureSuccessStatusCode()

                Dim payload As String = Await response.Content.ReadAsStringAsync(cancellationToken)
                Using doc As JsonDocument = JsonDocument.Parse(payload)
                    Dim root As JsonElement = doc.RootElement
                    Dim release As New UpdateReleaseInfo() With {
                        .TagName = GetJsonString(root, "tag_name"),
                        .Title = GetJsonString(root, "name"),
                        .Notes = GetJsonString(root, "body"),
                        .HtmlUrl = GetJsonString(root, "html_url")
                    }

                    release.Version = ParseVersionSafe(release.TagName)

                    Dim assets As New List(Of UpdateAssetInfo)()
                    Dim assetsElement As JsonElement
                    If root.TryGetProperty("assets", assetsElement) Then
                        For Each entry As JsonElement In assetsElement.EnumerateArray()
                            Dim name As String = GetJsonString(entry, "name")
                            Dim downloadUrl As String = GetJsonString(entry, "browser_download_url")
                            Dim size As Long = GetJsonLong(entry, "size")

                            If Not String.IsNullOrWhiteSpace(name) AndAlso Not String.IsNullOrWhiteSpace(downloadUrl) Then
                                assets.Add(New UpdateAssetInfo() With {
                                    .Name = name,
                                    .DownloadUrl = downloadUrl,
                                    .Size = size
                                })
                            End If
                        Next
                    End If

                    release.Assets = assets
                    Return release
                End Using
            End Using
        End Using
    End Function

    ''' <summary>
    ''' Computes GetLatestReleaseAsync.
    ''' </summary>
    Public Async Function GetLatestReleaseAsync(settings As AppSettings) As Task(Of UpdateReleaseInfo)
        Return Await GetLatestReleaseAsync(settings, CancellationToken.None)
    End Function

    ''' <summary>
    ''' Computes SelectBestAsset.
    ''' </summary>
    Public Function SelectBestAsset(release As UpdateReleaseInfo) As UpdateAssetInfo
        If release Is Nothing OrElse release.Assets Is Nothing OrElse release.Assets.Count = 0 Then
            Return Nothing
        End If

        Dim architectureToken As String = If(Environment.Is64BitProcess, "x64", "x86")
        Dim bestAsset As UpdateAssetInfo = Nothing
        Dim bestScore As Integer = -1

        For Each asset In release.Assets
            If asset Is Nothing OrElse String.IsNullOrWhiteSpace(asset.Name) Then
                Continue For
            End If

            Dim lowerName = asset.Name.ToLowerInvariant()
            Dim score As Integer = 0

            If lowerName.Contains("win") OrElse lowerName.Contains("windows") Then
                score += 2
            End If

            If lowerName.Contains(architectureToken) Then
                score += 3
            End If

            If lowerName.EndsWith(".zip", StringComparison.OrdinalIgnoreCase) Then
                score += 3
            ElseIf lowerName.EndsWith(".exe", StringComparison.OrdinalIgnoreCase) Then
                score += 2
            End If

            If score > bestScore Then
                bestScore = score
                bestAsset = asset
            End If
        Next

        Return bestAsset
    End Function

    ''' <summary>
    ''' Computes DownloadAssetAsync.
    ''' </summary>
    Public Async Function DownloadAssetAsync(
        settings As AppSettings,
        asset As UpdateAssetInfo,
        destinationPath As String,
        progress As IProgress(Of DownloadProgressInfo),
        cancellationToken As CancellationToken) As Task

        If settings Is Nothing Then
            Throw New ArgumentNullException(NameOf(settings))
        End If

        If asset Is Nothing OrElse String.IsNullOrWhiteSpace(asset.DownloadUrl) Then
            Throw New InvalidOperationException("No compatible update package is available.")
        End If

        If String.IsNullOrWhiteSpace(destinationPath) Then
            Throw New ArgumentException("Destination path must be provided.", NameOf(destinationPath))
        End If

        Dim destinationDirectory = Path.GetDirectoryName(destinationPath)
        If Not String.IsNullOrWhiteSpace(destinationDirectory) Then
            Directory.CreateDirectory(destinationDirectory)
        End If

        Using client As New HttpClient()
            client.Timeout = TimeSpan.FromMinutes(10)

            Dim userAgent As String = ResolveUserAgent(settings)
            If Not String.IsNullOrWhiteSpace(userAgent) Then
                client.DefaultRequestHeaders.UserAgent.ParseAdd(userAgent)
            End If

            Using response As HttpResponseMessage = Await client.GetAsync(asset.DownloadUrl, HttpCompletionOption.ResponseHeadersRead, cancellationToken)
                response.EnsureSuccessStatusCode()

                Dim totalBytes As Long = If(response.Content.Headers.ContentLength, asset.Size)
                Dim buffer(81919) As Byte
                Dim totalRead As Long = 0
                Dim downloadWatch As Stopwatch = Stopwatch.StartNew()

                Using networkStream As Stream = Await response.Content.ReadAsStreamAsync(cancellationToken)
                    Using fileStream As New FileStream(destinationPath, FileMode.Create, FileAccess.Write, FileShare.None)
                        Do
                            Dim read = Await networkStream.ReadAsync(buffer.AsMemory(0, buffer.Length), cancellationToken)
                            If read = 0 Then
                                Exit Do
                            End If

                            Await fileStream.WriteAsync(buffer.AsMemory(0, read), cancellationToken)
                            totalRead += read

                            If progress IsNot Nothing Then
                                Dim elapsed = downloadWatch.Elapsed
                                Dim speed = If(elapsed.TotalSeconds > 0, totalRead / elapsed.TotalSeconds, 0.0R)
                                progress.Report(New DownloadProgressInfo(totalRead, totalBytes, speed, elapsed))
                            End If
                        Loop
                    End Using
                End Using

                If progress IsNot Nothing Then
                    Dim elapsed = downloadWatch.Elapsed
                    Dim speed = If(elapsed.TotalSeconds > 0, totalRead / elapsed.TotalSeconds, 0.0R)
                    progress.Report(New DownloadProgressInfo(totalRead, totalBytes, speed, elapsed))
                End If
            End Using
        End Using
    End Function

    ''' <summary>
    ''' Computes IsUpdateAvailable.
    ''' </summary>
    Public Function IsUpdateAvailable(currentVersion As Version, latestVersion As Version) As Boolean
        If currentVersion Is Nothing OrElse latestVersion Is Nothing Then
            Return False
        End If

        Dim normalizedLatest As Version = NormalizeVersionForCompare(latestVersion)
        Dim normalizedCurrent As Version = NormalizeVersionForCompare(currentVersion)
        Return normalizedLatest.CompareTo(normalizedCurrent) > 0
    End Function

    ''' <summary>
    ''' Computes ParseVersionSafe.
    ''' </summary>
    Public Function ParseVersionSafe(text As String) As Version
        Dim normalized As String = NormalizeVersionText(text)
        If String.IsNullOrWhiteSpace(normalized) Then
            Return New Version(0, 0)
        End If

        Dim parts As String() = normalized.Split("."c)
        Dim major As Integer = If(parts.Length > 0, ToIntSafe(parts(0)), 0)
        Dim minor As Integer = If(parts.Length > 1, ToIntSafe(parts(1)), 0)

        If parts.Length <= 2 Then
            Return New Version(major, minor)
        End If

        Dim build As Integer = ToIntSafe(parts(2))
        If parts.Length = 3 Then
            Return New Version(major, minor, build)
        End If

        Dim revision As Integer = ToIntSafe(parts(3))
        Return New Version(major, minor, build, revision)
    End Function

    ''' <summary>
    ''' Computes NormalizeVersionForCompare.
    ''' </summary>
    Public Function NormalizeVersionForCompare(version As Version) As Version
        Dim build As Integer = If(version.Build >= 0, version.Build, 0)
        Dim revision As Integer = If(version.Revision >= 0, version.Revision, 0)
        Return New Version(version.Major, version.Minor, build, revision)
    End Function

    ''' <summary>
    ''' Computes FormatBytes.
    ''' </summary>
    Public Function FormatBytes(value As Long) As String
        Dim units = {"B", "KB", "MB", "GB", "TB"}
        Dim size As Double = value
        Dim unitIndex As Integer = 0

        While size >= 1024.0R AndAlso unitIndex < units.Length - 1
            size /= 1024.0R
            unitIndex += 1
        End While

        Return $"{size:F1} {units(unitIndex)}"
    End Function

    Private Function NormalizeVersionText(text As String) As String
        If String.IsNullOrWhiteSpace(text) Then
            Return String.Empty
        End If

        Dim clean As String = text.Trim()
        Dim startIndex As Integer = 0
        While startIndex < clean.Length AndAlso Not Char.IsDigit(clean(startIndex))
            startIndex += 1
        End While

        If startIndex >= clean.Length Then
            Return String.Empty
        End If

        clean = clean.Substring(startIndex)
        Dim endIndex As Integer = clean.IndexOfAny(New Char() {" "c, "-"c, "+"c})
        If endIndex >= 0 Then
            clean = clean.Substring(0, endIndex)
        End If

        Return clean
    End Function

    Private Function ToIntSafe(text As String) As Integer
        Dim value As Integer
        If Integer.TryParse(text, value) Then
            Return value
        End If

        Return 0
    End Function

    Private Function ResolveUserAgent(settings As AppSettings) As String
        If settings IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(settings.UserAgent) Then
            Return settings.UserAgent
        End If

        Dim versionText As String = If(String.IsNullOrWhiteSpace(Application.ProductVersion), "dev", Application.ProductVersion)
        Return $"XactCopy/{versionText}"
    End Function

    Private Function GetJsonString(root As JsonElement, name As String) As String
        Dim element As JsonElement
        If root.TryGetProperty(name, element) AndAlso element.ValueKind = JsonValueKind.String Then
            Return If(element.GetString(), String.Empty)
        End If

        Return String.Empty
    End Function

    Private Function GetJsonLong(root As JsonElement, name As String) As Long
        Dim element As JsonElement
        If root.TryGetProperty(name, element) AndAlso element.ValueKind = JsonValueKind.Number Then
            Dim value As Long
            If element.TryGetInt64(value) Then
                Return value
            End If
        End If

        Return 0
    End Function
End Module
