Imports System.Collections.Generic
Imports System.Net.Http
Imports System.Text.Json
Imports System.Threading
Imports System.Threading.Tasks

Imports XactCopy.Configuration

Friend Class UpdateAssetInfo
    Public Property Name As String = String.Empty
    Public Property DownloadUrl As String = String.Empty
    Public Property Size As Long
End Class

Friend Class UpdateReleaseInfo
    Public Property TagName As String = String.Empty
    Public Property Version As Version = New Version(0, 0)
    Public Property Title As String = String.Empty
    Public Property Notes As String = String.Empty
    Public Property HtmlUrl As String = String.Empty
    Public Property Assets As List(Of UpdateAssetInfo) = New List(Of UpdateAssetInfo)()
End Class

Friend Module UpdateService
    Private Const DefaultTimeoutSeconds As Integer = 20

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

    Public Async Function GetLatestReleaseAsync(settings As AppSettings) As Task(Of UpdateReleaseInfo)
        Return Await GetLatestReleaseAsync(settings, CancellationToken.None)
    End Function

    Public Function IsUpdateAvailable(currentVersion As Version, latestVersion As Version) As Boolean
        If currentVersion Is Nothing OrElse latestVersion Is Nothing Then
            Return False
        End If

        Dim normalizedLatest As Version = NormalizeVersionForCompare(latestVersion)
        Dim normalizedCurrent As Version = NormalizeVersionForCompare(currentVersion)
        Return normalizedLatest.CompareTo(normalizedCurrent) > 0
    End Function

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

    Public Function NormalizeVersionForCompare(version As Version) As Version
        Dim build As Integer = If(version.Build >= 0, version.Build, 0)
        Dim revision As Integer = If(version.Revision >= 0, version.Revision, 0)
        Return New Version(version.Major, version.Minor, build, revision)
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
