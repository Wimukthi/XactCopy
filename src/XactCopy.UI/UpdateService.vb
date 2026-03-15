' -----------------------------------------------------------------------------
' File: src\XactCopy.UI\UpdateService.vb
' Purpose: Source file for XactCopy runtime behavior.
' -----------------------------------------------------------------------------

Imports System.Collections.Generic
Imports System.Linq
Imports System.Net.Http
Imports System.Security.Cryptography
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
    ''' <summary>
    ''' Gets or sets ExpectedSha256Hex.
    ''' </summary>
    Public Property ExpectedSha256Hex As String = String.Empty
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
                            Dim digest As String = GetJsonString(entry, "digest")

                            If Not String.IsNullOrWhiteSpace(name) AndAlso Not String.IsNullOrWhiteSpace(downloadUrl) Then
                                assets.Add(New UpdateAssetInfo() With {
                                    .Name = name,
                                    .DownloadUrl = downloadUrl,
                                    .Size = size,
                                    .ExpectedSha256Hex = ParseSha256HexCandidate(digest)
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

        Dim downloadUri As Uri = Nothing
        If Not Uri.TryCreate(asset.DownloadUrl, UriKind.Absolute, downloadUri) OrElse
            Not String.Equals(downloadUri.Scheme, Uri.UriSchemeHttps, StringComparison.OrdinalIgnoreCase) Then
            Throw New InvalidOperationException("Update download URL must use HTTPS.")
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
    ''' Resolves SHA-256 checksum for a selected release asset. Uses direct digest metadata first,
    ''' then optional checksum sidecar assets from the same release.
    ''' </summary>
    Public Async Function ResolveAssetSha256Async(
        settings As AppSettings,
        release As UpdateReleaseInfo,
        asset As UpdateAssetInfo,
        cancellationToken As CancellationToken) As Task(Of String)

        If settings Is Nothing Then
            Throw New ArgumentNullException(NameOf(settings))
        End If

        If release Is Nothing Then
            Throw New ArgumentNullException(NameOf(release))
        End If

        If asset Is Nothing Then
            Throw New ArgumentNullException(NameOf(asset))
        End If

        Dim direct = ParseSha256HexCandidate(asset.ExpectedSha256Hex)
        If direct.Length > 0 Then
            asset.ExpectedSha256Hex = direct
            Return direct
        End If

        If release.Assets Is Nothing OrElse release.Assets.Count = 0 Then
            Return String.Empty
        End If

        Dim checksumAssets = release.Assets.
            Where(Function(candidate) candidate IsNot Nothing AndAlso
                Not String.IsNullOrWhiteSpace(candidate.DownloadUrl) AndAlso
                IsChecksumAssetName(candidate.Name)).
            OrderByDescending(Function(candidate) ScoreChecksumAssetName(candidate.Name, asset.Name)).
            ToList()

        For Each checksumAsset In checksumAssets
            cancellationToken.ThrowIfCancellationRequested()
            Try
                Dim checksumText = Await DownloadTextAssetAsync(settings, checksumAsset.DownloadUrl, cancellationToken).ConfigureAwait(False)
                Dim parsed = ParseSha256FromChecksumText(checksumText, asset.Name)
                If parsed.Length = 0 Then
                    Continue For
                End If

                asset.ExpectedSha256Hex = parsed
                Return parsed
            Catch ex As OperationCanceledException
                Throw
            Catch
                ' Ignore sidecar parse/download errors and continue with other candidates.
            End Try
        Next

        Return String.Empty
    End Function

    ''' <summary>
    ''' Computes SHA-256 over a local package file.
    ''' </summary>
    Public Async Function ComputeFileSha256Async(filePath As String, cancellationToken As CancellationToken) As Task(Of String)
        If String.IsNullOrWhiteSpace(filePath) OrElse Not File.Exists(filePath) Then
            Throw New FileNotFoundException("Package file was not found.", filePath)
        End If

        Using stream As New FileStream(
            filePath,
            FileMode.Open,
            FileAccess.Read,
            FileShare.Read,
            bufferSize:=131072,
            options:=FileOptions.Asynchronous Or FileOptions.SequentialScan)

            Using hash = IncrementalHash.CreateHash(HashAlgorithmName.SHA256)
                Dim buffer(131071) As Byte
                Do
                    Dim read = Await stream.ReadAsync(buffer.AsMemory(0, buffer.Length), cancellationToken).ConfigureAwait(False)
                    If read <= 0 Then
                        Exit Do
                    End If

                    hash.AppendData(buffer, 0, read)
                Loop

                Return Convert.ToHexString(hash.GetHashAndReset()).ToLowerInvariant()
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

    Private Function ParseSha256HexCandidate(value As String) As String
        Dim text = If(value, String.Empty).Trim()
        If text.Length = 0 Then
            Return String.Empty
        End If

        If text.StartsWith("sha256:", StringComparison.OrdinalIgnoreCase) Then
            text = text.Substring("sha256:".Length).Trim()
        End If

        If text.Length <> 64 Then
            Return String.Empty
        End If

        For Each c In text
            If Not Uri.IsHexDigit(c) Then
                Return String.Empty
            End If
        Next

        Return text.ToLowerInvariant()
    End Function

    Private Function IsChecksumAssetName(name As String) As Boolean
        Dim lower = If(name, String.Empty).Trim().ToLowerInvariant()
        If lower.Length = 0 Then
            Return False
        End If

        Return lower.EndsWith(".sha256", StringComparison.OrdinalIgnoreCase) OrElse
            lower.EndsWith(".sha256sum", StringComparison.OrdinalIgnoreCase) OrElse
            lower.EndsWith(".sha256.txt", StringComparison.OrdinalIgnoreCase) OrElse
            lower.EndsWith("checksums.txt", StringComparison.OrdinalIgnoreCase) OrElse
            lower.EndsWith("checksums.sha256", StringComparison.OrdinalIgnoreCase) OrElse
            lower.Contains("sha256")
    End Function

    Private Function ScoreChecksumAssetName(checksumAssetName As String, targetAssetName As String) As Integer
        Dim checksumName = If(checksumAssetName, String.Empty).Trim().ToLowerInvariant()
        Dim targetName = If(targetAssetName, String.Empty).Trim().ToLowerInvariant()
        Dim score As Integer = 0

        If checksumName.Length = 0 Then
            Return score
        End If

        If targetName.Length > 0 Then
            If checksumName = $"{targetName}.sha256" OrElse
                checksumName = $"{targetName}.sha256sum" OrElse
                checksumName = $"{targetName}.sha256.txt" Then
                score += 100
            End If

            If checksumName.Contains(targetName) Then
                score += 25
            End If
        End If

        If checksumName.Contains("checksums") Then
            score += 10
        End If

        If checksumName.EndsWith(".sha256", StringComparison.OrdinalIgnoreCase) OrElse
            checksumName.EndsWith(".sha256sum", StringComparison.OrdinalIgnoreCase) Then
            score += 10
        End If

        Return score
    End Function

    Private Async Function DownloadTextAssetAsync(settings As AppSettings, url As String, cancellationToken As CancellationToken) As Task(Of String)
        Dim assetUri As Uri = Nothing
        If Not Uri.TryCreate(url, UriKind.Absolute, assetUri) OrElse
            Not String.Equals(assetUri.Scheme, Uri.UriSchemeHttps, StringComparison.OrdinalIgnoreCase) Then
            Throw New InvalidOperationException("Checksum download URL must use HTTPS.")
        End If

        Using client As New HttpClient()
            client.Timeout = TimeSpan.FromSeconds(DefaultTimeoutSeconds)

            Dim userAgent As String = ResolveUserAgent(settings)
            If Not String.IsNullOrWhiteSpace(userAgent) Then
                client.DefaultRequestHeaders.UserAgent.ParseAdd(userAgent)
            End If

            Using response = Await client.GetAsync(url, HttpCompletionOption.ResponseHeadersRead, cancellationToken).ConfigureAwait(False)
                response.EnsureSuccessStatusCode()
                Return Await response.Content.ReadAsStringAsync(cancellationToken).ConfigureAwait(False)
            End Using
        End Using
    End Function

    Private Function ParseSha256FromChecksumText(checksumText As String, targetAssetName As String) As String
        If String.IsNullOrWhiteSpace(checksumText) Then
            Return String.Empty
        End If

        Dim targetFileName = Path.GetFileName(If(targetAssetName, String.Empty))
        Dim lines = checksumText.Replace(vbCrLf, vbLf).Replace(vbCr, vbLf).Split({vbLf}, StringSplitOptions.RemoveEmptyEntries)
        Dim fallbackHash As String = String.Empty

        For Each lineRaw In lines
            Dim line = lineRaw.Trim()
            If line.Length = 0 Then
                Continue For
            End If

            Dim firstToken = line.Split({" "c, vbTab}, StringSplitOptions.RemoveEmptyEntries).FirstOrDefault()
            If String.IsNullOrWhiteSpace(firstToken) Then
                Continue For
            End If
            Dim parsedHash = ParseSha256HexCandidate(firstToken)
            If parsedHash.Length > 0 Then
                If targetFileName.Length = 0 Then
                    If fallbackHash.Length = 0 Then
                        fallbackHash = parsedHash
                    End If
                    Continue For
                End If

                Dim remainder = line.Substring(line.IndexOf(firstToken, StringComparison.OrdinalIgnoreCase) + firstToken.Length).Trim()
                remainder = remainder.TrimStart("*"c)
                remainder = remainder.Trim(""""c)
                Dim remainderFile = Path.GetFileName(remainder)
                If remainderFile.Length > 0 AndAlso
                    String.Equals(remainderFile, targetFileName, StringComparison.OrdinalIgnoreCase) Then
                    Return parsedHash
                End If

                If fallbackHash.Length = 0 Then
                    fallbackHash = parsedHash
                End If
            End If

            ' Supports "SHA256 (filename.ext) = <hash>" style entries.
            Dim marker = "SHA256 ("
            Dim markerIndex = line.IndexOf(marker, StringComparison.OrdinalIgnoreCase)
            Dim equalsIndex = line.IndexOf("="c)
            If markerIndex >= 0 AndAlso equalsIndex > markerIndex Then
                Dim nameStart = markerIndex + marker.Length
                Dim nameEnd = line.IndexOf(")", nameStart, StringComparison.OrdinalIgnoreCase)
                If nameEnd > nameStart Then
                    Dim candidateName = line.Substring(nameStart, nameEnd - nameStart).Trim()
                    Dim candidateHash = ParseSha256HexCandidate(line.Substring(equalsIndex + 1).Trim())
                    If candidateHash.Length > 0 Then
                        If targetFileName.Length = 0 OrElse
                            String.Equals(Path.GetFileName(candidateName), targetFileName, StringComparison.OrdinalIgnoreCase) Then
                            Return candidateHash
                        End If

                        If fallbackHash.Length = 0 Then
                            fallbackHash = candidateHash
                        End If
                    End If
                End If
            End If
        Next

        If lines.Length = 1 Then
            Return fallbackHash
        End If

        Return String.Empty
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
