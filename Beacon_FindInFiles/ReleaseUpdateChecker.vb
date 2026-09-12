Imports System.IO
Imports System.Net.Http
Imports System.Text.Json
Imports System.Text.RegularExpressions
Imports System.Threading
Imports System.Threading.Tasks

Namespace Beacon
    Public Enum ReleaseCheckStatus
        Unavailable
        UpToDate
        UpdateAvailable
    End Enum

    Public NotInheritable Class ReleaseCheckResult
        Public ReadOnly Property Status As ReleaseCheckStatus
        Public ReadOnly Property LatestVersion As Version

        Public Sub New(status As ReleaseCheckStatus, Optional latestVersion As Version = Nothing)
            Me.Status = status
            Me.LatestVersion = latestVersion
        End Sub
    End Class

    Public NotInheritable Class ReleaseUpdateChecker
        Public Const LatestReleaseApi As String = "https://api.github.com/repos/GlitchedLoaiza/Beacon_FindInFiles/releases/latest"
        Public Const ReleasePage As String = "https://github.com/GlitchedLoaiza/Beacon_FindInFiles/releases/latest"
        Private Const MaximumResponseBytes As Integer = 262144
        Private Shared ReadOnly Client As New HttpClient(New HttpClientHandler With {.AllowAutoRedirect = False})

        Private Sub New()
        End Sub

        Public Shared Async Function CheckAsync(currentVersion As Version, token As CancellationToken,
                                                 Optional http As HttpClient = Nothing) As Task(Of Version)
            Dim result = Await CheckResultAsync(currentVersion, token, http).ConfigureAwait(False)
            Return If(result.Status = ReleaseCheckStatus.UpdateAvailable, result.LatestVersion, Nothing)
        End Function

        Public Shared Async Function CheckResultAsync(currentVersion As Version, token As CancellationToken,
                                                       Optional http As HttpClient = Nothing) As Task(Of ReleaseCheckResult)
            Dim unavailable As New ReleaseCheckResult(ReleaseCheckStatus.Unavailable)
            Try
                Using deadline = CancellationTokenSource.CreateLinkedTokenSource(token)
                    deadline.CancelAfter(TimeSpan.FromSeconds(8))
                    Using request As New HttpRequestMessage(HttpMethod.Get, LatestReleaseApi)
                        request.Headers.UserAgent.ParseAdd("Beacon-UpdateCheck/1.0")
                        request.Headers.Accept.ParseAdd("application/vnd.github+json")
                        request.Headers.Add("X-GitHub-Api-Version", "2022-11-28")
                        Using response = Await If(http, Client).SendAsync(request, HttpCompletionOption.ResponseHeadersRead, deadline.Token).ConfigureAwait(False)
                            If Not response.IsSuccessStatusCode Then Return unavailable
                            If response.Content.Headers.ContentLength.GetValueOrDefault() > MaximumResponseBytes Then Return unavailable
                            Using stream = Await response.Content.ReadAsStreamAsync(deadline.Token).ConfigureAwait(False)
                                Using buffer As New MemoryStream()
                                    Dim chunk(8191) As Byte
                                    Do
                                        Dim read = Await stream.ReadAsync(chunk.AsMemory(), deadline.Token).ConfigureAwait(False)
                                        If read = 0 Then Exit Do
                                        If buffer.Length + read > MaximumResponseBytes Then Return unavailable
                                        buffer.Write(chunk, 0, read)
                                    Loop
                                    Using document = JsonDocument.Parse(buffer.ToArray())
                                        Return EvaluateRelease(document.RootElement, currentVersion)
                                    End Using
                                End Using
                            End Using
                        End Using
                    End Using
                End Using
            Catch ex As Exception When TypeOf ex Is HttpRequestException OrElse TypeOf ex Is IOException OrElse
                                        TypeOf ex Is OperationCanceledException OrElse TypeOf ex Is JsonException
                Return unavailable
            End Try
        End Function

        Private Shared Function EvaluateRelease(release As JsonElement, currentVersion As Version) As ReleaseCheckResult
            Dim unavailable As New ReleaseCheckResult(ReleaseCheckStatus.Unavailable)
            If release.ValueKind <> JsonValueKind.Object Then Return unavailable
            Dim draft, prerelease, tag, title As JsonElement
            If Not release.TryGetProperty("draft", draft) OrElse draft.ValueKind <> JsonValueKind.False OrElse
               Not release.TryGetProperty("prerelease", prerelease) OrElse prerelease.ValueKind <> JsonValueKind.False Then Return unavailable
            Dim candidate As Version = Nothing
            If release.TryGetProperty("tag_name", tag) AndAlso tag.ValueKind = JsonValueKind.String Then
                candidate = ParseVersion(tag.GetString(), False)
            End If
            If candidate Is Nothing AndAlso release.TryGetProperty("name", title) AndAlso title.ValueKind = JsonValueKind.String Then
                candidate = ParseVersion(title.GetString(), True)
            End If
            If candidate Is Nothing OrElse currentVersion Is Nothing Then Return unavailable
            Dim current As New Version(currentVersion.Major, currentVersion.Minor, Math.Max(0, currentVersion.Build), Math.Max(0, currentVersion.Revision))
            Return New ReleaseCheckResult(If(candidate > current, ReleaseCheckStatus.UpdateAvailable, ReleaseCheckStatus.UpToDate), candidate)
        End Function

        Private Shared Function ParseVersion(text As String, allowProductName As Boolean) As Version
            If String.IsNullOrWhiteSpace(text) OrElse text.Length > 128 Then Return Nothing
            Dim prefix = If(allowProductName, "(?:Beacon\s+)?", "")
            Dim match = Regex.Match(text.Trim(), "\A" & prefix & "v?(\d+\.\d+(?:\.\d+){0,2})\z", RegexOptions.IgnoreCase Or RegexOptions.CultureInvariant, TimeSpan.FromMilliseconds(100))
            Dim parsed As Version = Nothing
            If Not match.Success OrElse Not Version.TryParse(match.Groups(1).Value, parsed) Then Return Nothing
            Return New Version(parsed.Major, parsed.Minor, Math.Max(0, parsed.Build), Math.Max(0, parsed.Revision))
        End Function
    End Class
End Namespace
