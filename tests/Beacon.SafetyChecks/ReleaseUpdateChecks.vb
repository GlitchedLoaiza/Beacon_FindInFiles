Imports System.Net
Imports System.Net.Http
Imports System.Text.Json
Imports System.Threading
Imports System.Threading.Tasks
Imports Beacon

Module ReleaseUpdateChecks
    Public Sub Run()
        For Each item In {
            ("v2.10.0", "Beacon v2.10.0", False, False, True),
            ("PublicRelease", "Beacon v2.0.2", False, False, True),
            ("v2.0.1", "Beacon v2.0.1", False, False, False),
            ("v1.9.0", "Beacon v1.9.0", False, False, False),
            ("v9.0.0", "Beacon v9.0.0", True, False, False),
            ("v9.0.0", "Beacon v9.0.0", False, True, False),
            ("v9.0.0-beta", "Beacon v9.0.0-beta", False, False, False),
            ("PublicRelease", "Unknown release", False, False, False)}
            Dim json = JsonSerializer.Serialize(New With {.tag_name = item.Item1, .name = item.Item2, .draft = item.Item3, .prerelease = item.Item4})
            Using handler As New ResponseHandler(json), client As New HttpClient(handler)
                Dim result = ReleaseUpdateChecker.CheckAsync(New Version(2, 0, 1, 0), CancellationToken.None, client).GetAwaiter().GetResult()
                Require((result IsNot Nothing) = item.Item5, "Incorrect release comparison: " & item.Item1)
                Require(handler.Requests = 1, "Release check sent extra requests.")
                Dim detailed = ReleaseUpdateChecker.CheckResultAsync(New Version(2, 0, 1, 0), CancellationToken.None, client).GetAwaiter().GetResult()
                Dim expected = If(item.Item5, ReleaseCheckStatus.UpdateAvailable,
                    If(item.Item1 = "v2.0.1" OrElse item.Item1 = "v1.9.0", ReleaseCheckStatus.UpToDate, ReleaseCheckStatus.Unavailable))
                Require(detailed.Status = expected, "Detailed release outcome is incorrect.")
            End Using
        Next
        For Each payload In {"not json", "[]", "{}", New String("x"c, 262145)}
            Using handler As New ResponseHandler(payload), client As New HttpClient(handler)
                Require(ReleaseUpdateChecker.CheckAsync(New Version(2, 0, 1), CancellationToken.None, client).GetAwaiter().GetResult() Is Nothing,
                        "Malformed/oversized response was accepted.")
                Require(ReleaseUpdateChecker.CheckResultAsync(New Version(2, 0, 1), CancellationToken.None, client).GetAwaiter().GetResult().Status = ReleaseCheckStatus.Unavailable,
                        "Invalid response was reported as up to date.")
            End Using
        Next
        For Each status In {HttpStatusCode.Forbidden, HttpStatusCode.NotFound, HttpStatusCode.TooManyRequests}
            Using handler As New ResponseHandler("{}", status), client As New HttpClient(handler)
                Require(ReleaseUpdateChecker.CheckAsync(New Version(2, 0, 1), CancellationToken.None, client).GetAwaiter().GetResult() Is Nothing, "HTTP failure was not silent.")
                Require(ReleaseUpdateChecker.CheckResultAsync(New Version(2, 0, 1), CancellationToken.None, client).GetAwaiter().GetResult().Status = ReleaseCheckStatus.Unavailable, "HTTP failure was reported as up to date.")
            End Using
        Next
        Using handler As New ResponseHandler("{}", offline:=True), client As New HttpClient(handler)
            Require(ReleaseUpdateChecker.CheckAsync(New Version(2, 0, 1), CancellationToken.None, client).GetAwaiter().GetResult() Is Nothing, "Offline failure was not silent.")
        End Using
        Using handler As New ResponseHandler("{}"), client As New HttpClient(handler), cancelled As New CancellationTokenSource()
            cancelled.Cancel()
            Require(ReleaseUpdateChecker.CheckAsync(New Version(2, 0, 1), cancelled.Token, client).GetAwaiter().GetResult() Is Nothing, "Cancellation was not silent.")
        End Using
    End Sub

    Public Sub SettingsResults()
        For Each dark In {False, True}
            For Each state In {ReleaseCheckStatus.UpdateAvailable, ReleaseCheckStatus.UpToDate, ReleaseCheckStatus.Unavailable}
                Dim completion As New TaskCompletionSource(Of ReleaseCheckResult)()
                Dim calls As Integer = 0
                Dim token As CancellationToken
                Dim window As New SettingsWindow(New BeaconSettings(), dark,
                    Function(ct)
                        calls += 1
                        token = ct
                        Return completion.Task
                    End Function)
                Try
                    Dim button = DirectCast(window.FindName("CheckUpdates_btn"), System.Windows.Controls.Button)
                    Dim message = DirectCast(window.FindName("UpdateCheckStatus_txt"), System.Windows.Controls.TextBlock)
                    button.RaiseEvent(New System.Windows.RoutedEventArgs(System.Windows.Controls.Button.ClickEvent))
                    Require(Not button.IsEnabled AndAlso message.Text.Contains("Checking"), "Manual check did not enter busy state.")
                    button.RaiseEvent(New System.Windows.RoutedEventArgs(System.Windows.Controls.Button.ClickEvent))
                    Require(calls = 1, "Duplicate manual update request started.")
                    completion.SetResult(New ReleaseCheckResult(state, New Version(2, 10, 0, 0)))
                    Dim timeout = System.Diagnostics.Stopwatch.StartNew()
                    While Not button.IsEnabled AndAlso timeout.Elapsed < TimeSpan.FromSeconds(2)
                        window.Dispatcher.Invoke(Sub()
                                                 End Sub, System.Windows.Threading.DispatcherPriority.ContextIdle)
                        Thread.Sleep(10)
                    End While
                    Dim expected = If(state = ReleaseCheckStatus.UpdateAvailable, "v2.10.0", If(state = ReleaseCheckStatus.UpToDate, "No updates pending", "Couldn't check"))
                    Require(button.IsEnabled AndAlso message.Text.Contains(expected), "Settings inline update result is incorrect.")
                    Require(window.FindName("UpdatePopup") Is Nothing, "Settings contains a popup notification.")
                Finally
                    window.Close()
                End Try
                Require(token.IsCancellationRequested, "Closing Settings did not cancel its update token.")
            Next
        Next
        Dim pending As New TaskCompletionSource(Of ReleaseCheckResult)()
        Dim pendingToken As CancellationToken
        Dim closing As New SettingsWindow(New BeaconSettings(), False,
            Function(ct)
                pendingToken = ct
                Return pending.Task
            End Function)
        DirectCast(closing.FindName("CheckUpdates_btn"), System.Windows.Controls.Button).RaiseEvent(New System.Windows.RoutedEventArgs(System.Windows.Controls.Button.ClickEvent))
        closing.Close()
        Require(pendingToken.IsCancellationRequested, "In-flight check was not cancelled on close.")
        pending.SetResult(New ReleaseCheckResult(ReleaseCheckStatus.UpToDate))
        closing.Dispatcher.Invoke(Sub()
                                 End Sub, System.Windows.Threading.DispatcherPriority.ContextIdle)
        Require(DirectCast(closing.FindName("UpdateCheckStatus_txt"), System.Windows.Controls.TextBlock).Text.Contains("Checking"), "Late result updated closed Settings.")
    End Sub

    Private Class ResponseHandler
        Inherits HttpMessageHandler
        Private ReadOnly _body As String
        Private ReadOnly _status As HttpStatusCode
        Private ReadOnly _offline As Boolean
        Public Property Requests As Integer

        Public Sub New(body As String, Optional status As HttpStatusCode = HttpStatusCode.OK, Optional offline As Boolean = False)
            _body = body
            _status = status
            _offline = offline
        End Sub

        Protected Overrides Function SendAsync(request As HttpRequestMessage, cancellationToken As CancellationToken) As Task(Of HttpResponseMessage)
            Requests += 1
            cancellationToken.ThrowIfCancellationRequested()
            Require(request.Method = HttpMethod.Get AndAlso request.RequestUri.AbsoluteUri = ReleaseUpdateChecker.LatestReleaseApi, "Unexpected update endpoint or method.")
            Require(request.Headers.UserAgent.Count > 0 AndAlso request.Headers.Contains("X-GitHub-Api-Version"), "GitHub request headers are missing.")
            If _offline Then Throw New HttpRequestException("Simulated offline machine")
            Return Task.FromResult(New HttpResponseMessage(_status) With {.Content = New StringContent(_body)})
        End Function
    End Class

    Private Sub Require(condition As Boolean, message As String)
        If Not condition Then Throw New InvalidOperationException(message)
    End Sub
End Module
