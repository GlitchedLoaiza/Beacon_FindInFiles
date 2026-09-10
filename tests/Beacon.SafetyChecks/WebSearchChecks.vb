Imports System.IO
Imports System.Text.Json
Imports System.Threading
Imports System.Threading.Tasks
Imports System.Windows
Imports System.Windows.Threading
Imports Microsoft.Web.WebView2.Core
Imports Microsoft.Web.WebView2.Wpf
Imports Beacon

Module WebSearchChecks
    Public Sub Run()
        Dim uiDispatcher = Dispatcher.CurrentDispatcher
        Dim previous = SynchronizationContext.Current
        SynchronizationContext.SetSynchronizationContext(New DispatcherSynchronizationContext(uiDispatcher))
        Try
            Dim task = ExerciseAsync()
            While Not task.IsCompleted
                uiDispatcher.Invoke(Sub()
                                  End Sub, DispatcherPriority.ContextIdle)
                Thread.Sleep(5)
            End While
            task.GetAwaiter().GetResult()
        Finally
            SynchronizationContext.SetSynchronizationContext(previous)
        End Try
    End Sub

    Private Async Function ExerciseAsync() As Task
        Dim profile = Path.Combine(Path.GetTempPath(), "BeaconWebSearchTests-" & Guid.NewGuid().ToString("N"))
        Dim browser As New WebView2()
        Dim window As New Window With {.Content = browser, .Width = 650, .Height = 400, .Left = -10000, .Top = -10000,
                                       .ShowActivated = False, .ShowInTaskbar = False, .WindowStartupLocation = WindowStartupLocation.Manual}
        Try
            window.Show()
            Dim environment = Await CoreWebView2Environment.CreateAsync(Nothing, profile).WaitAsync(TimeSpan.FromSeconds(20))
            Await browser.EnsureCoreWebView2Async(environment).WaitAsync(TimeSpan.FromSeconds(20))
            Await LoadHtml(browser, "<html><body><p>😀 <span>connection </span><b>failed</b>; code=123</p><p style='display:none'>hidden</p><script>var hidden='hidden';</script></body></html>")
            Dim original = Await browser.ExecuteScriptAsync("document.body.textContent")
            Dim query As New SearchQuery("connection failed", SearchMode.PlainText, False)
            Dim captured = JsonSerializer.Deserialize(Of String)(Await browser.ExecuteScriptAsync(WebSearchScripts.Capture(10000, "inline")))
            Require(captured.Contains("connection failed") AndAlso Not captured.Contains("hidden"), "Visible browser text capture failed.")
            Dim spans = query.FindHighlights(captured)
            Require(spans.Count = 1 AndAlso spans(0).Start = 3, "UTF-16 range offsets are incorrect.")
            Require(Await browser.ExecuteScriptAsync(WebSearchScripts.Apply(spans, "inline")) = "1", "Cross-node match was counted more than once.")
            Require(Await browser.ExecuteScriptAsync("Array.from(document.querySelectorAll('mark')).map(x=>x.textContent).join('')") = JsonSerializer.Serialize("connection failed"), "Cross-node highlighting lost text.")
            Require(Await browser.ExecuteScriptAsync("document.querySelectorAll('mark').length") = "2", "Both inline nodes were not highlighted.")
            Require(Await browser.ExecuteScriptAsync("document.body.textContent") = original, "Highlighting modified document content.")

            captured = JsonSerializer.Deserialize(Of String)(Await browser.ExecuteScriptAsync(WebSearchScripts.Capture(10000, "regex")))
            spans = New SearchQuery("(?<=code=)\d+", SearchMode.RegularExpression, False).FindHighlights(captured)
            Require(Await browser.ExecuteScriptAsync(WebSearchScripts.Apply(spans, "regex")) = "1", ".NET lookbehind range was not applied.")
            Require(Await browser.ExecuteScriptAsync("document.querySelector('mark').textContent") = """123""", "Browser highlight reinterpreted .NET regex.")

            Await browser.ExecuteScriptAsync(WebSearchScripts.Capture(10000, "old"))
            Await browser.ExecuteScriptAsync(WebSearchScripts.Capture(10000, "new"))
            Require(Await browser.ExecuteScriptAsync(WebSearchScripts.Apply(spans, "old")) = "null", "A stale capture can overwrite current match counts.")
            Require(Await browser.ExecuteScriptAsync("document.querySelectorAll('mark').length") = "0", "Stale capture modified the document.")
            Require(Await browser.ExecuteScriptAsync(WebSearchScripts.Apply(spans, "new")) = "1", "Latest capture failed to apply.")
            Require(Await browser.ExecuteScriptAsync(WebSearchScripts.Apply(spans, "old")) = "null", "Old completion did not remain stale.")
            Require(Await browser.ExecuteScriptAsync("document.querySelectorAll('mark').length") = "1", "Old capture erased current highlights.")
            captured = JsonSerializer.Deserialize(Of String)(Await browser.ExecuteScriptAsync(WebSearchScripts.Capture(5, "bounded")))
            Require(captured.Length <= 5, "Browser capture exceeded its configured bound.")
            Await browser.ExecuteScriptAsync(WebSearchScripts.Apply(New SearchQuery("co", SearchMode.PlainText, False).FindHighlights(captured), "bounded"))
            Require(Await browser.ExecuteScriptAsync("document.body.textContent") = original, "Bounded highlighting removed uncaptured text.")
            Await browser.ExecuteScriptAsync(WebSearchScripts.Capture(10000, "none"))
            Require(Await browser.ExecuteScriptAsync(WebSearchScripts.Apply(New List(Of SearchSpan)(), "none")) = "0", "A genuine zero-match result must differ from a stale result.")
            Await browser.ExecuteScriptAsync(WebSearchScripts.Capture(10000, "changed"))
            Await browser.ExecuteScriptAsync("document.querySelector('b').textContent = 'updated'")
            Require(Await browser.ExecuteScriptAsync(WebSearchScripts.Apply(spans, "changed")) = "null", "Changed DOM text accepted stale offsets.")
            Require(Await browser.ExecuteScriptAsync("document.querySelectorAll('mark').length") = "0", "Changed DOM was partially highlighted.")
            Dim report = ReportChecks.Sample(Path.GetTempPath())
            report.Run.QueryText = "error"
            Dim context = SearchDetail.FromRecord(New SearchQuery("error", SearchMode.PlainText, False),
                "before one" & vbLf & "before two" & vbLf & "error <img src=x onerror=alert(1)>" & vbLf & "after one" & vbLf & "after two", "Event text")
            report.Results(0).Matches(0).Context = context.Context
            report.Results(0).Matches(0).ContextKind = "Event text lines"
            Dim reportPath = Path.Combine(profile, "context-report.html")
            Await ScanReportWriter.SaveAsync(report, reportPath, ScanReportFormat.Html, True, False, CancellationToken.None)
            Await LoadHtml(browser, File.ReadAllText(reportPath))
            Require(Await browser.ExecuteScriptAsync("document.querySelectorAll('.snippet tr').length") = "5", "HTML report did not render five context rows.")
            Require(Await browser.ExecuteScriptAsync("document.querySelector('.snippet mark').textContent") = """error""", "HTML report did not highlight the match.")
            Require(Await browser.ExecuteScriptAsync("document.querySelectorAll('img,script').length") = "0", "Log text became active report content.")
            Require(Await browser.ExecuteScriptAsync("document.querySelectorAll('.snippet tr.focus').length") = "1", "The matching context line is not identified.")
            Require(Await browser.ExecuteScriptAsync("document.body.scrollWidth <= window.innerWidth + 1") = "true", "HTML report overflows a normal browser viewport.")
        Finally
            browser.Dispose()
            window.Close()
            For attempt = 1 To 5
                Try
                    If Directory.Exists(profile) Then Directory.Delete(profile, True)
                    Exit For
                Catch ex As IOException
                    Thread.Sleep(100)
                Catch ex As UnauthorizedAccessException
                    Thread.Sleep(100)
                End Try
            Next
            If Directory.Exists(profile) Then Console.WriteLine("WebView2 profile is still releasing: " & profile)
        End Try
    End Function

    Private Async Function LoadHtml(browser As WebView2, html As String) As Task
        Dim ready As New TaskCompletionSource(Of Boolean)(TaskCreationOptions.RunContinuationsAsynchronously)
        Dim handler As EventHandler(Of CoreWebView2NavigationCompletedEventArgs) =
            Sub(sender, args)
                If args.IsSuccess Then
                    ready.TrySetResult(True)
                Else
                    ready.TrySetException(New IOException("Test navigation failed: " & args.WebErrorStatus.ToString()))
                End If
            End Sub
        AddHandler browser.NavigationCompleted, handler
        Try
            browser.NavigateToString(html)
            Await ready.Task.WaitAsync(TimeSpan.FromSeconds(10))
        Finally
            RemoveHandler browser.NavigationCompleted, handler
        End Try
    End Function

    Private Sub Require(condition As Boolean, message As String)
        If Not condition Then Throw New InvalidOperationException(message)
    End Sub
End Module
