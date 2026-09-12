Imports System.Collections
Imports System.IO
Imports System.Reflection
Imports System.Text
Imports System.Text.Json
Imports System.Threading
Imports System.Threading.Tasks
Imports System.Windows
Imports System.Windows.Controls
Imports Beacon

Module HarChecks
    Private Function Entry(body As String, Optional encoding As String = "", Optional mime As String = "application/json") As String
        Return JsonSerializer.Serialize(New With {
            .startedDateTime = "2026-01-02T03:04:05Z", .time = 125.5,
            .request = New With {.method = "GET", .url = "https://user:pass@example.invalid/path?token=QUERY_SECRET&plain=QUERY_VALUE#fragment",
                .headers = {New With {.name = "Authorization", .value = "HEADER_SECRET"}},
                .postData = New With {.mimeType = "application/x-www-form-urlencoded", .text = "password=FORM_SECRET&safe=visible"}},
            .response = New With {.status = 503, .statusText = "Unavailable", .headers = {New With {.name = "Set-Cookie", .value = "COOKIE_SECRET"}},
                .content = New With {.text = body, .encoding = encoding, .mimeType = mime}}})
    End Function
    Private Function Read(json As String, Optional options As BeaconSettings = Nothing) As HarRecord
        Using doc = JsonDocument.Parse(json)
            Return HarReader.Read(doc.RootElement, If(options, New BeaconSettings()), CancellationToken.None)
        End Using
    End Function
    Public Sub Run()
        Dim json = Entry("{""token"":""BODY_SECRET"",""nested"":{""password"":""NESTED_SECRET""},""safe"":""error visible""}")
        Dim original = Read(json)
        Dim filtered As New HarFilter("get", "example.invalid", "500,503", "json", "125.5", "2026-01-02 03:04:05", "2026-01-02 03:04:05")
        Require(filtered.Matches(original), "Combined HAR filter or inclusive UTC boundary failed.")
        Require(Not New HarFilter("POST", "", "", "", "", "", "").Matches(original), "Method filter failed.")
        Require(Not New HarFilter("", "other.invalid", "", "", "", "", "").Matches(original), "Host filter failed.")
        Require(Not New HarFilter("", "", "200", "", "", "", "").Matches(original), "Status filter failed.")
        Require(Not New HarFilter("", "", "", "html", "", "", "").Matches(original), "MIME filter failed.")
        Require(Not New HarFilter("", "", "", "", "126", "", "").Matches(original), "Duration filter failed.")
        For Each invalid In {"-1", "600", "2xx", "200,"}
            Try
                Dim unused As New HarFilter("", "", invalid, "", "", "", "")
                Throw New InvalidOperationException("Invalid status filter was accepted.")
            Catch ex As ArgumentException
            End Try
        Next
        Dim safe = HarRedaction.Present(original, True, CancellationToken.None)
        For Each secret In {"HEADER_SECRET", "COOKIE_SECRET", "QUERY_SECRET", "QUERY_VALUE", "FORM_SECRET", "BODY_SECRET", "NESTED_SECRET", "user:pass"}
            Require(Not safe.SearchText().Contains(secret), "HAR redaction leaked " & secret)
        Next
        Require(safe.RedactionApplied AndAlso safe.ResponseBody.Contains("error visible"), "Redaction lost safe JSON content or state.")
        Require(HarRedaction.Present(original, False, CancellationToken.None).SearchText().Contains("BODY_SECRET"), "Disabled redaction still changed the data.")
        Dim query As New SearchQuery("BODY_SECRET", SearchMode.PlainText, False)
        Require(query.IsMatch(original.SearchText()), "Original sensitive data is not searchable.")
        Dim detail = HarRedaction.CaptureDetail(query, safe, "Request 1", 0, CancellationToken.None)
        Require(detail.VisibleMatches = 0 AndAlso Not detail.Excerpt.Contains("BODY_SECRET"), "Sensitive-only context leaked original data.")
        Dim encoded = Convert.ToBase64String(Encoding.UTF8.GetBytes("{""safe"":""decoded error""}"))
        Require(Read(Entry(encoded, "base64")).ResponseBody.Contains("decoded error"), "Base64 body was not decoded.")
        Require(Read(Entry(encoded, "base64"), New BeaconSettings With {.DecodeBase64HarBodies = False}).BodyIncomplete, "Disabled decoding was not reported.")
        Require(Read(Entry("???", "base64")).BodyIncomplete, "Malformed Base64 was not reported.")
        Require(Read(Entry("encoded payload", "unknown")).ResponseBody = "", "Unknown body encoding was treated as decoded text.")
        Require(Read(Entry(Convert.ToBase64String({CByte(255)}), "base64")).BodyIncomplete, "Binary response was not omitted.")
        Require(Read(Entry(New String("x"c, 1048577), mime:="text/plain"), New BeaconSettings With {.MaximumHarBodySizeMb = 1}).ResponseBody = "", "Oversized body was retained.")
        Require(HarRedaction.Present(Read(Entry("FREE_SECRET", mime:="text/plain")), True, CancellationToken.None).ResponseBody.Contains("withheld"), "Unstructured body was exposed with redaction enabled.")
        Pipeline(json)
    End Sub
    Private Sub Pipeline(entry As String)
        Dim root = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "BeaconHarChecks-" & Guid.NewGuid().ToString("N"))
        Directory.CreateDirectory(root)
        Dim path = System.IO.Path.Combine(root, "sample.har")
        File.WriteAllText(path, "{""log"":{""entries"":[" & entry & "]}}")
        Dim window As New MainWindow()
        Dim type = GetType(MainWindow), flags = BindingFlags.Instance Or BindingFlags.NonPublic
        window.RemoveHandler(FrameworkElement.LoadedEvent, type.GetMethod("MainWindow_Loaded", flags).CreateDelegate(GetType(RoutedEventHandler), window))
        RemoveHandler window.Closing, DirectCast(type.GetMethod("MainWindow_Closing", flags).CreateDelegate(GetType(ComponentModel.CancelEventHandler), window), ComponentModel.CancelEventHandler)
        Try
            Dim options As New BeaconSettings With {.StopAfterFirstMatchPerFile = False, .HarStatusCodes = "503"}
            type.GetField("_settings", flags).SetValue(window, options)
            type.GetField("_searchOptions", flags).SetValue(window, options)
            type.GetField("_activeQuery", flags).SetValue(window, New SearchQuery("BODY_SECRET", SearchMode.PlainText, False))
            type.GetField("_scanRootFolder", flags).SetValue(window, root)
            Dim work = Task.Run(Async Function()
                                    Await DirectCast(type.GetMethod("SearchSourceAsync", flags).Invoke(window, {path, CType(CancellationToken.None, Object)}), Task)
                                End Function)
            While Not work.IsCompleted
                window.Dispatcher.Invoke(Sub()
                                         End Sub, System.Windows.Threading.DispatcherPriority.ContextIdle)
                Thread.Sleep(5)
            End While
            work.GetAwaiter().GetResult()
            Dim hits = DirectCast(type.GetField("_hits", flags).GetValue(window), IList)
            Require(hits.Count = 1, "Sensitive-only match was dropped by the active HAR pipeline.")
            DirectCast(window.FindName("Results_lst"), ListBox).SelectedItem = hits(0)
            window.Left = -10000
            window.Top = -10000
            window.WindowStartupLocation = WindowStartupLocation.Manual
            window.WindowState = WindowState.Normal
            window.ShowActivated = False
            window.ShowInTaskbar = False
            window.Show()
            For Each dark In {True, False}
                type.GetMethod(If(dark, "ApplyDarkTheme", "ApplyLightTheme"), flags).Invoke(window, Nothing)
                Dim pane = DirectCast(window.FindName("HarToolsPane"), Expander)
                pane.IsExpanded = True
                window.UpdateLayout()
                Require(pane.Template.FindName("ToolsToggle", pane) IsNot Nothing, "HAR tools does not use the themed side pane.")
                Dim editor = DirectCast(DirectCast(window.FindName("HarPreviewFilters_host"), ContentControl).Content, HarFilterEditor)
                DirectCast(editor.FindName("Status_txt"), TextBox).Text = "200"
                DirectCast(window.FindName("ApplyHarFilter_btn"), Button).RaiseEvent(New RoutedEventArgs(Button.ClickEvent))
                Require(DirectCast(window.FindName("HarBodyNotice_txt"), TextBlock).Text.Contains("No captured"), "Empty preview filter retained request data.")
                DirectCast(window.FindName("ClearHarFilter_btn"), Button).RaiseEvent(New RoutedEventArgs(Button.ClickEvent))
                Require(DirectCast(window.FindName("HarBodyNotice_txt"), TextBlock).Text.Contains("redaction is enabled"), "Clearing HAR filters did not restore safe request preview.")
            Next
            Dim snapshot = DirectCast(type.GetMethod("BuildReportSnapshot", flags).Invoke(window, Nothing), ScanReportSnapshot)
            Require(snapshot.RedactionApplied AndAlso snapshot.ForExport(True, False).RedactionApplied, "Redaction state was lost in report snapshots.")
            For Each reportFormat As ScanReportFormat In {ScanReportFormat.Html, ScanReportFormat.Json, ScanReportFormat.Csv}
                Dim output = System.IO.Path.Combine(root, "report." & reportFormat.ToString())
                ScanReportWriter.SaveAsync(snapshot, output, reportFormat, True, False, CancellationToken.None).GetAwaiter().GetResult()
                Dim text = File.ReadAllText(output)
                For Each secret In {"BODY_SECRET", "HEADER_SECRET", "COOKIE_SECRET", "QUERY_SECRET", "FORM_SECRET"}
                    Require(Not text.Contains(secret), "Export leaked original HAR data: " & secret)
                Next
            Next
        Finally
            window.Close()
            Directory.Delete(root, True)
        End Try
    End Sub
    Private Sub Require(condition As Boolean, message As String)
        If Not condition Then Throw New InvalidOperationException(message)
    End Sub
End Module
