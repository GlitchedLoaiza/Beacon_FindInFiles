Imports System.Diagnostics
Imports System.IO
Imports System.IO.Compression
Imports System.Reflection
Imports System.Text
Imports System.Text.Json
Imports System.Threading
Imports System.Threading.Tasks
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Threading
Imports Beacon

Module PreviewAcceptanceChecks
    Private Const Flags As BindingFlags = BindingFlags.Instance Or BindingFlags.NonPublic

    Public Sub Directions()
        Fixture(Sub(root)
                    For Each newline In {vbCrLf, vbLf, vbCr}
                        Dim path = IO.Path.Combine(root, "unicode.log")
                        File.WriteAllText(path, "😀 café code=7 then code=12" & newline & "code=345" & newline & "日本語", New UTF8Encoding(True))
                        SearchChecks.CheckPipeline(path, "code=\d+", SearchMode.RegularExpression, New BeaconSettings(),
                            Sub(hits) Require(hits.Count = 1, "Direction fixture missing."),
                            Sub(window, hits)
                                DirectCast(window.FindName("Results_lst"), ListBox).SelectedItem = hits(0)
                                PreviewTestHelpers.WaitForPreview(window)
                                Dim preview = DirectCast(window.FindName("TextPreview_rtb"), RichTextBox)
                                Require(preview.Selection.Text = "code=7", "Initial UTF-16 selection was wrong.")
                                Click(window, "FindPrevious_btn")
                                Require(preview.Selection.Text = "code=345", "Previous did not wrap correctly with this newline encoding.")
                                Click(window, "FindNext_btn")
                                Require(preview.Selection.Text = "code=7", "Forward wrap was wrong.")
                                Click(window, "FindNext_btn")
                                Require(preview.Selection.Text = "code=12", "Variable-length selection was wrong.")
                                Click(window, "FindPrevious_btn")
                                Require(preview.Selection.Text = "code=7", "Reversing direction skipped a match.")
                                Dim capturedDetails = Details(hits(0))
                                DirectCast(window.FindName("Details_lst"), ListBox).SelectedItem = capturedDetails(1)
                                PreviewTestHelpers.WaitForPreview(window)
                                Require(preview.Selection.Text = "code=345", "Captured detail did not map to the source run.")
                            End Sub)
                    Next
                End Sub)
    End Sub

    Public Sub ArchivedSummary()
        Fixture(Sub(root)
                    Dim payload = IO.Path.Combine(root, "payload.log")
                    File.WriteAllText(payload, "before" & vbLf & "error one" & vbLf & "middle" & vbLf & "error two" & vbLf & "after")
                    Dim inner = IO.Path.Combine(root, "inner.zip")
                    Using archive = ZipFile.Open(inner, ZipArchiveMode.Create)
                        archive.CreateEntryFromFile(payload, "payload.log")
                    End Using
                    Dim outer = IO.Path.Combine(root, "outer.zip")
                    Using archive = ZipFile.Open(outer, ZipArchiveMode.Create)
                        archive.CreateEntryFromFile(inner, "inner.zip")
                    End Using
                    Dim cab = IO.Path.Combine(root, "payload.cab")
                    Dim start As New ProcessStartInfo(IO.Path.Combine(Environment.SystemDirectory, "makecab.exe")) With {.UseShellExecute = False, .CreateNoWindow = True}
                    start.ArgumentList.Add(payload)
                    start.ArgumentList.Add(cab)
                    Using child = Process.Start(start)
                        If Not child.WaitForExit(15000) Then
                            child.Kill(True)
                            Throw New TimeoutException("CAB fixture creation timed out.")
                        End If
                        Require(child.ExitCode = 0, "CAB fixture creation failed.")
                    End Using
                    For Each source In {outer, cab}
                        SearchChecks.CheckPipeline(source, "error", SearchMode.PlainText, New BeaconSettings(),
                            Sub(hits) Require(hits.Count = 1, "Archived Summary fixture missing."),
                            Sub(window, hits)
                                Dim hit = hits(0)
                                Dim capturedDetails = Details(hit)
                                Dim before = JsonSerializer.Serialize(capturedDetails)
                                Dim physical = CStr(hit.GetType().GetProperty("FilePath").GetValue(hit))
                                Dim archive = CStr(hit.GetType().GetProperty("ZipPath").GetValue(hit))
                                If Not String.IsNullOrEmpty(physical) AndAlso File.Exists(physical) Then File.Delete(physical)
                                If Not String.IsNullOrEmpty(archive) AndAlso File.Exists(archive) Then File.Delete(archive)
                                If File.Exists(source) Then File.Delete(source)
                                GetType(MainWindow).GetField("_summaryPreferred", Flags).SetValue(window, True)
                                DirectCast(window.FindName("Results_lst"), ListBox).SelectedItem = hit
                                PreviewTestHelpers.WaitForPreview(window)
                                Dim summary = DirectCast(window.FindName("SummaryPreview_lst"), ListBox)
                                Require(summary.IsVisible AndAlso summary.Items.Count = capturedDetails.Count, "Summary tried to reopen a removed archive or extracted source.")
                                Click(window, "FindNext_btn")
                                Require(DirectCast(summary.SelectedItem, TextSummaryBlock).DetailIndex = 1, "Archived Summary lost record ordering.")
                                Click(window, "FindPrevious_btn")
                                Require(DirectCast(summary.SelectedItem, TextSummaryBlock).DetailIndex = 0, "Archived Summary reverse navigation failed.")
                                Require(JsonSerializer.Serialize(capturedDetails) = before, "Archived Summary changed captured/export data.")
                            End Sub, expectedCount:=1)
                    Next
                End Sub)
    End Sub

    Public Sub LimitsAndMetadata()
        Fixture(Sub(root)
                    Dim path = IO.Path.Combine(root, "preview.log")
                    File.WriteAllText(path, String.Join(vbLf, Enumerable.Repeat(New String("x"c, 512), 2200)) & vbLf & "preview marker")
                    Dim options As New BeaconSettings With {.MaximumPreviewSizeMb = 1, .SearchFileNames = True}
                    SearchChecks.CheckPipeline(path, "preview", SearchMode.PlainText, options,
                        Sub(hits) Require(hits.Count = 1 AndAlso Details(hits(0)).Count = 2, "Mixed metadata/content fixture missing."),
                        Sub(window, hits)
                            DirectCast(window.FindName("Results_lst"), ListBox).SelectedItem = hits(0)
                            PreviewTestHelpers.WaitForPreview(window)
                            Dim nextMatch = DirectCast(window.FindName("FindNext_btn"), Button)
                            Require(Not nextMatch.IsEnabled, "The truncation notice became a fabricated source match.")
                            Dim mode = DirectCast(window.FindName("TextPreviewMode_cmb"), ComboBox)
                            mode.SelectedIndex = 1
                            PreviewTestHelpers.WaitForPreview(window)
                            Dim summary = DirectCast(window.FindName("SummaryPreview_lst"), ListBox)
                            Dim selected = DirectCast(summary.SelectedItem, TextSummaryBlock)
                            Require(selected.DetailIndex = 1 AndAlso selected.Focus.LineNumber = 2201 AndAlso nextMatch.IsEnabled,
                                    "Summary lost a captured match beyond the Full preview limit.")
                            summary.SelectedIndex = 0
                            PreviewTestHelpers.WaitForPreview(window)
                            Require(Not summary.IsVisible AndAlso DirectCast(window.FindName("TextPreviewModePanel"), FrameworkElement).Visibility = Visibility.Collapsed,
                                    "A metadata entry was displayed as numbered file context.")
                            Dim capturedDetails = Details(hits(0))
                            DirectCast(window.FindName("Details_lst"), ListBox).SelectedItem = capturedDetails(1)
                            PreviewTestHelpers.WaitForPreview(window)
                            Require(summary.IsVisible, "Content detail selection did not restore the session's Summary mode.")
                            mode.SelectedIndex = 0
                            PreviewTestHelpers.WaitForPreview(window)
                            Require(Not nextMatch.IsEnabled AndAlso DirectCast(window.FindName("Status_lbl"), TextBlock).Text.Contains("outside the loaded preview"),
                                    "Full view pretended to display a match outside its loaded prefix.")
                        End Sub)
                    File.WriteAllText(path, "before" & vbLf & "error" & vbLf & "after")
                    SearchChecks.CheckPipeline(path, "(?=error)", SearchMode.RegularExpression, New BeaconSettings(),
                        Sub(hits) Require(hits.Count = 1, "Zero-width fixture missing."),
                        Sub(window, hits)
                            GetType(MainWindow).GetField("_summaryPreferred", Flags).SetValue(window, True)
                            DirectCast(window.FindName("Results_lst"), ListBox).SelectedItem = hits(0)
                            PreviewTestHelpers.WaitForPreview(window)
                            Click(window, "FindPrevious_btn")
                            Click(window, "FindNext_btn")
                            Dim block = DirectCast(DirectCast(window.FindName("SummaryPreview_lst"), ListBox).SelectedItem, TextSummaryBlock)
                            Require(block.Focus.LineNumber = 2 AndAlso block.CurrentSpan Is Nothing AndAlso block.IsCurrent, "Zero-width Summary anchor was lost or fabricated a highlight.")
                        End Sub)
                End Sub)
    End Sub

    Public Sub StaleLoads()
        Fixture(Sub(root)
                    File.WriteAllText(IO.Path.Combine(root, "one.log"), "error one")
                    File.WriteAllText(IO.Path.Combine(root, "two.log"), "error two")
                    SearchChecks.CheckPipeline(root, "error", SearchMode.PlainText, New BeaconSettings(),
                        Sub(hits) Require(hits.Count = 2, "Stale-preview fixture missing."),
                        Sub(window, hits)
                            Dim results = DirectCast(window.FindName("Results_lst"), ListBox)
                            Dim expected = File.ReadAllText(CStr(hits(1).GetType().GetProperty("FilePath").GetValue(hits(1))))
                            results.SelectedItem = hits(0)
                            PreviewTestHelpers.WaitForPreview(window)
                            Dim entered As New TaskCompletionSource(Of Boolean)(TaskCreationOptions.RunContinuationsAsynchronously)
                            Dim release As New TaskCompletionSource(Of Boolean)(TaskCreationOptions.RunContinuationsAsynchronously)
                            Dim reader As Func(Of PreviewContentService, CancellationToken, PreviewText) =
                                Function(service, token)
                                    entered.TrySetResult(True)
                                    release.Task.GetAwaiter().GetResult()
                                    Return New PreviewText("STALE error", False, 100)
                                End Function
                            Try
                                GetType(MainWindow).GetMethod("StartPlainTextPreview", Flags).Invoke(window, {reader})
                                Wait(window, entered.Task)
                                results.SelectedItem = hits(1)
                                PreviewTestHelpers.WaitForPreview(window)
                                release.TrySetResult(True)
                                Dim drain = DirectCast(GetType(MainWindow).GetMethod("DrainPreviewOperationsAsync", Flags).Invoke(window, Nothing), Task)
                                Wait(window, drain)
                                Dim preview = DirectCast(window.FindName("TextPreview_rtb"), RichTextBox)
                                Dim text = New Documents.TextRange(preview.Document.ContentStart, preview.Document.ContentEnd).Text
                                Require(text.Contains(expected) AndAlso Not text.Contains("STALE"), "An obsolete reader overwrote a newer preview.")
                            Finally
                                release.TrySetResult(True)
                            End Try
                        End Sub)
                End Sub)
    End Sub

    Private Sub Click(window As MainWindow, name As String)
        DirectCast(window.FindName(name), Button).RaiseEvent(New RoutedEventArgs(Button.ClickEvent))
        PreviewTestHelpers.WaitForPreview(window)
    End Sub

    Private Sub Wait(window As MainWindow, task As Task)
        Dim timeout = Stopwatch.StartNew()
        While Not task.IsCompleted
            window.Dispatcher.Invoke(Sub()
                                     End Sub, DispatcherPriority.ContextIdle)
            If timeout.Elapsed > TimeSpan.FromSeconds(20) Then Throw New TimeoutException("Preview test did not make progress.")
            Thread.Sleep(1)
        End While
        task.GetAwaiter().GetResult()
        window.Dispatcher.Invoke(Sub()
                                 End Sub, DispatcherPriority.ContextIdle)
    End Sub

    Private Function Details(hit As Object) As List(Of SearchDetail)
        Return DirectCast(hit.GetType().GetProperty("Details").GetValue(hit), List(Of SearchDetail))
    End Function

    Private Sub Fixture(action As Action(Of String))
        Dim root = IO.Path.Combine(IO.Path.GetTempPath(), "BeaconPreviewAcceptance-" & Guid.NewGuid().ToString("N"))
        Directory.CreateDirectory(root)
        Try
            action(root)
        Finally
            Directory.Delete(root, True)
        End Try
    End Sub

    Private Sub Require(value As Boolean, message As String)
        If Not value Then Throw New InvalidOperationException(message)
    End Sub
End Module
