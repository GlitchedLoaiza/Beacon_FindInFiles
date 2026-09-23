Imports System.IO
Imports System.Text
Imports System.Text.Json
Imports System.Threading
Imports System.Reflection
Imports System.Windows
Imports System.Windows.Controls
Imports Beacon

Module TextSummaryChecks
    Public Sub Projection()
        Dim query As New SearchQuery("code=\d+", SearchMode.RegularExpression, False)
        Dim text = String.Join(vbCrLf, Enumerable.Range(1, 8).Select(Function(number) $"code={number} then code={number + 100}"))
        Dim captured As TextSearchResult
        Using stream As New MemoryStream(Encoding.UTF8.GetBytes(text))
            captured = SearchTextCollector.CollectAsync(stream, query, 100, False, False, CancellationToken.None).GetAwaiter().GetResult()
        End Using
        Dim before = JsonSerializer.Serialize(captured.Details)
        Dim projection As New TextSummaryProjection(captured.Details, "captured limit")
        Require(projection.Blocks.Count = 8 AndAlso projection.ContentCount = 8 AndAlso projection.PartialReason = "captured limit", "Summary changed record counts or coverage.")
        Require(projection.NavigationTargetCount = 16, "Repeated neighboring context inflated the Summary match total.")
        For index = 0 To projection.Blocks.Count - 1
            Dim block = projection.Blocks(index)
            Dim detail = captured.Details(index)
            Require(block.DetailIndex = index AndAlso block.Lines.Count = 5 AndAlso
                    block.Lines.Select(Function(line) line.LineNumber).SequenceEqual(detail.Context.Select(Function(line) line.LineNumber)),
                    "Summary merged or changed an exported context window.")
            Require(block.Focus.LineNumber = detail.LineNumber, "Summary focus moved to neighboring context.")
            Require(block.NavigationOffset = CLng(index) * 2 AndAlso block.NavigationTargetCount = 2, "Summary counter offsets differ from navigation targets.")
            block.SelectEdge(True)
            Require(ReadSpan(block) = $"code={index + 1}", "Summary did not select the record's first occurrence.")
            Require(block.MoveWithin(True) AndAlso ReadSpan(block) = $"code={index + 101}", "Summary lost a second occurrence on the focus line.")
            Require(Not block.MoveWithin(True), "Neighbor context became a duplicate navigation target.")
            Require(block.MoveWithin(False) AndAlso ReadSpan(block) = $"code={index + 1}", "Summary reverse navigation failed.")
            block.Deactivate()
            Require(block.CurrentSpan Is Nothing, "An inactive block retained its navigation index.")
        Next
        Require(JsonSerializer.Serialize(captured.Details) = before, "Summary navigation mutated captured/exported data.")
        Using stream As New MemoryStream(Encoding.UTF8.GetBytes("before" & vbLf & "error" & vbLf & "after"))
            Dim result = SearchTextCollector.CollectAsync(stream, New SearchQuery("(?=error)", SearchMode.RegularExpression, False), 10, False, False, CancellationToken.None).GetAwaiter().GetResult()
            Dim block = New TextSummaryProjection(result.Details, "").Blocks(0)
            block.SelectEdge(True)
            Require(block.Lines.Count = 3 AndAlso block.Focus.LineNumber = 2 AndAlso block.CurrentSpan Is Nothing AndAlso Not block.MoveWithin(True),
                    "A zero-width match lost its record anchor or invented a visible match.")
            Require(block.NavigationTargetCount = 1, "A zero-width record lost its counter anchor.")
        End Using
        Dim longDetail = SearchDetail.FromRecord(New SearchQuery("error", SearchMode.PlainText, False), New String("x"c, 1200) & " error 😀 " & New String("y"c, 1200), "Line 1", lineNumber:=1)
        Dim shortened = New TextSummaryProjection({longDetail}, "").Blocks(0)
        Require(shortened.IsShortened AndAlso shortened.Lines(0).Text.Length <= MatchContext.CharacterLimit, "Summary removed line-length safeguards.")
        Dim fallback As New TextSummaryProjection({New SearchDetail With {.Location = "Line 12", .LineNumber = 12, .Excerpt = "captured error"}}, "")
        Require(fallback.Blocks(0).Lines.Count = 0 AndAlso fallback.Blocks(0).ToString().Contains("captured error"), "Missing context did not preserve its excerpt.")
        Dim metadata As New TextSummaryProjection({New SearchDetail With {.IsMetadata = True, .Location = "File name", .Excerpt = "error.log"}}, "")
        Require(metadata.ContentCount = 0 AndAlso metadata.Blocks(0).IsMetadata, "Metadata was treated as file content.")
        Require(metadata.NavigationTargetCount = 0 AndAlso metadata.Blocks(0).NavigationTargetCount = 0, "Metadata inflated the Summary navigation count.")
    End Sub

    Public Sub Preview()
        Dim root = Path.Combine(Path.GetTempPath(), "BeaconSummaryUi-" & Guid.NewGuid().ToString("N"))
        Directory.CreateDirectory(root)
        Try
            Dim path = IO.Path.Combine(root, "sample.log")
            File.WriteAllText(path, "before" & vbCrLf & "code=7 then code=12" & vbCrLf & "middle" & vbCrLf & "code=99" & vbCrLf & "after")
            SearchChecks.CheckPipeline(path, "code=\d+", SearchMode.RegularExpression, New BeaconSettings(),
                Sub(hits) Require(hits.Count = 1, "Summary fixture did not match."),
                Sub(window, hits)
                    Dim results = DirectCast(window.FindName("Results_lst"), ListBox)
                    Dim mode = DirectCast(window.FindName("TextPreviewMode_cmb"), ComboBox)
                    Dim summary = DirectCast(window.FindName("SummaryPreview_lst"), ListBox)
                    Dim textPreview = DirectCast(window.FindName("TextPreview_rtb"), RichTextBox)
                    Require(mode.SelectedIndex = 0, "A fresh application did not start in Full.")
                    results.SelectedItem = hits(0)
                    PreviewTestHelpers.WaitForPreview(window)
                    Require(mode.IsVisible AndAlso DirectCast(window.FindName("TextPreviewModePanel"), FrameworkElement).IsVisible,
                            "The View Full/Summary selector is missing for a selected plain-text result.")
                    Require(textPreview.Selection.Text = "code=7", "Full preview did not select its first match.")
                    File.Delete(path)
                    mode.SelectedIndex = 1
                    PreviewTestHelpers.WaitForPreview(window)
                    Require(summary.IsVisible AndAlso Not textPreview.IsVisible AndAlso summary.Items.Count = 2, "Summary did not render separate captured records without the source.")
                    Dim nextMatch = DirectCast(window.FindName("FindNext_btn"), Button)
                    Dim previous = DirectCast(window.FindName("FindPrevious_btn"), Button)
                    Require(previous.IsVisible AndAlso previous.IsEnabled, "Summary reverse navigation is unavailable.")
                    Require(ReadSpan(DirectCast(summary.SelectedItem, TextSummaryBlock)) = "code=7", "Summary did not start at the first occurrence.")
                    nextMatch.RaiseEvent(New RoutedEventArgs(Button.ClickEvent))
                    PreviewTestHelpers.WaitForPreview(window)
                    Require(ReadSpan(DirectCast(summary.SelectedItem, TextSummaryBlock)) = "code=12", "Summary skipped the second focus-line occurrence.")
                    nextMatch.RaiseEvent(New RoutedEventArgs(Button.ClickEvent))
                    PreviewTestHelpers.WaitForPreview(window)
                    Require(ReadSpan(DirectCast(summary.SelectedItem, TextSummaryBlock)) = "code=99", "Summary navigated duplicate neighboring context.")
                    previous.RaiseEvent(New RoutedEventArgs(Button.ClickEvent))
                    PreviewTestHelpers.WaitForPreview(window)
                    Require(ReadSpan(DirectCast(summary.SelectedItem, TextSummaryBlock)) = "code=12", "Summary previous did not select the last occurrence of the previous record.")
                    previous.RaiseEvent(New RoutedEventArgs(Button.ClickEvent))
                    previous.RaiseEvent(New RoutedEventArgs(Button.ClickEvent))
                    PreviewTestHelpers.WaitForPreview(window)
                    Require(ReadSpan(DirectCast(summary.SelectedItem, TextSummaryBlock)) = "code=99", "Summary previous did not wrap.")
                    mode.SelectedIndex = 0
                    PreviewTestHelpers.WaitForPreview(window)
                    Require(textPreview.IsVisible AndAlso textPreview.Selection.Text = "code=99", "Returning to cached Full lost the captured record or reread a deleted source.")
                    mode.SelectedIndex = 1
                    PreviewTestHelpers.WaitForPreview(window)
                    Dim flags = BindingFlags.Instance Or BindingFlags.NonPublic
                    GetType(MainWindow).GetMethod("Reset_btn_Click", flags).Invoke(window, {Nothing, New RoutedEventArgs()})
                    PreviewTestHelpers.WaitForPreview(window)
                    Require(CBool(GetType(MainWindow).GetField("_summaryPreferred", flags).GetValue(window)), "Reset cleared the session-only Summary choice.")
                    Require(results.Items.Count = 0, "Reset retained summary/search results.")
                End Sub)
        Finally
            Directory.Delete(root, True)
        End Try
    End Sub

    Private Function ReadSpan(block As TextSummaryBlock) As String
        Dim span = block.CurrentSpan
        Return block.Focus.Text.Substring(span.Start, span.Length)
    End Function

    Private Sub Require(value As Boolean, message As String)
        If Not value Then Throw New InvalidOperationException(message)
    End Sub
End Module
