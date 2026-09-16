Imports System.IO
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Threading
Imports Beacon

Module ContextReportChecks
    Public Sub Run()
        Dim query As New SearchQuery("error", SearchMode.PlainText, False)
        For Each focus In {1, 4, 8}
            Dim lines = Enumerable.Range(1, 8).Select(Function(number) If(number = focus, "error at " & number, "context " & number)).ToArray()
            For Each firstOnly In {False, True}
                Using stream As New MemoryStream(Encoding.UTF8.GetBytes(String.Join(vbCrLf, lines)))
                    Dim result = SearchTextCollector.CollectAsync(stream, query, 20, firstOnly, False, CancellationToken.None).GetAwaiter().GetResult()
                    Dim detail = result.Details.Single()
                    Ensure(detail.Context.Count = 5 AndAlso detail.Context.Single(Function(line) line.IsMatchLine).LineNumber = focus, "Text context lost the matched line or five-line window.")
                    Dim expectedStart = If(focus = 1, 1, If(focus = 8, 4, 2))
                    Ensure(detail.Context.First().LineNumber = expectedStart AndAlso detail.Context.Last().LineNumber = expectedStart + 4, "Context did not adjust correctly at a file boundary.")
                    Ensure(detail.Context.Single(Function(line) line.IsMatchLine).Highlights.Count = 1, "Matching text was not captured for highlighting.")
                End Using
            Next
            Dim record = SearchDetail.FromRecord(query, String.Join(vbLf, lines), "Synthetic event/request")
            Ensure(record.Context.Count = 5 AndAlso record.Context.Single(Function(line) line.IsMatchLine).LineNumber = focus, "Record-local context is incorrect.")
        Next
        Using stream As New MemoryStream(Encoding.UTF8.GetBytes("before" & vbLf & "error"))
            Dim detail = SearchTextCollector.CollectAsync(stream, query, 1, True, False, CancellationToken.None).GetAwaiter().GetResult().Details.Single()
            Ensure(detail.Context.Count = 2, "Short files should not invent extra context lines.")
        End Using
        Dim longRecord = SearchDetail.FromRecord(query, New String("x"c, 1200) & " error " & New String("y"c, 1200), "Request")
        Ensure(longRecord.Context.Single().Shortened AndAlso longRecord.Context.Single().Text.Length <= MatchContext.CharacterLimit, "Long context lines were not bounded.")
        Ensure(longRecord.Context.Single().Highlights.Any(Function(span) longRecord.Context.Single().Text.Substring(span.Start, span.Length) = "error"), "Line shortening hid the match.")
        Using stream As New BoundedReadStream(New MemoryStream(Encoding.UTF8.GetBytes("error" & vbLf & New String("x"c, 12000))), 8192, CancellationToken.None)
            Dim result = SearchTextCollector.CollectAsync(stream, query, 1, True, False, CancellationToken.None).GetAwaiter().GetResult()
            Ensure(result.Details.Count = 1 AndAlso result.PartialReason.Contains("read"), "Context lookahead discarded an already-found match when a byte limit was reached.")
        End Using
        Dim root = Path.Combine(Path.GetTempPath(), "BeaconContextTests-" & Guid.NewGuid().ToString("N"))
        Directory.CreateDirectory(root)
        Try
            Dim source = Path.Combine(root, "context.log")
            File.WriteAllText(source, "CONTEXT_ONLY_SECRET" & vbLf & "before" & vbLf & "error <script>alert(1)</script>" & vbLf & "after one" & vbLf & "after two" & vbLf & "outside snippet")
            Dim captured As TextSearchResult
            Using input = File.OpenRead(source)
                captured = SearchTextCollector.CollectAsync(input, query, 1, True, False, CancellationToken.None).GetAwaiter().GetResult()
            End Using
            Dim detail = captured.Details.Single()
            Dim report As New ScanReportSnapshot With {.Run = New ScanRunInfo With {.SourceRoot = root, .QueryText = "error", .State = ScanReportState.Completed, .Options = New BeaconSettings()}}
            report.Results.Add(New ReportFile With {.DisplayName = "context.log", .SourcePath = source, .PartialReason = captured.PartialReason,
                .Matches = New List(Of ReportMatch) From {New ReportMatch With {.Location = detail.Location, .LineNumber = detail.LineNumber,
                    .Excerpt = detail.Excerpt, .ContextKind = detail.ContextKind, .Context = detail.Context.Select(Function(line) line.Copy()).ToList()}}})
            Dim detached = report.ForExport(True, False)
            report.Results(0).Matches(0).Context(0).Text = "changed"
            Ensure(detached.Results(0).Matches(0).Context(0).Text = "CONTEXT_ONLY_SECRET", "Report context was not detached.")
            Ensure(detached.ForExport(False, False).Results(0).Matches(0).Context.Count = 0, "Excerpt-free exports leaked context.")
            File.Delete(source)
            Dim output = Path.Combine(root, "results.html")
            ScanReportWriter.SaveAsync(detached, output, ScanReportFormat.Html, True, False, CancellationToken.None).GetAwaiter().GetResult()
            Dim html = File.ReadAllText(output)
            Ensure(html.Contains("context.log") AndAlso html.Contains("href='#file-0'"), "HTML report is not grouped by source file.")
            Ensure(Regex.Matches(html, "<tr(?: class='focus')?>").Count = 5, "HTML snippet does not contain exactly the captured five lines.")
            Ensure(html.Contains("<mark>error</mark>") AndAlso html.Contains("&lt;script&gt;") AndAlso Not html.Contains("<script>"), "HTML highlighting or encoding failed.")
            Ensure(Not html.Contains("outside snippet"), "Export reread or included data outside the context window.")
        Finally
            Directory.Delete(root, True)
        End Try
    End Sub

    Private Sub Ensure(condition As Boolean, message As String)
        If Not condition Then Throw New InvalidOperationException(message)
    End Sub
End Module
