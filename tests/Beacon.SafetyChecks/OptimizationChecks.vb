Imports System.IO
Imports System.Text
Imports System.Threading
Imports Beacon

Module OptimizationChecks
    Public Sub TextContextParity()
        For Each mode In {SearchMode.PlainText, SearchMode.ExactWord, SearchMode.RegularExpression, SearchMode.AnyTerm, SearchMode.AllTerms}
            Dim term = If(mode = SearchMode.RegularExpression, "err(or)?", If(mode = SearchMode.AnyTerm OrElse mode = SearchMode.AllTerms, "error code", "error"))
            Dim query As New SearchQuery(term, mode, False)
            Dim lines = Enumerable.Range(1, 8).Select(Function(number) $"error code={number}").ToArray()
            Using stream As New MemoryStream(Encoding.UTF8.GetBytes(String.Join(vbCrLf, lines)))
                Dim result = SearchTextCollector.CollectAsync(stream, query, 100, False, False, CancellationToken.None).GetAwaiter().GetResult()
                Require(result.Details.Count = 8 AndAlso result.PartialReason = "", "Dense match collection changed.")
                For index = 0 To 7
                    Dim detail = result.Details(index)
                    Dim expected = SearchDetail.FromRecord(query, lines(index), $"Line {index + 1:N0}", index + 1)
                    Require(detail.Excerpt = expected.Excerpt AndAlso detail.MatchOffset = expected.MatchOffset AndAlso
                            detail.VisibleMatches = expected.VisibleMatches AndAlso detail.ContextKind = expected.ContextKind,
                            "Line-specialized detail differs from the general record helper.")
                    Require(detail.Location = $"Line {index + 1:N0}", "Formatted matching-line label changed.")
                    Dim start = Math.Clamp(index - 2, 0, 3)
                    Require(detail.Context.Select(Function(line) line.LineNumber).SequenceEqual(Enumerable.Range(start + 1, 5)), "Overlapping or EOF context order changed.")
                    Require(detail.Context.Single(Function(line) line.IsMatchLine).LineNumber = index + 1, "Focus moved to another match.")
                    For Each line In detail.Context
                        Require(line.Text = lines(line.LineNumber - 1) AndAlso line.Highlights.Count > 0, "Neighboring matching-line text or highlights changed.")
                    Next
                Next
                result.Details(0).Context(1).Text = "changed"
                result.Details(0).Context(1).Highlights(0).Start = 999
                Require(result.Details(1).Context(1).Text = lines(1) AndAlso result.Details(1).Context(1).Highlights(0).Start <> 999, "Result snippets share mutable context state.")
            End Using
        Next
        Using stream As New MemoryStream(Encoding.UTF8.GetBytes("before" & vbLf & "error" & vbLf & "after"))
            Dim query As New SearchQuery("(?=error)", SearchMode.RegularExpression, False)
            Dim detail = SearchTextCollector.CollectAsync(stream, query, 1, True, False, CancellationToken.None).GetAwaiter().GetResult().Details.Single()
            Require(detail.VisibleMatches = 0 AndAlso detail.Location = "Line 2" AndAlso detail.Context.Count = 3, "Zero-width matching or first-match context changed.")
        End Using
        For Each line In {"error" & vbTab & "code", New String("x"c, 1200) & " error 😀 " & New String("y"c, 600), ""}
            Dim query As New SearchQuery(If(line = "", "^$", "error"), SearchMode.RegularExpression, False)
            Using stream As New MemoryStream(Encoding.UTF8.GetBytes(line & vbLf))
                Dim result = SearchTextCollector.CollectAsync(stream, query, 10, False, False, CancellationToken.None).GetAwaiter().GetResult()
                Dim expected = SearchDetail.FromRecord(query, line, "Line 1", 1)
                Dim actual = result.Details.Single()
                Require(actual.Excerpt = expected.Excerpt AndAlso actual.MatchOffset = expected.MatchOffset AndAlso actual.VisibleMatches = expected.VisibleMatches,
                        "Long, tabbed or empty matching line changed detail semantics.")
                If line.Length > 0 Then
                    Require(actual.Context.Single().Text = expected.Context.Single().Text AndAlso actual.Context.Single().Shortened = expected.Context.Single().Shortened,
                            "Long-line context cropping changed.")
                End If
            End Using
        Next
    End Sub

    Private Sub Require(value As Boolean, message As String)
        If Not value Then Throw New InvalidOperationException(message)
    End Sub
End Module
