Imports System.Threading
Imports Beacon

Module TextNavigationChecks
    Public Sub Run()
        Dim text = "😀 code=7" & vbCrLf & "code=123 then code=99"
        Dim query As New SearchQuery("code=\d+", SearchMode.RegularExpression, False)
        Dim ranges = query.FindHighlights(text)
        Dim navigation As New TextPreviewNavigation(text.Length, ranges)
        ranges(0).Start = 999
        Require(navigation.Move(True) AndAlso ReadMatch(text, navigation) = "code=7", "First match or detached snapshot failed.")
        Require(navigation.Move(True) AndAlso ReadMatch(text, navigation) = "code=123", "Variable-length next match failed.")
        Require(navigation.Move(False) AndAlso ReadMatch(text, navigation) = "code=7", "Changing direction selected the same match.")
        Require(navigation.Move(False) AndAlso navigation.Wrapped AndAlso ReadMatch(text, navigation) = "code=99", "Previous did not wrap to the last match.")
        Require(navigation.Move(True) AndAlso navigation.Wrapped AndAlso ReadMatch(text, navigation) = "code=7", "Next did not wrap to the first match.")
        Require(navigation.SelectAtOrAfter(text.IndexOf("code=123", StringComparison.Ordinal)) AndAlso ReadMatch(text, navigation) = "code=123", "Detail-offset selection failed.")
        Dim copy = navigation.Current
        copy.Start = 999
        Require(ReadMatch(text, navigation) = "code=123", "Returning a match exposed mutable cursor state.")
        Dim adjacent As New TextPreviewNavigation(2, New SearchQuery("a", SearchMode.PlainText, False).FindHighlights("aa"))
        adjacent.Move(True)
        adjacent.Move(True)
        adjacent.Move(False)
        Require(adjacent.Current.Start = 0 AndAlso Not adjacent.Wrapped, "Adjacent matches were skipped.")
        Dim singleMatch As New TextPreviewNavigation(1, {New SearchSpan With {.Start = 0, .Length = 1}})
        singleMatch.Move(True)
        Require(singleMatch.Move(False) AndAlso singleMatch.Wrapped AndAlso singleMatch.Index = 0, "Single-match wrap failed.")
        For Each value In {"", "error"}
            Dim zero As New TextPreviewNavigation(value.Length, New SearchQuery("(?=error)", SearchMode.RegularExpression, False).FindHighlights(value))
            Require(Not zero.Move(True) AndAlso Not zero.Move(False), "Zero-width navigation did not terminate.")
        Next
        Dim malformed As New TextPreviewNavigation(5, {New SearchSpan With {.Start = -1, .Length = 1}, New SearchSpan With {.Start = 4, .Length = Integer.MaxValue}})
        Require(malformed.Count = 0, "Invalid ranges escaped bounds checks.")
        Dim overlap As New TextPreviewNavigation(5, New SearchQuery("err error", SearchMode.AnyTerm, False).FindHighlights("error"))
        Require(overlap.Count = 1, "Overlapping query ranges became duplicate navigation targets.")
        For Each newline In {vbCrLf, vbLf, vbCr}
            Dim source = "first" & newline & "code=12" & newline & "last"
            Dim detail As New SearchDetail With {.LineNumber = 2, .Location = "Line 2"}
            Dim map As New TextPreviewLineMap(source, {detail}, CancellationToken.None)
            Dim location = map.ForDetail(detail)
            Require(location IsNot Nothing AndAlso source.Substring(location.Start, location.Length) = "code=12" AndAlso map.AtPosition(location.Start) Is detail,
                    "Captured-line mapping confused source lines with rendered positions.")
            Require(New TextPreviewLineMap("first", {detail}, CancellationToken.None).ForDetail(detail) Is Nothing, "A truncated prefix invented a later line.")
        Next
        For Each mode In {SearchMode.PlainText, SearchMode.ExactWord, SearchMode.RegularExpression, SearchMode.AnyTerm, SearchMode.AllTerms}
            Dim sample = "error code error code"
            Dim term = If(mode = SearchMode.AnyTerm OrElse mode = SearchMode.AllTerms, "error code", "error")
            Dim cursor As New TextPreviewNavigation(sample.Length, New SearchQuery(term, mode, False).FindHighlights(sample))
            Require(cursor.Move(True) AndAlso cursor.Move(False) AndAlso cursor.Index = cursor.Count - 1, "A query mode lost reverse navigation.")
        Next
        Using cancelled As New CancellationTokenSource()
            cancelled.Cancel()
            Try
                query.FindHighlights(text, token:=cancelled.Token)
                Throw New InvalidOperationException("Highlight preparation ignored cancellation.")
            Catch ex As OperationCanceledException
            End Try
        End Using
    End Sub

    Private Function ReadMatch(text As String, navigation As TextPreviewNavigation) As String
        Dim span = navigation.Current
        Return text.Substring(span.Start, span.Length)
    End Function

    Private Sub Require(value As Boolean, message As String)
        If Not value Then Throw New InvalidOperationException(message)
    End Sub
End Module
