Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Threading

Namespace Beacon
    Public NotInheritable Class SearchSpan
        Public Property Start As Integer
        Public Property Length As Integer
    End Class

    Public NotInheritable Class SearchQuery
        Private ReadOnly _patterns As Regex()
        Public ReadOnly Property Mode As SearchMode
        Public ReadOnly Property Text As String
        Public ReadOnly Property CaseSensitive As Boolean
        Public Const MaximumHighlights As Integer = 10000

        Public Sub New(text As String, mode As SearchMode, caseSensitive As Boolean)
            If String.IsNullOrWhiteSpace(text) Then Throw New ArgumentException("Enter a search term.")
            If text.Length > 4096 Then Throw New ArgumentException("Search queries are limited to 4,096 characters.")
            If Not [Enum].IsDefined(mode) Then Throw New ArgumentException("Select a supported search mode.")
            Me.Text = text
            Me.Mode = mode
            Me.CaseSensitive = caseSensitive
            Dim terms = If(mode = SearchMode.AnyTerm OrElse mode = SearchMode.AllTerms, ParseTerms(text), New String() {text})
            Dim options = RegexOptions.CultureInvariant Or RegexOptions.Multiline
            If Not caseSensitive Then options = options Or RegexOptions.IgnoreCase
            _patterns = terms.Distinct(If(caseSensitive, StringComparer.Ordinal, StringComparer.OrdinalIgnoreCase)).
                Select(Function(term)
                           Dim pattern = If(mode = SearchMode.RegularExpression, term, Regex.Escape(term))
                           If mode = SearchMode.ExactWord Then pattern = "(?<!\w)" & pattern & "(?!\w)"
                           Return New Regex(pattern, options, TimeSpan.FromMilliseconds(100))
                       End Function).ToArray()
        End Sub

        Public Function IsMatch(record As String, Optional token As CancellationToken = Nothing) As Boolean
            If record Is Nothing Then Return False
            Dim found As Boolean = False
            For Each pattern In _patterns
                token.ThrowIfCancellationRequested()
                Dim matches = pattern.IsMatch(record)
                If Mode = SearchMode.AllTerms AndAlso Not matches Then Return False
                found = found OrElse matches
                If Mode <> SearchMode.AllTerms AndAlso found Then Return True
            Next
            Return found
        End Function

        Public Function FindHighlights(text As String, Optional limit As Integer = MaximumHighlights,
                                       Optional token As CancellationToken = Nothing) As List(Of SearchSpan)
            Dim spans As New List(Of SearchSpan)()
            Dim examined As Integer = 0
            If String.IsNullOrEmpty(text) OrElse limit <= 0 Then Return spans
            For Each pattern In _patterns
                token.ThrowIfCancellationRequested()
                Dim match = pattern.Match(text)
                While match.Success
                    token.ThrowIfCancellationRequested()
                    examined += 1
                    If match.Length > 0 Then spans.Add(New SearchSpan With {.Start = match.Index, .Length = match.Length})
                    If examined >= limit Then Exit While
                    match = match.NextMatch()
                End While
                If examined >= limit Then Exit For
            Next
            Dim merged As New List(Of SearchSpan)()
            For Each span In spans.OrderBy(Function(item) item.Start).ThenByDescending(Function(item) item.Length)
                If merged.Count > 0 AndAlso span.Start < merged.Last().Start + merged.Last().Length Then
                    Dim previous = merged.Last()
                    previous.Length = Math.Max(previous.Start + previous.Length, span.Start + span.Length) - previous.Start
                Else
                    merged.Add(span)
                End If
            Next
            Return merged
        End Function

        Private Shared Function ParseTerms(text As String) As String()
            Dim terms As New List(Of String)()
            Dim current As New StringBuilder()
            Dim quoted As Boolean = False
            For Each character In text
                If character = """"c Then
                    quoted = Not quoted
                ElseIf Char.IsWhiteSpace(character) AndAlso Not quoted Then
                    If current.Length > 0 Then
                        terms.Add(current.ToString())
                        current.Clear()
                    End If
                Else
                    current.Append(character)
                End If
            Next
            If quoted Then Throw New ArgumentException("Close the double quote around the search phrase.")
            If current.Length > 0 Then terms.Add(current.ToString())
            If terms.Count = 0 Then Throw New ArgumentException("Enter at least one nonempty term.")
            If terms.Count > 32 Then Throw New ArgumentException("Any/all searches support up to 32 terms.")
            Return terms.ToArray()
        End Function
    End Class

    Public NotInheritable Class SearchModeChoice
        Public Property Value As SearchMode
        Public Property Label As String
        Public Shared Function All() As SearchModeChoice()
            Return {New SearchModeChoice With {.Value = SearchMode.PlainText, .Label = "Literal text"},
                    New SearchModeChoice With {.Value = SearchMode.ExactWord, .Label = "Whole word"},
                    New SearchModeChoice With {.Value = SearchMode.RegularExpression, .Label = "Regular expression"},
                    New SearchModeChoice With {.Value = SearchMode.AnyTerm, .Label = "Any term"},
                    New SearchModeChoice With {.Value = SearchMode.AllTerms, .Label = "All terms"}}
        End Function
    End Class
End Namespace
