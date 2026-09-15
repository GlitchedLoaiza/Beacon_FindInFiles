Imports System.IO
Imports System.Text.RegularExpressions
Imports System.Threading
Imports System.Threading.Tasks

Namespace Beacon
    Public NotInheritable Class SearchDetail
        Public Property Location As String
        Public Property Excerpt As String
        Public Property LineNumber As Integer
        Public Property RecordIndex As Integer = -1
        Public Property IsMetadata As Boolean
        Public Property VisibleMatches As Integer
        Public Property MatchOffset As Integer
        Public Property ContextKind As String = "Record text"
        Public Property Context As New List(Of ContextLine)()
        Public ReadOnly Property CountLabel As String
            Get
                Return If(VisibleMatches >= SearchQuery.MaximumHighlights, $"{VisibleMatches:N0}+", VisibleMatches.ToString("N0"))
            End Get
        End Property

        Public Shared Function FromRecord(query As SearchQuery, text As String, location As String,
                                          Optional lineNumber As Integer = 0, Optional recordIndex As Integer = -1,
                                          Optional metadata As Boolean = False, Optional token As CancellationToken = Nothing) As SearchDetail
            If Not query.IsMatch(text, token) Then Return Nothing
            Dim spans = query.FindHighlights(text, token:=token)
            Dim offset = If(spans.Count > 0, spans(0).Start, 0)
            Dim start = Math.Max(0, offset - 70)
            Dim excerpt = text.Substring(start, Math.Min(240, text.Length - start)).Replace(vbCr, " ").Replace(vbLf, " ").Replace(vbTab, " ")
            Return New SearchDetail With {
                .Location = location, .Excerpt = If(start > 0, "…", "") & excerpt & If(start + 240 < text.Length, "…", ""),
                .LineNumber = lineNumber, .RecordIndex = recordIndex, .IsMetadata = metadata,
                .VisibleMatches = spans.Count, .MatchOffset = offset,
                .ContextKind = If(metadata, "Name/path match", If(lineNumber > 0, "Source lines", "Record text")),
                .Context = MatchContext.FromText(text, spans, token, If(lineNumber > 0, lineNumber, 1))
            }
        End Function
    End Class

    Public NotInheritable Class TextSearchResult
        Public ReadOnly Property Details As New List(Of SearchDetail)()
        Public Property PartialReason As String = ""
    End Class

    Public NotInheritable Class SearchTextCollector
        Private Sub New()
        End Sub

        Public Shared Async Function CollectAsync(stream As Stream, query As SearchQuery, limit As Integer,
                                                   firstOnly As Boolean, document As Boolean, token As CancellationToken) As Task(Of TextSearchResult)
            Dim result As New TextSearchResult()
            Using reader As New StreamReader(stream, detectEncodingFromByteOrderMarks:=True, bufferSize:=4096, leaveOpen:=True)
                If document Then
                    Dim raw = Await reader.ReadToEndAsync(token)
                    Dim content = ExtractDocumentText(raw)
                    Dim detail = SearchDetail.FromRecord(query, content, "Document text", token:=token)
                    If detail IsNot Nothing Then
                        Dim readable = ExtractDocumentText(raw, normalizeWhitespace:=False)
                        Dim spans = query.FindHighlights(readable, token:=token)
                        detail.ContextKind = "Normalized document text"
                        If spans.Count > 0 Then
                            detail.Context = MatchContext.FromText(readable, spans, token)
                            detail.ContextKind = "Document text lines"
                        End If
                        result.Details.Add(detail)
                    End If
                    Return result
                End If
                Dim lineNumber As Integer = 0
                Dim history As New Queue(Of ContextLine)()
                Dim pending As New List(Of SearchDetail)()
                Dim stoppedMatching As Boolean = False
                Try
                    While True
                        token.ThrowIfCancellationRequested()
                        Dim line = Await reader.ReadLineAsync(token)
                        If line Is Nothing Then
                            For Each item In pending
                                item.Context = CopyHistory(history, item.LineNumber, MatchContext.LineLimit)
                            Next
                            Exit While
                        End If
                        lineNumber += 1
                        Dim detail As SearchDetail = Nothing
                        Dim captured As ContextLine
                        If Not stoppedMatching AndAlso query.IsMatch(line, token) Then
                            Dim spans = query.FindHighlights(line, token:=token)
                            Dim offset = If(spans.Count > 0, spans(0).Start, 0)
                            Dim start = Math.Max(0, offset - 70)
                            Dim excerpt = line.Substring(start, Math.Min(240, line.Length - start)).Replace(vbTab, " ")
                            detail = New SearchDetail With {
                                .Location = $"Line {lineNumber:N0}", .LineNumber = lineNumber,
                                .Excerpt = If(start > 0, "…", "") & excerpt & If(start + 240 < line.Length, "…", ""),
                                .VisibleMatches = spans.Count, .MatchOffset = offset, .ContextKind = "Source lines"}
                            ' ReadLine has removed line endings; capture once, then build independent file-context windows below.
                            captured = MatchContext.CaptureLine(line, lineNumber, spans)
                        Else
                            captured = MatchContext.CaptureLine(line, lineNumber)
                        End If
                        history.Enqueue(captured)
                        If history.Count > MatchContext.LineLimit Then history.Dequeue()
                        For index = pending.Count - 1 To 0 Step -1
                            Dim item = pending(index)
                            Dim following = CopyContextLine(captured, False)
                            item.Context.Add(following)
                            If item.Context.Count >= MatchContext.LineLimit Then pending.RemoveAt(index)
                        Next
                        If detail IsNot Nothing Then
                            detail.Context = CopyHistory(history, lineNumber, 3)
                            result.Details.Add(detail)
                            pending.Add(detail)
                            If firstOnly OrElse result.Details.Count >= limit Then
                                result.PartialReason = If(firstOnly, "first matching line only", "detail limit reached")
                                stoppedMatching = True
                            End If
                        End If
                        If stoppedMatching AndAlso pending.Count = 0 Then Exit While
                    End While
                Catch ex As Exception When result.Details.Count > 0 AndAlso (TypeOf ex Is IOException OrElse TypeOf ex Is InvalidDataException)
                    result.PartialReason = "read stopped; some context may be unavailable"
                End Try
            End Using
            Return result
        End Function

        Private Shared Function CopyHistory(history As Queue(Of ContextLine), focus As Integer, maximumLines As Integer) As List(Of ContextLine)
            Dim result As New List(Of ContextLine)(MatchContext.LineLimit)
            Dim skip = Math.Max(0, history.Count - maximumLines)
            For Each line In history
                If skip > 0 Then
                    skip -= 1
                Else
                    result.Add(CopyContextLine(line, line.LineNumber = focus))
                End If
            Next
            Return result
        End Function

        Private Shared Function CopyContextLine(line As ContextLine, isMatch As Boolean) As ContextLine
            Dim result As New ContextLine With {.LineNumber = line.LineNumber, .Text = line.Text,
                .Shortened = line.Shortened, .IsMatchLine = isMatch}
            For Each span In line.Highlights
                result.Highlights.Add(New SearchSpan With {.Start = span.Start, .Length = span.Length})
            Next
            Return result
        End Function

        Public Shared Function ExtractDocumentText(content As String, Optional normalizeWhitespace As Boolean = True) As String
            Dim timeout = TimeSpan.FromMilliseconds(100)
            Dim options = RegexOptions.Singleline Or RegexOptions.IgnoreCase Or RegexOptions.CultureInvariant
            content = Regex.Replace(content, "<(script|style)\b[^>]*>.*?</\1\s*>", " ", options, timeout)
            content = Regex.Replace(content, "<[^>]+>", " ", RegexOptions.None, timeout)
            content = System.Net.WebUtility.HtmlDecode(content)
            Return If(normalizeWhitespace, Regex.Replace(content, "\s+", " ", RegexOptions.None, timeout).Trim(), content.Trim())
        End Function
    End Class
End Namespace
