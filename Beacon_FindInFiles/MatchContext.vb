Imports System.IO
Imports System.Threading

Namespace Beacon
    Public NotInheritable Class ContextLine
        Public Property LineNumber As Integer
        Public Property Text As String
        Public Property IsMatchLine As Boolean
        Public Property Shortened As Boolean
        Public Property Highlights As New List(Of SearchSpan)()

        Public Function Copy() As ContextLine
            Return New ContextLine With {.LineNumber = LineNumber, .Text = Text, .IsMatchLine = IsMatchLine, .Shortened = Shortened,
                .Highlights = Highlights.Select(Function(span) New SearchSpan With {.Start = span.Start, .Length = span.Length}).ToList()}
        End Function
    End Class

    Public NotInheritable Class MatchContext
        Public Const LineLimit As Integer = 5
        Public Const CharacterLimit As Integer = 500
        Private Sub New()
        End Sub

        Public Shared Function CaptureLine(text As String, number As Integer, Optional highlights As IEnumerable(Of SearchSpan) = Nothing) As ContextLine
            Dim spans = If(highlights, Enumerable.Empty(Of SearchSpan)()).ToList()
            Dim start = If(text.Length > CharacterLimit AndAlso spans.Count > 0, Math.Max(0, spans(0).Start - 70), 0)
            If start > 0 AndAlso Char.IsLowSurrogate(text(start)) AndAlso Char.IsHighSurrogate(text(start - 1)) Then start -= 1
            Dim length = If(text.Length <= CharacterLimit, text.Length, Math.Min(CharacterLimit - 2, text.Length - start))
            If length > 0 AndAlso start + length < text.Length AndAlso Char.IsHighSurrogate(text(start + length - 1)) Then length -= 1
            Dim prefix = If(start > 0, "…", "")
            Dim result As New ContextLine With {.LineNumber = number, .Shortened = start > 0 OrElse length < text.Length,
                .Text = prefix & text.Substring(start, length) & If(start + length < text.Length, "…", "")}
            For Each span In spans
                Dim first = Math.Max(start, span.Start)
                Dim last = Math.Min(start + length, span.Start + span.Length)
                If last > first Then result.Highlights.Add(New SearchSpan With {.Start = first - start + prefix.Length, .Length = last - first})
            Next
            Return result
        End Function

        Public Shared Function WithFocus(lines As IEnumerable(Of ContextLine), focus As Integer) As List(Of ContextLine)
            Return lines.TakeLast(LineLimit).Select(Function(line)
                                                       Dim copy = line.Copy()
                                                       copy.IsMatchLine = copy.LineNumber = focus
                                                       Return copy
                                                   End Function).ToList()
        End Function

        Public Shared Function FromText(text As String, spans As List(Of SearchSpan), token As CancellationToken,
                                        Optional firstLineNumber As Integer = 1) As List(Of ContextLine)
            Dim history As New Queue(Of ContextLine)()
            Dim result As New List(Of ContextLine)()
            Dim focusOffset = If(spans.Count > 0, spans(0).Start, 0)
            Dim focusLine As Integer = 0
            Dim offset As Integer = 0
            Dim number = firstLineNumber - 1
            Using reader As New StringReader(text)
                While True
                    token.ThrowIfCancellationRequested()
                    Dim line = reader.ReadLine()
                    If line Is Nothing Then Exit While
                    number += 1
                    Dim local As New List(Of SearchSpan)()
                    If offset + line.Length >= focusOffset Then
                        For Each span In spans
                            If span.Start >= offset + line.Length Then Exit For
                            Dim first = Math.Max(offset, span.Start)
                            Dim last = Math.Min(offset + line.Length, span.Start + span.Length)
                            If last > first Then local.Add(New SearchSpan With {.Start = first - offset, .Length = last - first})
                        Next
                    End If
                    Dim captured = CaptureLine(line, number, local)
                    history.Enqueue(captured)
                    If history.Count > LineLimit Then history.Dequeue()
                    If focusLine > 0 Then
                        result.Add(captured.Copy())
                    ElseIf focusOffset <= offset + line.Length Then
                        focusLine = number
                        result = history.TakeLast(3).Select(Function(item) item.Copy()).ToList()
                    End If
                    If result.Count >= LineLimit Then Return WithFocus(result, focusLine)
                    offset += line.Length
                    If offset < text.Length AndAlso text(offset) = ControlChars.Cr Then offset += 1
                    If offset < text.Length AndAlso text(offset) = ControlChars.Lf Then offset += 1
                End While
            End Using
            Return WithFocus(history, If(focusLine > 0, focusLine, firstLineNumber))
        End Function
    End Class
End Namespace
