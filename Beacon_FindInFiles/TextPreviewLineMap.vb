Imports System.Threading

Namespace Beacon
    Friend NotInheritable Class TextPreviewLineMap
        Private Shared ReadOnly Breaks As Char() = {ControlChars.Cr, ControlChars.Lf}
        Private ReadOnly _lines As TextPreviewLine()
        Private ReadOnly _details As New Dictionary(Of SearchDetail, TextPreviewLine)()

        Public Sub New(text As String, details As IEnumerable(Of SearchDetail), token As CancellationToken)
            Dim requested = details.Where(Function(detail) Not detail.IsMetadata AndAlso detail.LineNumber > 0).
                OrderBy(Function(detail) detail.LineNumber).ToArray()
            Dim lines As New List(Of TextPreviewLine)()
            Dim request As Integer = 0
            Dim position As Integer = 0
            Dim number As Integer = 1
            While request < requested.Length AndAlso position < text.Length
                token.ThrowIfCancellationRequested()
                Dim finish = text.IndexOfAny(Breaks, position)
                If finish < 0 Then finish = text.Length
                While request < requested.Length AndAlso requested(request).LineNumber = number
                    Dim detail = requested(request)
                    Dim focus = detail.Context.FirstOrDefault(Function(item) item.IsMatchLine)
                    Dim matches = focus Is Nothing OrElse focus.Shortened OrElse focus.Text = text.Substring(position, finish - position)
                    Dim line As New TextPreviewLine(position, finish - position, detail, matches)
                    lines.Add(line)
                    _details(requested(request)) = line
                    request += 1
                End While
                position = finish
                If position < text.Length AndAlso text(position) = ControlChars.Cr Then position += 1
                If position < text.Length AndAlso text(position) = ControlChars.Lf Then position += 1
                number += 1
            End While
            _lines = lines.ToArray()
        End Sub

        Public Function ForDetail(detail As SearchDetail) As TextPreviewLine
            Dim line As TextPreviewLine = Nothing
            _details.TryGetValue(detail, line)
            Return line
        End Function

        Public Function AtPosition(position As Integer) As SearchDetail
            Dim low As Integer = 0
            Dim high = _lines.Length - 1
            While low <= high
                Dim middle = low + (high - low) \ 2
                Dim line = _lines(middle)
                If position < line.Start Then
                    high = middle - 1
                ElseIf position >= line.Start + line.Length Then
                    low = middle + 1
                Else
                    Return If(line.MatchesCapture, line.Detail, Nothing)
                End If
            End While
            Return Nothing
        End Function
    End Class

    Friend NotInheritable Class TextPreviewLine
        Public ReadOnly Start As Integer
        Public ReadOnly Length As Integer
        Public ReadOnly Detail As SearchDetail
        Public ReadOnly MatchesCapture As Boolean
        Public Sub New(start As Integer, length As Integer, detail As SearchDetail, matchesCapture As Boolean)
            Me.Start = start
            Me.Length = length
            Me.Detail = detail
            Me.MatchesCapture = matchesCapture
        End Sub
    End Class
End Namespace
