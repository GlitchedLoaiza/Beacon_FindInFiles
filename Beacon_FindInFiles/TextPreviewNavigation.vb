Namespace Beacon
    Friend NotInheritable Class TextPreviewNavigation
        Private ReadOnly _ranges As SearchSpan()
        Private _index As Integer = -1
        Public ReadOnly Property Wrapped As Boolean

        Public Sub New(textLength As Integer, ranges As IEnumerable(Of SearchSpan))
            _ranges = ValidRanges(textLength, ranges).OrderBy(Function(span) span.Start).
                Select(Function(span) New SearchSpan With {.Start = span.Start, .Length = span.Length}).ToArray()
        End Sub

        Public Shared Function CountRanges(textLength As Integer, ranges As IEnumerable(Of SearchSpan)) As Integer
            Return ValidRanges(textLength, ranges).Count()
        End Function

        Private Shared Function ValidRanges(textLength As Integer, ranges As IEnumerable(Of SearchSpan)) As IEnumerable(Of SearchSpan)
            If textLength < 0 Then Throw New ArgumentOutOfRangeException(NameOf(textLength))
            Return If(ranges, Enumerable.Empty(Of SearchSpan)()).
                Where(Function(span) span IsNot Nothing AndAlso span.Start >= 0 AndAlso span.Length > 0 AndAlso
                          span.Start < textLength AndAlso span.Length <= textLength - span.Start).
                Take(SearchQuery.MaximumHighlights)
        End Function

        Public ReadOnly Property Count As Integer
            Get
                Return _ranges.Length
            End Get
        End Property
        Public ReadOnly Property Index As Integer
            Get
                Return _index
            End Get
        End Property
        Public ReadOnly Property Current As SearchSpan
            Get
                If _index < 0 Then Return Nothing
                Dim span = _ranges(_index)
                Return New SearchSpan With {.Start = span.Start, .Length = span.Length}
            End Get
        End Property

        Public Function Move(forward As Boolean) As Boolean
            _Wrapped = False
            If Count = 0 Then Return False
            If _index < 0 Then
                _index = If(forward, 0, Count - 1)
            Else
                _index += If(forward, 1, -1)
                If _index < 0 OrElse _index >= Count Then
                    _index = If(forward, 0, Count - 1)
                    _Wrapped = True
                End If
            End If
            Return True
        End Function

        Public Function SelectIndex(index As Integer) As Boolean
            If index < 0 OrElse index >= Count Then Return False
            _index = index
            _Wrapped = False
            Return True
        End Function

        Public Function SelectAtOrAfter(position As Integer) As Boolean
            For candidate As Integer = 0 To Count - 1
                If _ranges(candidate).Start >= position Then Return SelectIndex(candidate)
            Next
            Return False
        End Function

        Public Function MoveFromPosition(position As Integer, forward As Boolean) As Boolean
            If Count = 0 Then Return False
            If forward Then
                If SelectAtOrAfter(position) Then Return True
            Else
                For candidate As Integer = Count - 1 To 0 Step -1
                    If _ranges(candidate).Start + _ranges(candidate).Length <= position Then Return SelectIndex(candidate)
                Next
            End If
            SelectIndex(If(forward, 0, Count - 1))
            _Wrapped = True
            Return True
        End Function
    End Class
End Namespace
