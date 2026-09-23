Imports System.ComponentModel

Namespace Beacon
    Friend NotInheritable Class TextSummaryProjection
        Public ReadOnly Property Blocks As IReadOnlyList(Of TextSummaryBlock)
        Public ReadOnly Property PartialReason As String
        Public ReadOnly Property ContentCount As Integer
        Public ReadOnly Property NavigationTargetCount As Long

        Public Sub New(details As IEnumerable(Of SearchDetail), partialReason As String)
            Dim blocks As New List(Of TextSummaryBlock)()
            For Each detail In details
                Dim block As New TextSummaryBlock(detail, blocks.Count, NavigationTargetCount)
                blocks.Add(block)
                _NavigationTargetCount += block.NavigationTargetCount
            Next
            _Blocks = blocks.AsReadOnly()
            _PartialReason = If(partialReason, "")
            _ContentCount = blocks.Where(Function(block) Not block.IsMetadata).Count()
        End Sub
    End Class

    Friend NotInheritable Class TextSummaryBlock
        Implements INotifyPropertyChanged
        Private ReadOnly _lines As IReadOnlyList(Of ContextLine)
        Private ReadOnly _focus As ContextLine
        Private _navigation As TextPreviewNavigation
        Private _activeOccurrence As Integer = -1
        Private _isCurrent As Boolean
        Public ReadOnly Property DetailIndex As Integer
        Public ReadOnly Property Location As String
        Public ReadOnly Property ContextKind As String
        Public ReadOnly Property Excerpt As String
        Public ReadOnly Property IsMetadata As Boolean
        Public ReadOnly Property IsShortened As Boolean
        Public ReadOnly Property NavigationOffset As Long
        Public ReadOnly Property NavigationTargetCount As Integer
        Public ReadOnly Property Lines As IReadOnlyList(Of ContextLine)
            Get
                Return _lines
            End Get
        End Property
        Public ReadOnly Property Focus As ContextLine
            Get
                Return _focus
            End Get
        End Property
        Public ReadOnly Property ActiveOccurrence As Integer
            Get
                Return _activeOccurrence
            End Get
        End Property
        Public ReadOnly Property IsCurrent As Boolean
            Get
                Return _isCurrent
            End Get
        End Property
        Public ReadOnly Property CurrentSpan As SearchSpan
            Get
                Return If(_navigation Is Nothing, Nothing, _navigation.Current)
            End Get
        End Property

        Public Sub New(detail As SearchDetail, detailIndex As Integer, Optional navigationOffset As Long = 0)
            _DetailIndex = detailIndex
            _Location = detail.Location
            _ContextKind = detail.ContextKind
            _Excerpt = detail.Excerpt
            _IsMetadata = detail.IsMetadata
            _lines = Array.AsReadOnly(detail.Context.Take(MatchContext.LineLimit).ToArray())
            _focus = _lines.FirstOrDefault(Function(line) line.IsMatchLine)
            _IsShortened = _lines.Any(Function(line) line.Shortened)
            _NavigationOffset = navigationOffset
            If Not IsMetadata Then
                _NavigationTargetCount = Math.Max(1, If(_focus Is Nothing, 0,
                    TextPreviewNavigation.CountRanges(If(_focus.Text, "").Length, _focus.Highlights)))
            End If
        End Sub

        Private Function Navigation() As TextPreviewNavigation
            If _navigation Is Nothing Then
                _navigation = If(_focus Is Nothing OrElse IsMetadata,
                    New TextPreviewNavigation(0, Nothing), New TextPreviewNavigation(If(_focus.Text, "").Length, _focus.Highlights))
            End If
            Return _navigation
        End Function

        Public Sub SelectEdge(forward As Boolean)
            _isCurrent = True
            Dim cursor = Navigation()
            cursor.SelectIndex(If(forward, 0, cursor.Count - 1))
            SetOccurrence(cursor.Index)
        End Sub

        Public Function MoveWithin(forward As Boolean) As Boolean
            Dim cursor = Navigation()
            Dim nextIndex = cursor.Index + If(forward, 1, -1)
            If nextIndex < 0 OrElse nextIndex >= cursor.Count Then Return False
            cursor.SelectIndex(nextIndex)
            SetOccurrence(cursor.Index)
            Return True
        End Function

        Public Sub Deactivate()
            _navigation = Nothing
            _isCurrent = False
            SetOccurrence(-1)
        End Sub

        Private Sub SetOccurrence(value As Integer)
            _activeOccurrence = value
            RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(NameOf(ActiveOccurrence)))
            RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(NameOf(IsCurrent)))
        End Sub

        Public Overrides Function ToString() As String
            If IsMetadata OrElse Lines.Count = 0 Then Return Location & Environment.NewLine & Excerpt
            Return Location & Environment.NewLine & String.Join(Environment.NewLine, Lines.Select(Function(line) $"{line.LineNumber}: {line.Text}"))
        End Function

        Public Event PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged
    End Class
End Namespace
