Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Documents
Imports System.Windows.Media
Imports System.Windows.Threading

Namespace Beacon
    Partial Public Class MainWindow
        Private Sub SearchModeChanged(sender As Object, e As SelectionChangedEventArgs)
            If SearchMode_cmb.SelectedValue Is Nothing Then Return
            _selectedSearchMode = DirectCast(SearchMode_cmb.SelectedValue, SearchMode)
            UpdateSearchHint()
            UpdateButtonsState()
        End Sub

        Private Sub UpdateSearchHint()
            If _settings Is Nothing Then Return
            Dim instructions As String
            Select Case _selectedSearchMode
                Case SearchMode.AnyTerm, SearchMode.AllTerms
                    instructions = "Separate terms with spaces; wrap phrases in double quotes. All terms must occur in the same line, event or request. HTML/XML/JSON extracted document text is one record."
                Case SearchMode.RegularExpression
                    instructions = ".NET regular expression, with a 100 ms matching timeout. Zero-width matches select records but have no visible highlight."
                Case SearchMode.ExactWord
                    instructions = "Match the literal term at Unicode word boundaries."
                Case Else
                    instructions = "Find this literal text. Regular-expression characters are treated literally."
            End Select
            Search_txt.ToolTip = instructions
            SearchMode_cmb.ToolTip = instructions
        End Sub

        Private Sub ShowMetadataPreview(hit As SearchHit)
            ShowTextPreviewMode()
            _showingTextMetadata = True
            FindNextEvent_btn.Visibility = Visibility.Collapsed
            FindPreviousEvent_btn.Visibility = Visibility.Collapsed
            FindNextHarRequest_btn.Visibility = Visibility.Collapsed
            FindPreviousHarRequest_btn.Visibility = Visibility.Collapsed
            FindNext_btn.Visibility = Visibility.Visible
            SetTextPreview("Name/path match" & vbCrLf & hit.LogicalPath & vbCrLf & vbCrLf &
                           String.Join(vbCrLf & vbCrLf, hit.Details.Where(Function(detail) detail.IsMetadata).
                                       Select(Function(detail) detail.Location & ": " & detail.Excerpt)))
            UpdateTextNavigationState()
            UpdateTextPreviewModeState()
        End Sub

        Private Sub SearchDetailSelected(sender As Object, e As SelectionChangedEventArgs)
            If _syncingTextDetails Then Return
            Dim hit = TryCast(Results_lst.SelectedItem, SearchHit)
            Dim detail = TryCast(Details_lst.SelectedItem, SearchDetail)
            If hit Is Nothing OrElse detail Is Nothing Then Return
            If detail.IsMetadata Then
                ShowMetadataPreview(hit)
            ElseIf hit.MatchingEvents.Count > 0 Then
                ClearEventFilter(Nothing, Nothing)
                FindNext_btn.Visibility = Visibility.Collapsed
                hit.CurrentEventIndex = detail.RecordIndex
                _currentEventMessageMatchIndex = 0
                ShowEventPreviewMode()
                RenderEvent(hit.MatchingEvents(hit.CurrentEventIndex))
                UpdateEventCounter(hit)
                FindNextEvent_btn.Visibility = Visibility.Visible
                FindPreviousEvent_btn.Visibility = Visibility.Visible
            ElseIf hit.MatchingRequests.Count > 0 Then
                FindNext_btn.Visibility = Visibility.Collapsed
                ClearHarFilter(Nothing, Nothing)
                hit.CurrentRequestIndex = detail.RecordIndex
                ShowHarPreviewMode()
                RenderHarRequest(hit)
                FindNextHarRequest_btn.Visibility = Visibility.Visible
                FindPreviousHarRequest_btn.Visibility = Visibility.Visible
            Else
                If detail.LineNumber > 0 Then
                    RequestTextDetail(hit, detail)
                Else
                    Results_lst_SelectionChanged(Results_lst, Nothing)
                End If
            End If
        End Sub
        Private Function PreviewSpans(text As String) As List(Of SearchSpan)
            If _activeQuery Is Nothing Then Return New List(Of SearchSpan)()
            Try
                Return _activeQuery.FindHighlights(text)
            Catch ex As System.Text.RegularExpressions.RegexMatchTimeoutException
                If Not _isScanning Then Status("Preview highlighting exceeded the regex time limit.")
                Return New List(Of SearchSpan)()
            End Try
        End Function

        Private Sub RenderMatchedText(block As TextBlock, text As String, Optional currentStart As Integer = -1)
            block.Inlines.Clear()
            If String.IsNullOrEmpty(text) Then Return
            Dim position As Integer = 0
            For Each span In PreviewSpans(text)
                If span.Start > position Then block.Inlines.Add(New Run(text.Substring(position, span.Start - position)))
                block.Inlines.Add(New Run(text.Substring(span.Start, span.Length)) With {
                    .Background = If(span.Start = currentStart, Brushes.Orange, Brushes.Yellow),
                    .Foreground = Brushes.Black, .FontWeight = FontWeights.SemiBold
                })
                position = span.Start + span.Length
            Next
            If position < text.Length Then block.Inlines.Add(New Run(text.Substring(position)))
        End Sub

        Private Sub NavigateEventMatch(forward As Boolean)
            Dim hit = TryCast(Results_lst.SelectedItem, SearchHit)
            If hit Is Nothing OrElse hit.MatchingEvents.Count = 0 Then Return
            Dim visible = VisibleEventIndexes(hit)
            If visible.Count = 0 Then Return
            If Not visible.Contains(hit.CurrentEventIndex) Then hit.CurrentEventIndex = visible(0)
            Dim spans = PreviewSpans(hit.MatchingEvents(hit.CurrentEventIndex).Message)
            Dim span = If(forward,
                          spans.FirstOrDefault(Function(item) item.Start >= _currentEventMessageMatchIndex),
                          spans.LastOrDefault(Function(item) item.Start + item.Length < _currentEventMessageMatchIndex))
            If span Is Nothing Then
                Dim position = visible.IndexOf(hit.CurrentEventIndex)
                hit.CurrentEventIndex = visible((position + If(forward, 1, visible.Count - 1)) Mod visible.Count)
                RenderEvent(hit.MatchingEvents(hit.CurrentEventIndex))
                spans = PreviewSpans(hit.MatchingEvents(hit.CurrentEventIndex).Message)
                span = If(forward, spans.FirstOrDefault(), spans.LastOrDefault())
            End If
            If span IsNot Nothing Then
                _currentEventMessageMatchIndex = span.Start + span.Length
                RenderEventWithHighlight(hit.MatchingEvents(hit.CurrentEventIndex), span.Start, span.Length)
            End If
            UpdateEventCounter(hit)
        End Sub
    End Class
End Namespace
