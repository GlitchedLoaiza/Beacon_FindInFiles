Imports System.Threading
Imports System.Threading.Tasks
Imports System.Windows
Imports System.Windows.Documents
Imports System.Windows.Media
Imports System.Windows.Threading

Namespace Beacon
    Partial Public Class MainWindow
        Private _textNavigation As New TextPreviewNavigation(0, Nothing)
        Private _textSourceRun As Run
        Private _textContent As String = ""
        Private _textLines As TextPreviewLineMap
        Private _textHit As SearchHit
        Private _pendingTextDetail As SearchDetail
        Private _textAnchorDetail As SearchDetail
        Private _textCounterStatus As String = ""
        Private _textPreviewTruncated As Boolean
        Private _fullPreviewSnapshot As PreviewText
        Private _fullPreviewHit As SearchHit
        Private _fullPreviewQuery As SearchQuery

        Private Sub StartPlainTextPreview(read As Func(Of PreviewContentService, CancellationToken, PreviewText))
            If _isClosing OrElse _isResetting Then Return
            _showingTextMetadata = False
            Dim operation = BeginPreviewOperation()
            ResetTextNavigation(operation.Hit, loading:=True)
            TextPreview_rtb.Document = New FlowDocument(New Paragraph(New Run("[Loading preview…]")))
            Dim options = BeaconSettingsService.Clone(_settings)
            Dim service As New PreviewContentService(options,
                Sub(path, ex)
                    Dispatcher.BeginInvoke(Sub()
                                               If IsPreviewCurrent(operation) Then RecordFileSystemIssue(path, ex)
                                           End Sub, DispatcherPriority.Background)
                End Sub)
            Dim pending = PrepareTextPreviewAsync(operation, Function(token) read(service, token), Nothing, True)
        End Sub

        Private Sub StartInlineTextPreview(preview As PreviewText, Optional searchable As Boolean = True)
            If _isClosing OrElse _isResetting Then Return
            Dim operation = BeginPreviewOperation()
            ResetTextNavigation(operation.Hit, loading:=True)
            Dim pending = PrepareTextPreviewAsync(operation, Nothing, preview, searchable)
        End Sub

        Private Sub ResetTextNavigation(hit As SearchHit, Optional loading As Boolean = False)
            _textHit = hit
            _textSourceRun = Nothing
            _textContent = ""
            _textLines = Nothing
            _textNavigation = New TextPreviewNavigation(0, Nothing)
            _pendingTextDetail = Nothing
            _textAnchorDetail = Nothing
            _textCounterStatus = If(loading, "Finding matches…", "")
            _textPreviewTruncated = False
            _currentTextFindStart = 0
            UpdateTextNavigationState()
        End Sub

        Private Async Function PrepareTextPreviewAsync(operation As PreviewOperation,
                                                       read As Func(Of CancellationToken, PreviewText), supplied As PreviewText,
                                                       searchable As Boolean) As Task
            Dim failure As Exception = Nothing
            Try
                Try
                    Dim token = operation.Cancellation.Token
                    Dim preview = supplied
                    If read IsNot Nothing Then preview = Await Task.Run(Function() read(token), token).ConfigureAwait(False)
                    Dim rendered As String = Nothing
                    Dim document As FlowDocument = Nothing
                    Dim details As SearchDetail() = Array.Empty(Of SearchDetail)()
                    Await Dispatcher.InvokeAsync(Sub()
                                                     If Not IsPreviewCurrent(operation) Then Return
                                                     _textSourceRun = New Run(preview.Text)
                                                     document = New FlowDocument(New Paragraph(_textSourceRun))
                                                     If preview.IsTruncated Then
                                                         Dim notice As New Paragraph(New Run(preview.DisplayText.Substring(preview.Text.Length).TrimStart(ControlChars.Cr, ControlChars.Lf)))
                                                         notice.SetResourceReference(TextElement.ForegroundProperty, "TextSecondaryBrush")
                                                         document.Blocks.Add(notice)
                                                     End If
                                                     TextPreview_rtb.Document = document
                                                     _textPreviewTruncated = preview.IsTruncated
                                                     _textContent = _textSourceRun.Text
                                                     rendered = _textContent
                                                     If Not _showingTextMetadata AndAlso CanSummarize(operation.Hit) Then details = operation.Hit.Details.ToArray()
                                                     If read IsNot Nothing Then
                                                         _fullPreviewSnapshot = preview
                                                         _fullPreviewHit = operation.Hit
                                                         _fullPreviewQuery = operation.Query
                                                     End If
                                                 End Sub).Task.ConfigureAwait(False)
                    If rendered Is Nothing Then Return
                    Dim prepared = Await Task.Run(Function()
                                                      Dim spans = If(searchable AndAlso operation.Query IsNot Nothing,
                                                          operation.Query.FindHighlights(rendered, token:=token), New List(Of SearchSpan)())
                                                      Return (Spans:=spans, Lines:=New TextPreviewLineMap(rendered, details, token))
                                                  End Function, token).ConfigureAwait(False)
                    Await Dispatcher.InvokeAsync(Sub()
                                                     If Not IsPreviewCurrent(operation) OrElse TextPreview_rtb.Document IsNot document Then Return
                                                     _textNavigation = New TextPreviewNavigation(rendered.Length, prepared.Spans)
                                                     _textLines = prepared.Lines
                                                     _textCounterStatus = ""
                                                     If Not SelectPendingTextDetail() AndAlso _textNavigation.SelectIndex(0) Then SelectCurrentTextMatch()
                                                     UpdateTextNavigationState()
                                                 End Sub).Task.ConfigureAwait(False)
                Catch ex As OperationCanceledException When operation.Cancellation.IsCancellationRequested
                Catch ex As Exception
                    failure = ex
                End Try
                If failure IsNot Nothing AndAlso Not Dispatcher.HasShutdownStarted Then
                    Await Dispatcher.InvokeAsync(Sub()
                                                     If Not IsPreviewCurrent(operation) Then Return
                                                     _textNavigation = New TextPreviewNavigation(0, Nothing)
                                                     If TypeOf failure Is System.Text.RegularExpressions.RegexMatchTimeoutException Then
                                                         _textCounterStatus = "Highlighting unavailable"
                                                         If Not _isScanning Then Status("Preview highlighting exceeded the regex time limit.")
                                                     Else
                                                         _textCounterStatus = "Preview unavailable"
                                                         _textSourceRun = Nothing
                                                         _textContent = ""
                                                         TextPreview_rtb.Document = New FlowDocument(New Paragraph(New Run($"[Preview unavailable: {failure.Message}]")))
                                                     End If
                                                     UpdateTextNavigationState()
                                                 End Sub).Task.ConfigureAwait(False)
                End If
            Finally
                operation.Completion.TrySetResult(True)
            End Try
        End Function

        Private Sub UpdateTextNavigationState()
            Dim visible = TextPreview_grp.Visibility = Visibility.Visible AndAlso FindNext_btn.Visibility = Visibility.Visible
            Dim hasMatches = If(_summaryActive, _summaryProjection IsNot Nothing AndAlso _summaryProjection.ContentCount > 0, _textNavigation.Count > 0)
            FindPrevious_btn.Visibility = If(visible, Visibility.Visible, Visibility.Collapsed)
            FindPrevious_btn.IsEnabled = visible AndAlso hasMatches AndAlso Not (_isResetting OrElse _isClosing)
            If TextPreview_grp.Visibility = Visibility.Visible Then
                FindNext_btn.IsEnabled = hasMatches AndAlso Not (_isResetting OrElse _isClosing)
            Else
                FindNext_btn.IsEnabled = True
            End If
            UpdateTextMatchCounter()
        End Sub

        Private Sub UpdateTextMatchCounter()
            If _summaryActive AndAlso _summaryProjection IsNot Nothing Then
                Dim position As Long = 0
                Dim block = _currentSummaryBlock
                If block IsNot Nothing AndAlso block.IsCurrent AndAlso Not block.IsMetadata Then
                    position = block.NavigationOffset + Math.Max(0, block.ActiveOccurrence) + 1
                End If
                TextMatchCounter_lbl.Text = $"Match {position:N0} out of {_summaryProjection.NavigationTargetCount:N0}"
                TextMatchScope_lbl.Text = If(position > 0 AndAlso block.CurrentSpan Is Nothing,
                    "Summary · record anchor", "Summary · captured matches")
                Return
            End If

            Dim total = _textNavigation.Count
            Dim limited = _textPreviewTruncated OrElse total >= SearchQuery.MaximumHighlights
            TextMatchScope_lbl.Text = If(_showingTextMetadata, "Name/path preview", If(limited, "Full · limited preview", "Full · indexed matches"))
            If _textCounterStatus.Length > 0 Then
                TextMatchCounter_lbl.Text = _textCounterStatus
            ElseIf _textAnchorDetail IsNot Nothing Then
                TextMatchCounter_lbl.Text = "Record match · no visible highlight"
                TextMatchScope_lbl.Text = $"{total:N0} indexed preview matches"
            Else
                Dim position = If(_textNavigation.Index < 0, 0, _textNavigation.Index + 1)
                TextMatchCounter_lbl.Text = $"Match {position:N0} out of {total:N0}{If(total >= SearchQuery.MaximumHighlights, "+", "")}"
            End If
        End Sub

        Private Sub NavigateTextMatch(forward As Boolean)
            If _isResetting OrElse _isClosing OrElse TextPreview_grp.Visibility <> Visibility.Visible Then Return
            If _summaryActive Then
                NavigateSummary(forward)
                Return
            End If
            Dim moved As Boolean
            If _textAnchorDetail IsNot Nothing AndAlso _textLines IsNot Nothing Then
                Dim anchor = _textLines.ForDetail(_textAnchorDetail)
                moved = If(anchor Is Nothing, _textNavigation.Move(forward), _textNavigation.MoveFromPosition(anchor.Start, forward))
            Else
                moved = _textNavigation.Move(forward)
            End If
            If Not moved Then Return
            Dim wrapped = _textNavigation.Wrapped
            _pendingTextDetail = Nothing
            SelectCurrentTextMatch()
            If Not _isScanning Then Status(If(wrapped, If(forward, "Wrapped to top", "Wrapped to bottom"), "Ready"))
        End Sub

        Private Sub SelectCurrentTextMatch()
            Dim span = _textNavigation.Current
            If span Is Nothing Then Return
            _textAnchorDetail = Nothing
            SelectInRichTextBox(span.Start, span.Length)
            Dim detail = CapturedDetailAtTextCursor()
            If detail IsNot Nothing Then
                _syncingTextDetails = True
                Try
                    Details_lst.SelectedItem = detail
                Finally
                    _syncingTextDetails = False
                End Try
            End If
        End Sub

        Private Function CapturedDetailAtTextCursor() As SearchDetail
            If _textAnchorDetail IsNot Nothing Then Return _textAnchorDetail
            Dim span = _textNavigation.Current
            If _textLines Is Nothing OrElse span Is Nothing OrElse _showingTextMetadata Then Return Nothing
            Return _textLines.AtPosition(span.Start)
        End Function

        Private Sub FindPrevious_btn_Click(sender As Object, e As RoutedEventArgs)
            FindPreviousInText()
        End Sub

        Private Sub FindPreviousInText()
            NavigateTextMatch(False)
        End Sub

        Private Sub RequestTextDetail(hit As SearchHit, detail As SearchDetail)
            If Results_lst.SelectedItem IsNot hit OrElse _isResetting OrElse _isClosing Then Return
            If _summaryPreferred AndAlso CanSummarize(hit) Then
                ShowSummaryPreview(hit, detail)
                Return
            End If
            If _showingTextMetadata Then
                ShowFullTextPreview(hit, detail)
                Return
            End If
            If _textHit IsNot hit Then Results_lst_SelectionChanged(Results_lst, Nothing)
            _pendingTextDetail = detail
            If PendingPreviewTask.IsCompleted Then SelectPendingTextDetail()
        End Sub

        Private Function SelectPendingTextDetail() As Boolean
            Dim detail = _pendingTextDetail
            If detail Is Nothing OrElse _textSourceRun Is Nothing OrElse detail.LineNumber <= 0 Then Return False
            _pendingTextDetail = Nothing
            Dim location = _textLines?.ForDetail(detail)
            If location Is Nothing Then
                If Not _isScanning Then Status("This captured match is outside the loaded preview. Its captured context remains available.")
                Return False
            End If
            Dim start = location.Start
            Dim finish = start + location.Length
            Dim captured = detail.Context.FirstOrDefault(Function(line) line.IsMatchLine)
            If captured IsNot Nothing AndAlso Not captured.Shortened AndAlso captured.Text <> _textContent.Substring(start, finish - start) Then
                If Not _isScanning Then Status("Captured text is not fully available here; the source may have changed or the preview may be truncated.")
                Return False
            End If
            If _textNavigation.SelectAtOrAfter(start + detail.MatchOffset) AndAlso _textNavigation.Current.Start < finish Then
                SelectCurrentTextMatch()
                Return True
            End If
            _textAnchorDetail = detail
            SelectInRichTextBox(start, 0)
            Return True
        End Function

        Private Shared Function TextLineStart(text As String, number As Integer) As Integer
            If number < 1 Then Return -1
            Dim line As Integer = 1
            Dim position As Integer = 0
            While line < number AndAlso position < text.Length
                Dim character = text(position)
                position += 1
                If character = ControlChars.Cr Then
                    If position < text.Length AndAlso text(position) = ControlChars.Lf Then position += 1
                    line += 1
                ElseIf character = ControlChars.Lf Then
                    line += 1
                End If
            End While
            Return If(line = number, position, -1)
        End Function
    End Class
End Namespace
