Imports System.IO
Imports System.Threading
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Data
Imports System.Windows.Documents
Imports System.Windows.Input
Imports System.Windows.Threading

Namespace Beacon
    Partial Public Class MainWindow
        Private _summaryPreferred As Boolean
        Private _summaryActive As Boolean
        Private _summaryProjection As TextSummaryProjection
        Private _summaryHit As SearchHit
        Private _summaryQuery As SearchQuery
        Private _currentSummaryBlock As TextSummaryBlock
        Private _summarySelectionChanging As Boolean
        Private _syncingTextDetails As Boolean
        Private _settingPreviewMode As Boolean
        Private _showingTextMetadata As Boolean

        Private Sub InitializeTextPreviewUi()
            TextPreviewMode_cmb.ItemsSource = New String() {"Full", "Summary"}
            TextPreviewMode_cmb.SelectedIndex = 0
            AddHandler TextPreviewMode_cmb.SelectionChanged, AddressOf TextPreviewModeChanged
            AddHandler SummaryPreview_lst.SelectionChanged, AddressOf SummarySelectionChanged
            Dim factory As New FrameworkElementFactory(GetType(TextSummaryBlockView))
            factory.SetBinding(TextSummaryBlockView.BlockProperty, New Binding())
            factory.SetBinding(TextSummaryBlockView.ActiveOccurrenceProperty, New Binding(NameOf(TextSummaryBlock.ActiveOccurrence)))
            factory.SetBinding(TextSummaryBlockView.IsCurrentProperty, New Binding(NameOf(TextSummaryBlock.IsCurrent)))
            factory.SetBinding(TextSummaryBlockView.WrapLinesProperty, New Binding("Tag") With {
                .RelativeSource = New RelativeSource(RelativeSourceMode.FindAncestor, GetType(ListBox), 1), .TargetNullValue = False, .FallbackValue = False})
            SummaryPreview_lst.ItemTemplate = New DataTemplate With {.VisualTree = factory}
            SummaryPreview_lst.CommandBindings.Add(New CommandBinding(ApplicationCommands.Copy,
                Sub(sender, e)
                    Dim block = TryCast(SummaryPreview_lst.SelectedItem, TextSummaryBlock)
                    If block Is Nothing Then Return
                    Try
                        Clipboard.SetText(block.ToString())
                    Catch ex As Exception
                        If Not _isScanning Then Status("Could not copy captured context: " & ex.Message)
                    End Try
                    e.Handled = True
                End Sub,
                Sub(sender, e)
                    e.CanExecute = _summaryActive AndAlso SummaryPreview_lst.SelectedItem IsNot Nothing
                    e.Handled = True
                End Sub))
        End Sub

        Private Function CanSummarize(hit As SearchHit) As Boolean
            If hit Is Nothing OrElse hit.MetadataOnly OrElse hit.MatchingEvents.Count > 0 OrElse hit.MatchingRequests.Count > 0 Then Return False
            If hit.Kind <> HitKind.DiskTextFile AndAlso hit.Kind <> HitKind.ZipTextEntry AndAlso hit.Kind <> HitKind.CabTextEntry Then Return False
            Dim extension = Path.GetExtension(If(hit.ZipEntryName, hit.FilePath)).ToLowerInvariant()
            Return extension <> ".html" AndAlso extension <> ".xml" AndAlso extension <> ".json" AndAlso extension <> ".har" AndAlso extension <> ".evtx"
        End Function

        Private Sub UpdateTextPreviewModeState()
            Dim eligible = CanSummarize(TryCast(Results_lst.SelectedItem, SearchHit)) AndAlso Not _showingTextMetadata
            TextPreviewModePanel.Visibility = If(eligible AndAlso TextPreview_grp.Visibility = Visibility.Visible, Visibility.Visible, Visibility.Collapsed)
            _settingPreviewMode = True
            Try
                TextPreviewMode_cmb.SelectedIndex = If(_summaryPreferred, 1, 0)
            Finally
                _settingPreviewMode = False
            End Try
        End Sub

        Private Sub TextPreviewModeChanged(sender As Object, e As SelectionChangedEventArgs)
            If _settingPreviewMode OrElse _isClosing OrElse _isResetting Then Return
            Dim hit = TryCast(Results_lst.SelectedItem, SearchHit)
            If Not CanSummarize(hit) OrElse _showingTextMetadata Then Return
            _summaryPreferred = TextPreviewMode_cmb.SelectedIndex = 1
            TextPreviewMode_cmb.IsDropDownOpen = False
            Dim detail = If(_summaryActive, TryCast(Details_lst.SelectedItem, SearchDetail), CapturedDetailAtTextCursor())
            If _summaryPreferred Then
                ShowSummaryPreview(hit, detail)
            Else
                ShowFullTextPreview(hit, detail)
            End If
        End Sub

        Private Sub ShowFullTextPreview(hit As SearchHit, preferred As SearchDetail)
            _showingTextMetadata = False
            ShowTextPreviewMode()
            FindNext_btn.Visibility = Visibility.Visible
            If _fullPreviewSnapshot IsNot Nothing AndAlso _fullPreviewHit Is hit AndAlso _fullPreviewQuery Is _activeQuery Then
                StartInlineTextPreview(_fullPreviewSnapshot)
            Else
                Select Case hit.Kind
                    Case HitKind.DiskTextFile
                        LoadTextFromDisk(hit.FilePath)
                    Case HitKind.ZipTextEntry
                        LoadTextFromArchive(hit.ZipPath, hit.ZipEntryName)
                    Case HitKind.CabTextEntry
                        LoadTextFromCab(hit.ZipPath, hit.ZipEntryName)
                End Select
            End If
            _pendingTextDetail = If(preferred IsNot Nothing AndAlso Not preferred.IsMetadata, preferred, Nothing)
            UpdateTextPreviewModeState()
            UpdateTextNavigationState()
        End Sub

        Private Sub ShowSummaryPreview(hit As SearchHit, Optional preferred As SearchDetail = Nothing)
            If Not CanSummarize(hit) Then Return
            InvalidatePreviewRequests()
            _showingTextMetadata = False
            ShowTextPreviewMode()
            _summaryActive = True
            _textSourceRun = Nothing
            _textContent = ""
            _textNavigation = New TextPreviewNavigation(0, Nothing)
            _textHit = hit
            TextPreview_rtb.Document = New FlowDocument()
            TextPreview_rtb.Visibility = Visibility.Collapsed
            SummaryPreview_lst.Visibility = Visibility.Visible
            TextSummaryNotice_txt.Visibility = Visibility.Visible
            FindNext_btn.Visibility = Visibility.Visible
            If _summaryHit IsNot hit OrElse _summaryQuery IsNot _activeQuery OrElse _summaryProjection Is Nothing Then
                _currentSummaryBlock?.Deactivate()
                _currentSummaryBlock = Nothing
                _summaryHit = hit
                _summaryQuery = _activeQuery
                _summaryProjection = New TextSummaryProjection(hit.Details, hit.PartialReason)
                _summarySelectionChanging = True
                Try
                    SummaryPreview_lst.ItemsSource = _summaryProjection.Blocks
                Finally
                    _summarySelectionChanging = False
                End Try
            End If
            TextSummaryNotice_txt.Text = $"Captured search context · {_summaryProjection.ContentCount:N0} matching record(s). Unrelated lines are omitted." &
                If(_summaryProjection.PartialReason.Length > 0, " " & _summaryProjection.PartialReason & "; additional matches may exist.", "")
            Dim block As TextSummaryBlock = Nothing
            If preferred IsNot Nothing Then
                Dim index = hit.Details.IndexOf(preferred)
                block = _summaryProjection.Blocks.FirstOrDefault(Function(item) item.DetailIndex = index AndAlso Not item.IsMetadata)
            End If
            If block Is Nothing Then block = If(_currentSummaryBlock, _summaryProjection.Blocks.FirstOrDefault(Function(item) Not item.IsMetadata))
            If block IsNot Nothing Then SelectSummaryBlock(block, True, False)
            UpdateTextPreviewModeState()
            UpdateTextNavigationState()
        End Sub

        Private Sub HideTextSummary()
            _summaryActive = False
            _currentSummaryBlock?.Deactivate()
            SummaryPreview_lst.Visibility = Visibility.Collapsed
            TextSummaryNotice_txt.Visibility = Visibility.Collapsed
            TextPreview_rtb.Visibility = Visibility.Visible
        End Sub

        Private Sub ClearSummaryForSelection(hit As SearchHit)
            HideTextSummary()
            _showingTextMetadata = False
            If hit IsNot _summaryHit OrElse _activeQuery IsNot _summaryQuery Then
                _currentSummaryBlock = Nothing
                _summaryProjection = Nothing
                _summaryHit = Nothing
                _summaryQuery = Nothing
                _summarySelectionChanging = True
                Try
                    SummaryPreview_lst.ItemsSource = Nothing
                Finally
                    _summarySelectionChanging = False
                End Try
            End If
            UpdateTextPreviewModeState()
        End Sub

        Private Sub SummarySelectionChanged(sender As Object, e As SelectionChangedEventArgs)
            If _summarySelectionChanging OrElse Not _summaryActive Then Return
            Dim block = TryCast(SummaryPreview_lst.SelectedItem, TextSummaryBlock)
            If block Is Nothing Then Return
            If block.IsMetadata Then
                SynchronizeTextDetail(block)
                ShowMetadataPreview(_summaryHit)
            Else
                SelectSummaryBlock(block, True, False)
            End If
        End Sub

        Private Sub SelectSummaryBlock(block As TextSummaryBlock, forward As Boolean, focus As Boolean)
            _currentSummaryBlock?.Deactivate()
            _currentSummaryBlock = block
            block.SelectEdge(forward)
            RefreshSummarySelection(focus)
        End Sub

        Private Sub RefreshSummarySelection(focus As Boolean)
            Dim block = _currentSummaryBlock
            If block Is Nothing Then Return
            _summarySelectionChanging = True
            Try
                SummaryPreview_lst.SelectedItem = block
            Finally
                _summarySelectionChanging = False
            End Try
            SynchronizeTextDetail(block)
            UpdateTextMatchCounter()
            Dim generation = Volatile.Read(_previewGeneration)
            Dim request = Interlocked.Increment(_textNavigationGeneration)
            Dispatcher.BeginInvoke(Sub()
                                       If Not _summaryActive OrElse generation <> Volatile.Read(_previewGeneration) OrElse
                                          request <> Volatile.Read(_textNavigationGeneration) OrElse _isClosing OrElse _isResetting Then Return
                                       SummaryPreview_lst.ScrollIntoView(block)
                                       If focus Then SummaryPreview_lst.Focus()
                                   End Sub, DispatcherPriority.Loaded)
        End Sub

        Private Sub SynchronizeTextDetail(block As TextSummaryBlock)
            Dim hit = TryCast(Results_lst.SelectedItem, SearchHit)
            If hit Is Nothing OrElse hit IsNot _summaryHit OrElse block.DetailIndex < 0 OrElse block.DetailIndex >= hit.Details.Count Then Return
            _syncingTextDetails = True
            Try
                Details_lst.SelectedItem = hit.Details(block.DetailIndex)
            Finally
                _syncingTextDetails = False
            End Try
        End Sub

        Private Sub NavigateSummary(forward As Boolean)
            If _summaryProjection Is Nothing OrElse _summaryProjection.ContentCount = 0 OrElse _currentSummaryBlock Is Nothing Then Return
            If _currentSummaryBlock.MoveWithin(forward) Then
                RefreshSummarySelection(True)
                Return
            End If
            Dim index = _currentSummaryBlock.DetailIndex
            Dim wrapped As Boolean = False
            Do
                index += If(forward, 1, -1)
                If index < 0 OrElse index >= _summaryProjection.Blocks.Count Then
                    index = If(forward, 0, _summaryProjection.Blocks.Count - 1)
                    wrapped = True
                End If
            Loop While _summaryProjection.Blocks(index).IsMetadata
            SelectSummaryBlock(_summaryProjection.Blocks(index), forward, True)
            If Not _isScanning Then Status(If(wrapped, If(forward, "Wrapped to top", "Wrapped to bottom"), "Ready"))
        End Sub
    End Class
End Namespace
