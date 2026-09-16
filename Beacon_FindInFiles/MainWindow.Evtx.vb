Imports System.Windows

Namespace Beacon
    Partial Public Class MainWindow
        Private ReadOnly EventFilterLevels_txt As New BeaconFilterControls.SeveritySelector()
        Private ReadOnly EventFilterFrom_txt As New BeaconFilterControls.UtcDateTimePicker()
        Private ReadOnly EventFilterTo_txt As New BeaconFilterControls.UtcDateTimePicker()

        Private Sub InitializeEvtxControls()
            EventFilterLevels_host.Content = EventFilterLevels_txt
            EventFilterFrom_host.Content = EventFilterFrom_txt
            EventFilterTo_host.Content = EventFilterTo_txt
        End Sub

        Private _eventPreviewFilter As New EvtxFilter("", "", "", "", "")

        Private Function VisibleEventIndexes(hit As SearchHit) As List(Of Integer)
            Return Enumerable.Range(0, hit.MatchingEvents.Count).Where(
                Function(index)
                    Dim ev = hit.MatchingEvents(index)
                    Return _eventPreviewFilter.Matches(ev.EventId, ev.Provider, ev.LevelNumber, ev.TimeCreated)
                End Function).ToList()
        End Function

        Private Sub ApplyEventFilter(sender As Object, e As RoutedEventArgs) Handles ApplyEventFilter_btn.Click
            Try
                Dim filter As New EvtxFilter(EventFilterIds_txt.Text, EventFilterProvider_cmb.Text, EventFilterLevels_txt.Text,
                                             EventFilterFrom_txt.Text, EventFilterTo_txt.Text)
                _eventPreviewFilter = filter
                EventFilterStatus_txt.Text = "Preview filters applied; captured results and exports are unchanged."
                Dim hit = TryCast(Results_lst.SelectedItem, SearchHit)
                If hit IsNot Nothing Then RenderFirstEvent(hit)
            Catch ex As ArgumentException
                EventFilterStatus_txt.Text = ex.Message
            End Try
        End Sub

        Private Sub ClearEventFilter(sender As Object, e As RoutedEventArgs) Handles ClearEventFilter_btn.Click
            EventFilterIds_txt.Clear()
            EventFilterProvider_cmb.Text = ""
            EventFilterLevels_txt.Clear()
            EventFilterFrom_txt.Clear()
            EventFilterTo_txt.Clear()
            ApplyEventFilter(sender, e)
        End Sub

        Private Sub RenderEventXml(ev As EventRecordSummary)
            EventXml_txt.Text = If(ev.XmlShortened, ev.RawXml, PreviewFormatting.FormatXml(ev.RawXml, True, 65536))
            EventXmlStatus_txt.Text = If(ev.XmlShortened, "XML shortened to the capture limit; copying is disabled to avoid an incomplete XML document.",
                If(String.IsNullOrEmpty(ev.RawXml), "Raw XML is unavailable for this captured event.", "Captured event XML. Copying includes unredacted event data."))
            CopyEventXml_btn.IsEnabled = Not ev.XmlShortened AndAlso Not String.IsNullOrEmpty(ev.RawXml)
            EventResourceGuidance_txt.Text = If(ev.MessageUnavailable, "This event's message could not be rendered." & vbCrLf & vbCrLf, "") & EvtxFilter.ResourceGuidance
        End Sub

        Private Sub CopyEventXml(sender As Object, e As RoutedEventArgs) Handles CopyEventXml_btn.Click
            If Not CopyEventXml_btn.IsEnabled Then Return
            Try
                Clipboard.SetText(EventXml_txt.Text)
                EventXmlStatus_txt.Text = "Event XML copied. Review unredacted data before sharing."
            Catch ex As Exception
                EventXmlStatus_txt.Text = "Could not copy XML: " & ex.Message
            End Try
        End Sub

        Private Sub ClearEventPreview()
            EventLevel_txt.Text = "[No events match the preview filters]"
            EventId_txt.Text = ""
            EventProvider_txt.Text = ""
            EventTime_txt.Text = ""
            EventMessage_txt.Inlines.Clear()
            EventXml_txt.Clear()
            CopyEventXml_btn.IsEnabled = False
            EventXmlStatus_txt.Text = "No event selected."
            EventMatchCounter_lbl.Text = "0 visible matches"
            EventResourceGuidance_txt.Text = EvtxFilter.ResourceGuidance
        End Sub

        Private Function CapturedEventProviders() As String()
            Return _hits.SelectMany(Function(hit) hit.MatchingEvents).Select(Function(ev) ev.Provider).
                Where(Function(provider) Not String.IsNullOrWhiteSpace(provider)).
                Distinct(StringComparer.OrdinalIgnoreCase).OrderBy(Function(provider) provider, StringComparer.OrdinalIgnoreCase).ToArray()
        End Function

        Private Sub RefreshEventProviders(sender As Object, e As EventArgs) Handles EventFilterProvider_cmb.DropDownOpened
            Dim typed = EventFilterProvider_cmb.Text
            EventFilterProvider_cmb.ItemsSource = New String() {""}.Concat(CapturedEventProviders()).ToArray()
            EventFilterProvider_cmb.Text = typed
        End Sub
    End Class
End Namespace
