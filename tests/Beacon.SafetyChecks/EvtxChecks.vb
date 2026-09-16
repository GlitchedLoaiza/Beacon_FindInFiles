Imports System.Collections
Imports System.Reflection
Imports System.Windows
Imports System.Windows.Controls
Imports Beacon

Module EvtxChecks
    Public Sub Run()
        Dim instant As New DateTime(2026, 1, 2, 3, 4, 5, DateTimeKind.Utc)
        Dim filter As New EvtxFilter("1000,1001", "Provider", "2,3", "2026-01-02 03:04:05", "2026-01-02T03:04:05Z")
        Require(filter.Matches(1000, "provider", CByte(2), instant), "Inclusive UTC boundary or case-insensitive provider failed.")
        Require(Not filter.Matches(1000, "provider", CByte(2), instant.AddSeconds(1)), "UTC end boundary failed.")
        Require(Not filter.Matches(1002, "provider", CByte(2), instant), "ID filter failed.")
        Require(Not filter.Matches(1000, "other", CByte(2), instant), "Provider filter failed.")
        Require(Not filter.Matches(1000, "provider", CByte(4), instant), "Severity filter failed.")
        Require(Not filter.Matches(1000, "provider", Nothing, instant), "Missing severity passed a level filter.")
        Require(Not filter.Matches(1000, "provider", CByte(2), Nothing), "Missing timestamp passed a date filter.")
        Require(New EvtxFilter("", "", "", "", "").Matches(0, Nothing, Nothing, Nothing), "Blank filters are not unrestricted.")
        For Each invalid In {"-1", "65536", "1,", "1-5", "x", New String("1"c, 2049)}
            Reject(Sub()
                       Dim unused As New EvtxFilter(invalid, "", "", "", "")
                   End Sub)
        Next
        Reject(Sub()
                   Dim unused As New EvtxFilter("", "", "6", "", "")
               End Sub)
        Reject(Sub()
                   Dim unused As New EvtxFilter("", "", "", "01/02/2026", "")
               End Sub)
        Reject(Sub()
                   Dim unused As New EvtxFilter("", "", "", "2026-01-02 00:00:00", "2026-01-01 00:00:00")
               End Sub)
        Dim settings = BeaconSettingsService.Clone(New BeaconSettings With {.EvtxEventIds = "1000", .EvtxProvider = "Provider", .EvtxLevels = "2", .EvtxFromUtc = "2026-01-02 03:04:05"})
        Require(EvtxFilter.FromSettings(settings).Matches(1000, "Provider", CByte(2), instant), "Filter settings did not survive serialization.")
        Require(EvtxFilter.Describe(settings).Contains("Provider: Provider"), "Report filter description is missing scope.")
        For Each dark In {False, True}
            CheckPreview(dark)
        Next
    End Sub

    Private Sub CheckPreview(dark As Boolean)
        Dim window As New MainWindow()
        Dim type = GetType(MainWindow)
        Dim flags = BindingFlags.Instance Or BindingFlags.NonPublic
        window.RemoveHandler(FrameworkElement.LoadedEvent, type.GetMethod("MainWindow_Loaded", flags).CreateDelegate(GetType(RoutedEventHandler), window))
        RemoveHandler window.Closing, DirectCast(type.GetMethod("MainWindow_Closing", flags).CreateDelegate(GetType(ComponentModel.CancelEventHandler), window), ComponentModel.CancelEventHandler)
        Try
            type.GetMethod(If(dark, "ApplyDarkTheme", "ApplyLightTheme"), flags).Invoke(window, Nothing)
            type.GetField("_activeQuery", flags).SetValue(window, New SearchQuery("error", SearchMode.PlainText, False))
            Dim hitType = type.GetNestedType("SearchHit", BindingFlags.NonPublic)
            Dim eventType = type.GetNestedType("EventSummary", BindingFlags.NonPublic)
            Dim hit = Activator.CreateInstance(hitType, True)
            Dim events = DirectCast(hitType.GetProperty("MatchingEvents").GetValue(hit), IList)
            For Each id In {1000, 1001}
                Dim ev = Activator.CreateInstance(eventType, True)
                eventType.GetProperty("EventId").SetValue(ev, id)
                eventType.GetProperty("LevelNumber").SetValue(ev, CByte(2))
                eventType.GetProperty("Provider").SetValue(ev, "Provider")
                eventType.GetProperty("Message").SetValue(ev, "error")
                eventType.GetProperty("RawXml").SetValue(ev, $"<Event><Id>{id}</Id></Event>")
                events.Add(ev)
            Next
            hitType.GetProperty("DisplayName").SetValue(hit, "captured.evtx")
            Dim hits = DirectCast(type.GetField("_hits", flags).GetValue(window), IList)
            hits.Add(hit)
            DirectCast(window.FindName("Results_lst"), ListBox).SelectedItem = hit
            type.GetMethod("ShowEventPreviewMode", flags).Invoke(window, Nothing)
            window.WindowStartupLocation = WindowStartupLocation.Manual
            window.WindowState = WindowState.Normal
            window.Left = -10000
            window.Top = -10000
            window.ShowActivated = False
            window.ShowInTaskbar = False
            window.Show()
            DirectCast(window.FindName("EventXml_exp"), Expander).IsExpanded = True
            window.UpdateLayout()
            Require(DirectCast(window.FindName("EventXml_exp"), Expander).Template.FindName("DetailsToggle", DirectCast(window.FindName("EventXml_exp"), Expander)) IsNot Nothing,
                    "XML tools did not use the themed expander template.")
            Dim tools = DirectCast(window.FindName("EventToolsPane"), Expander)
            Dim content = DirectCast(window.FindName("EventContentScroll"), ScrollViewer)
            Require(Not tools.IsExpanded, "Event tools should initially leave the preview space available.")
            For Each width In {1000.0, 1400.0, 2200.0}
                window.Width = width
                tools.IsExpanded = False
                window.UpdateLayout()
                Dim collapsedWidth = content.ActualWidth
                Dim toggle = DirectCast(tools.Template.FindName("ToolsToggle", tools), System.Windows.Controls.Primitives.ToggleButton)
                toggle.IsChecked = True
                window.UpdateLayout()
                Require(tools.IsExpanded AndAlso content.ActualWidth < collapsedWidth - 250 AndAlso content.ActualWidth > 100,
                        "Expanding event tools did not reserve a usable right-side pane.")
                Dim toolsPosition = tools.TranslatePoint(New Point(), window)
                Dim contentPosition = content.TranslatePoint(New Point(), window)
                Require(toolsPosition.X >= contentPosition.X + content.ActualWidth - 1, "Event tools overlaps the main event content.")
                Require(DirectCast(window.FindName("EventResourceGuidance_txt"), FrameworkElement).IsDescendantOf(tools) AndAlso
                        DirectCast(window.FindName("EventFilterIds_txt"), FrameworkElement).IsDescendantOf(tools), "Filters or guidance are outside the side pane.")
                Require(DirectCast(window.FindName("EventXml_exp"), FrameworkElement).IsDescendantOf(content), "Raw XML moved out of the main preview.")
                toggle.IsChecked = False
                window.UpdateLayout()
                Require(Not tools.IsExpanded AndAlso Math.Abs(content.ActualWidth - collapsedWidth) <= 1, "Collapsing event tools did not restore preview space.")
            Next
            tools.IsExpanded = True
            Dim provider = DirectCast(window.FindName("EventFilterProvider_cmb"), ComboBox)
            provider.Text = "Typed.Provider"
            type.GetMethod("RefreshEventProviders", flags).Invoke(window, {Nothing, EventArgs.Empty})
            Require(provider.Items.Cast(Of String)().SequenceEqual({"", "Provider"}) AndAlso provider.Text = "Typed.Provider", "Preview provider suggestions lost captured providers or typed text.")
            provider.Text = ""
            DirectCast(window.FindName("EventFilterIds_txt"), TextBox).Text = "1001"
            DirectCast(window.FindName("ApplyEventFilter_btn"), Button).RaiseEvent(New RoutedEventArgs(Button.ClickEvent))
            Require(CInt(hitType.GetProperty("CurrentEventIndex").GetValue(hit)) = 1 AndAlso events.Count = 2, "Preview filtering changed captured events or selected the wrong event.")
            tools.IsExpanded = False
            tools.IsExpanded = True
            Require(DirectCast(window.FindName("EventFilterIds_txt"), TextBox).Text = "1001" AndAlso CInt(hitType.GetProperty("CurrentEventIndex").GetValue(hit)) = 1,
                    "Collapsing the side pane cleared the user's filters.")
            Require(DirectCast(window.FindName("EventXml_txt"), TextBox).Text.Contains("1001"), "XML did not track the selected event.")
            Require(DirectCast(window.FindName("CopyEventXml_btn"), Button).IsEnabled, "Complete XML cannot be copied.")
            type.GetMethod("NavigateEventMatch", flags).Invoke(window, {True})
            Require(CInt(hitType.GetProperty("CurrentEventIndex").GetValue(hit)) = 1, "Navigation escaped the preview filter.")
            DirectCast(window.FindName("EventFilterIds_txt"), TextBox).Text = "invalid"
            DirectCast(window.FindName("ApplyEventFilter_btn"), Button).RaiseEvent(New RoutedEventArgs(Button.ClickEvent))
            Require(CInt(hitType.GetProperty("CurrentEventIndex").GetValue(hit)) = 1, "Invalid filter replaced the current preview filter.")
            DirectCast(window.FindName("EventFilterIds_txt"), TextBox).Text = "999"
            DirectCast(window.FindName("ApplyEventFilter_btn"), Button).RaiseEvent(New RoutedEventArgs(Button.ClickEvent))
            Require(CInt(hitType.GetProperty("CurrentEventIndex").GetValue(hit)) = -1 AndAlso Not DirectCast(window.FindName("CopyEventXml_btn"), Button).IsEnabled, "Empty filter retained stale XML state.")
            DirectCast(window.FindName("ClearEventFilter_btn"), Button).RaiseEvent(New RoutedEventArgs(Button.ClickEvent))
            Require(CInt(hitType.GetProperty("CurrentEventIndex").GetValue(hit)) = 0 AndAlso events.Count = 2, "Clearing filters did not restore captured events.")
            eventType.GetProperty("XmlShortened").SetValue(events(0), True)
            type.GetMethod("RenderEvent", flags).Invoke(window, {events(0)})
            Require(Not DirectCast(window.FindName("CopyEventXml_btn"), Button).IsEnabled, "Truncated XML was offered as a complete copy.")
        Finally
            window.Close()
        End Try
    End Sub

    Private Sub Reject(action As Action)
        Try
            action()
        Catch ex As ArgumentException
            Return
        End Try
        Throw New InvalidOperationException("Invalid EVTX filter was accepted.")
    End Sub

    Private Sub Require(condition As Boolean, message As String)
        If Not condition Then Throw New InvalidOperationException(message)
    End Sub
End Module
