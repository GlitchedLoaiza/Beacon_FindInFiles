Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Media
Imports Beacon

Module HelpChecks
    Public Sub Run()
        Require(Not BeaconSettings.CreateDefaults().StopAfterFirstMatchPerFile, "First-match must be unchecked by default.")
        Require(Not System.Text.Json.JsonSerializer.Deserialize(Of BeaconSettings)("{}").StopAfterFirstMatchPerFile, "Missing persisted preference must use the new default.")
        Require(BeaconSettingsService.Clone(New BeaconSettings With {.StopAfterFirstMatchPerFile = True}).StopAfterFirstMatchPerFile, "Explicit saved first-match preference was overwritten.")
        Dim topics = HelpContent.Topics()
        Require(topics.Count >= 14 AndAlso topics.All(Function(topic) topic.Body.Length > 200), "Help topics are missing meaningful instructions.")
        Dim quickStart = topics.Single(Function(topic) topic.Title = "Start here — your first search").Body
        Require(quickStart.Contains("Path shows your selection; it cannot be edited directly.") AndAlso quickStart.Contains("""error"""), "Quick start has inaccurate Path guidance or an unquoted example.")
        Require(Not quickStart.Contains("Export") AndAlso Not quickStart.Contains("Reset"), "Quick start contains instructions belonging to other sections.")
        Dim modes = topics.Single(Function(topic) topic.Title = "Search modes").Body
        Require(Not modes.Contains("Settings → Search") AndAlso modes.Contains("""code=\d+"""), "Search modes contains unrelated settings instructions or an unquoted example.")
        Require(Not topics.Any(Function(topic) topic.Body.Contains("First-match mode is unchecked by default") OrElse topic.Body.Contains("First-match mode defaults to unchecked")), "Help retained unnecessary first-match default wording.")
        Require(topics.Any(Function(topic) topic.Title = "Changing a search, cancelling and Reset") AndAlso topics.Any(Function(topic) topic.Title = "Exporting, copying and Diagnostics"), "Task-specific help sections are missing.")
        Require(topics.Where(Function(topic) topic.Title.StartsWith("Regex ")).Count() = 5 AndAlso quickStart.Contains("Do not type those surrounding quotes"), "Regex learning sequence or example quoting guidance is missing.")
        For Each example In RegexHelpContent.Examples()
            Dim query As New SearchQuery(example.Pattern, SearchMode.RegularExpression, False)
            Require(query.IsMatch(example.Input) AndAlso Not query.IsMatch(example.NonMatchingInput), "Regex recipe has incorrect matching behavior: " & example.Title)
            Dim spans = query.FindHighlights(example.Input)
            Require(spans.Count = 1 AndAlso example.Input.Substring(spans(0).Start, spans(0).Length) = example.Expected, "Regex recipe highlights different text than documented: " & example.Title)
        Next
        Require(Not New SearchQuery("error", SearchMode.RegularExpression, True).IsMatch("ERROR"), "Case-sensitive lesson contradicts actual behavior.")
        Require(New SearchQuery("^ERROR\b.*", SearchMode.RegularExpression, False).IsMatch("INFO ready" & vbLf & "ERROR failure"), "Multiline anchor lesson contradicts actual behavior.")
        Require(Not New SearchQuery("error.*timeout", SearchMode.RegularExpression, False).IsMatch("error" & vbLf & "timeout") AndAlso
                New SearchQuery("(?s)error.*timeout", SearchMode.RegularExpression, False).IsMatch("error" & vbLf & "timeout"), "Dot/newline lesson contradicts actual behavior.")
        Require(New SearchQuery("^[0-9]{4}-[0-9]{2}-[0-9]{2}\r?$", SearchMode.RegularExpression, False).IsMatch("2026-01-02" & vbCrLf), "Windows line-ending example failed.")
        Dim zeroWidth As New SearchQuery("(?=error)", SearchMode.RegularExpression, False)
        Require(zeroWidth.IsMatch("error") AndAlso zeroWidth.FindHighlights("error").Count = 0, "Zero-width explanation is incorrect.")
        For Each dark In {False, True}
            Dim owner As New SettingsWindow(New BeaconSettings(), dark) With {.WindowStartupLocation = WindowStartupLocation.Manual, .Left = -10000, .Top = -10000, .ShowActivated = False, .ShowInTaskbar = False}
            Dim window As HelpWindow = Nothing
            Try
                owner.Show()
                Require(Not DirectCast(owner.FindName("StopAfterFirstMatch_chk"), CheckBox).IsChecked.GetValueOrDefault(), "Settings checkbox does not show the new default.")
                DirectCast(owner.FindName("StopAfterFirstMatch_chk"), CheckBox).IsChecked = True
                DirectCast(owner.FindName("RestoreDefaults_btn"), Button).RaiseEvent(New RoutedEventArgs(Button.ClickEvent))
                Require(Not DirectCast(owner.FindName("StopAfterFirstMatch_chk"), CheckBox).IsChecked.GetValueOrDefault(), "Restore defaults did not uncheck first-match.")
                Dim opened As New List(Of String)()
                Dim failBrowser As Boolean = False
                window = New HelpWindow(owner, dark, Sub(url)
                                                        If failBrowser Then Throw New InvalidOperationException("Test browser failure")
                                                        opened.Add(url)
                                                    End Sub) With {.WindowStartupLocation = WindowStartupLocation.Manual, .Left = -10000, .Top = -10000, .ShowActivated = False, .ShowInTaskbar = False}
                window.Show()
                Pump(window)
                NativeCaptionChecks.Verify(window, dark)
                Dim list = DirectCast(window.FindName("HelpTopics_lst"), ListBox)
                Dim search = DirectCast(window.FindName("HelpSearch_txt"), TextBox)
                Dim body = DirectCast(window.FindName("HelpBody_txt"), TextBox)
                Require(list.Items.Count = topics.Count AndAlso body.Text.Contains("1. Click Folder"), "Help does not start with the quick-start guide.")
                Require(body.IsReadOnly AndAlso body.TextWrapping = TextWrapping.Wrap, "Help must be readable and protected from editing.")
                Require(DirectCast(body.Foreground, SolidColorBrush).Color = DirectCast(owner.Resources("TextPrimaryBrush"), SolidColorBrush).Color, "Help text does not follow the theme.")
                search.Text = "only one event"
                Require(list.Items.Count > 0, "Common one-event question has no help result.")
                search.Text = "zzzz-no-help-topic"
                Require(list.Items.Count = 0 AndAlso body.Text.Contains("Click Clear"), "Empty help search lacks recovery instructions.")
                DirectCast(window.FindName("ClearHelpSearch_btn"), Button).RaiseEvent(New RoutedEventArgs(Button.ClickEvent))
                Require(list.Items.Count = topics.Count, "Clear did not restore help topics.")
                window.Width = window.MinWidth
                window.Height = window.MinHeight
                Pump(window)
                Require(body.ActualWidth > 200 AndAlso body.ActualHeight > 150, "Help is not readable at minimum size.")
                For Each topic In list.Items.Cast(Of HelpTopic)().ToArray()
                    list.SelectedItem = topic
                    Require(body.Text = topic.Body, "Help topic selection did not update the description.")
                Next
                search.Text = "regex digits"
                Require(list.Items.Count > 0, "Beginner regex search has no results.")
                Dim regexTopic = list.Items.Cast(Of HelpTopic)().First(Function(topic) topic.DocumentationUrl IsNot Nothing)
                list.SelectedItem = regexTopic
                Pump(window)
                Dim linkRow = DirectCast(window.FindName("DocumentationLinkRow"), TextBlock)
                Require(linkRow.IsVisible AndAlso body.ActualHeight > 100, "Regex documentation link hides the reading area.")
                Dim link = DirectCast(window.FindName("RegexDocumentation_lnk"), System.Windows.Documents.Hyperlink)
                Require(opened.Count = 0, "Help opened an external page without a click.")
                link.RaiseEvent(New RoutedEventArgs(System.Windows.Documents.Hyperlink.ClickEvent))
                Require(opened.SequenceEqual({RegexHelpContent.DocumentationUrl}) AndAlso New Uri(opened(0)).Host = "learn.microsoft.com", "Documentation action opened an unexpected destination.")
                failBrowser = True
                link.RaiseEvent(New RoutedEventArgs(System.Windows.Documents.Hyperlink.ClickEvent))
                Require(DirectCast(window.FindName("DocumentationStatus_txt"), TextBlock).Text.Contains(RegexHelpContent.DocumentationUrl), "Browser failure lacks recovery instructions.")
                search.Text = "zzzz-no-help-topic"
                Require(linkRow.Visibility = Visibility.Collapsed, "No-results view retained a stale documentation link.")
                DirectCast(window.FindName("CloseHelp_btn"), Button).RaiseEvent(New RoutedEventArgs(Button.ClickEvent))
                Require(Not window.IsVisible, "Close did not dismiss Help.")
            Finally
                If window IsNot Nothing AndAlso window.IsVisible Then window.Close()
                owner.Close()
            End Try
        Next
    End Sub
    Private Sub Pump(window As Window)
        window.UpdateLayout()
        window.Dispatcher.Invoke(Sub()
                                 End Sub, System.Windows.Threading.DispatcherPriority.ContextIdle)
        window.UpdateLayout()
    End Sub
    Private Sub Require(condition As Boolean, message As String)
        If Not condition Then Throw New InvalidOperationException(message)
    End Sub
End Module
