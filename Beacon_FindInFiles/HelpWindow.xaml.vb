Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Media
Imports System.Diagnostics

Namespace Beacon
    Partial Public Class HelpWindow
        Private ReadOnly _topics As List(Of HelpTopic) = HelpContent.Topics()
        Private ReadOnly _openDocumentation As Action(Of String)

        Public Sub New(owner As Window, dark As Boolean)
            Me.New(owner, dark, Sub(url) Process.Start(New ProcessStartInfo(url) With {.UseShellExecute = True}))
        End Sub

        Friend Sub New(owner As Window, dark As Boolean, openDocumentation As Action(Of String))
            InitializeComponent()
            _openDocumentation = openDocumentation
            Me.Owner = owner
            BeaconThemePalette.CopyOwnerColors(owner, Me, dark)
            NativeCaptionTheme.Apply(Me, dark)
            AddHandler HelpSearch_txt.TextChanged, Sub() FilterTopics()
            AddHandler HelpTopics_lst.SelectionChanged, AddressOf TopicSelected
            AddHandler ClearHelpSearch_btn.Click, Sub()
                                                      HelpSearch_txt.Clear()
                                                      HelpSearch_txt.Focus()
                                                  End Sub
            AddHandler CloseHelp_btn.Click, Sub() Close()
            FilterTopics()
        End Sub

        Private Sub FilterTopics()
            Dim previous = TryCast(HelpTopics_lst.SelectedItem, HelpTopic)
            Dim terms = HelpSearch_txt.Text.Split({" "c, ControlChars.Tab}, StringSplitOptions.RemoveEmptyEntries)
            Dim matches = _topics.Where(Function(topic) terms.All(Function(term) (topic.Title & " " & topic.Body).Contains(term, StringComparison.OrdinalIgnoreCase))).ToList()
            HelpTopics_lst.ItemsSource = matches
            HelpTopics_lst.SelectedItem = If(matches.Contains(previous), previous, matches.FirstOrDefault())
            HelpCount_txt.Text = $"{matches.Count} of {_topics.Count} topics"
            If matches.Count = 0 Then
                DocumentationLinkRow.Visibility = Visibility.Collapsed
                DocumentationStatus_txt.Visibility = Visibility.Collapsed
                HelpTitle_txt.Text = "No matching topics"
                HelpBody_txt.Text = "Try a shorter search such as events, filters, UTC, export, or redaction. Click Clear to see all topics."
            End If
        End Sub

        Private Sub TopicSelected(sender As Object, e As SelectionChangedEventArgs)
            Dim topic = TryCast(HelpTopics_lst.SelectedItem, HelpTopic)
            If topic Is Nothing Then Return
            HelpTitle_txt.Text = topic.Title
            HelpBody_txt.Text = topic.Body
            DocumentationLinkRow.Visibility = If(topic.DocumentationUrl = RegexHelpContent.DocumentationUrl, Visibility.Visible, Visibility.Collapsed)
            DocumentationStatus_txt.Visibility = Visibility.Collapsed
            HelpBody_txt.ScrollToHome()
        End Sub

        Private Sub OpenDocumentation(sender As Object, e As RoutedEventArgs)
            Dim topic = TryCast(HelpTopics_lst.SelectedItem, HelpTopic)
            If topic Is Nothing OrElse topic.DocumentationUrl <> RegexHelpContent.DocumentationUrl Then Return
            Try
                _openDocumentation(RegexHelpContent.DocumentationUrl)
                DocumentationStatus_txt.Visibility = Visibility.Collapsed
            Catch ex As Exception
                DocumentationStatus_txt.Text = "Could not open your browser. Open this address manually: " & RegexHelpContent.DocumentationUrl
                DocumentationStatus_txt.Visibility = Visibility.Visible
            End Try
            e.Handled = True
        End Sub
    End Class
End Namespace
