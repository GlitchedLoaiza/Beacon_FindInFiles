Imports System.IO
Imports System.Reflection
Imports System.Threading
Imports System.Threading.Tasks
Imports System.Windows
Imports System.Windows.Automation
Imports System.Windows.Controls
Imports System.Windows.Media
Imports Beacon

Module TextCounterChecks
    Private Const Flags As BindingFlags = BindingFlags.Instance Or BindingFlags.NonPublic

    Public Sub NavigationAndThemes()
        WithSource("matches.log", "before" & vbCrLf & "code=1 code=22" & vbCrLf & "between" & vbCrLf & "code=333" & vbCrLf & "after",
            Sub(path)
                For Each theme In {AppTheme.Light, AppTheme.Dark, AppTheme.Beacon}
                    SearchChecks.CheckPipeline(path, "code=\d+", SearchMode.RegularExpression, New BeaconSettings With {.Theme = theme},
                        Sub(hits) Require(hits.Count = 1, "Counter fixture is missing."),
                        Sub(window, hits)
                            Dim results = DirectCast(window.FindName("Results_lst"), ListBox)
                            Dim mode = DirectCast(window.FindName("TextPreviewMode_cmb"), ComboBox)
                            Dim details = DirectCast(hits(0).GetType().GetProperty("Details").GetValue(hits(0)), List(Of SearchDetail))
                            Expect(window, "Match 0 out of 0")
                            results.SelectedItem = hits(0)
                            PreviewTestHelpers.WaitForPreview(window)
                            Expect(window, "Match 1 out of 3")
                            VerifyBar(window)
                            Click(window, "FindNext_btn") : Expect(window, "Match 2 out of 3")
                            Click(window, "FindNext_btn") : Expect(window, "Match 3 out of 3")
                            Click(window, "FindPrevious_btn") : Expect(window, "Match 2 out of 3")
                            Click(window, "FindPrevious_btn") : Expect(window, "Match 1 out of 3")
                            Click(window, "FindPrevious_btn") : Expect(window, "Match 3 out of 3")
                            Click(window, "FindNext_btn") : Expect(window, "Match 1 out of 3")
                            mode.SelectedIndex = 1
                            PreviewTestHelpers.WaitForPreview(window)
                            Expect(window, "Match 1 out of 3")
                            Require(Scope(window) = "Summary · captured matches", "Summary counter scope is missing.")
                            Click(window, "FindNext_btn") : Expect(window, "Match 2 out of 3")
                            Click(window, "FindNext_btn") : Expect(window, "Match 3 out of 3")
                            Click(window, "FindPrevious_btn") : Expect(window, "Match 2 out of 3")
                            DirectCast(window.FindName("Details_lst"), ListBox).SelectedItem = details(1)
                            PreviewTestHelpers.WaitForPreview(window)
                            Expect(window, "Match 3 out of 3")
                            If theme = AppTheme.Beacon Then
                                For Each dark In {False, True}
                                    GetType(MainWindow).GetMethod("ApplyThemeSelection", Flags).Invoke(window, {dark})
                                    PreviewTestHelpers.WaitForPreview(window)
                                    Expect(window, "Match 3 out of 3")
                                    VerifyBar(window)
                                Next
                            End If
                            mode.SelectedIndex = 0
                            PreviewTestHelpers.WaitForPreview(window)
                            Expect(window, "Match 3 out of 3")
                            results.SelectedIndex = -1
                            PreviewTestHelpers.WaitForPreview(window)
                            Expect(window, "Match 0 out of 0")
                        End Sub)
                Next
            End Sub)
    End Sub

    Public Sub LoadingLimitsAndErrors()
        WithSource("preview.log", "preview",
            Sub(path)
                SearchChecks.CheckPipeline(path, "preview", SearchMode.PlainText, New BeaconSettings(),
                    Sub(hits) Require(hits.Count = 1, "Loading counter fixture is missing."),
                    Sub(window, hits)
                        DirectCast(window.FindName("Results_lst"), ListBox).SelectedItem = hits(0)
                        PreviewTestHelpers.WaitForPreview(window)
                        Expect(window, "Match 1 out of 1")
                        Dim release As New TaskCompletionSource(Of Boolean)(TaskCreationOptions.RunContinuationsAsynchronously)
                        Dim read As Func(Of PreviewContentService, CancellationToken, PreviewText) =
                            Function(service, token)
                                release.Task.WaitAsync(token).GetAwaiter().GetResult()
                                Return New PreviewText("preview preview", True, 15)
                            End Function
                        Try
                            GetType(MainWindow).GetMethod("StartPlainTextPreview", Flags).Invoke(window, {read})
                            Expect(window, "Finding matches…")
                            release.TrySetResult(True)
                            PreviewTestHelpers.WaitForPreview(window)
                            Expect(window, "Match 1 out of 2")
                            Require(Scope(window) = "Full · limited preview", "The counter presented a bounded prefix as the full source.")
                            GetType(MainWindow).GetMethod("LoadTextFromDisk", Flags).Invoke(window, {path & ".missing"})
                            PreviewTestHelpers.WaitForPreview(window)
                            Expect(window, "Preview unavailable")
                            GetType(MainWindow).GetMethod("SetTextPreview", Flags).Invoke(window, {"no occurrences"})
                            PreviewTestHelpers.WaitForPreview(window)
                            Expect(window, "Match 0 out of 0")
                            Dim many = String.Concat(Enumerable.Repeat("preview ", SearchQuery.MaximumHighlights + 1))
                            GetType(MainWindow).GetMethod("SetTextPreview", Flags).Invoke(window, {many})
                            PreviewTestHelpers.WaitForPreview(window)
                            Expect(window, $"Match 1 out of {SearchQuery.MaximumHighlights:N0}+")
                            Click(window, "FindPrevious_btn")
                            Expect(window, $"Match {SearchQuery.MaximumHighlights:N0} out of {SearchQuery.MaximumHighlights:N0}+")
                            GetType(MainWindow).GetMethod("ClearTextPreview", Flags).Invoke(window, Nothing)
                            Expect(window, "Match 0 out of 0")
                        Finally
                            release.TrySetResult(True)
                        End Try
                    End Sub)
            End Sub)
    End Sub

    Public Sub RecordAnchorsAndMetadata()
        WithSource("error.log", "error" & vbLf & "anchor",
            Sub(path)
                SearchChecks.CheckPipeline(path, "error|(?=anchor)", SearchMode.RegularExpression,
                    New BeaconSettings With {.SearchFileNames = True},
                    Sub(hits) Require(hits.Count = 1, "Anchor counter fixture is missing."),
                    Sub(window, hits)
                        Dim details = DirectCast(hits(0).GetType().GetProperty("Details").GetValue(hits(0)), List(Of SearchDetail))
                        DirectCast(window.FindName("Results_lst"), ListBox).SelectedItem = hits(0)
                        PreviewTestHelpers.WaitForPreview(window)
                        Expect(window, "Match 1 out of 1")
                        Dim mode = DirectCast(window.FindName("TextPreviewMode_cmb"), ComboBox)
                        mode.SelectedIndex = 1
                        PreviewTestHelpers.WaitForPreview(window)
                        Expect(window, "Match 1 out of 2")
                        Click(window, "FindNext_btn")
                        Expect(window, "Match 2 out of 2")
                        Require(Scope(window) = "Summary · record anchor", "Zero-width Summary records were counted as visible glyphs.")
                        mode.SelectedIndex = 0
                        PreviewTestHelpers.WaitForPreview(window)
                        Require(Counter(window).Text.Contains("no visible highlight"), "Full mode retained an unrelated active index for a zero-width record.")
                        Click(window, "FindNext_btn")
                        Expect(window, "Match 1 out of 1")
                        mode.SelectedIndex = 1
                        PreviewTestHelpers.WaitForPreview(window)
                        DirectCast(window.FindName("SummaryPreview_lst"), ListBox).SelectedIndex = 0
                        PreviewTestHelpers.WaitForPreview(window)
                        Require(Scope(window) = "Name/path preview", "Metadata preview retained a Summary counter.")
                        DirectCast(window.FindName("Details_lst"), ListBox).SelectedItem = details.First(Function(detail) Not detail.IsMetadata)
                        PreviewTestHelpers.WaitForPreview(window)
                        Expect(window, "Match 1 out of 2")
                    End Sub)
            End Sub)
    End Sub

    Private Sub VerifyBar(window As MainWindow)
        Dim bar = DirectCast(window.FindName("TextCounterBar"), Border)
        Require(bar.IsVisible AndAlso bar.ActualHeight >= 34 AndAlso bar.CornerRadius.BottomLeft = 4, "Text counter bar is missing, clipped, or styled differently from EVTX.")
        Require(DirectCast(bar.Background, SolidColorBrush).Color = DirectCast(window.Resources("ButtonBackgroundBrush"), SolidColorBrush).Color,
                "Text counter surface did not follow the theme.")
        Require(DirectCast(Counter(window).Foreground, SolidColorBrush).Color = DirectCast(window.Resources("TextPrimaryBrush"), SolidColorBrush).Color,
                "Text counter foreground did not follow the theme.")
        Require(AutomationProperties.GetLiveSetting(Counter(window)) = AutomationLiveSetting.Polite, "Match position changes are not accessible.")
    End Sub

    Private Function Counter(window As MainWindow) As TextBlock
        Return DirectCast(window.FindName("TextMatchCounter_lbl"), TextBlock)
    End Function

    Private Function Scope(window As MainWindow) As String
        Return DirectCast(window.FindName("TextMatchScope_lbl"), TextBlock).Text
    End Function

    Private Sub Expect(window As MainWindow, expected As String)
        Require(Counter(window).Text = expected, $"Expected '{expected}', got '{Counter(window).Text}'.")
    End Sub

    Private Sub Click(window As MainWindow, name As String)
        DirectCast(window.FindName(name), Button).RaiseEvent(New RoutedEventArgs(Button.ClickEvent))
        PreviewTestHelpers.WaitForPreview(window)
    End Sub

    Private Sub WithSource(name As String, text As String, action As Action(Of String))
        Dim root = Path.Combine(Path.GetTempPath(), "BeaconTextCounter-" & Guid.NewGuid().ToString("N"))
        Directory.CreateDirectory(root)
        Try
            Dim path = IO.Path.Combine(root, name)
            File.WriteAllText(path, text)
            action(path)
        Finally
            Directory.Delete(root, True)
        End Try
    End Sub

    Private Sub Require(value As Boolean, message As String)
        If Not value Then Throw New InvalidOperationException(message)
    End Sub
End Module
