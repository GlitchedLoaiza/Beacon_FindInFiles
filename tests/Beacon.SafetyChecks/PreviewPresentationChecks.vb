Imports System.Diagnostics
Imports System.IO
Imports System.Reflection
Imports System.Windows
Imports System.Windows.Automation
Imports System.Windows.Controls
Imports System.Windows.Media
Imports Beacon

Module PreviewPresentationChecks
    Public Sub Run()
        Dim root = Path.Combine(Path.GetTempPath(), "BeaconSummaryLayout-" & Guid.NewGuid().ToString("N"))
        Directory.CreateDirectory(root)
        Try
            Dim path = IO.Path.Combine(root, "dense.log")
            File.WriteAllLines(path, Enumerable.Range(1, 10000).Select(Function(number) $"error {number}: " & New String("x"c, 80)))
            For Each theme In {AppTheme.Light, AppTheme.Dark, AppTheme.Beacon}
                Dim options As New BeaconSettings With {.Theme = theme, .MaximumStructuredMatches = 10000,
                    .PreviewFontFamily = "Consolas", .PreviewFontSize = 14, .PreviewWordWrap = theme <> AppTheme.Dark}
                SearchChecks.CheckPipeline(path, "error", SearchMode.PlainText, options,
                    Sub(hits) Require(hits.Count = 1, "Large-summary fixture missing."),
                    Sub(window, hits)
                        window.Width = 1000
                        window.Height = 700
                        GetType(MainWindow).GetField("_summaryPreferred", BindingFlags.Instance Or BindingFlags.NonPublic).SetValue(window, True)
                        Dim allocated = GC.GetTotalAllocatedBytes(True)
                        Dim watch = Stopwatch.StartNew()
                        DirectCast(window.FindName("Results_lst"), ListBox).SelectedItem = hits(0)
                        PreviewTestHelpers.WaitForPreview(window)
                        Dim summary = DirectCast(window.FindName("SummaryPreview_lst"), ListBox)
                        Dim selector = DirectCast(window.FindName("TextPreviewMode_cmb"), ComboBox)
                        Require(summary.Items.Count = 10000 AndAlso summary.ActualHeight > 0 AndAlso summary.ActualWidth > 0, "Large Summary was truncated or clipped out of the preview.")
                        Require(summary.FontSize = 14 AndAlso summary.FontFamily.Source = "Consolas" AndAlso CBool(summary.Tag) = options.PreviewWordWrap,
                                "Summary ignored font or wrapping preferences.")
                        Require(VirtualizingPanel.GetIsVirtualizing(summary) AndAlso ScrollViewer.GetCanContentScroll(summary), "Summary virtualization was disabled.")
                        Require(Not String.IsNullOrWhiteSpace(AutomationProperties.GetName(selector)), "Preview mode has no accessible name.")
                        VerifyRealized(window, summary)
                        Dim initialMs = watch.Elapsed.TotalMilliseconds
                        summary.SelectedIndex = summary.Items.Count - 1
                        PreviewTestHelpers.WaitForPreview(window)
                        summary.UpdateLayout()
                        VerifyRealized(window, summary)
                        Dim block = DirectCast(summary.SelectedItem, TextSummaryBlock)
                        Require(block.DetailIndex = 9999 AndAlso block.IsCurrent, "The last captured record is not reachable.")
                        DirectCast(window.FindName("FindNext_btn"), Button).RaiseEvent(New RoutedEventArgs(Button.ClickEvent))
                        PreviewTestHelpers.WaitForPreview(window)
                        Require(DirectCast(summary.SelectedItem, TextSummaryBlock).DetailIndex = 0, "Large-summary wrap lost the first record.")
                        VerifyRealized(window, summary)
                        Console.WriteLine($"SUMMARY {theme}: 10,000 records; initial={initialMs:F1} ms; render/scroll={watch.Elapsed.TotalMilliseconds:F1} ms; allocated={GC.GetTotalAllocatedBytes(True) - allocated:N0} bytes; realized={Descendants(summary).OfType(Of TextSummaryBlockView)().Count()}.")
                    End Sub)
            Next
        Finally
            Directory.Delete(root, True)
        End Try
    End Sub

    Private Sub VerifyRealized(window As MainWindow, summary As ListBox)
        Dim views = Descendants(summary).OfType(Of TextSummaryBlockView)().ToArray()
        Require(views.Length > 0 AndAlso views.Length < 100, "Summary eagerly realized its captured records.")
        Dim expected = DirectCast(window.Resources("TextPrimaryBrush"), SolidColorBrush).Color
        Dim secondary = DirectCast(window.Resources("TextSecondaryBrush"), SolidColorBrush).Color
        For Each view In views
            Dim border = TryCast(view.Content, Border)
            Require(border IsNot Nothing AndAlso border.CornerRadius.TopLeft = 4, "Summary does not preserve rounded theme-aware cards.")
            Require(DirectCast(border.Background, SolidColorBrush).Color = DirectCast(window.Resources("CardBackgroundBrush"), SolidColorBrush).Color,
                    "Summary background did not follow the selected theme.")
            For Each label In Descendants(view).OfType(Of TextBlock)()
                Dim foreground = TryCast(label.Foreground, SolidColorBrush)
                Require(foreground IsNot Nothing AndAlso (foreground.Color = expected OrElse foreground.Color = secondary), "Summary text lost theme-aware foreground colors.")
            Next
        Next
    End Sub

    Private Iterator Function Descendants(parent As DependencyObject) As IEnumerable(Of DependencyObject)
        For index = 0 To VisualTreeHelper.GetChildrenCount(parent) - 1
            Dim child = VisualTreeHelper.GetChild(parent, index)
            Yield child
            For Each descendant In Descendants(child)
                Yield descendant
            Next
        Next
    End Function

    Private Sub Require(value As Boolean, message As String)
        If Not value Then Throw New InvalidOperationException(message)
    End Sub
End Module
