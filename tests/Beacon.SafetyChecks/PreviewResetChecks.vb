Imports System.Collections
Imports System.Diagnostics
Imports System.IO
Imports System.Reflection
Imports System.Threading
Imports System.Threading.Tasks
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Threading
Imports Beacon

Module PreviewResetChecks
    Private Const Flags As BindingFlags = BindingFlags.Instance Or BindingFlags.NonPublic
    Private ReadOnly PreviewButtons As String() = {"FindPrevious_btn", "FindNext_btn", "FindPreviousEvent_btn", "FindNextEvent_btn",
                                                   "FindPreviousHarRequest_btn", "FindNextHarRequest_btn"}

    Public Sub Run()
        Dim root = Path.Combine(Path.GetTempPath(), "BeaconPreviewReset-" & Guid.NewGuid().ToString("N"))
        Directory.CreateDirectory(root)
        Try
            Dim textPath = Path.Combine(root, "sample.log")
            File.WriteAllText(textPath, "before" & vbLf & "error" & vbLf & "after")
            Dim harPath = Path.Combine(root, "sample.har")
            File.WriteAllText(harPath, "{""log"":{""entries"":[{""request"":{""method"":""GET"",""url"":""https://example.invalid/""},""response"":{""status"":503,""content"":{""text"":""error""}}}]}}")
            For Each kind In {"DiskTextFile", "EvtxFileOnDisk", "HarFileOnDisk"}
                Dim source = If(kind = "HarFileOnDisk", harPath, textPath)
                SearchChecks.CheckPipeline(source, "error", SearchMode.PlainText, New BeaconSettings(),
                    Sub(hits) Require(hits.Count = 1, "Reset fixture did not produce one result."),
                    Sub(window, hits)
                        Dim hit = hits(0)
                        If kind = "EvtxFileOnDisk" Then
                            Dim kindProperty = hit.GetType().GetProperty("Kind")
                            kindProperty.SetValue(hit, [Enum].Parse(kindProperty.PropertyType, kind))
                            Dim events = DirectCast(hit.GetType().GetProperty("MatchingEvents").GetValue(hit), IList)
                            events.Add(New EventRecordSummary With {.EventId = 42, .Provider = "Synthetic reset fixture",
                                .Message = "error", .RawXml = "<Event>error</Event>"})
                        End If
                        Dim results = DirectCast(window.FindName("Results_lst"), ListBox)
                        results.SelectedItem = hit
                        PreviewTestHelpers.WaitForPreview(window)
                        Dim previousName = If(kind = "EvtxFileOnDisk", "FindPreviousEvent_btn", If(kind = "HarFileOnDisk", "FindPreviousHarRequest_btn", "FindPrevious_btn"))
                        Require(DirectCast(window.FindName(previousName), Button).IsVisible, "The expected previous button was not displayed before Reset.")
                        Dim modePanel = DirectCast(window.FindName("TextPreviewModePanel"), FrameworkElement)
                        Require(modePanel.IsVisible = (kind = "DiskTextFile"), "Summary availability does not match the selected preview type.")

                        results.SelectedIndex = -1
                        PreviewTestHelpers.WaitForPreview(window)
                        RequireHidden(window)
                        results.SelectedItem = hit
                        PreviewTestHelpers.WaitForPreview(window)
                        Require(DirectCast(window.FindName(previousName), Button).IsVisible, "Reselecting a result did not restore its navigation.")
                        DirectCast(window.FindName("NextFile_btn"), Button).Visibility = Visibility.Visible
                        Reset(window)
                        VerifyEmpty(window)

                        For Each name In PreviewButtons
                            DirectCast(window.FindName(name), Button).Visibility = Visibility.Visible
                        Next
                        DirectCast(window.FindName("Path_txt"), TextBox).Text = root
                        DirectCast(window.FindName("Search_txt"), TextBox).Text = "error"
                        GetType(MainWindow).GetMethod("StartScan", Flags).Invoke(window, Nothing)
                        RequireHidden(window)
                        Dim scan = DirectCast(GetType(MainWindow).GetField("_scanTask", Flags).GetValue(window), Task)
                        Require(scan IsNot Nothing, "New-scan reset fixture did not start.")
                        WaitUntil(window, Function() scan.IsCompleted)
                        scan.GetAwaiter().GetResult()
                        PreviewTestHelpers.WaitForPreview(window)
                        Reset(window)
                        VerifyEmpty(window)
                    End Sub)
            Next
        Finally
            Directory.Delete(root, True)
        End Try
    End Sub

    Private Sub Reset(window As MainWindow)
        GetType(MainWindow).GetMethod("Reset_btn_Click", Flags).Invoke(window, {Nothing, New RoutedEventArgs()})
        RequireHidden(window)
        WaitUntil(window, Function() Not CBool(GetType(MainWindow).GetField("_isResetting", Flags).GetValue(window)))
        PreviewTestHelpers.WaitForPreview(window)
    End Sub

    Private Sub VerifyEmpty(window As MainWindow)
        RequireHidden(window)
        Require(DirectCast(window.FindName("NextFile_btn"), Button).Visibility = Visibility.Collapsed, "Next File survived Reset.")
        Require(DirectCast(window.FindName("Results_lst"), ListBox).Items.Count = 0 AndAlso
                DirectCast(window.FindName("Details_lst"), ListBox).Items.Count = 0, "Reset retained results or details.")
        Require(DirectCast(window.FindName("TextMatchCounter_lbl"), TextBlock).Text = "Match 0 out of 0", "Reset retained the match counter.")
        Require(DirectCast(window.FindName("TextPreviewModePanel"), FrameworkElement).Visibility = Visibility.Collapsed, "An empty preview retained its mode selector.")
    End Sub

    Private Sub RequireHidden(window As MainWindow)
        For Each name In PreviewButtons
            Require(DirectCast(window.FindName(name), Button).Visibility = Visibility.Collapsed, name & " remained visible after clearing navigation.")
        Next
    End Sub

    Private Sub WaitUntil(window As MainWindow, ready As Func(Of Boolean))
        Dim timeout = Stopwatch.StartNew()
        While Not ready()
            window.Dispatcher.Invoke(Sub()
                                     End Sub, DispatcherPriority.Background)
            If timeout.Elapsed > TimeSpan.FromSeconds(20) Then Throw New TimeoutException("Reset/scan lifecycle did not complete.")
            Thread.Sleep(1)
        End While
        window.Dispatcher.Invoke(Sub()
                                 End Sub, DispatcherPriority.Background)
    End Sub

    Private Sub Require(value As Boolean, message As String)
        If Not value Then Throw New InvalidOperationException(message)
    End Sub
End Module
