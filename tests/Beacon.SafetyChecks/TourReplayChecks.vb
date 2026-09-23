Imports System.Diagnostics
Imports System.IO
Imports System.Reflection
Imports System.Threading
Imports System.Threading.Tasks
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Threading
Imports Beacon

Module TourReplayChecks
    Private Const Flags As BindingFlags = BindingFlags.Instance Or BindingFlags.NonPublic

    Public Sub MarkerIsolation()
        Dim root = Path.Combine(Path.GetTempPath(), "BeaconReplay-" & Guid.NewGuid().ToString("N"))
        Directory.CreateDirectory(root)
        Try
            For Each scenario In {"existing", "missing", "blocked"}
                Dim marker = Path.Combine(root, scenario, "welcome-tour.offered")
                If scenario = "existing" Then Require(WelcomeTourState.TryClaimOffer(marker), "Marker fixture was not created.")
                If scenario = "blocked" Then File.WriteAllText(Path.Combine(root, scenario), "not a directory")
                Dim before = If(File.Exists(marker), File.ReadAllBytes(marker), Array.Empty(Of Byte)())
                Dim written = If(File.Exists(marker), File.GetLastWriteTimeUtc(marker), DateTime.MinValue)
                Dim claims As Integer = 0
                Dim window = NewWindow()
                Dim claim As Func(Of Boolean) = Function()
                                                    claims += 1
                                                    Return WelcomeTourState.TryClaimOffer(marker)
                                                End Function
                GetType(MainWindow).GetField("_claimWelcomeOffer", Flags).SetValue(window, claim)
                Try
                    DirectCast(window.FindName("Search_txt"), TextBox).Text = "unchanged search"
                    For iteration = 0 To 2
                        ReplayThroughHelp(window)
                        Dim first = Tour(window)
                        Require(first IsNot Nothing AndAlso first.IsVisible AndAlso first.StepIndex = -1, "Manual replay did not start a fresh tour.")
                        DirectCast(first.FindName("NextTour_btn"), Button).RaiseEvent(New RoutedEventArgs(Button.ClickEvent))
                        ReplayThroughHelp(window)
                        Require(Tour(window) Is first AndAlso first.StepIndex = 0, "Repeated replay opened a duplicate or reset an active tour.")
                        If iteration = 1 Then
                            For index = first.StepIndex To WelcomeTour.Steps().Length - 1
                                DirectCast(first.FindName("NextTour_btn"), Button).RaiseEvent(New RoutedEventArgs(Button.ClickEvent))
                            Next
                        Else
                            first.Close()
                        End If
                        Pump(window)
                        Require(Tour(window) Is Nothing, "Closed/completed tour was retained.")
                    Next
                    Require(claims = 0 AndAlso DirectCast(window.FindName("Search_txt"), TextBox).Text = "unchanged search", "Manual replay invoked startup eligibility or changed search text.")
                    If scenario = "existing" Then
                        Require(File.ReadAllBytes(marker).SequenceEqual(before) AndAlso File.GetLastWriteTimeUtc(marker) = written, "Manual replay rewrote the marker.")
                    Else
                        Require(Not File.Exists(marker), "Manual replay created a marker.")
                    End If
                    GetType(MainWindow).GetMethod("OfferWelcomeTour", Flags).Invoke(window, Nothing)
                    GetType(MainWindow).GetMethod("OfferWelcomeTour", Flags).Invoke(window, Nothing)
                    Require(claims = 1, "Automatic offering no longer checks once per window.")
                    Require((Tour(window) IsNot Nothing) = (scenario = "missing"), "Automatic marker semantics changed.")
                    ReplayThroughHelp(window)
                    Dim openTour = Tour(window)
                    window.Close()
                    WaitUntil(window, Function() Not window.IsVisible)
                    Require(Not openTour.IsVisible, "Tour outlived its owner.")
                Finally
                    If window.IsVisible Then
                        window.Close()
                        WaitUntil(window, Function() Not window.IsVisible)
                    End If
                End Try
            Next
        Finally
            Directory.Delete(root, True)
        End Try
    End Sub

    Public Sub ActiveSearch()
        Dim root = Path.Combine(Path.GetTempPath(), "BeaconReplayScan-" & Guid.NewGuid().ToString("N"))
        Directory.CreateDirectory(root)
        Dim release As New TaskCompletionSource(Of Boolean)(TaskCreationOptions.RunContinuationsAsynchronously)
        Dim window As MainWindow = Nothing
        Try
            For index = 0 To 15
                File.WriteAllText(Path.Combine(root, $"{index:D2}.log"), $"before{vbLf}error {index}{vbLf}after")
            Next
            window = NewWindow()
            Dim claims As Integer = 0
            Dim claim As Func(Of Boolean) = Function()
                                                claims += 1
                                                Return False
                                            End Function
            GetType(MainWindow).GetField("_claimWelcomeOffer", Flags).SetValue(window, claim)
            window.BeforeSourceForChecks = Async Function(path, ordinal, token)
                                               If ordinal > 0 Then Await release.Task.WaitAsync(token).ConfigureAwait(False)
                                           End Function
            DirectCast(window.FindName("Path_txt"), TextBox).Text = root
            DirectCast(window.FindName("Search_txt"), TextBox).Text = "error"
            GetType(MainWindow).GetMethod("StartScan", Flags).Invoke(window, Nothing)
            Dim scan = DirectCast(GetType(MainWindow).GetField("_scanTask", Flags).GetValue(window), Task)
            Dim results = DirectCast(window.FindName("Results_lst"), ListBox)
            WaitUntil(window, Function() results.Items.Count > 0 OrElse scan.IsCompleted)
            Require(Not scan.IsCompleted AndAlso results.Items.Count = 1, "Controlled search did not remain active.")
            results.SelectedIndex = 0
            PreviewTestHelpers.WaitForPreview(window)
            DirectCast(window.FindName("TextPreviewMode_cmb"), ComboBox).SelectedIndex = 1
            PreviewTestHelpers.WaitForPreview(window)
            ReplayThroughHelp(window)
            Require(Not scan.IsCompleted AndAlso Tour(window).IsVisible AndAlso claims = 0, "Replay disturbed the active scan or invoked startup state.")
            Require(DirectCast(window.FindName("SummaryPreview_lst"), ListBox).IsVisible AndAlso
                    DirectCast(window.FindName("Search_txt"), TextBox).Text = "error", "Replay changed preview mode or search input.")
            Dim source = DirectCast(GetType(MainWindow).GetField("_scanCts", Flags).GetValue(window), CancellationTokenSource)
            Require(Not source.IsCancellationRequested, "Replay cancelled the search.")
            release.TrySetResult(True)
            WaitUntil(window, Function() scan.IsCompleted)
            scan.GetAwaiter().GetResult()
            Require(results.Items.Count = 16 AndAlso CInt(GetType(MainWindow).GetField("_filesScanned", Flags).GetValue(window)) = 16,
                    "Concurrent replay changed results or completed progress.")
            Require(DirectCast(window.FindName("SummaryPreview_lst"), ListBox).IsVisible, "Search completion reset the session preview mode.")
            window.Close()
            WaitUntil(window, Function() Not window.IsVisible)
        Finally
            release.TrySetResult(True)
            If window IsNot Nothing AndAlso window.IsVisible Then
                DirectCast(GetType(MainWindow).GetField("_scanCts", Flags).GetValue(window), CancellationTokenSource)?.Cancel()
                window.Close()
                WaitUntil(window, Function() Not window.IsVisible)
            End If
            Directory.Delete(root, True)
        End Try
    End Sub

    Private Function NewWindow() As MainWindow
        Dim window As New MainWindow()
        window.ApplicationExitForChecks = Sub() window.Dispatcher.BeginInvoke(New Action(AddressOf window.Close), DispatcherPriority.Background)
        window.RemoveHandler(FrameworkElement.LoadedEvent,
            GetType(MainWindow).GetMethod("MainWindow_Loaded", Flags).CreateDelegate(GetType(RoutedEventHandler), window))
        Dim options As New BeaconSettings()
        GetType(MainWindow).GetField("_settings", Flags).SetValue(window, options)
        GetType(MainWindow).GetMethod("ApplySettings", Flags).Invoke(window, Nothing)
        window.WindowStartupLocation = WindowStartupLocation.Manual
        window.WindowState = WindowState.Normal
        window.Width = 1000
        window.Height = 700
        window.Left = -10000
        window.Top = -10000
        window.ShowActivated = False
        window.ShowInTaskbar = False
        window.Show()
        Pump(window)
        Return window
    End Function

    Private Sub ReplayThroughHelp(window As MainWindow)
        GetType(MainWindow).GetMethod("OpenHelp", Flags).Invoke(window, {Nothing, New RoutedEventArgs()})
        Dim help = DirectCast(GetType(MainWindow).GetField("_helpWindow", Flags).GetValue(window), HelpWindow)
        DirectCast(help.FindName("ReplayTour_btn"), Button).RaiseEvent(New RoutedEventArgs(Button.ClickEvent))
        Pump(window)
    End Sub

    Private Function Tour(window As MainWindow) As TourWindow
        Return DirectCast(GetType(MainWindow).GetField("_tourWindow", Flags).GetValue(window), TourWindow)
    End Function

    Private Sub WaitUntil(window As MainWindow, predicate As Func(Of Boolean))
        Dim timeout = Stopwatch.StartNew()
        While Not predicate()
            Pump(window)
            If timeout.Elapsed > TimeSpan.FromSeconds(30) Then Throw New TimeoutException("Combined tour/search lifecycle did not finish.")
            Thread.Sleep(1)
        End While
        Pump(window)
    End Sub

    Private Sub Pump(window As Window)
        window.Dispatcher.Invoke(Sub()
                                 End Sub, DispatcherPriority.ContextIdle)
    End Sub

    Private Sub Require(value As Boolean, message As String)
        If Not value Then Throw New InvalidOperationException(message)
    End Sub
End Module
