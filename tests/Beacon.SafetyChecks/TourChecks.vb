Imports System.IO
Imports System.Reflection
Imports System.Threading.Tasks
Imports System.Windows
Imports System.Windows.Controls
Imports Beacon

Module TourChecks
    Public Sub Run()
        Dim root = Path.Combine(Path.GetTempPath(), "BeaconTourTests-" & Guid.NewGuid().ToString("N"))
        Try
            Dim marker = Path.Combine(root, "welcome-tour.offered")
            Require(WelcomeTourState.TryClaimOffer(marker), "First launch did not claim the welcome offer.")
            Require(Not WelcomeTourState.TryClaimOffer(marker), "Second launch offered the tour again.")
            Dim concurrent = Path.Combine(root, "concurrent", "offered")
            Dim claims = Enumerable.Range(0, 8).Select(Function(index) Task.Run(Function() WelcomeTourState.TryClaimOffer(concurrent))).ToArray()
            Task.WaitAll(claims)
            Require(claims.Count(Function(task) task.Result) = 1, "Concurrent launches claimed multiple offers.")
            Dim blocked = Path.Combine(root, "blocked")
            File.WriteAllText(blocked, "file, not directory")
            Require(Not WelcomeTourState.TryClaimOffer(Path.Combine(blocked, "marker")), "Unwritable onboarding state should skip safely.")
        Finally
            If Directory.Exists(root) Then Directory.Delete(root, True)
        End Try
        For Each dark In {False, True}
            CheckWindow(dark)
        Next
        Dim splash As New SplashWindow(False)
        Try
            Dim version = GetType(SplashWindow).Assembly.GetName().Version
            Dim label = DirectCast(splash.FindName("SplashVersion_txt"), TextBlock)
            Require(label.Text = "Version " & version.ToString(If(version.Build > 0, 3, 2)), "Splash version is not assembly-derived.")
            Require(label.HorizontalAlignment = HorizontalAlignment.Right AndAlso label.VerticalAlignment = VerticalAlignment.Bottom, "Splash version is not bottom-right.")
        Finally
            splash.Close()
        End Try
    End Sub

    Private Sub CheckWindow(dark As Boolean)
        Dim owner As New MainWindow()
        Dim flags = BindingFlags.Instance Or BindingFlags.NonPublic
        Dim type = GetType(MainWindow)
        owner.RemoveHandler(FrameworkElement.LoadedEvent, type.GetMethod("MainWindow_Loaded", flags).CreateDelegate(GetType(RoutedEventHandler), owner))
        RemoveHandler owner.Closing, DirectCast(type.GetMethod("MainWindow_Closing", flags).CreateDelegate(GetType(ComponentModel.CancelEventHandler), owner), ComponentModel.CancelEventHandler)
        Dim tour As TourWindow = Nothing
        Try
            type.GetMethod(If(dark, "ApplyDarkTheme", "ApplyLightTheme"), flags).Invoke(owner, Nothing)
            owner.WindowStartupLocation = WindowStartupLocation.Manual
            owner.WindowState = WindowState.Normal
            owner.Width = 1000
            owner.Height = 700
            owner.Left = -10000
            owner.Top = -10000
            owner.ShowActivated = False
            owner.ShowInTaskbar = False
            owner.Show()
            For stopAt = -1 To WelcomeTour.Steps().Length - 1
                tour = New TourWindow(owner, dark)
                tour.Show()
                Pump(tour)
                Require(tour.WindowStyle = WindowStyle.None AndAlso tour.AllowsTransparency AndAlso Not tour.ShowInTaskbar,
                        "Tour must display as a borderless coaching card, not a native mini window.")
                Require(DirectCast(tour.Background, System.Windows.Media.SolidColorBrush).Color.A = 0,
                        "Tour window background must preserve transparent rounded corners.")
                Dim card = DirectCast(tour.FindName("TourCard"), Border)
                Require(card.BorderThickness = New Thickness(0) AndAlso card.CornerRadius.TopLeft = 6,
                        "Tour card must keep rounded corners without a visible border.")
                Require(DirectCast(card.Background, System.Windows.Media.SolidColorBrush).Color =
                        DirectCast(owner.Resources("CardBackgroundBrush"), System.Windows.Media.SolidColorBrush).Color,
                        "Borderless tour card must follow the owner theme.")
                Require(tour.StepIndex = -1 AndAlso CStr(DirectCast(tour.FindName("ExitTour_btn"), Button).Content) = "Not now", "Welcome must be optional.")
                Dim reminder = DirectCast(tour.FindName("TourReminder_txt"), TextBlock)
                Require(reminder.IsVisible, "Welcome should explain that the tour is optional.")
                Dim nextButton = DirectCast(tour.FindName("NextTour_btn"), Button)
                For index = 0 To stopAt
                    nextButton.RaiseEvent(New RoutedEventArgs(Button.ClickEvent))
                    Pump(tour)
                    Require(tour.StepIndex = index, "Tour step progression failed.")
                    Require(reminder.Visibility = Visibility.Collapsed, "Optional-tour reminder repeated after the welcome.")
                    Require(Not DirectCast(tour.FindName("TourBody_txt"), TextBlock).Text.Contains("you cannot type into it"), "Tour retained unnecessary Path instruction.")
                    Require(owner.FindName(WelcomeTour.Steps()(index).Target) IsNot Nothing, "Tour target is not a real control.")
                Next
                If stopAt > 0 Then
                    DirectCast(tour.FindName("BackTour_btn"), Button).RaiseEvent(New RoutedEventArgs(Button.ClickEvent))
                    Require(tour.StepIndex = stopAt - 1, "Back navigation failed.")
                    Require(reminder.Visibility = Visibility.Collapsed, "Back navigation restored the welcome-only reminder.")
                    nextButton.RaiseEvent(New RoutedEventArgs(Button.ClickEvent))
                End If
                Require(owner.IsEnabled AndAlso DirectCast(owner.FindName("Search_txt"), TextBox).Text = "", "Tour changed input or blocked the app.")
                If stopAt = WelcomeTour.Steps().Length - 1 Then
                    Require(CStr(nextButton.Content) = "Finish", "Final tour action is not Finish.")
                    nextButton.RaiseEvent(New RoutedEventArgs(Button.ClickEvent))
                Else
                    DirectCast(tour.FindName("ExitTour_btn"), Button).RaiseEvent(New RoutedEventArgs(Button.ClickEvent))
                End If
                Require(Not tour.IsVisible, "Tour cannot be exited at this stage.")
            Next
            tour = New TourWindow(owner, dark)
            tour.Show()
            owner.Close()
            Require(Not tour.IsVisible, "Tour remained open after owner close.")
        Finally
            If tour IsNot Nothing AndAlso tour.IsVisible Then tour.Close()
            If owner.IsVisible Then owner.Close()
        End Try
    End Sub
    Private Sub Pump(window As Window)
        window.UpdateLayout()
        window.Dispatcher.Invoke(Sub()
                                 End Sub, System.Windows.Threading.DispatcherPriority.ContextIdle)
    End Sub
    Private Sub Require(value As Boolean, message As String)
        If Not value Then Throw New InvalidOperationException(message)
    End Sub
End Module
