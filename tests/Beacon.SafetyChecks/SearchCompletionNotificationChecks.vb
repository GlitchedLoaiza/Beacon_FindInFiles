Imports System.ComponentModel
Imports System.Diagnostics
Imports System.IO
Imports System.Reflection
Imports System.Runtime.InteropServices
Imports System.Threading
Imports System.Threading.Tasks
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Threading
Imports Beacon

Module SearchCompletionNotificationChecks
    Private Const PrivateInstance As BindingFlags = BindingFlags.Instance Or BindingFlags.NonPublic

    Public Sub Policy()
        Require(SearchCompletionNotifier.NotificationSoundAlias = "Notification.IM", "Completion alerts no longer use the Windows messaging notification event.")
        Require(SearchCompletionNotifier.NotificationSoundFlags = &H210003UI, "System-sound playback must be asynchronous, alias-based, system-volume controlled, and have no default fallback.")
        Dim native = GetType(SearchCompletionNotifier).GetNestedType("WindowsNotificationPlatform", BindingFlags.NonPublic)
        For Each method In native.GetMethods(BindingFlags.Static Or BindingFlags.NonPublic)
            If method.GetCustomAttribute(Of DllImportAttribute)() Is Nothing Then Continue For
            Dim searchPaths = method.GetCustomAttribute(Of DefaultDllImportSearchPathsAttribute)()
            Require(searchPaths IsNot Nothing AndAlso searchPaths.Paths = DllImportSearchPath.System32, "Notification APIs must be loaded from Windows System32.")
        Next
        Dim platform As New FakePlatform()
        Dim settings As New BeaconSettings()
        Using notifier As New SearchCompletionNotifier(Nothing, platform)
            notifier.Complete(1, ScanReportState.Completed, settings)
            Require(platform.SoundCalls = 0, "A completion without an active search notified.")
            notifier.BeginSearch(1)
            notifier.Complete(1, ScanReportState.Running, settings)
            Require(platform.SoundCalls = 0, "A running search notified.")
            notifier.Complete(1, ScanReportState.Completed, settings)
            notifier.Complete(1, ScanReportState.Completed, settings)
            notifier.BeginSearch(1)
            notifier.Complete(1, ScanReportState.Completed, settings)
            Require(platform.SoundCalls = 1 AndAlso platform.FlashCalls = 1, "One search produced duplicate completion alerts.")

            notifier.BeginSearch(2)
            Require(platform.StopCalls = 1, "A new search did not clear prior taskbar attention.")
            notifier.Complete(1, ScanReportState.Completed, settings)
            notifier.Complete(2, ScanReportState.Cancelled, settings)
            notifier.Complete(2, ScanReportState.Completed, settings)
            notifier.BeginSearch(3)
            notifier.Complete(3, ScanReportState.Failed, settings)
            Require(platform.SoundCalls = 1 AndAlso platform.FlashCalls = 1, "A stale, canceled, or failed search notified.")

            notifier.BeginSearch(4)
            notifier.Complete(4, ScanReportState.ResultLimitReached, settings)
            Require(platform.SoundCalls = 2 AndAlso platform.FlashCalls = 2, "A result-limit completion was not announced.")
            notifier.CancelPending()
            notifier.BeginSearch(5)
            notifier.CancelPending()
            notifier.Complete(5, ScanReportState.Completed, settings)
            Require(platform.SoundCalls = 2 AndAlso platform.StopCalls = 2, "Resetting did not suppress pending completion attention.")

            settings.NotifyOnSearchCompletion = False
            notifier.BeginSearch(6)
            notifier.Complete(6, ScanReportState.Completed, settings)
            settings.NotifyOnSearchCompletion = True
            platform.Foreground = True
            notifier.BeginSearch(7)
            notifier.Complete(7, ScanReportState.Completed, settings)
            platform.Foreground = False
            notifier.Complete(7, ScanReportState.Completed, settings)
            Require(platform.SoundCalls = 2, "Disabled or foreground-suppressed completion was replayed later.")

            platform.Foreground = True
            settings.NotifyOnlyInBackground = False
            notifier.BeginSearch(8)
            notifier.Complete(8, ScanReportState.Completed, settings)
            Require(platform.SoundCalls = 3 AndAlso platform.FlashCalls = 2, "Foreground alerts should play only the sound when explicitly enabled.")
            platform.Foreground = False
            platform.Available = False
            notifier.BeginSearch(9)
            notifier.Complete(9, ScanReportState.Completed, settings)
            platform.Available = True
            notifier.Complete(9, ScanReportState.Completed, settings)
            Require(platform.SoundCalls = 3, "A hidden/headless completion notified or was replayed.")
            notifier.BeginSearch(10)
            notifier.Complete(10, ScanReportState.Completed, settings)
            Require(platform.SoundCalls = 4 AndAlso platform.FlashCalls = 3, "A later search lost its completion notification.")
        End Using
        Require(platform.StopCalls = 3, "Disposal did not clear taskbar attention.")
    End Sub

    Public Sub LifecycleAndFailures()
        EnsureApplication()
        Dim window As New Window With {.ShowInTaskbar = False, .ShowActivated = False, .Width = 200, .Height = 100}
        Prepare(window)
        Dim platform As New FakePlatform With {.SoundFault = True}
        Dim notifier As New SearchCompletionNotifier(window, platform)
        Try
            window.Show()
            notifier.BeginSearch(1)
            notifier.Complete(1, ScanReportState.Completed, New BeaconSettings())
            Require(platform.SoundCalls = 1 AndAlso platform.FlashCalls = 1, "A sound failure prevented taskbar notification.")
            GetType(Window).GetMethod("OnActivated", PrivateInstance).Invoke(window, {EventArgs.Empty})
            Require(platform.StopCalls = 1, "Window activation did not stop taskbar flashing.")
            notifier.BeginSearch(2)
            notifier.Complete(2, ScanReportState.Completed, New BeaconSettings())
            GetType(Application).GetMethod("OnActivated", PrivateInstance).Invoke(Application.Current, {EventArgs.Empty})
            Require(platform.StopCalls = 2, "Activating an owned/application window did not acknowledge attention.")
            notifier.BeginSearch(3)
            notifier.Complete(3, ScanReportState.Completed, New BeaconSettings())
            window.Close()
            Require(platform.StopCalls = 3, "Closing the window did not clear attention.")
            notifier.BeginSearch(4)
            notifier.Complete(4, ScanReportState.Completed, New BeaconSettings())
            GetType(Application).GetMethod("OnActivated", PrivateInstance).Invoke(Application.Current, {EventArgs.Empty})
            Require(platform.SoundCalls = 3 AndAlso platform.StopCalls = 3, "Disposed notifications retained activation handlers or emitted an alert.")
        Finally
            notifier.Dispose()
            window.Close()
        End Try

        Dim failing As New FakePlatform With {.FlashFault = True}
        Using alerts As New SearchCompletionNotifier(Nothing, failing)
            alerts.BeginSearch(1)
            alerts.Complete(1, ScanReportState.Completed, New BeaconSettings())
            Require(failing.SoundCalls = 1 AndAlso failing.FlashCalls = 1, "A taskbar failure changed sound behavior.")
            failing.FlashFault = False
            alerts.BeginSearch(2)
            alerts.Complete(2, ScanReportState.Completed, New BeaconSettings())
            failing.StopFault = True
            alerts.CancelPending()
            failing.PolicyFault = True
            alerts.BeginSearch(3)
            alerts.Complete(3, ScanReportState.Completed, New BeaconSettings())
            Require(failing.SoundCalls = 2, "An eligibility failure emitted an alert.")
        End Using

        Dim hidden As New Window With {.ShowInTaskbar = False}
        Using alerts As New SearchCompletionNotifier(hidden)
            Dim native = DirectCast(GetType(SearchCompletionNotifier).GetField("_platform", PrivateInstance).GetValue(alerts), ISearchCompletionNotificationPlatform)
            Require(Not native.CanNotify, "An unshown window is eligible for native notifications.")
            Prepare(hidden)
            hidden.Show()
            Require(Not native.CanNotify, "An off-screen non-taskbar window is eligible for native notifications.")
            hidden.Close()
        End Using
    End Sub

    Public Sub Settings()
        EnsureApplication()
        Dim defaults = BeaconSettings.CreateDefaults()
        Dim legacy = BeaconSettingsService.Deserialize("{}")
        Require(defaults.NotifyOnSearchCompletion AndAlso defaults.NotifyOnlyInBackground AndAlso
                legacy.NotifyOnSearchCompletion AndAlso legacy.NotifyOnlyInBackground, "New and existing settings should default to background-only alerts.")
        For Each enabled In {False, True}
            For Each backgroundOnly In {False, True}
                Dim settings = BeaconSettingsService.Validate(BeaconSettingsService.Clone(New BeaconSettings With {
                    .NotifyOnSearchCompletion = enabled, .NotifyOnlyInBackground = backgroundOnly
                }))
                Require(settings.NotifyOnSearchCompletion = enabled AndAlso settings.NotifyOnlyInBackground = backgroundOnly, "Notification preferences did not survive serialization.")
            Next
        Next
        For Each theme In {AppTheme.Light, AppTheme.Dark, AppTheme.Aero}
            Dim original As New BeaconSettings With {.Theme = theme, .NotifyOnSearchCompletion = False, .NotifyOnlyInBackground = False}
            Dim window As New SettingsWindow(original, theme <> AppTheme.Light)
            Prepare(window)
            Try
                window.Show()
                DirectCast(window.FindName("SettingsTabs"), TabControl).SelectedIndex = 4
                Pump(window)
                Dim enabled = DirectCast(window.FindName("NotifyOnCompletion_chk"), CheckBox)
                Dim backgroundOnly = DirectCast(window.FindName("NotifyOnlyInBackground_chk"), CheckBox)
                Require(Not enabled.IsChecked.GetValueOrDefault() AndAlso Not backgroundOnly.IsChecked.GetValueOrDefault() AndAlso Not backgroundOnly.IsEnabled,
                        "Settings did not load disabled notification preferences.")
                enabled.IsChecked = True
                Pump(window)
                Require(backgroundOnly.IsEnabled, "Enabling alerts did not enable the background-only option.")
                backgroundOnly.IsChecked = True
                Require(Not original.NotifyOnSearchCompletion AndAlso Not original.NotifyOnlyInBackground AndAlso window.SavedSettings Is Nothing,
                        "Editing notification controls changed saved settings before Save.")
                DirectCast(window.FindName("RestoreDefaults_btn"), Button).RaiseEvent(New RoutedEventArgs(Button.ClickEvent))
                Pump(window)
                Require(enabled.IsChecked.GetValueOrDefault() AndAlso backgroundOnly.IsChecked.GetValueOrDefault(), "Restore defaults did not restore background-only completion alerts.")
            Finally
                window.Close()
            End Try
        Next
    End Sub

    Public Sub ScanIntegration()
        EnsureApplication()
        Dim root = Path.Combine(Path.GetTempPath(), "BeaconCompletionTests-" & Guid.NewGuid().ToString("N"))
        Directory.CreateDirectory(root)
        Dim platform As New FakePlatform()
        Dim window As New MainWindow(platform)
        Prepare(window)
        DetachStartup(window)
        Dim settings As New BeaconSettings With {.IncludedExtensions = ".log", .SelectFirstResult = False, .NotifyOnSearchCompletion = True, .NotifyOnlyInBackground = True}
        GetType(MainWindow).GetField("_settings", PrivateInstance).SetValue(window, settings)
        GetType(MainWindow).GetMethod("ApplySettings", PrivateInstance).Invoke(window, Nothing)
        Dim observedFinalState As Boolean
        platform.OnSound = Sub()
                               Dim run = DirectCast(GetType(MainWindow).GetField("_scanRun", PrivateInstance).GetValue(window), ScanRunInfo)
                               observedFinalState = Not CBool(GetType(MainWindow).GetField("_isScanning", PrivateInstance).GetValue(window)) AndAlso
                                   run.CompletedUtc.HasValue AndAlso DirectCast(window.FindName("ScanProgress_pb"), ProgressBar).Visibility = Visibility.Collapsed
                           End Sub
        Try
            File.WriteAllText(Path.Combine(root, "one.log"), "needle")
            File.WriteAllText(Path.Combine(root, "two.log"), "needle again")
            window.Show()
            Pump(window)
            StartSearch(window, root, "needle")
            AwaitSearch(window)
            Require(platform.SoundCalls = 1 AndAlso platform.FlashCalls = 1 AndAlso observedFinalState, "A completed scan did not notify after finalizing the UI.")
            Require(Not window.IsActive, "A background scan completion activated the main window.")

            StartSearch(window, root, "absent")
            AwaitSearch(window)
            Require(DirectCast(window.FindName("Results_lst"), ListBox).Items.Count = 0 AndAlso platform.SoundCalls = 2, "A no-match completion did not notify.")
            settings.MaximumTotalResults = 1
            StartSearch(window, root, "needle")
            AwaitSearch(window)
            Require(CurrentRun(window).State = ScanReportState.ResultLimitReached AndAlso platform.SoundCalls = 3, "A result-limit completion did not notify.")
            settings.MaximumTotalResults = 10000

            StartSearch(window, root, "needle")
            GetType(MainWindow).GetMethod("CancelScan", PrivateInstance).Invoke(window, Nothing)
            AwaitSearch(window)
            Require(CurrentRun(window).State = ScanReportState.Cancelled AndAlso platform.SoundCalls = 3, "A canceled scan emitted a completion sound.")
            StartSearch(window, root, "needle")
            window.Dispatcher.Invoke(Sub() GetType(MainWindow).GetMethod("Reset_btn_Click", PrivateInstance).Invoke(window, {Nothing, New RoutedEventArgs()}), DispatcherPriority.Normal)
            AwaitSearch(window)
            WaitUntil(window, Function() Not CBool(GetType(MainWindow).GetField("_isResetting", PrivateInstance).GetValue(window)))
            Require(platform.SoundCalls = 3 AndAlso DirectCast(window.FindName("Results_lst"), ListBox).Items.Count = 0, "Resetting a scan emitted a completion alert.")

            platform.Foreground = True
            StartSearch(window, root, "needle")
            AwaitSearch(window)
            Require(platform.SoundCalls = 3, "The default background-only policy notified while the application was active.")
            settings.NotifyOnlyInBackground = False
            StartSearch(window, root, "needle")
            AwaitSearch(window)
            Require(platform.SoundCalls = 4 AndAlso platform.FlashCalls = 3, "Foreground opt-in should play a sound without flashing the taskbar.")

            platform.Foreground = False
            window.ApplicationExitForChecks = Sub() window.Close()
            StartSearch(window, root, "needle")
            window.Dispatcher.Invoke(Sub() GetType(MainWindow).GetMethod("MainWindow_Closing", PrivateInstance).Invoke(window, {window, New CancelEventArgs()}), DispatcherPriority.Normal)
            AwaitSearch(window)
            WaitUntil(window, Function() Not window.IsVisible)
            Require(platform.SoundCalls = 4, "Shutting down emitted a completion notification.")
        Finally
            window.Close()
            Directory.Delete(root, True)
        End Try
    End Sub

    Private Sub StartSearch(window As MainWindow, root As String, query As String)
        DirectCast(window.FindName("Path_txt"), TextBox).Text = root
        DirectCast(window.FindName("Search_txt"), TextBox).Text = query
        Require(DirectCast(window.FindName("Scan_btn"), Button).IsEnabled, "The test search was not eligible to start.")
        GetType(MainWindow).GetMethod("StartScan", PrivateInstance).Invoke(window, Nothing)
    End Sub

    Private Sub AwaitSearch(window As MainWindow)
        Dim pending = DirectCast(GetType(MainWindow).GetField("_scanTask", PrivateInstance).GetValue(window), Task)
        Require(pending IsNot Nothing, "The scan task was not started.")
        WaitUntil(window, Function() pending.IsCompleted)
        pending.GetAwaiter().GetResult()
        Pump(window)
    End Sub

    Private Function CurrentRun(window As MainWindow) As ScanRunInfo
        Return DirectCast(GetType(MainWindow).GetField("_scanRun", PrivateInstance).GetValue(window), ScanRunInfo)
    End Function

    Private Sub DetachStartup(window As MainWindow)
        Dim type = GetType(MainWindow)
        window.RemoveHandler(FrameworkElement.LoadedEvent, type.GetMethod("MainWindow_Loaded", PrivateInstance).CreateDelegate(GetType(RoutedEventHandler), window))
        RemoveHandler window.Closing, DirectCast(type.GetMethod("MainWindow_Closing", PrivateInstance).CreateDelegate(GetType(CancelEventHandler), window), CancelEventHandler)
        RemoveHandler Microsoft.Win32.SystemEvents.UserPreferenceChanged,
            DirectCast(type.GetMethod("SystemThemeChanged", PrivateInstance).CreateDelegate(GetType(Microsoft.Win32.UserPreferenceChangedEventHandler), window), Microsoft.Win32.UserPreferenceChangedEventHandler)
    End Sub

    Private Sub Prepare(window As Window)
        window.WindowStartupLocation = WindowStartupLocation.Manual
        window.WindowState = WindowState.Normal
        window.Left = -10000
        window.Top = -10000
        window.ShowActivated = False
        window.ShowInTaskbar = False
    End Sub

    Private Sub EnsureApplication()
        If Application.Current Is Nothing Then
            Dim application As New Application With {.ShutdownMode = ShutdownMode.OnExplicitShutdown}
        Else
            Application.Current.ShutdownMode = ShutdownMode.OnExplicitShutdown
        End If
    End Sub

    Private Sub Pump(window As Window)
        window.UpdateLayout()
        window.Dispatcher.Invoke(Sub()
                                 End Sub, DispatcherPriority.ContextIdle)
    End Sub

    Private Sub WaitUntil(window As Window, ready As Func(Of Boolean))
        Dim watch = Stopwatch.StartNew()
        While Not ready()
            If watch.Elapsed > TimeSpan.FromSeconds(30) Then Throw New TimeoutException("The notification integration scan did not finish.")
            Pump(window)
            Thread.Sleep(2)
        End While
    End Sub

    Private Sub Require(condition As Boolean, message As String)
        If Not condition Then Throw New InvalidOperationException(message)
    End Sub

    Private NotInheritable Class FakePlatform
        Implements ISearchCompletionNotificationPlatform

        Public Property Available As Boolean = True
        Public Property Foreground As Boolean
        Public Property SoundFault As Boolean
        Public Property FlashFault As Boolean
        Public Property StopFault As Boolean
        Public Property PolicyFault As Boolean
        Public Property OnSound As Action
        Public Property SoundCalls As Integer
        Public Property FlashCalls As Integer
        Public Property StopCalls As Integer

        Public ReadOnly Property CanNotify As Boolean Implements ISearchCompletionNotificationPlatform.CanNotify
            Get
                If PolicyFault Then Throw New InvalidOperationException("Simulated notification eligibility failure.")
                Return Available
            End Get
        End Property

        Public ReadOnly Property IsApplicationForeground As Boolean Implements ISearchCompletionNotificationPlatform.IsApplicationForeground
            Get
                Return Foreground
            End Get
        End Property

        Public Sub PlayCompletionSound() Implements ISearchCompletionNotificationPlatform.PlayCompletionSound
            SoundCalls += 1
            OnSound?.Invoke()
            If SoundFault Then Throw New InvalidOperationException("Simulated Windows sound failure.")
        End Sub

        Public Sub FlashTaskbar() Implements ISearchCompletionNotificationPlatform.FlashTaskbar
            FlashCalls += 1
            If FlashFault Then Throw New InvalidOperationException("Simulated taskbar notification failure.")
        End Sub

        Public Sub StopFlashing() Implements ISearchCompletionNotificationPlatform.StopFlashing
            StopCalls += 1
            If StopFault Then Throw New InvalidOperationException("Simulated taskbar stop failure.")
        End Sub
    End Class
End Module
