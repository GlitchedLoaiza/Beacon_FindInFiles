Imports System.Diagnostics
Imports System.IO
Imports System.Reflection
Imports System.Threading
Imports System.Threading.Tasks
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Input
Imports System.Windows.Threading
Imports Beacon

Module ShutdownChecks
    Private Const PrivateInstance As BindingFlags = BindingFlags.Instance Or BindingFlags.NonPublic

    Public Sub NativeCloseWithPendingBrowser()
        CheckPendingBrowser(AppTheme.Dark)
    End Sub

    Public Sub AeroCloseWithPendingBrowser()
        CheckPendingBrowser(AppTheme.Aero)
    End Sub

    Private Sub CheckPendingBrowser(theme As AppTheme)
        EnsureApplication()
        Dim startup As New TaskCompletionSource(Of Boolean)(TaskCreationOptions.RunContinuationsAsynchronously)
        Dim window = CreateWindow(theme)
        Dim closed As Boolean
        AddHandler window.Closed, Sub() closed = True
        window.ApplicationExitForChecks = Sub() window.Dispatcher.BeginInvoke(New Action(AddressOf window.Close), DispatcherPriority.Background)
        GetType(MainWindow).GetField("_webViewInitializationTask", PrivateInstance).SetValue(window, startup.Task)
        Dim waiter = DirectCast(GetType(MainWindow).GetMethod("EnsureWebView2InitializedAsync", PrivateInstance).Invoke(window, Nothing), Task(Of Boolean))
        Try
            window.Show()
            Pump(window)
            RequestClose(window, theme)
            Pump(window)
            If Not closed Then RequestClose(window, theme)
            Require(WaitUntil(window, Function() closed, TimeSpan.FromSeconds(2)),
                    $"{theme} close remained blocked by unfinished optional WebView2 initialization.")
            Require(Not startup.Task.IsCompleted, "Closing should not require optional native browser startup to finish.")
            Require(WaitUntil(window, Function() waiter.IsCompleted, TimeSpan.FromSeconds(2)) AndAlso Not waiter.GetAwaiter().GetResult(),
                    "A pending preview caller did not receive cancellation during shutdown.")
            startup.TrySetResult(True)
            Pump(window)
            Require(Not CBool(GetType(MainWindow).GetField("_webViewInitialized", PrivateInstance).GetValue(window)) AndAlso
                    Not window.IsVisible, "Late browser completion revived a closed preview.")
        Finally
            startup.TrySetResult(False)
            WaitUntil(window, Function() closed, TimeSpan.FromSeconds(5))
            If Not closed Then window.Dispatcher.Invoke(New Action(AddressOf window.Close))
        End Try
    End Sub

    Public Sub CompletedBrowserStates()
        EnsureApplication()
        For Each theme In {AppTheme.Dark, AppTheme.Aero}
            For Each failed In {False, True}
                Dim window = CreateWindow(theme)
                Dim closed As Boolean
                AddHandler window.Closed, Sub() closed = True
                window.ApplicationExitForChecks = Sub() window.Dispatcher.BeginInvoke(New Action(AddressOf window.Close), DispatcherPriority.Background)
                Dim startup = If(failed, Task.FromException(Of Boolean)(New IOException("Synthetic browser startup failure.")), Task.FromResult(False))
                GetType(MainWindow).GetField("_webViewInitializationTask", PrivateInstance).SetValue(window, startup)
                Try
                    window.Show()
                    Pump(window)
                    RequestClose(window, theme)
                    Require(WaitUntil(window, Function() closed, TimeSpan.FromSeconds(2)), "Completed/failed browser startup prevented closing.")
                Finally
                    Dim observed = startup.Exception
                    If Not closed Then window.Dispatcher.Invoke(New Action(AddressOf window.Close))
                End Try
            Next
        Next
    End Sub

    Public Sub ProcessExit()
        For Each theme In {AppTheme.Dark, AppTheme.Aero}
            Dim start As New ProcessStartInfo(Environment.ProcessPath) With {
                .UseShellExecute = False, .CreateNoWindow = True, .RedirectStandardOutput = True, .RedirectStandardError = True
            }
            If String.Equals(Path.GetFileNameWithoutExtension(Environment.ProcessPath), "dotnet", StringComparison.OrdinalIgnoreCase) Then
                start.ArgumentList.Add(GetType(ShutdownChecks).Assembly.Location)
            End If
            start.ArgumentList.Add("--shutdown-child")
            start.ArgumentList.Add(theme.ToString())
            Using child = Process.Start(start)
                Dim output = child.StandardOutput.ReadToEndAsync()
                Dim errors = child.StandardError.ReadToEndAsync()
                Try
                    If Not child.WaitForExit(15000) Then
                        child.Kill(True)
                        child.WaitForExit(5000)
                        Throw New TimeoutException($"{theme} child process did not exit after its close command. " & output.GetAwaiter().GetResult() & errors.GetAwaiter().GetResult())
                    End If
                    Dim trace = output.GetAwaiter().GetResult() & errors.GetAwaiter().GetResult()
                    Require(child.ExitCode = 0 AndAlso trace.Contains("WINDOW CLOSED") AndAlso trace.Contains("APPLICATION EXITED"),
                            $"{theme} close did not terminate the application cleanly: " & trace)
                    Console.WriteLine($"Shutdown child {theme}: " & trace.Trim().Replace(Environment.NewLine, " | "))
                Finally
                    If Not child.HasExited Then
                        child.Kill(True)
                        child.WaitForExit(5000)
                    End If
                End Try
            End Using
        Next
    End Sub

    Public Function RunChildMode(args As String()) As Integer
        Dim theme = [Enum].Parse(Of AppTheme)(args(1))
        Dim application As New Application With {.ShutdownMode = ShutdownMode.OnExplicitShutdown}
        Dim window = CreateWindow(theme)
        Dim requested As Boolean
        Dim closed As Boolean
        AddHandler window.Closed, Sub()
                                     closed = True
                                     Console.WriteLine("WINDOW CLOSED")
                                 End Sub
        AddHandler application.DispatcherUnhandledException,
            Sub(sender, e)
                Console.Error.WriteLine(e.Exception.ToString())
                e.Handled = True
                application.Shutdown(2)
            End Sub
        Dim timer As New DispatcherTimer(DispatcherPriority.Background) With {.Interval = TimeSpan.FromMilliseconds(250)}
        AddHandler timer.Tick,
            Sub()
                timer.Stop()
                requested = True
                Dim startup = TryCast(GetType(MainWindow).GetField("_webViewInitializationTask", PrivateInstance).GetValue(window), Task)
                Console.WriteLine("CLOSE REQUEST; browser startup pending=" & (startup IsNot Nothing AndAlso Not startup.IsCompleted).ToString())
                RequestClose(window, theme)
            End Sub
        AddHandler window.Loaded,
            Sub()
                GetType(MainWindow).GetMethod("InitializeWebView2Async", PrivateInstance).Invoke(window, Nothing)
                timer.Start()
            End Sub
        Dim result = application.Run(window)
        timer.Stop()
        Console.WriteLine("APPLICATION EXITED")
        Return If(result = 0 AndAlso requested AndAlso closed, 0, 1)
    End Function

    Private Function CreateWindow(theme As AppTheme) As MainWindow
        Dim window As New MainWindow()
        Dim type = GetType(MainWindow)
        window.RemoveHandler(FrameworkElement.LoadedEvent, type.GetMethod("MainWindow_Loaded", PrivateInstance).CreateDelegate(GetType(RoutedEventHandler), window))
        type.GetField("_settings", PrivateInstance).SetValue(window, New BeaconSettings With {.Theme = theme, .NotifyOnSearchCompletion = False})
        type.GetMethod("ApplyThemeSelection", PrivateInstance).Invoke(window, {False})
        window.WindowStartupLocation = WindowStartupLocation.Manual
        window.WindowState = WindowState.Normal
        window.Left = -10000
        window.Top = -10000
        window.ShowActivated = False
        window.ShowInTaskbar = False
        Return window
    End Function

    Private Sub RequestClose(window As MainWindow, theme As AppTheme)
        window.Dispatcher.Invoke(Sub()
                                     If theme = AppTheme.Aero AndAlso Not SystemParameters.HighContrast Then
                                         Dim button = DirectCast(window.Template.FindName("AeroCloseButton", window), Button)
                                         Require(button IsNot Nothing AndAlso button.CommandTarget Is window, "The Aero caption close button is not connected to its window.")
                                         Dim command = DirectCast(button.Command, RoutedCommand)
                                         Require(command.CanExecute(button.CommandParameter, window), "The Aero close command is disabled.")
                                         command.Execute(button.CommandParameter, window)
                                     Else
                                         SystemCommands.CloseWindow(window)
                                     End If
                                 End Sub, DispatcherPriority.Send)
    End Sub

    Private Sub EnsureApplication()
        If Application.Current Is Nothing Then
            Dim application As New Application With {.ShutdownMode = ShutdownMode.OnExplicitShutdown}
        Else
            Application.Current.ShutdownMode = ShutdownMode.OnExplicitShutdown
        End If
    End Sub

    Private Sub Pump(window As Window)
        window.Dispatcher.Invoke(Sub()
                                 End Sub, DispatcherPriority.ContextIdle)
    End Sub

    Private Function WaitUntil(window As Window, condition As Func(Of Boolean), timeout As TimeSpan) As Boolean
        Dim watch = Stopwatch.StartNew()
        While Not condition() AndAlso watch.Elapsed < timeout
            Pump(window)
            Thread.Sleep(5)
        End While
        Return condition()
    End Function

    Private Sub Require(condition As Boolean, message As String)
        If Not condition Then Throw New InvalidOperationException(message)
    End Sub
End Module
