Class Application

    ' Application-level events, such as Startup, Exit, and DispatcherUnhandledException
    ' can be handled in this file.

    Private Sub Application_Startup(sender As Object, e As StartupEventArgs) Handles Me.Startup
        ' SIMPLEST APPROACH: Stay in OnExplicitShutdown mode forever
        ' Manually handle shutdown when MainWindow closes
        Me.ShutdownMode = ShutdownMode.OnExplicitShutdown

        ' Show splash screen
        Dim splash As New Beacon.SplashWindow()
        splash.Show()
    End Sub

End Class
