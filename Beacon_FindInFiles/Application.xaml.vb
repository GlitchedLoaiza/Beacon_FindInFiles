Class Application

    ' Application-level events, such as Startup, Exit, and DispatcherUnhandledException
    ' can be handled in this file.

    Private Sub Application_Startup(sender As Object, e As StartupEventArgs) Handles Me.Startup
        ' Show splash screen instead of main window
        Dim splash As New Beacon.SplashWindow()
        splash.Show()
    End Sub

End Class
