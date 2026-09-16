Class Application
    Private _instance As Beacon.SingleInstanceCoordinator
    Private _exiting As Boolean

    Private Sub Application_Startup(sender As Object, e As StartupEventArgs) Handles Me.Startup
        Me.ShutdownMode = ShutdownMode.OnExplicitShutdown
        Try
            _instance = New Beacon.SingleInstanceCoordinator(Beacon.SingleInstanceCoordinator.InstanceName(), AddressOf RequestActivation)
            If Not _instance.IsPrimary Then
                Shutdown()
                Return
            End If
            Dim splash As New Beacon.SplashWindow()
            splash.Show()
        Catch ex As Exception
            MessageBox.Show($"Beacon could not start: {ex.Message}", "Beacon", MessageBoxButton.OK, MessageBoxImage.Error)
            Shutdown(1)
        End Try
    End Sub

    Private Sub RequestActivation()
        If Dispatcher.HasShutdownStarted OrElse Dispatcher.HasShutdownFinished Then Return
        Dispatcher.BeginInvoke(Sub()
                                   If Not _exiting Then Beacon.InstanceWindowActivation.Activate(Me)
                               End Sub)
    End Sub

    Protected Overrides Sub OnExit(e As ExitEventArgs)
        _exiting = True
        _instance?.Dispose()
        MyBase.OnExit(e)
    End Sub
End Class
