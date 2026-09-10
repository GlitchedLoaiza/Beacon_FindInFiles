Imports System.ComponentModel
Imports System.Windows
Imports System.Windows.Automation.Peers
Imports System.Windows.Threading
Imports System.Windows.Media.Animation
Imports System.Windows.Media
Imports System.Linq
Imports System.Diagnostics

Namespace Beacon

    Public Partial Class SplashWindow

        Private ReadOnly _autoTransition As Boolean
        Private _autoCloseTimer As DispatcherTimer
        Private _motionStoryboard As Storyboard
        Private _fadeStoryboard As Storyboard
        Private _loadedOnce As Boolean
        Private _transitioning As Boolean
        Private _mainWindowShown As Boolean
        Private _closed As Boolean

        Public Sub New()
            Me.New(True)
        End Sub

        Friend Sub New(autoTransition As Boolean)
            InitializeComponent()
            _autoTransition = autoTransition
        End Sub

        Private Sub Window_Loaded(sender As Object, e As RoutedEventArgs)
            If _loadedOnce OrElse _closed Then Return
            _loadedOnce = True
            If _autoTransition Then
                Dim area = SystemParameters.WorkArea
                Dim size = CalculateSplashSize(New Size(area.Width, area.Height))
                Width = size.Width
                Height = size.Height
                Left = area.Left + (area.Width - Width) / 2
                Top = area.Top + (area.Height - Height) / 2
            End If
            If (RenderCapability.Tier >> 16) = 0 Then
                FarLogLayer.Effect = Nothing
                MiddleLogLayer.Effect = Nothing
            End If
            AddHandler SystemParameters.StaticPropertyChanged, AddressOf SystemAnimationPreferenceChanged
            UpdateAnimationState(SystemParameters.ClientAreaAnimation)
            If _autoTransition Then
                _autoCloseTimer = New DispatcherTimer(DispatcherPriority.Background) With {.Interval = TimeSpan.FromSeconds(3.2)}
                AddHandler _autoCloseTimer.Tick, AddressOf AutoCloseTimer_Tick
                _autoCloseTimer.Start()
            End If
        End Sub

        Friend Shared Function CalculateSplashSize(workArea As Size) As Size
            Dim scale = Math.Min(1.0, Math.Min(Math.Max(1, workArea.Width - 32) / 900,
                                             Math.Max(1, workArea.Height - 32) / 600))
            Return New Size(900 * scale, 600 * scale)
        End Function

        Private Sub SystemAnimationPreferenceChanged(sender As Object, e As PropertyChangedEventArgs)
            If e.PropertyName <> NameOf(SystemParameters.ClientAreaAnimation) OrElse _closed Then Return
            If Dispatcher.CheckAccess() Then
                UpdateAnimationState(SystemParameters.ClientAreaAnimation)
            Else
                Dispatcher.BeginInvoke(New Action(Sub() UpdateAnimationState(SystemParameters.ClientAreaAnimation)))
            End If
        End Sub

        Private Sub UpdateAnimationState(enabled As Boolean)
            If _closed OrElse Not _loadedOnce OrElse _transitioning Then Return
            If enabled Then
                If _motionStoryboard IsNot Nothing Then Return
                _motionStoryboard = DirectCast(FindResource("AmbientMotion"), Storyboard).Clone()
                _motionStoryboard.Begin(Me, HandoffBehavior.SnapshotAndReplace, True)
            ElseIf _motionStoryboard IsNot Nothing Then
                _motionStoryboard.Remove(Me)
                _motionStoryboard = Nothing
            End If
        End Sub

        Private Sub AutoCloseTimer_Tick(sender As Object, e As EventArgs)
            TransitionToMainWindow()
        End Sub

        Private Sub StopStartupTimer()
            If _autoCloseTimer Is Nothing Then Return
            _autoCloseTimer.Stop()
            RemoveHandler _autoCloseTimer.Tick, AddressOf AutoCloseTimer_Tick
            _autoCloseTimer = Nothing
        End Sub

        ' -----------------------------
        ' FADE-OUT + MAIN WINDOW TRANSITION
        ' -----------------------------
        Public Sub TransitionToMainWindow()
            If _transitioning OrElse _closed Then Return
            _transitioning = True
            StopStartupTimer()
            LoadingStatus_txt.Text = "Opening Beacon…"
            _motionStoryboard?.Pause(Me)
            If Not SystemParameters.ClientAreaAnimation Then
                ShowMainAndClose()
                Return
            End If
            _fadeStoryboard = DirectCast(FindResource("FadeOutStoryboard"), Storyboard).Clone()
            AddHandler _fadeStoryboard.Completed, AddressOf FadeCompleted
            _fadeStoryboard.Begin(Me, HandoffBehavior.SnapshotAndReplace, True)
        End Sub

        Private Sub FadeCompleted(sender As Object, e As EventArgs)
            ShowMainAndClose()
        End Sub

        Private Sub ShowMainAndClose()
            If _closed OrElse _mainWindowShown Then Return
            Try
                Dim main = Application.Current.Windows.OfType(Of MainWindow)().FirstOrDefault()
                If main Is Nothing Then main = New MainWindow()
                Application.Current.MainWindow = main
                main.Show()
                _mainWindowShown = True
                Close()
            Catch ex As Exception
                Debug.WriteLine($"Beacon startup failed: {ex}")
                MessageBox.Show(Me, $"Beacon could not start: {ex.Message}", "Beacon", MessageBoxButton.OK, MessageBoxImage.Error)
                Close()
            End Try
        End Sub

        Protected Overrides Function OnCreateAutomationPeer() As AutomationPeer
            Return New SplashAutomationPeer(Me)
        End Function

        Private NotInheritable Class SplashAutomationPeer
            Inherits WindowAutomationPeer

            Public Sub New(owner As SplashWindow)
                MyBase.New(owner)
            End Sub

            Protected Overrides Function GetChildrenCore() As List(Of AutomationPeer)
                Dim splash = DirectCast(Owner, SplashWindow)
                Dim status = UIElementAutomationPeer.CreatePeerForElement(splash.LoadingStatus_txt)
                Return If(status Is Nothing, New List(Of AutomationPeer)(), New List(Of AutomationPeer) From {status})
            End Function
        End Class

        Protected Overrides Sub OnClosed(e As EventArgs)
            _closed = True
            StopStartupTimer()
            RemoveHandler SystemParameters.StaticPropertyChanged, AddressOf SystemAnimationPreferenceChanged
            If _motionStoryboard IsNot Nothing Then
                _motionStoryboard.Remove(Me)
                _motionStoryboard = Nothing
            End If
            If _fadeStoryboard IsNot Nothing Then
                RemoveHandler _fadeStoryboard.Completed, AddressOf FadeCompleted
                _fadeStoryboard.Remove(Me)
                _fadeStoryboard = Nothing
            End If
            MyBase.OnClosed(e)
            If _autoTransition AndAlso Not _mainWindowShown Then Application.Current?.Shutdown()
        End Sub

    End Class

End Namespace
