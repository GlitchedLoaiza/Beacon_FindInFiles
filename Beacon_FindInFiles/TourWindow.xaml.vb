Imports System.Windows
Imports System.Windows.Media
Imports System.Windows.Threading

Namespace Beacon
    Partial Public Class TourWindow
        Private ReadOnly _main As Window
        Private ReadOnly _steps As TourStep() = WelcomeTour.Steps()
        Private _index As Integer = -1
        Private _closed As Boolean

        Public Sub New(owner As Window, dark As Boolean)
            InitializeComponent()
            _main = owner
            Me.Owner = owner
            BeaconThemePalette.CopyOwnerColors(owner, Me, dark)
            AddHandler owner.LocationChanged, AddressOf OwnerChanged
            AddHandler owner.SizeChanged, AddressOf OwnerChanged
            AddHandler owner.StateChanged, AddressOf OwnerChanged
            AddHandler owner.Closed, AddressOf OwnerClosed
            AddHandler Loaded, AddressOf OwnerChanged
            AddHandler SizeChanged, AddressOf OwnerChanged
            ShowStep()
        End Sub

        Friend ReadOnly Property StepIndex As Integer
            Get
                Return _index
            End Get
        End Property

        Private Sub ShowStep()
            If _index < 0 Then
                TourProgress_txt.Text = "An optional introduction"
                TourTitle_txt.Text = "Welcome to Beacon"
                TourBody_txt.Text = "Find information in files without opening them one by one. Would you like a short tour of a simple search and exporting its results?" & vbCrLf & vbCrLf &
                    "Choose Start tour to follow the steps, or Not now to use the app immediately. You can exit at any time. This welcome offer will not appear automatically on later launches. You can replay the tour from Help at any time."
            Else
                TourProgress_txt.Text = $"Step {_index + 1} of {_steps.Length}"
                TourTitle_txt.Text = _steps(_index).Title
                TourBody_txt.Text = _steps(_index).Instructions
            End If
            TourReminder_txt.Visibility = If(_index < 0, Visibility.Visible, Visibility.Collapsed)
            BackTour_btn.IsEnabled = _index > 0
            ExitTour_btn.Content = If(_index < 0, "Not now", "Exit tour")
            NextTour_btn.Content = If(_index < 0, "Start tour", If(_index = _steps.Length - 1, "Finish", "Next"))
            Dispatcher.BeginInvoke(Sub() PositionBesideTarget(), DispatcherPriority.Loaded)
        End Sub

        Private Sub NextStep(sender As Object, e As RoutedEventArgs)
            If _index = _steps.Length - 1 Then
                Close()
                Return
            End If
            _index += 1
            ShowStep()
        End Sub
        Private Sub PreviousStep(sender As Object, e As RoutedEventArgs)
            If _index <= 0 Then Return
            _index -= 1
            ShowStep()
        End Sub
        Private Sub ExitTour(sender As Object, e As RoutedEventArgs)
            Close()
        End Sub
        Private Sub OwnerClosed(sender As Object, e As EventArgs)
            If Not _closed Then Close()
        End Sub
        Private Sub OwnerChanged(sender As Object, e As EventArgs)
            If _closed Then Return
            Dispatcher.BeginInvoke(Sub() PositionBesideTarget(), DispatcherPriority.Loaded)
        End Sub

        Private Sub PositionBesideTarget()
            If _closed OrElse Not IsVisible OrElse _main.WindowState = WindowState.Minimized Then Return
            Dim element = If(_index < 0, TryCast(_main.Content, FrameworkElement), TryCast(_main.FindName(_steps(_index).Target), FrameworkElement))
            If element Is Nothing OrElse Not element.IsVisible Then element = TryCast(_main.Content, FrameworkElement)
            If element Is Nothing OrElse PresentationSource.FromVisual(element) Is Nothing Then Return
            Dim source = TryCast(PresentationSource.FromVisual(Me), System.Windows.Interop.HwndSource)
            If source Is Nothing Then Return
            Dim topLeft = element.PointToScreen(New Point())
            Dim bottomRight = element.PointToScreen(New Point(element.ActualWidth, element.ActualHeight))
            Dim screen = System.Windows.Forms.Screen.FromPoint(New System.Drawing.Point(CInt(topLeft.X), CInt(topLeft.Y)))
            Dim scale = source.CompositionTarget.TransformFromDevice
            Dim areaStart = scale.Transform(New Point(screen.WorkingArea.Left, screen.WorkingArea.Top))
            Dim areaEnd = scale.Transform(New Point(screen.WorkingArea.Right, screen.WorkingArea.Bottom))
            Dim first = scale.Transform(topLeft), last = scale.Transform(bottomRight)
            Dim x As Double, y As Double
            If _index < 0 Then
                x = (first.X + last.X - ActualWidth) / 2
                y = first.Y + 36
            Else
                x = last.X + 10
                y = first.Y
                If x + ActualWidth > areaEnd.X Then
                    x = first.X
                    y = last.Y + 10
                    If y + ActualHeight > areaEnd.Y Then y = first.Y - ActualHeight - 10
                End If
            End If
            Left = Math.Max(areaStart.X + 8, Math.Min(x, areaEnd.X - ActualWidth - 8))
            Top = Math.Max(areaStart.Y + 8, Math.Min(y, areaEnd.Y - ActualHeight - 8))
        End Sub

        Protected Overrides Sub OnClosed(e As EventArgs)
            _closed = True
            RemoveHandler _main.LocationChanged, AddressOf OwnerChanged
            RemoveHandler _main.SizeChanged, AddressOf OwnerChanged
            RemoveHandler _main.StateChanged, AddressOf OwnerChanged
            RemoveHandler _main.Closed, AddressOf OwnerClosed
            RemoveHandler Loaded, AddressOf OwnerChanged
            RemoveHandler SizeChanged, AddressOf OwnerChanged
            MyBase.OnClosed(e)
        End Sub
    End Class
End Namespace
