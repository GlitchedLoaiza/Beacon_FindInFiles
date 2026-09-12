Imports System.Diagnostics
Imports System.Threading
Imports System.Windows
Imports System.Windows.Media
Imports System.Windows.Media.Animation
Imports System.Runtime.InteropServices
Imports System.Windows.Interop

Namespace Beacon
    Partial Public Class MainWindow
        Private ReadOnly _updateCheckCancellation As New CancellationTokenSource()
        Private _updateCheckStarted As Boolean
        Private Const NotificationDurationSeconds As Integer = 5
        Private WithEvents _updateNotificationTimer As New System.Windows.Threading.DispatcherTimer With {.Interval = TimeSpan.FromMilliseconds(100)}
        Private ReadOnly _notificationElapsed As New Stopwatch()
        Private _notificationGeneration As Integer
        Private _notificationSeconds As Integer = -1

        <StructLayout(LayoutKind.Sequential)>
        Private Structure DesktopRectangle
            Public Left, Top, Right, Bottom As Integer
        End Structure

        <StructLayout(LayoutKind.Sequential)>
        Private Structure TaskbarData
            Public Size As UInteger
            Public Window As IntPtr
            Public CallbackMessage As UInteger
            Public Edge As UInteger
            Public Bounds As DesktopRectangle
            Public Parameter As IntPtr
        End Structure

        <DllImport("shell32.dll")>
        Private Shared Function SHAppBarMessage(message As UInteger, ByRef data As TaskbarData) As UIntPtr
        End Function

        <DllImport("user32.dll", SetLastError:=True)>
        Private Shared Function GetWindowRect(handle As IntPtr, ByRef bounds As DesktopRectangle) As <MarshalAs(UnmanagedType.Bool)> Boolean
        End Function

        <DllImport("user32.dll", SetLastError:=True)>
        Private Shared Function SetWindowPos(handle As IntPtr, insertAfter As IntPtr, x As Integer, y As Integer,
                                             width As Integer, height As Integer, flags As UInteger) As <MarshalAs(UnmanagedType.Bool)> Boolean
        End Function

        Private Async Sub CheckForUpdatesAtStartup()
            If _updateCheckStarted Then Return
            _updateCheckStarted = True
            Try
                Dim current = GetType(MainWindow).Assembly.GetName().Version
                Dim newer = Await ReleaseUpdateChecker.CheckAsync(current, _updateCheckCancellation.Token)
                If newer IsNot Nothing AndAlso Not _updateCheckCancellation.IsCancellationRequested AndAlso Not _isClosing Then
                    ShowUpdateNotification(newer)
                End If
            Catch ex As Exception
                Debug.WriteLine($"Startup release check unavailable: {ex.Message}")
            End Try
        End Sub

        Private Sub ShowUpdateNotification(version As Version)
            _notificationGeneration += 1
            _updateNotificationTimer.Stop()
            UpdateSlide.BeginAnimation(TranslateTransform.YProperty, Nothing)
            Dim label = If(version.Revision > 0, version.ToString(4), version.ToString(3))
            UpdateVersion_txt.Text = $"Version {label}"
            _notificationSeconds = -1
            SetNotificationCountdown(NotificationDurationSeconds)
            UpdateNotification.Visibility = Visibility.Visible
            UpdateNotification.IsHitTestVisible = True
            Dim workArea = SystemParameters.WorkArea
            UpdatePopup.HorizontalOffset = workArea.Left + (workArea.Width - UpdatePopupHost.Width) / 2
            UpdatePopup.VerticalOffset = workArea.Bottom - UpdatePopupHost.Height
            UpdatePopup.IsOpen = True
            PositionUpdateNotification()
            UpdateSlide.Y = 0
            If SystemParameters.ClientAreaAnimation Then
                UpdateSlide.BeginAnimation(TranslateTransform.YProperty,
                    New DoubleAnimation(UpdatePopupHost.Height, 0, TimeSpan.FromMilliseconds(280)) With {
                        .EasingFunction = New CubicEase With {.EasingMode = EasingMode.EaseOut}, .FillBehavior = FillBehavior.Stop})
            End If
            _notificationElapsed.Restart()
            _updateNotificationTimer.Start()
        End Sub

        Private Sub DismissUpdateNotification(sender As Object, e As RoutedEventArgs)
            _updateNotificationTimer.Stop()
            _notificationElapsed.Stop()
            _notificationGeneration += 1
            Dim generation = _notificationGeneration
            UpdateNotification.IsHitTestVisible = False
            If Not UpdatePopup.IsOpen OrElse Not SystemParameters.ClientAreaAnimation Then
                HideUpdateNotification()
                Return
            End If
            Dim animation As New DoubleAnimation(UpdateSlide.Y, UpdatePopupHost.Height, TimeSpan.FromMilliseconds(240)) With {
                .EasingFunction = New CubicEase With {.EasingMode = EasingMode.EaseIn}}
            AddHandler animation.Completed, Sub()
                                                If generation = _notificationGeneration Then HideUpdateNotification()
                                            End Sub
            UpdateSlide.BeginAnimation(TranslateTransform.YProperty, animation)
        End Sub

        Private Sub UpdateNotificationExpired(sender As Object, e As EventArgs) Handles _updateNotificationTimer.Tick
            PositionUpdateNotification()
            Dim remaining = Math.Max(0, CInt(Math.Ceiling(NotificationDurationSeconds - _notificationElapsed.Elapsed.TotalSeconds)))
            SetNotificationCountdown(remaining)
            If remaining = 0 Then DismissUpdateNotification(Nothing, Nothing)
        End Sub

        Private Sub SetNotificationCountdown(seconds As Integer)
            If seconds = _notificationSeconds Then Return
            _notificationSeconds = seconds
            UpdateMessage_txt.Text = $"A newer version of Beacon is available ({seconds}s)"
        End Sub

        Private Shared Function NotificationDesktopPosition(width As Integer, height As Integer) As Point
            Dim data As New TaskbarData With {.Size = CUInt(Marshal.SizeOf(Of TaskbarData)())}
            Dim hasTaskbar = SHAppBarMessage(5UI, data) <> UIntPtr.Zero AndAlso
                data.Bounds.Right > data.Bounds.Left AndAlso data.Bounds.Bottom > data.Bounds.Top
            Dim screen = If(hasTaskbar,
                System.Windows.Forms.Screen.FromRectangle(System.Drawing.Rectangle.FromLTRB(data.Bounds.Left, data.Bounds.Top, data.Bounds.Right, data.Bounds.Bottom)),
                System.Windows.Forms.Screen.PrimaryScreen)
            Dim area = screen.WorkingArea
            Dim centerX As Double = area.Left + area.Width / 2.0
            Dim bottom = area.Bottom
            If hasTaskbar Then
                Select Case data.Edge
                    Case 1UI, 3UI
                        centerX = data.Bounds.Left + (data.Bounds.Right - data.Bounds.Left) / 2.0
                        bottom = If(data.Edge = 3UI, Math.Min(data.Bounds.Top, screen.Bounds.Bottom), area.Bottom)
                End Select
            End If
            Dim x = Math.Max(area.Left, Math.Min(centerX - width / 2.0, area.Right - width))
            Dim y = Math.Max(area.Top, bottom - height)
            Return New Point(Math.Round(x), Math.Round(y))
        End Function

        Private Sub PositionUpdateNotification()
            If Not UpdatePopup.IsOpen Then Return
            Dim source = TryCast(PresentationSource.FromVisual(UpdatePopupHost), HwndSource)
            If source Is Nothing Then Return
            Dim bounds As DesktopRectangle
            If Not GetWindowRect(source.Handle, bounds) Then Return
            Dim target = NotificationDesktopPosition(bounds.Right - bounds.Left, bounds.Bottom - bounds.Top)
            If bounds.Left = CInt(target.X) AndAlso bounds.Top = CInt(target.Y) Then Return
            ' Keep native screen pixels throughout; WPF offsets can be adjusted by screen-edge placement and DPI.
            If Not SetWindowPos(source.Handle, IntPtr.Zero, CInt(target.X), CInt(target.Y), 0, 0, &H15UI) Then
                Debug.WriteLine($"Could not position update notification: {Marshal.GetLastWin32Error()}")
            End If
        End Sub

        Private Sub UpdateNotificationOpened(sender As Object, e As EventArgs) Handles UpdatePopup.Opened, UpdatePopupHost.SizeChanged
            PositionUpdateNotification()
        End Sub

        Private Sub UpdateNotificationWindowMoved(sender As Object, e As EventArgs) Handles Me.LocationChanged, Me.SizeChanged, Me.StateChanged
            If Not UpdatePopup.IsOpen Then Return
            If WindowState = WindowState.Minimized Then
                StopUpdateNotificationTimer(Nothing, EventArgs.Empty)
                Return
            End If
            PositionUpdateNotification()
        End Sub

        Private Sub HideUpdateNotification()
            UpdatePopup.IsOpen = False
            UpdateNotification.Visibility = Visibility.Collapsed
            UpdateSlide.BeginAnimation(TranslateTransform.YProperty, Nothing)
        End Sub

        Private Sub StopUpdateNotificationTimer(sender As Object, e As EventArgs) Handles Me.Closed
            _updateNotificationTimer.Stop()
            _notificationElapsed.Stop()
            _notificationGeneration += 1
            HideUpdateNotification()
        End Sub

        Private Sub OpenUpdateRelease(sender As Object, e As RoutedEventArgs)
            Try
                Process.Start(New ProcessStartInfo(ReleaseUpdateChecker.ReleasePage) With {.UseShellExecute = True})
            Catch ex As Exception
                UpdateVersion_txt.Text = "Couldn't open your browser. Try the release page on GitHub."
            End Try
        End Sub
    End Class
End Namespace
