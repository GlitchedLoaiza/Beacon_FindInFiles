Imports System.Windows.Threading
Imports System.Windows.Media.Animation
Imports System.Windows.Controls
Imports System.Windows.Media
Imports System.Windows.Shapes

Namespace Beacon

    Public Partial Class SplashWindow

        Private ReadOnly _rng As New Random()

        Private _logTimer As DispatcherTimer
        Private _particleTimer As DispatcherTimer
        Private _autoCloseTimer As DispatcherTimer

        Private Const LogLineCount As Integer = 18

        Private ReadOnly _fakeLogPool As String() = {
            "INFO  [Scanner] Enumerating files…",
            "INFO  [Zip] Opening archive…",
            "INFO  [EVTX] Parsing events…",
            "WARN  [IO] Access denied, skipping…",
            "INFO  [Match] Pattern hit: 0x80070005",
            "INFO  [Match] Keyword hit: 'EnrollmentStatus' ",
            "INFO  [Index] Building in-memory map…",
            "INFO  [Search] Searching: 'Error' ",
            "INFO  [Search] Searching: 'Failed' ",
            "INFO  [Search] Searching: 'Timeout' ",
            "INFO  [Path] C:\Logs\Device\…",
            "INFO  [Path] C:\Windows\CCM\Logs\…",
            "INFO  [Path] %TEMP%\Beacon\…",
            "INFO  [Done] Rendering preview…"
        }

        Private Sub Window_Loaded(sender As Object, e As RoutedEventArgs)

            ' Start the lupe sweep storyboard
            Dim lupeSb = TryCast(Me.FindResource("LupeSweep"), Storyboard)
            lupeSb?.Begin()

            ' Build initial moving log lines
            CreateLogLines()

            ' Timers:
            '  - Log refresh cadence (occasionally re-randomize a line's text)
            _logTimer = New DispatcherTimer With {.Interval = TimeSpan.FromMilliseconds(450)}
            AddHandler _logTimer.Tick, AddressOf LogTimer_Tick
            _logTimer.Start()

            '  - Particle spawn cadence
            _particleTimer = New DispatcherTimer With {.Interval = TimeSpan.FromMilliseconds(180)}
            AddHandler _particleTimer.Tick, AddressOf ParticleTimer_Tick
            _particleTimer.Start()

            ' OPTIONAL: auto transition after a short delay
            _autoCloseTimer = New DispatcherTimer With {.Interval = TimeSpan.FromSeconds(3.2)}
            AddHandler _autoCloseTimer.Tick,
                Sub()
                    _autoCloseTimer.Stop()
                    TransitionToMainWindow()
                End Sub
            _autoCloseTimer.Start()

        End Sub

        ' -----------------------------
        ' REAL MOVING "TEXT-LIKE LOGS"
        ' -----------------------------
        Private Sub CreateLogLines()
            LogCanvas.Children.Clear()

            For i = 0 To LogLineCount - 1
                Dim tb As New TextBlock() With {
                    .Text = RandomLog(),
                    .FontFamily = New FontFamily("Consolas"),
                    .FontSize = 12,
                    .Foreground = New SolidColorBrush(Color.FromArgb(110, 180, 220, 255)),
                    .Opacity = 0.55
                }

                ' Random start positions within the log area
                Dim x = _rng.Next(0, 260)
                Dim y = _rng.Next(0, 140)

                Canvas.SetLeft(tb, x)
                Canvas.SetTop(tb, y)

                ' Animate: drift upward + slight horizontal sway, forever
                Dim tg As New TransformGroup()
                Dim tt As New TranslateTransform()
                tg.Children.Add(tt)
                tb.RenderTransform = tg

                ' Upward motion (wrap effect simulated by reset in tick)
                Dim dur = TimeSpan.FromSeconds(1.6 + _rng.NextDouble() * 1.8)

                Dim animY As New DoubleAnimation With {
                    .From = 0,
                    .To = -(60 + _rng.Next(0, 90)),
                    .Duration = New Duration(dur),
                    .RepeatBehavior = RepeatBehavior.Forever
                }

                ' Horizontal subtle motion
                Dim animX As New DoubleAnimation With {
                    .From = 0,
                    .To = (-12 + _rng.Next(0, 25)),
                    .Duration = New Duration(TimeSpan.FromSeconds(1.2 + _rng.NextDouble() * 1.2)),
                    .AutoReverse = True,
                    .RepeatBehavior = RepeatBehavior.Forever
                }

                tt.BeginAnimation(TranslateTransform.YProperty, animY)
                tt.BeginAnimation(TranslateTransform.XProperty, animX)

                LogCanvas.Children.Add(tb)
            Next
        End Sub

        Private Sub LogTimer_Tick(sender As Object, e As EventArgs)
            ' Occasionally re-randomize one line text to keep it "alive"
            If LogCanvas.Children.Count = 0 Then Return
            Dim idx = _rng.Next(0, LogCanvas.Children.Count)

            Dim tb = TryCast(LogCanvas.Children(idx), TextBlock)
            If tb Is Nothing Then Return

            If _rng.NextDouble() < 0.55 Then
                tb.Text = RandomLog()
            End If
        End Sub

        Private Function RandomLog() As String
            Dim baseMsg = _fakeLogPool(_rng.Next(0, _fakeLogPool.Length))
            Dim ms = _rng.Next(10, 9999).ToString("0000")
            Return $"{DateTime.Now:HH:mm:ss}.{ms}  {baseMsg}"
        End Function

        ' -----------------------------
        ' PARTICLE "MATCH" SPAWNING
        ' -----------------------------
        Private Sub ParticleTimer_Tick(sender As Object, e As EventArgs)
            ' Spawn fewer particles randomly for subtlety
            If _rng.NextDouble() < 0.35 Then Return

            Dim dot As New Ellipse() With {
                .Width = 4 + _rng.Next(0, 5),
                .Height = 4 + _rng.Next(0, 5),
                .Fill = New SolidColorBrush(Color.FromArgb(170, 255, 60, 60)),
                .Opacity = 0.0
            }

            Dim x = _rng.Next(20, 460)
            Dim y = _rng.Next(35, 165)

            Canvas.SetLeft(dot, x)
            Canvas.SetTop(dot, y)

            ' Scale + fade in/out, then remove
            Dim st As New ScaleTransform(1, 1)
            dot.RenderTransform = st
            dot.RenderTransformOrigin = New Point(0.5, 0.5)

            ParticleCanvas.Children.Add(dot)

            Dim sb As New Storyboard()

            Dim fadeIn As New DoubleAnimation With {.From = 0, .To = 0.95, .Duration = New Duration(TimeSpan.FromMilliseconds(90))}
            Storyboard.SetTarget(fadeIn, dot)
            Storyboard.SetTargetProperty(fadeIn, New PropertyPath("Opacity"))
            sb.Children.Add(fadeIn)

            Dim fadeOut As New DoubleAnimation With {.From = 0.95, .To = 0, .BeginTime = TimeSpan.FromMilliseconds(220), .Duration = New Duration(TimeSpan.FromMilliseconds(380))}
            Storyboard.SetTarget(fadeOut, dot)
            Storyboard.SetTargetProperty(fadeOut, New PropertyPath("Opacity"))
            sb.Children.Add(fadeOut)

            Dim grow As New DoubleAnimation With {.From = 1, .To = 3.1, .Duration = New Duration(TimeSpan.FromMilliseconds(520))}
            Storyboard.SetTarget(grow, dot)
            Storyboard.SetTargetProperty(grow, New PropertyPath("RenderTransform.ScaleX"))
            sb.Children.Add(grow)

            Dim growY As New DoubleAnimation With {.From = 1, .To = 3.1, .Duration = New Duration(TimeSpan.FromMilliseconds(520))}
            Storyboard.SetTarget(growY, dot)
            Storyboard.SetTargetProperty(growY, New PropertyPath("RenderTransform.ScaleY"))
            sb.Children.Add(growY)

            AddHandler sb.Completed,
                Sub()
                    ParticleCanvas.Children.Remove(dot)
                End Sub

            sb.Begin()
        End Sub

        ' -----------------------------
        ' FADE-OUT + MAIN WINDOW TRANSITION
        ' -----------------------------
        Public Sub TransitionToMainWindow()
            ' Stop timers
            _logTimer?.Stop()
            _particleTimer?.Stop()

            Dim fadeSb = TryCast(Me.FindResource("FadeOutStoryboard"), Storyboard)
            If fadeSb Is Nothing Then
                ShowMainAndClose()
                Return
            End If

            AddHandler fadeSb.Completed,
                Sub()
                    ShowMainAndClose()
                End Sub

            fadeSb.Begin(Me)
        End Sub

        Private Sub ShowMainAndClose()
            Dim mw As New MainWindow()
            mw.Show()
            Me.Close()
        End Sub

    End Class

End Namespace
