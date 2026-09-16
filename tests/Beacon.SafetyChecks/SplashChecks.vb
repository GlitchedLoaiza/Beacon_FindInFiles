Imports System.Diagnostics
Imports System.IO
Imports System.Reflection
Imports System.Threading
Imports System.Windows
Imports System.Windows.Automation.Peers
Imports System.Windows.Controls
Imports System.Windows.Media
Imports System.Windows.Media.Animation
Imports System.Windows.Media.Imaging
Imports System.Windows.Threading
Imports Beacon

Module SplashChecks
    Public Sub SceneAndMotion()
        Dim window As New SplashWindow(False) With {
            .WindowStartupLocation = WindowStartupLocation.Manual,
            .Left = -10000, .Top = -10000, .Topmost = False,
            .ShowActivated = False, .ShowInTaskbar = False
        }
        Dim flags = BindingFlags.Instance Or BindingFlags.NonPublic
        Dim motionField = GetType(SplashWindow).GetField("_motionStoryboard", flags)
        Dim setMotion = GetType(SplashWindow).GetMethod("UpdateAnimationState", flags)
        Try
            window.Show()
            Pump(window)
            Require(Not window.AllowsTransparency, "Splash uses a per-pixel transparent window.")
            Dim scene = DirectCast(window.FindName("SceneRoot"), Grid)
            Dim logo = DirectCast(window.FindName("BeaconLogo"), Image)
            Require(logo.Source IsNot Nothing AndAlso logo.ActualWidth = 180, "Beacon logo did not load.")
            Dim timeline = DirectCast(window.FindResource("AmbientMotion"), Storyboard)
            Require(System.Windows.Media.Animation.Timeline.GetDesiredFrameRate(timeline).GetValueOrDefault() = 60, "Animation frame-rate target changed.")
            Require(timeline.Children.Count <= 16, "Splash animation count grew unexpectedly.")
            For Each child In timeline.Children
                Require(window.FindName(Storyboard.GetTargetName(child)) IsNot Nothing, "Unresolved animation target.")
                Require(child.RepeatBehavior = RepeatBehavior.Forever, "Ambient motion does not loop.")
                Require({"X", "Y", "Opacity", "Angle"}.Contains(Storyboard.GetTargetProperty(child).Path), "Animation changes layout rather than transforms/opacity.")
            Next
            Dim floatAnimation = DirectCast(timeline.Children.Single(Function(child) Storyboard.GetTargetName(child) = "LogoMotion"), DoubleAnimation)
            Require(floatAnimation.To.HasValue AndAlso floatAnimation.To.Value = -3 AndAlso floatAnimation.AutoReverse AndAlso floatAnimation.Duration.TimeSpan.TotalSeconds * 2 = 4.5,
                    "Logo float no longer matches the requested amplitude/cycle.")
            Dim beam = timeline.Children.Single(Function(child) Storyboard.GetTargetName(child) = "BeamMotion")
            Require(beam.Duration.TimeSpan.TotalSeconds = 3, "Beam cycle is not three seconds.")
            Dim rotation = DirectCast(timeline.Children.Single(Function(child) Storyboard.GetTargetName(child) = "LoadingRotation"), DoubleAnimation)
            Require(rotation.From.HasValue AndAlso rotation.To.HasValue AndAlso rotation.From.Value = 0 AndAlso rotation.To.Value = 360 AndAlso Not rotation.AutoReverse AndAlso rotation.EasingFunction Is Nothing,
                    "Spinner must rotate continuously at a steady angular speed.")
            Dim far = timeline.Children.Single(Function(child) Storyboard.GetTargetName(child) = "FarDrift").Duration.TimeSpan
            Dim middle = timeline.Children.Single(Function(child) Storyboard.GetTargetName(child) = "MiddleDrift").Duration.TimeSpan
            Dim near = timeline.Children.Single(Function(child) Storyboard.GetTargetName(child) = "NearDrift").Duration.TimeSpan
            Require(far > middle AndAlso middle > near, "Parallax layers do not have distinct speeds.")
            Dim labels = Descendants(scene).OfType(Of TextBlock)().Select(Function(label) label.Text).ToArray()
            For Each kind In {"EVTX", "TXT", "LOG", "XML", "JSON", "ZIP"}
                Require(labels.Contains(kind), "Missing file-type badge: " & kind)
            Next
            Dim status = DirectCast(window.FindName("LoadingStatus_txt"), TextBlock)
            Require(status.Text = "Preparing your workspace…", "Decorative logs are presented as actual scan progress.")
            Dim peer = UIElementAutomationPeer.CreatePeerForElement(window)
            Dim children = peer.GetChildren()
            Require(children.Count = 1 AndAlso children(0).GetName() = status.Text, "Decorative logs leaked into the accessibility tree.")

            setMotion.Invoke(window, {False})
            WaitUntil(window, Function() Not DirectCast(window.FindName("LogoMotion"), TranslateTransform).HasAnimatedProperties)
            Require(motionField.GetValue(window) Is Nothing, "Reduced-motion state retained a storyboard.")
            Require(DirectCast(window.FindName("ScanBeam"), Canvas).Opacity = 0, "Reduced-motion state retained a scanning beam.")
            setMotion.Invoke(window, {True})
            Dim active = DirectCast(motionField.GetValue(window), Storyboard)
            Pump(window)
            Dim count = Descendants(scene).Count()
            Require(count < 180, "Splash visual tree is unexpectedly large.")
            For Each seconds In {0.5, 1.5, 3.0, 4.5, 12.0}
                active.SeekAlignedToLastTick(window, TimeSpan.FromSeconds(seconds), TimeSeekOrigin.BeginTime)
                Pump(window)
                Dim y = DirectCast(window.FindName("LogoMotion"), TranslateTransform).Y
                Require(y >= -3.01 AndAlso y <= 0.01, "Logo moved outside the subtle float range.")
                Require(Descendants(scene).Count() = count, "Animation created additional visual objects.")
            Next
            active.SeekAlignedToLastTick(window, TimeSpan.FromSeconds(1.5), TimeSeekOrigin.BeginTime)
            Pump(window)
            Dim image As New RenderTargetBitmap(900, 600, 96, 96, PixelFormats.Pbgra32)
            image.Render(scene)
            Dim encoder As New PngBitmapEncoder()
            encoder.Frames.Add(BitmapFrame.Create(image))
            Dim outputDirectory = Path.Combine(AppContext.BaseDirectory, "artifacts")
            Directory.CreateDirectory(outputDirectory)
            Using output = File.Create(Path.Combine(outputDirectory, "splash-preview.png"))
                encoder.Save(output)
            End Using
            window.Width = 600
            window.Height = 400
            Pump(window)
            Require(scene.ActualWidth = 900 AndAlso scene.ActualHeight = 600, "Viewbox scaling changed the scene's layout geometry.")
        Finally
            window.Close()
        End Try
        WaitUntil(window, Function() Not DirectCast(window.FindName("LogoMotion"), TranslateTransform).HasAnimatedProperties)
        Require(motionField.GetValue(window) Is Nothing AndAlso GetType(SplashWindow).GetField("_autoCloseTimer", flags).GetValue(window) Is Nothing,
                "Closed splash retained animation or timer state.")
        window.TransitionToMainWindow()
        Require(GetType(SplashWindow).GetField("_fadeStoryboard", flags).GetValue(window) Is Nothing, "Closed splash restarted a transition.")
    End Sub

    Public Sub WorkAreaSizing()
        For Each area In {New Size(1920, 1080), New Size(800, 600), New Size(640, 480), New Size(3840, 2160)}
            Dim size = SplashWindow.CalculateSplashSize(area)
            Require(size.Width <= area.Width - 32 AndAlso size.Height <= area.Height - 32, "Splash does not fit the work area.")
            Require(size.Width <= 900 AndAlso size.Height <= 600 AndAlso Math.Abs(size.Width / size.Height - 1.5) < 0.001,
                    "Splash size lost its aspect ratio or was unnecessarily enlarged.")
        Next
    End Sub

    Private Sub WaitUntil(window As Window, condition As Func(Of Boolean))
        Dim timer = Stopwatch.StartNew()
        While Not condition()
            If timer.Elapsed > TimeSpan.FromSeconds(2) Then Throw New TimeoutException("Animation cleanup did not complete.")
            Pump(window)
            Thread.Sleep(20)
        End While
    End Sub

    Private Sub Pump(window As Window)
        window.UpdateLayout()
        window.Dispatcher.Invoke(Sub()
                                 End Sub, DispatcherPriority.ContextIdle)
    End Sub

    Private Iterator Function Descendants(parent As DependencyObject) As IEnumerable(Of DependencyObject)
        For index = 0 To VisualTreeHelper.GetChildrenCount(parent) - 1
            Dim child = VisualTreeHelper.GetChild(parent, index)
            Yield child
            For Each descendant In Descendants(child)
                Yield descendant
            Next
        Next
    End Function

    Private Sub Require(condition As Boolean, message As String)
        If Not condition Then Throw New InvalidOperationException(message)
    End Sub
End Module
