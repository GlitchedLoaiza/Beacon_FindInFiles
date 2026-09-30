Imports System.ComponentModel
Imports System.IO
Imports System.Linq
Imports System.Reflection
Imports System.Windows
Imports System.Windows.Automation
Imports System.Windows.Controls
Imports System.Windows.Documents
Imports System.Windows.Input
Imports System.Windows.Media
Imports System.Windows.Media.Imaging
Imports System.Windows.Shell
Imports System.Windows.Threading
Imports Beacon

Module AeroThemeChecks
    Private Const PrivateInstance As BindingFlags = BindingFlags.Instance Or BindingFlags.NonPublic

    Public Sub PersistenceAndContrast()
        EnsureApplication()
        Require(CInt(AppTheme.Aero) = 4 AndAlso BeaconSettings.CreateDefaults().Theme = AppTheme.System,
                "Aero must be appended without changing existing theme values or defaults.")
        Dim settings = BeaconSettingsService.Validate(BeaconSettingsService.Clone(New BeaconSettings With {.Theme = AppTheme.Aero}))
        Require(settings.Theme = AppTheme.Aero, "Aero did not survive settings serialization.")
        Require(BeaconSettingsService.Deserialize("{""Theme"":""Aero""}").Theme = AppTheme.Aero, "Named Aero settings could not be loaded.")
        For Each dark In {False, True}
            Require(BeaconThemePalette.UsesDarkBackground(AppTheme.Aero, dark), "Aero must remain dark in both system modes.")
        Next
        Require(Not BeaconThemePalette.FollowsSystem(AppTheme.Aero), "Aero incorrectly follows system dark/light mode.")

        Dim palette = AeroTheme.CreateResources(False)
        Require(AeroTheme.IsEnabled(palette) AndAlso Not BeaconThemePalette.IsEnabled(palette), "Aero was confused with Beacon Theme.")
        Require(TypeOf palette("ActionButtonBackgroundBrush") Is LinearGradientBrush, "Aero glossy buttons are missing.")
        Require(TypeOf palette("CardBackgroundBrush") Is SolidColorBrush, "Log previews must retain opaque solid backgrounds.")
        For Each key In {"ActionButtonBackgroundBrush", "ActionButtonHoverBrush", "ActionButtonPressedBrush", "PrimaryButtonBrush"}
            CheckContrast(DirectCast(palette("TextPrimaryBrush"), Brush), DirectCast(palette(key), Brush), key)
        Next
        CheckContrast(DirectCast(palette("TextPrimaryBrush"), Brush), DirectCast(palette("InputBackgroundBrush"), Brush), "Input text")
        CheckContrast(DirectCast(palette("TextSecondaryBrush"), Brush), DirectCast(palette("CardBackgroundBrush"), Brush), "Secondary text")
        CheckContrast(DirectCast(palette("TextSecondaryBrush"), Brush), DirectCast(palette("SelectionBackgroundBrush"), Brush), "Selected result metadata")
        CheckContrast(DirectCast(palette("AeroLabelBrush"), Brush), DirectCast(palette("AeroPanelHeaderBrush"), Brush), "Panel heading")
        CheckContrast(DirectCast(palette("TextSelectionForegroundBrush"), Brush), DirectCast(palette("TextSelectionBrush"), Brush), "Selected text")
        CheckContrast(DirectCast(palette("CheckmarkBrush"), Brush), DirectCast(palette("AccentBrush"), Brush), "Check mark")
        CheckContrast(DirectCast(palette("EventLevelForegroundBrush"), Brush), DirectCast(palette("CardBackgroundBrush"), Brush), "Event level")

        Dim accessible = AeroTheme.CreateResources(True)
        Require(DirectCast(accessible("CardBackgroundBrush"), SolidColorBrush).Color = SystemColors.WindowColor, "High contrast ignored the system background.")
        Require(DirectCast(accessible("TextPrimaryBrush"), SolidColorBrush).Color = SystemColors.WindowTextColor, "High contrast ignored system text colors.")
        Require(DirectCast(accessible("SelectionBackgroundBrush"), SolidColorBrush).Color = SystemColors.WindowColor AndAlso
                DirectCast(accessible("SelectionBorderBrush"), SolidColorBrush).Color = SystemColors.WindowTextColor,
                "High-contrast row selection must remain visible without obscuring explicitly colored metadata.")
        Require(DirectCast(accessible("TextSelectionForegroundBrush"), SolidColorBrush).Color = SystemColors.HighlightTextColor, "High contrast selection text is incorrect.")
        Require(DirectCast(accessible("ActionButtonBackgroundBrush"), SolidColorBrush).Color = SystemColors.WindowColor, "High contrast kept glossy gradients.")
        Require(DirectCast(accessible("EventLevelForegroundBrush"), SolidColorBrush).Color = SystemColors.WindowTextColor, "High contrast did not replace event-level text colors.")
        Require(DirectCast(accessible("AeroDecorationVisibility"), Visibility) = Visibility.Collapsed, "High contrast retained decorative glow.")
    End Sub

    Public Sub MainWindowTransitions()
        EnsureApplication()
        Dim window As New MainWindow()
        DetachStartup(window)
        Prepare(window)
        Dim settings As New BeaconSettings With {.Theme = AppTheme.Dark}
        GetType(MainWindow).GetField("_settings", PrivateInstance).SetValue(window, settings)
        Dim apply = GetType(MainWindow).GetMethod("ApplyThemeSelection", PrivateInstance)
        Try
            apply.Invoke(window, {False})
            Dim originalButtonStyle = window.Resources(GetType(Button))
            Dim originalTemplate = window.ReadLocalValue(Control.TemplateProperty)
            Dim originalIcon = window.Icon
            Dim originalTitle = window.Title
            window.Show()
            Pump(window)
            Dim search = DirectCast(window.FindName("Search_txt"), TextBox)
            Dim preview = DirectCast(window.FindName("TextPreview_rtb"), RichTextBox)
            Dim results = DirectCast(window.FindName("Results_lst"), ListBox)
            Dim source = results.ItemsSource
            Dim document As New FlowDocument(New Paragraph(New Run("Aero theme test: readable log content.")))
            preview.Document = document
            search.Text = "error OR timeout"
            settings.Theme = AppTheme.Aero
            apply.Invoke(window, {False})
            Pump(window)
            Require(AeroTheme.IsEnabled(window.Resources), "Main window did not apply Aero.")
            Require(CBool(GetType(MainWindow).GetField("_isDarkMode", PrivateInstance).GetValue(window)), "Aero previews were treated as light mode.")
            Require(Not window.AllowsTransparency AndAlso window.WindowStyle <> WindowStyle.None, "Aero replaced the native window contract with a transparent window.")
            Require(window.Title = originalTitle AndAlso window.Icon Is originalIcon, "Aero changed the application title or taskbar icon.")
            Require(preview.Document Is document AndAlso results.ItemsSource Is source AndAlso search.Text = "error OR timeout", "Changing themes reset search or preview content.")
            Require(preview.Template.FindName("PART_ContentHost", preview) IsNot Nothing, "Aero removed the text preview content host.")
            Require(search.Template.FindName("PART_ContentHost", search) IsNot Nothing, "Aero removed the text editor content host.")
            Dim mode = DirectCast(window.FindName("SearchMode_cmb"), ComboBox)
            Require(AeroTheme.IsEnabled(mode.Resources), "Locally scoped dropdown styles did not inherit Aero.")
            Require(mode.Template.FindName("PART_Popup", mode) IsNot Nothing, "Aero dropdown popup is missing.")
            If Not SystemParameters.HighContrast Then
                Require(WindowChrome.GetWindowChrome(window) IsNot Nothing, "Aero custom chrome is missing.")
                Dim folder = DirectCast(window.FindName("BrowseFolder_btn"), Button)
                Require(DirectCast(folder.Template.FindName("ActionIcon", folder), FrameworkElement).Visibility = Visibility.Visible, "Aero folder icon is missing.")
                Require(CStr(DirectCast(folder.Template.FindName("ButtonContent", folder), ContentPresenter).Content) = "Folder", "Aero retained the emoji instead of its vector icon.")
                Require(DirectCast(window.FindName("ToolbarWatermark"), Image).Visibility = Visibility.Collapsed AndAlso
                        DirectCast(window.FindName("AeroWatermark"), Border).Visibility = Visibility.Visible, "The cyan emblem did not replace the legacy watermark.")
            Else
                Require(WindowChrome.GetWindowChrome(window) Is Nothing, "High contrast must use native captions.")
            End If

            For Each width In {1000.0, 1440.0}
                window.Width = width
                window.Height = 740
                Pump(window)
                For Each name In {"SourceActionsPanel", "SearchActionsPanel", "ReportActionsPanel"}
                    Dim panel = DirectCast(window.FindName(name), StackPanel)
                    Dim buttons = panel.Children.OfType(Of Button)().ToArray()
                    For index = 1 To buttons.Length - 1
                        Dim previous = buttons(index - 1).TranslatePoint(New Point(), panel)
                        Dim current = buttons(index).TranslatePoint(New Point(), panel)
                        Require(current.X >= previous.X + buttons(index - 1).ActualWidth + 7, "Aero action buttons overlap: " & name)
                    Next
                    Dim position = panel.TranslatePoint(New Point(), window)
                    Require(position.X >= 0 AndAlso position.X + panel.ActualWidth <= window.ActualWidth + 1, "Aero toolbar clipped at minimum width: " & name)
                Next
                Require(preview.ActualWidth > 400 AndAlso preview.ActualHeight > 200, "Aero frame crowded out the preview.")
            Next
            window.Width = 1000
            window.Height = 700
            Pump(window)
            If Not SystemParameters.HighContrast Then SavePreview(window)

            For Each theme In {AppTheme.Light, AppTheme.Aero, AppTheme.Dark, AppTheme.Aero, AppTheme.Beacon, AppTheme.System}
                settings.Theme = theme
                apply.Invoke(window, {False})
                Pump(window)
                Require(preview.Document Is document AndAlso results.ItemsSource Is source AndAlso search.Text = "error OR timeout", "Theme transition discarded user state.")
                If theme = AppTheme.Aero Then
                    Require(AeroTheme.IsEnabled(window.Resources), "Aero failed when reselected.")
                Else
                    Require(Not AeroTheme.IsEnabled(window.Resources) AndAlso Not AeroTheme.IsEnabled(mode.Resources), "Aero resource overlays leaked after switching themes.")
                    Require(WindowChrome.GetWindowChrome(window) Is Nothing, "Custom chrome remained after leaving Aero.")
                    Require(window.ReadLocalValue(Control.TemplateProperty) Is originalTemplate, "Original window template was not restored.")
                    Require(window.Resources(GetType(Button)) Is originalButtonStyle, "Original main button style was not restored.")
                    Require(TypeOf DirectCast(window.FindName("BrowseFolder_btn"), Button).Background Is SolidColorBrush, "Aero gradients leaked into another theme.")
                    NativeCaptionChecks.Verify(window, BeaconThemePalette.UsesDarkBackground(theme, False))
                End If
            Next
        Finally
            window.Close()
        End Try
    End Sub

    Public Sub OwnedWindowsAndChrome()
        EnsureApplication()
        Dim owner As New SettingsWindow(New BeaconSettings With {.Theme = AppTheme.Aero}, False)
        Prepare(owner)
        Dim help As HelpWindow = Nothing
        Dim report As ScanReportWindow = Nothing
        Dim tour As TourWindow = Nothing
        Try
            owner.Show()
            Pump(owner)
            Require(CStr(DirectCast(owner.FindName("Theme_cmb"), ComboBox).SelectedValue) = "Aero", "Settings did not select saved Aero mode.")
            Require(AeroTheme.IsEnabled(owner.Resources), "Settings did not apply Aero in a light system environment.")
            Dim tabs = DirectCast(owner.FindName("SettingsTabs"), TabControl)
            tabs.SelectedIndex = 4
            Pump(owner)
            Dim selector = DirectCast(owner.FindName("Theme_cmb"), ComboBox)
            selector.IsDropDownOpen = True
            Pump(owner)
            Require(DirectCast(selector.Template.FindName("PART_Popup", selector), System.Windows.Controls.Primitives.Popup).IsOpen, "Aero theme choices cannot be opened.")
            selector.IsDropDownOpen = False
            Require(owner.SavedSettings Is Nothing, "Opening Aero settings changed the saved configuration.")

            help = New HelpWindow(owner, True)
            report = New ScanReportWindow(owner, ReportChecks.Sample(Path.GetTempPath()), Nothing, False, True)
            tour = New TourWindow(owner, True)
            For Each child As Window In New Window() {help, report, tour}
                Prepare(child)
                child.Show()
                Pump(child)
                Require(AeroTheme.IsEnabled(child.Resources), "An owned window did not inherit Aero.")
                Require(child.Resources("ActionButtonBackgroundBrush").GetType() Is If(SystemParameters.HighContrast, GetType(SolidColorBrush), GetType(LinearGradientBrush)), "An owned window lost its button palette.")
                Require(DirectCast(child.Resources("TextSelectionBrush"), SolidColorBrush).Color = DirectCast(owner.Resources("TextSelectionBrush"), SolidColorBrush).Color, "Owned-window selection colors differ.")
            Next
            Require(tour.WindowStyle = WindowStyle.None AndAlso tour.AllowsTransparency AndAlso WindowChrome.GetWindowChrome(tour) Is Nothing,
                    "Aero added a frame to the borderless welcome tour.")
            If Not SystemParameters.HighContrast Then CheckCaptionCommands(owner)
            AeroTheme.Apply(owner, False)
            BeaconThemePalette.ApplyButtons(owner.Resources, False, True)
            For Each child As Window In New Window() {help, report, tour}
                BeaconThemePalette.CopyOwnerColors(owner, child, True)
                Pump(child)
                Require(Not AeroTheme.IsEnabled(child.Resources) AndAlso WindowChrome.GetWindowChrome(child) Is Nothing,
                        "Owned window retained Aero after the owner switched themes.")
            Next
        Finally
            tour?.Close()
            report?.Close()
            help?.Close()
            owner.Close()
        End Try
    End Sub

    Public Sub ProgressAppearance()
        EnsureApplication()
        Dim animated = AeroTheme.CreateResources(False, True)
        Dim reduced = AeroTheme.CreateResources(False, False)
        Dim accessible = AeroTheme.CreateResources(True, True)
        Require(animated("AeroProgressTemplate") Is animated("AeroProgressAnimatedTemplate"), "Aero did not select its native moving-glow template.")
        Require(reduced("AeroProgressTemplate") Is reduced("AeroProgressStaticTemplate") AndAlso
                accessible("AeroProgressTemplate") Is accessible("AeroProgressStaticTemplate"), "Motion-disabled progress must use the static template.")
        Dim green = DirectCast(animated("AeroProgressFillBrush"), LinearGradientBrush)
        Require(green.GradientStops.All(Function(item) item.Color.G > item.Color.R AndAlso item.Color.G > item.Color.B), "Aero progress lost its classic-green palette.")
        Require(DirectCast(accessible("AeroProgressFillBrush"), SolidColorBrush).Color = SystemColors.HighlightColor AndAlso
                DirectCast(accessible("AeroProgressTrackBrush"), SolidColorBrush).Color = SystemColors.WindowColor, "High-contrast progress did not use system colors.")

        Dim window As New MainWindow()
        DetachStartup(window)
        Prepare(window)
        Dim settings As New BeaconSettings With {.Theme = AppTheme.Dark}
        GetType(MainWindow).GetField("_settings", PrivateInstance).SetValue(window, settings)
        Dim apply = GetType(MainWindow).GetMethod("ApplyThemeSelection", PrivateInstance)
        Dim showProgress = GetType(MainWindow).GetMethod("ShowProgress", PrivateInstance)
        Try
            apply.Invoke(window, {False})
            window.Show()
            Pump(window)
            Dim bar = DirectCast(window.FindName("ScanProgress_pb"), ProgressBar)
            Dim originalStyle = bar.Style
            settings.Theme = AppTheme.Aero
            apply.Invoke(window, {False})
            window.Resources("AeroProgressTemplate") = animated("AeroProgressTemplate")
            bar.Minimum = 10
            bar.Maximum = 110
            bar.Value = 35
            showProgress.Invoke(window, {True})
            Pump(window)
            Require(bar.Width = 180 AndAlso bar.Height = 16, "Aero changed the status-bar layout.")
            CheckProgressWidth(bar, 0.25)
            Dim glow = DirectCast(bar.Template.FindName("PART_GlowRect", bar), FrameworkElement)
            Require(glow IsNot Nothing AndAlso glow.HasAnimatedProperties, "Aero determinate progress has no moving shine.")
            WaitForProgressShine(window, glow)
            Require(bar.Value = 35 AndAlso Not DependencyPropertyHelper.GetValueSource(bar, System.Windows.Controls.Primitives.RangeBase.ValueProperty).IsAnimated,
                    "The decorative animation changed the reported progress value.")
            SaveProgressPreview(bar, "aero-progress-determinate")

            showProgress.Invoke(window, {False})
            Pump(window)
            Require(Not glow.HasAnimatedProperties, "The progress shine kept animating while hidden.")
            showProgress.Invoke(window, {True})
            Pump(window)
            Require(glow.HasAnimatedProperties, "The progress shine did not resume when shown.")
            bar.Value = bar.Minimum
            Pump(window)
            CheckProgressWidth(bar, 0)
            Require(Not glow.HasAnimatedProperties, "Empty progress should not animate a shine.")
            bar.Value = 70
            bar.IsIndeterminate = True
            Pump(window)
            CheckProgressWidth(bar, 1)
            Require(DirectCast(bar.Template.FindName("ProgressFill", bar), Border).Visibility = Visibility.Collapsed AndAlso glow.HasAnimatedProperties,
                    "Indeterminate progress must show a moving segment rather than a completed fill.")
            Require(bar.Value = 70, "Indeterminate styling changed the stored progress value.")
            SaveProgressPreview(bar, "aero-progress-indeterminate")
            bar.IsIndeterminate = False
            Pump(window)
            CheckProgressWidth(bar, 0.6)

            window.Resources("AeroProgressTemplate") = reduced("AeroProgressTemplate")
            Pump(window)
            Require(bar.Template.FindName("PART_GlowRect", bar) Is Nothing, "Reduced motion retained a glow animation target.")
            CheckProgressWidth(bar, 0.6)
            bar.IsIndeterminate = True
            Pump(window)
            Dim segment = DirectCast(bar.Template.FindName("ProgressFill", bar), Border)
            Require(segment.ActualWidth = 64 AndAlso segment.HorizontalAlignment = HorizontalAlignment.Center,
                    "Static indeterminate progress looks like a completed operation.")
            SaveProgressPreview(bar, "aero-progress-reduced-motion")

            window.Resources("AeroProgressTemplate") = accessible("AeroProgressTemplate")
            For Each key In {"AeroProgressFillBrush", "AeroProgressTrackBrush", "AeroProgressBorderBrush", "AeroProgressFillBorderBrush"}
                window.Resources(key) = accessible(key)
            Next
            Pump(window)
            Require(bar.Template.FindName("PART_GlowRect", bar) Is Nothing AndAlso
                    DirectCast(bar.Foreground, SolidColorBrush).Color = SystemColors.HighlightColor, "High contrast retained decorative animation or invisible fill colors.")
            SaveProgressPreview(bar, "aero-progress-high-contrast")

            Dim stateProperty = DirectCast(GetType(AeroTheme).GetField("StateProperty", BindingFlags.Static Or BindingFlags.NonPublic).GetValue(Nothing), DependencyProperty)
            Dim state = window.GetValue(stateProperty)
            state.GetType().GetMethod("PreferenceChanged", PrivateInstance).Invoke(state,
                {Nothing, New PropertyChangedEventArgs(NameOf(SystemParameters.ClientAreaAnimation))})
            Pump(window)
            Require((bar.Template.FindName("PART_GlowRect", bar) IsNot Nothing) = (SystemParameters.ClientAreaAnimation AndAlso Not SystemParameters.HighContrast),
                    "Aero did not refresh the progress template after an animation-preference notification.")
            Require(bar.Value = 70 AndAlso bar.IsIndeterminate, "Refreshing animation preferences changed the progress state.")
            bar.IsIndeterminate = False
            bar.Orientation = Orientation.Vertical
            bar.Width = 16
            bar.Height = 180
            Pump(window)
            CheckProgressWidth(bar, 0.6)

            settings.Theme = AppTheme.Dark
            apply.Invoke(window, {False})
            Pump(window)
            Require(bar.Style Is originalStyle AndAlso Not window.Resources.Contains("AeroProgressTemplate"), "Aero progress styling leaked into another theme.")
            Require(bar.Value = 70 AndAlso bar.Minimum = 10 AndAlso bar.Maximum = 110, "Restoring the standard progress style changed its range or value.")
        Finally
            window.Close()
        End Try
    End Sub

    Private Sub CheckProgressWidth(bar As ProgressBar, fraction As Double)
        Dim track = DirectCast(bar.Template.FindName("PART_Track", bar), FrameworkElement)
        Dim indicator = DirectCast(bar.Template.FindName("PART_Indicator", bar), FrameworkElement)
        Require(track IsNot Nothing AndAlso indicator IsNot Nothing AndAlso track.ActualWidth > 0, "Aero progress template parts are missing or unmeasured.")
        Require(Math.Abs(indicator.ActualWidth - track.ActualWidth * fraction) <= 1, "The visible progress fill no longer matches the real value.")
    End Sub

    Private Sub WaitForProgressShine(window As Window, glow As FrameworkElement)
        Dim initial = glow.Margin.Left
        Dim watch = System.Diagnostics.Stopwatch.StartNew()
        While Math.Abs(glow.Margin.Left - initial) < 0.5 AndAlso watch.Elapsed < TimeSpan.FromSeconds(3)
            System.Threading.Thread.Sleep(20)
            Pump(window)
        End While
        Require(Math.Abs(glow.Margin.Left - initial) >= 0.5, "The native Aero progress shine is not moving.")
    End Sub

    Private Sub SaveProgressPreview(bar As ProgressBar, name As String)
        Dim drawing As New DrawingVisual()
        Using context = drawing.RenderOpen()
            context.DrawRectangle(New VisualBrush(bar), Nothing, New Rect(0, 0, bar.ActualWidth, bar.ActualHeight))
        End Using
        Dim image As New RenderTargetBitmap(CInt(Math.Ceiling(bar.ActualWidth * 3)), CInt(Math.Ceiling(bar.ActualHeight * 3)), 288, 288, PixelFormats.Pbgra32)
        image.Render(drawing)
        Dim encoder As New PngBitmapEncoder()
        encoder.Frames.Add(BitmapFrame.Create(image))
        Dim directory = Path.Combine(AppContext.BaseDirectory, "artifacts")
        IO.Directory.CreateDirectory(directory)
        Dim outputPath = Path.Combine(directory, name & ".png")
        Using output = File.Create(outputPath)
            encoder.Save(output)
        End Using
        Require(New FileInfo(outputPath).Length > 500, "The progress preview render is unexpectedly empty.")
        Console.WriteLine("Aero progress preview: " & outputPath)
    End Sub

    Public Sub SelectionReadability()
        EnsureApplication()
        Dim window As New MainWindow()
        DetachStartup(window)
        Prepare(window)
        Dim settings As New BeaconSettings With {.Theme = AppTheme.Aero}
        GetType(MainWindow).GetField("_settings", PrivateInstance).SetValue(window, settings)
        Dim apply = GetType(MainWindow).GetMethod("ApplyThemeSelection", PrivateInstance)
        Try
            apply.Invoke(window, {False})
            window.Width = 1000
            window.Height = 700
            window.Show()
            Pump(window)
            Dim preview = DirectCast(window.FindName("TextPreview_rtb"), RichTextBox)
            Dim search = DirectCast(window.FindName("Search_txt"), TextBox)
            Dim sample = "GPO: PS-WINDOWS11" & vbLf &
                         "Folder Id: Cryptography\AutoEnrollment\OfflineExpirationPercent" & vbLf &
                         "Value: 15, 0, 0, 0"
            Dim sourceRun As New Run(sample)
            preview.FontFamily = New FontFamily("Consolas")
            preview.FontSize = 14
            preview.Document = New FlowDocument(New Paragraph(sourceRun)) With {.PagePadding = New Thickness(8)}
            preview.IsInactiveSelectionHighlightEnabled = True
            preview.Selection.Select(sourceRun.ContentStart, sourceRun.ContentStart)
            Pump(window)
            Const selectedWord As String = "AutoEnrollment"
            Dim offset = sample.IndexOf(selectedWord, StringComparison.Ordinal)
            WithActiveSelection(preview,
                Sub()
                    Pump(window)
                    Dim before As Byte() = Nothing
                    If Not SystemParameters.HighContrast Then before = CaptureFramePixels(window)
                    preview.Selection.Select(sourceRun.ContentStart.GetPositionAtOffset(offset), sourceRun.ContentStart.GetPositionAtOffset(offset + selectedWord.Length))
                    Pump(window)
                    Require(preview.IsReadOnly AndAlso preview.Selection.Text = selectedWord, "Preview highlighting changed the selected range or read-only behavior.")
                    If before IsNot Nothing Then
                        SavePreview(window, "aero-selection")
                        CheckSelectedGlyphs(before, CaptureFramePixels(window))
                    End If
                End Sub)

            For Each theme In {AppTheme.Aero, AppTheme.Dark, AppTheme.Aero}
                settings.Theme = theme
                apply.Invoke(window, {False})
                Pump(window)
                Require(preview.Selection.Text = selectedWord, "Theme switching discarded the selected log text.")
                If theme <> AppTheme.Aero Then Continue For
                Require(preview.SelectionOpacity > 0 AndAlso preview.SelectionOpacity < 1, "Aero RichTextBox selection must not paint an opaque rectangle over the glyphs.")
                Dim overlay = DirectCast(preview.SelectionBrush, SolidColorBrush).Color
                CheckContrast(New SolidColorBrush(BlendSelection(DirectCast(preview.Foreground, SolidColorBrush).Color, overlay, preview.SelectionOpacity)),
                              New SolidColorBrush(BlendSelection(DirectCast(preview.Background, SolidColorBrush).Color, overlay, preview.SelectionOpacity)),
                              "Selected Aero preview glyphs through the overlay")
                Require(search.SelectionOpacity = 1 AndAlso
                        DirectCast(search.SelectionBrush, SolidColorBrush).Color = DirectCast(window.Resources("TextSelectionBrush"), SolidColorBrush).Color AndAlso
                        DirectCast(search.SelectionTextBrush, SolidColorBrush).Color = DirectCast(window.Resources("TextSelectionForegroundBrush"), SolidColorBrush).Color,
                        "Fixing the preview changed editable-input selection.")
            Next
        Finally
            window.Close()
        End Try
    End Sub

    Private Sub WithActiveSelection(preview As RichTextBox, action As Action)
        ' Exercise the active-selection adorner without changing the desktop's keyboard focus.
        Dim key = DirectCast(GetType(UIElement).GetField("IsKeyboardFocusedPropertyKey", BindingFlags.Static Or BindingFlags.NonPublic).GetValue(Nothing), DependencyPropertyKey)
        Dim original = preview.GetValue(key.DependencyProperty)
        preview.SetValue(key, True)
        preview.RaiseEvent(New KeyboardFocusChangedEventArgs(Keyboard.PrimaryDevice, Environment.TickCount, Nothing, preview) With {.RoutedEvent = Keyboard.GotKeyboardFocusEvent})
        Try
            action()
        Finally
            preview.SetValue(key, original)
            preview.RaiseEvent(New KeyboardFocusChangedEventArgs(Keyboard.PrimaryDevice, Environment.TickCount, preview, Nothing) With {.RoutedEvent = Keyboard.LostKeyboardFocusEvent})
        End Try
    End Sub

    Private Function CaptureFramePixels(window As Window) As Byte()
        Dim visual = DirectCast(window.Template.FindName("AeroFrame", window), FrameworkElement)
        Dim image As New RenderTargetBitmap(CInt(Math.Ceiling(visual.ActualWidth)), CInt(Math.Ceiling(visual.ActualHeight)), 96, 96, PixelFormats.Pbgra32)
        image.Render(visual)
        Dim stride = image.PixelWidth * 4
        Dim pixels(stride * image.PixelHeight - 1) As Byte
        image.CopyPixels(pixels, stride, 0)
        Return pixels
    End Function

    Private Sub CheckSelectedGlyphs(before As Byte(), after As Byte())
        Require(before.Length = after.Length, "Selection changed the rendered preview bounds.")
        Dim highlightedGlyphPixels As Integer
        Dim readableGlyphPixels As Integer
        Dim changedPixels As Integer
        For index = 0 To before.Length - 4 Step 4
            If before(index + 3) <> 255 OrElse after(index + 3) <> 255 Then Continue For
            If before(index) = after(index) AndAlso before(index + 1) = after(index + 1) AndAlso before(index + 2) = after(index + 2) Then Continue For
            changedPixels += 1
            Dim original = Color.FromRgb(before(index + 2), before(index + 1), before(index))
            If Luminance(original) < 0.3 Then Continue For
            highlightedGlyphPixels += 1
            Dim selected = Color.FromRgb(after(index + 2), after(index + 1), after(index))
            If Luminance(selected) >= 0.25 Then readableGlyphPixels += 1
        Next
        Require(highlightedGlyphPixels >= 30, $"The rendered check did not exercise a visible text-selection overlay: {changedPixels} changed pixels, {highlightedGlyphPixels} glyph pixels.")
        Require(readableGlyphPixels >= highlightedGlyphPixels * 0.8,
                $"Aero selection obscured the rendered text: only {readableGlyphPixels} of {highlightedGlyphPixels} affected glyph pixels remain readable.")
    End Sub

    Private Function BlendSelection(underlying As Color, overlay As Color, opacity As Double) As Color
        Return Color.FromRgb(CByte(Math.Round(underlying.R * (1 - opacity) + overlay.R * opacity)),
                             CByte(Math.Round(underlying.G * (1 - opacity) + overlay.G * opacity)),
                             CByte(Math.Round(underlying.B * (1 - opacity) + overlay.B * opacity)))
    End Function

    Private Sub CheckCaptionCommands(window As Window)
        Dim minimize = DirectCast(window.Template.FindName("AeroMinimizeButton", window), Button)
        Dim maximize = DirectCast(window.Template.FindName("AeroMaximizeButton", window), Button)
        Dim close = DirectCast(window.Template.FindName("AeroCloseButton", window), Button)
        Require(AutomationProperties.GetName(minimize) = "Minimize" AndAlso AutomationProperties.GetName(close) = "Close", "Caption controls have no accessible names.")
        Require(WindowChrome.GetIsHitTestVisibleInChrome(maximize), "Caption controls are not interactive in the non-client area.")
        Dim command = DirectCast(maximize.Command, RoutedCommand)
        Require(maximize.CommandTarget Is window AndAlso command.CanExecute(Nothing, window), "Caption maximize command is not wired to the window.")
        command.Execute(Nothing, window)
        Pump(window)
        Require(window.WindowState = WindowState.Maximized AndAlso AutomationProperties.GetName(maximize) = "Restore", "Caption maximize did not update state and accessible text.")
        command.Execute(Nothing, window)
        Pump(window)
        Require(window.WindowState = WindowState.Normal, "Caption restore did not return to normal bounds.")
        DirectCast(minimize.Command, RoutedCommand).Execute(Nothing, window)
        Pump(window)
        Require(window.WindowState = WindowState.Minimized, "Caption minimize command failed.")
        SystemCommands.RestoreWindowCommand.Execute(Nothing, window)
        Pump(window)
        Require(window.WindowState = WindowState.Normal, "Restore command did not restore a minimized window.")
        window.ResizeMode = ResizeMode.NoResize
        Require(Not command.CanExecute(Nothing, window), "A fixed-size window still exposes a maximize command.")
        window.ResizeMode = ResizeMode.CanResize
        Dim closeRequested As Boolean
        Dim cancelClose As CancelEventHandler = Sub(sender, e)
                                                    closeRequested = True
                                                    e.Cancel = True
                                                End Sub
        AddHandler window.Closing, cancelClose
        Try
            DirectCast(close.Command, RoutedCommand).Execute(Nothing, window)
            Pump(window)
            Require(closeRequested AndAlso window.IsVisible, "Caption close bypassed the cancellable closing event.")
        Finally
            RemoveHandler window.Closing, cancelClose
        End Try
    End Sub

    Private Sub DetachStartup(window As MainWindow)
        Dim type = GetType(MainWindow)
        window.RemoveHandler(FrameworkElement.LoadedEvent, type.GetMethod("MainWindow_Loaded", PrivateInstance).CreateDelegate(GetType(RoutedEventHandler), window))
        RemoveHandler window.Closing, DirectCast(type.GetMethod("MainWindow_Closing", PrivateInstance).CreateDelegate(GetType(CancelEventHandler), window), CancelEventHandler)
        RemoveHandler Microsoft.Win32.SystemEvents.UserPreferenceChanged,
            DirectCast(type.GetMethod("SystemThemeChanged", PrivateInstance).CreateDelegate(GetType(Microsoft.Win32.UserPreferenceChangedEventHandler), window), Microsoft.Win32.UserPreferenceChangedEventHandler)
    End Sub

    Private Sub SavePreview(window As Window, Optional name As String = "aero-theme")
        Dim visual = DirectCast(window.Template.FindName("AeroFrame", window), FrameworkElement)
        Require(visual IsNot Nothing AndAlso visual.ActualWidth > 900, "Aero frame was not laid out for rendering.")
        Dim outputDirectory = Path.Combine(AppContext.BaseDirectory, "artifacts")
        Directory.CreateDirectory(outputDirectory)
        For Each dpi In {96, 144}
            Dim image As New RenderTargetBitmap(CInt(Math.Ceiling(visual.ActualWidth * dpi / 96)), CInt(Math.Ceiling(visual.ActualHeight * dpi / 96)), dpi, dpi, PixelFormats.Pbgra32)
            image.Render(visual)
            Dim encoder As New PngBitmapEncoder()
            encoder.Frames.Add(BitmapFrame.Create(image))
            Dim path = IO.Path.Combine(outputDirectory, $"{name}-{dpi}dpi.png")
            Using output = File.Create(path)
                encoder.Save(output)
            End Using
            Require(New FileInfo(path).Length > 5000, "Aero preview render is unexpectedly empty.")
            Console.WriteLine("Aero preview: " & path)
        Next
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
        window.UpdateLayout()
    End Sub

    Private Sub CheckContrast(foreground As Brush, background As Brush, name As String)
        For Each foregroundColor In BrushColors(foreground)
            For Each backgroundColor In BrushColors(background)
                Dim first = Luminance(foregroundColor), second = Luminance(backgroundColor)
                Dim ratio = (Math.Max(first, second) + 0.05) / (Math.Min(first, second) + 0.05)
                Require(ratio >= 4.5, $"{name} contrast is {ratio:F2}:1; expected at least 4.5:1 at every gradient stop.")
            Next
        Next
    End Sub

    Private Function BrushColors(brush As Brush) As IEnumerable(Of Color)
        Require(brush.Opacity = 1, "Interactive Aero surfaces must be opaque.")
        Dim solid = TryCast(brush, SolidColorBrush)
        If solid IsNot Nothing Then
            Require(solid.Color.A = 255, "Aero text/background is translucent.")
            Return {solid.Color}
        End If
        Dim gradient = TryCast(brush, GradientBrush)
        Require(gradient IsNot Nothing AndAlso gradient.GradientStops.All(Function(item) item.Color.A = 255), "Expected an opaque gradient.")
        Return gradient.GradientStops.Select(Function(item) item.Color)
    End Function

    Private Function Luminance(color As Color) As Double
        Return 0.2126 * Linear(color.R) + 0.7152 * Linear(color.G) + 0.0722 * Linear(color.B)
    End Function

    Private Function Linear(channel As Byte) As Double
        Dim value = channel / 255.0
        Return If(value <= 0.04045, value / 12.92, Math.Pow((value + 0.055) / 1.055, 2.4))
    End Function

    Private Sub Require(condition As Boolean, message As String)
        If Not condition Then Throw New InvalidOperationException(message)
    End Sub
End Module
