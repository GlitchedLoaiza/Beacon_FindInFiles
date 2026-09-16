Imports System.IO
Imports System.Collections
Imports System.Reflection
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Controls.Primitives
Imports System.Windows.Documents
Imports System.Windows.Media
Imports System.Windows.Media.Imaging
Imports System.Windows.Threading
Imports Beacon
Imports BeaconFilterControls

Module BeaconThemeChecks
    Public Sub Run()
        Require(CInt(AppTheme.System) = 0 AndAlso CInt(AppTheme.Light) = 1 AndAlso CInt(AppTheme.Dark) = 2 AndAlso CInt(AppTheme.Beacon) = 3,
                "Existing serialized theme values changed.")
        Require(BeaconSettings.CreateDefaults().Theme = AppTheme.System, "Adding Beacon Theme changed the default.")
        For Each theme In [Enum].GetValues(Of AppTheme)()
            Require(BeaconSettingsService.Validate(BeaconSettingsService.Clone(New BeaconSettings With {.Theme = theme})).Theme = theme,
                    "Theme did not survive validated settings serialization.")
            For Each systemDark In {False, True}
                Dim expected = theme = AppTheme.Dark OrElse ((theme = AppTheme.System OrElse theme = AppTheme.Beacon) AndAlso systemDark)
                Require(BeaconThemePalette.UsesDarkBackground(theme, systemDark) = expected, "Theme background resolution is incorrect.")
            Next
            Require(BeaconThemePalette.FollowsSystem(theme) = (theme = AppTheme.System OrElse theme = AppTheme.Beacon), "System-following modes are incorrect.")
        Next
        CheckLogoColor()
        CheckMainTransitions()
        For Each dark In {False, True}
            CheckSettingsAndChildren(dark)
        Next
    End Sub

    Private Sub CheckLogoColor()
        Using source = GetType(ScanReportWriter).Assembly.GetManifestResourceStream("Beacon.ReportLogo.png")
            Require(source IsNot Nothing, "Beacon logo resource is missing.")
            Dim decoder As New PngBitmapDecoder(source, BitmapCreateOptions.PreservePixelFormat, BitmapCacheOption.OnLoad)
            Dim image As New FormatConvertedBitmap(decoder.Frames(0), PixelFormats.Bgra32, Nothing, 0)
            Dim stride = image.PixelWidth * 4
            Dim bytes(stride * image.PixelHeight - 1) As Byte
            image.CopyPixels(bytes, stride, 0)
            Dim counts As New Dictionary(Of Color, Integer)()
            For index = 0 To bytes.Length - 4 Step 4
                Dim red = CInt(bytes(index + 2)), green = CInt(bytes(index + 1)), blue = CInt(bytes(index))
                If bytes(index + 3) >= 240 AndAlso red > green + 35 AndAlso red > blue + 35 Then
                    Dim pixel = Color.FromRgb(CByte(red), CByte(green), CByte(blue))
                    If Not counts.ContainsKey(pixel) Then counts(pixel) = 0
                    counts(pixel) += 1
                End If
            Next
            Require(counts.Count > 0 AndAlso counts.OrderByDescending(Function(pair) pair.Value).First().Key = BeaconThemePalette.LogoRed,
                    "Beacon Theme no longer matches the logo's dominant opaque red.")
        End Using
    End Sub

    Private Sub CheckMainTransitions()
        Dim window As New MainWindow()
        Dim type = GetType(MainWindow)
        Dim flags = BindingFlags.Instance Or BindingFlags.NonPublic
        window.RemoveHandler(FrameworkElement.LoadedEvent, type.GetMethod("MainWindow_Loaded", flags).CreateDelegate(GetType(RoutedEventHandler), window))
        RemoveHandler window.Closing, DirectCast(type.GetMethod("MainWindow_Closing", flags).CreateDelegate(GetType(ComponentModel.CancelEventHandler), window), ComponentModel.CancelEventHandler)
        Dim settings As New BeaconSettings With {.Theme = AppTheme.Beacon}
        type.GetField("_settings", flags).SetValue(window, settings)
        type.GetField("_searchOptions", flags).SetValue(window, BeaconSettingsService.Clone(settings))
        PrepareWindow(window)
        Try
            window.Show()
            Dim title = window.Title
            Dim icon = window.Icon
            Dim hitType = type.GetNestedType("SearchHit", BindingFlags.NonPublic)
            Dim hit = Activator.CreateInstance(hitType, True)
            hitType.GetProperty("DisplayName").SetValue(hit, "sample.log")
            hitType.GetProperty("LogicalPath").SetValue(hit, Path.Combine(Path.GetTempPath(), "sample.log"))
            DirectCast(hitType.GetProperty("Details").GetValue(hit), List(Of SearchDetail)).Add(
                New SearchDetail With {.IsMetadata = True, .Location = "File name", .Excerpt = "sample.log", .VisibleMatches = 1})
            DirectCast(type.GetField("_hits", flags).GetValue(window), IList).Add(hit)
            Dim results = DirectCast(window.FindName("Results_lst"), ListBox)
            results.SelectedItem = hit
            Pump(window)
            Dim row = DirectCast(results.ItemContainerGenerator.ContainerFromIndex(0), ListBoxItem)
            For Each theme In {AppTheme.Beacon, AppTheme.Dark, AppTheme.Beacon, AppTheme.Light, AppTheme.System, AppTheme.Beacon}
                settings.Theme = theme
                For Each systemDark In {False, True}
                    type.GetMethod("ApplyThemeSelection", flags).Invoke(window, {systemDark})
                    Pump(window)
                    Dim dark = BeaconThemePalette.UsesDarkBackground(theme, systemDark)
                    Require(CBool(type.GetField("_isDarkMode", flags).GetValue(window)) = dark, "Main window ignored the resolved system background.")
                    Require(BrushColor(window.Resources("WindowBackgroundBrush")) = If(dark, Color.FromRgb(&H20, &H20, &H20), Color.FromRgb(&HF3, &HF3, &HF3)),
                            "Main background did not follow the selected theme.")
                    Require(BrushColor(window.Resources("InputBackgroundBrush")) = If(dark, Color.FromRgb(&H2B, &H2B, &H2B), Colors.White), "Input background was recolored.")
                    Dim neutral = If(dark, Color.FromRgb(&H3A, &H3A, &H3A), Color.FromRgb(&HF0, &HF0, &HF0))
                    Require(BrushColor(window.Resources("ButtonBackgroundBrush")) = neutral, "Neutral shared surface brush was replaced by red.")
                    For Each name In {"EventCounterBar", "HarCounterBar"}
                        Require(BrushColor(DirectCast(window.FindName(name), Border).Background) = neutral, "Counter strip became an action-button color.")
                    Next
                    Require(BeaconThemePalette.IsEnabled(window.Resources) = (theme = AppTheme.Beacon), "Beacon palette was not reset after switching themes.")
                    CheckMainAccents(window, row, theme = AppTheme.Beacon, dark)
                    For Each name In {"Help_btn", "Settings_btn", "Scan_btn", "ExportResults_btn", "Reset_btn"}
                        Dim button = DirectCast(window.FindName(name), Button)
                        If theme = AppTheme.Beacon Then
                            CheckBeaconButton(button, window)
                        Else
                            Dim enabled = button.IsEnabled
                            button.IsEnabled = True
                            button.ApplyTemplate()
                            Require(BrushColor(button.Background) = neutral AndAlso BrushColor(button.Foreground) = BrushColor(window.Resources("TextPrimaryBrush")),
                                    "Standard theme button colors did not return.")
                            button.IsEnabled = enabled
                        End If
                    Next
                    NativeCaptionChecks.Verify(window, dark)
                Next
            Next
            settings.Theme = AppTheme.Beacon
            Dim actualSystemDark = CBool(type.GetMethod("IsWindowsDarkModeEnabled", flags).Invoke(window, Nothing))
            type.GetMethod("ApplyThemeSelection", flags).Invoke(window, {Not actualSystemDark})
            type.GetMethod("SystemThemeChanged", flags).Invoke(window, {Nothing, New Microsoft.Win32.UserPreferenceChangedEventArgs(Microsoft.Win32.UserPreferenceCategory.General)})
            Pump(window)
            Require(CBool(type.GetField("_isDarkMode", flags).GetValue(window)) = actualSystemDark AndAlso BeaconThemePalette.IsEnabled(window.Resources),
                    "Windows preference notification did not refresh Beacon Theme.")
            Require(window.Title = title AndAlso window.Icon Is icon AndAlso window.WindowStyle <> WindowStyle.None, "Beacon Theme altered the native frame, title or logo.")
        Finally
            window.Close()
        End Try
    End Sub

    Private Sub CheckSettingsAndChildren(dark As Boolean)
        Dim settings As New BeaconSettings With {.Theme = AppTheme.Beacon}
        Dim window As New SettingsWindow(settings, dark)
        PrepareWindow(window)
        Dim help As HelpWindow = Nothing
        Dim diagnostics As ScanReportWindow = Nothing
        Dim tour As TourWindow = Nothing
        Try
            window.Show()
            Pump(window)
            Dim selector = DirectCast(window.FindName("Theme_cmb"), ComboBox)
            Require(CStr(selector.SelectedValue) = "Beacon" AndAlso CStr(DirectCast(selector.SelectedItem, ComboBoxItem).Content) = "Beacon Theme", "Settings did not restore Beacon Theme.")
            Require(BrushColor(window.Resources("CardBackgroundBrush")) = If(dark, Color.FromRgb(&H2B, &H2B, &H2B), Colors.White), "Settings changed the system-dependent card background.")
            CheckSelectionContrast(window.Resources)
            For Each name In {"Save_btn", "Cancel_btn", "RestoreDefaults_btn", "CheckUpdates_btn"}
                CheckBeaconButton(DirectCast(window.FindName(name), Button), window)
            Next
            Dim tabs = DirectCast(window.FindName("SettingsTabs"), TabControl)
            tabs.SelectedIndex = 3
            Pump(window)
            Dim picker = DirectCast(DirectCast(window.FindName("EvtxFrom_host"), ContentControl).Content, UtcDateTimePicker)
            picker.Text = "2026-01-02 03:04:05"
            CheckBeaconButton(DirectCast(picker.FindName("OpenPicker"), Button), window)
            picker.OpenEditor()
            Pump(window)
            Dim popup = DirectCast(picker.FindName("PickerPopup"), Popup)
            Try
                Require(popup.IsOpen, "Beacon date picker did not open.")
                Require(BrushColor(DirectCast(popup.Child, Border).Background) = BrushColor(window.Resources("CardBackgroundBrush")), "Picker popup background is not neutral.")
                For Each button In Descendants(popup.Child).OfType(Of Button)().ToArray()
                    CheckBeaconButton(button, window)
                Next
            Finally
                popup.IsOpen = False
            End Try
            Dim levels = DirectCast(DirectCast(window.FindName("EvtxLevels_host"), ContentControl).Content, SeveritySelector)
            CheckBeaconButton(DirectCast(levels.FindName("OpenLevels"), Button), window)
            Dim provider = DirectCast(window.FindName("EvtxProvider_cmb"), ComboBox)
            WithState(provider, GetType(UIElement), "IsKeyboardFocusWithinPropertyKey", True,
                Sub() Require(BrushColor(provider.BorderBrush) = BrushColor(window.Resources("AccentBrush")), "Settings dropdown kept a blue focus border."))
            Dim input = DirectCast(picker.FindName("DateText"), TextBox)
            Require(BrushColor(input.SelectionBrush) = BrushColor(window.Resources("SelectionBackgroundBrush")) AndAlso input.SelectionOpacity = 1,
                    "Picker text selection did not inherit the readable red selection palette.")

            help = New HelpWindow(window, dark)
            PrepareWindow(help)
            help.Show()
            Pump(help)
            CheckBeaconButton(DirectCast(help.FindName("ClearHelpSearch_btn"), Button), help)
            CheckBeaconButton(DirectCast(help.FindName("CloseHelp_btn"), Button), help)
            diagnostics = New ScanReportWindow(window, ReportChecks.Sample(Path.GetTempPath()), Nothing, False, dark)
            PrepareWindow(diagnostics)
            diagnostics.Show()
            Pump(diagnostics)
            CheckBeaconButton(DirectCast(diagnostics.FindName("CopyIssues_btn"), Button), diagnostics)
            CheckBeaconButton(DirectCast(diagnostics.FindName("Close_btn"), Button), diagnostics)
            tour = New TourWindow(window, dark)
            tour.Show()
            Pump(tour)
            For Each name In {"ExitTour_btn", "BackTour_btn", "NextTour_btn"}
                CheckBeaconButton(DirectCast(tour.FindName(name), Button), tour)
            Next
            Require(tour.WindowStyle = WindowStyle.None, "Beacon Theme changed the borderless tour.")
            For Each child As Window In New Window() {help, diagnostics, tour}
                Require(BeaconThemePalette.IsEnabled(child.Resources) AndAlso BrushColor(child.Resources("CardBackgroundBrush")) = BrushColor(window.Resources("CardBackgroundBrush")),
                        "Owned window did not inherit Beacon button/background colors.")
                Require(BrushColor(child.Resources("SelectionBackgroundBrush")) = BrushColor(window.Resources("SelectionBackgroundBrush")) AndAlso
                        BrushColor(child.Resources("AccentBrush")) = BrushColor(window.Resources("AccentBrush")), "Owned window replaced red selection/focus accents with blue.")
                CheckSelectionContrast(child.Resources)
            Next
            DirectCast(window.FindName("RestoreDefaults_btn"), Button).RaiseEvent(New RoutedEventArgs(Button.ClickEvent))
            Require(CStr(selector.SelectedValue) = "System" AndAlso settings.Theme = AppTheme.Beacon AndAlso window.SavedSettings Is Nothing,
                    "Restore defaults altered saved settings or failed to restore the System selection.")
        Finally
            tour?.Close()
            diagnostics?.Close()
            help?.Close()
            window.Close()
        End Try
    End Sub

    Private Sub CheckBeaconButton(button As Button, window As Window)
        Dim enabled = button.IsEnabled
        Try
            button.IsEnabled = True
            button.ApplyTemplate()
            Dim border = TryCast(button.Template.FindName("ButtonBorder", button), Border)
            If border Is Nothing Then border = DirectCast(button.Template.FindName("border", button), Border)
            Require(BrushColor(border.Background) = Color.FromRgb(&HB8, &H32, &H43) AndAlso BrushColor(button.Foreground) = Color.FromRgb(&HFF, &HF5, &HF5),
                    button.Name & ": normal button color is not muted crimson with off-white text.")
            CheckContrast(button.Foreground, border.Background, button.Name & " normal")
            For Each label As TextBlock In Descendants(button).OfType(Of TextBlock)()
                CheckContrast(label.Foreground, border.Background, button.Name & " rendered label")
            Next
            For Each key In {"ActionButtonHoverBrush", "ActionButtonPressedBrush", "PrimaryButtonHoverBrush", "PrimaryButtonPressedBrush"}
                CheckContrast(button.Foreground, DirectCast(window.Resources(key), Brush), button.Name & " " & key)
            Next
            WithState(button, GetType(UIElement), "IsMouseOverPropertyKey", True,
                Sub()
                    Require(BrushColor(border.Background) = Color.FromRgb(&HC4, &H3D, &H4D), button.Name & ": actual hover template did not use the softer palette.")
                    CheckContrast(button.Foreground, border.Background, button.Name & " actual hover")
                End Sub)
            WithState(button, GetType(ButtonBase), "IsPressedPropertyKey", True,
                Sub()
                    Require(BrushColor(border.Background) = Color.FromRgb(&H9E, &H29, &H39), button.Name & ": actual pressed template did not use the softer palette.")
                    CheckContrast(button.Foreground, border.Background, button.Name & " actual pressed")
                End Sub)
            WithState(button, GetType(UIElement), "IsKeyboardFocusedPropertyKey", True,
                Sub() Require(BrushColor(border.BorderBrush) = BeaconThemePalette.ButtonText, button.Name & ": keyboard focus cue is missing."))
            button.IsEnabled = False
            Require(button.Opacity = 1 AndAlso BrushColor(border.Background) <> BeaconThemePalette.ButtonCrimson, button.Name & ": disabled state is not distinct/readable.")
            CheckContrast(button.Foreground, border.Background, button.Name & " disabled")
        Finally
            button.IsEnabled = enabled
        End Try
    End Sub

    Private Sub CheckMainAccents(window As MainWindow, row As ListBoxItem, beacon As Boolean, dark As Boolean)
        Dim accent = If(beacon, If(dark, Color.FromRgb(&HE5, &H8E, &H9A), Color.FromRgb(&HB8, &H32, &H43)),
                        If(dark, Color.FromRgb(&H60, &HCF, &HFF), Color.FromRgb(0, &H66, &HCC)))
        Dim selection = If(beacon, If(dark, Color.FromRgb(&H48, &H2B, &H31), Color.FromRgb(&HF7, &HE6, &HE9)),
                           If(dark, Color.FromRgb(&H18, &H3C, &H50), Color.FromRgb(&HE5, &HF1, &HFF)))
        Require(BrushColor(window.Resources("AccentBrush")) = accent AndAlso BrushColor(window.Resources("SelectionBackgroundBrush")) = selection,
                "Accent/selection colors did not follow or reset with the theme.")
        row.ApplyTemplate()
        Dim selectedBorder = DirectCast(row.Template.FindName("ResultBorder", row), Border)
        Require(BrushColor(selectedBorder.Background) = selection, "Selected result row did not use the selected-theme color.")
        For Each label As TextBlock In Descendants(row).OfType(Of TextBlock)()
            CheckContrast(label.Foreground, selectedBorder.Background, "Selected filename/count text")
        Next
        Dim search = DirectCast(window.FindName("Search_txt"), TextBox)
        search.ApplyTemplate()
        Dim inputBorder = DirectCast(search.Template.FindName("InputBorder", search), Border)
        WithState(search, GetType(UIElement), "IsKeyboardFocusWithinPropertyKey", True,
            Sub() Require(BrushColor(inputBorder.BorderBrush) = accent, "Search field focus outline does not follow the accent."))
        Dim preview = DirectCast(window.FindName("TextPreview_rtb"), RichTextBox)
        preview.ApplyTemplate()
        If Not beacon Then
            Require(preview.ReadLocalValue(FrameworkElement.StyleProperty) Is DependencyProperty.UnsetValue AndAlso
                    preview.Style Is window.FindResource(GetType(RichTextBox)), "Standard preview style was not restored.")
            Require(BrushColor(search.SelectionBrush) = accent AndAlso search.SelectionOpacity = CDbl(TextBoxBase.SelectionOpacityProperty.GetMetadata(GetType(TextBox)).DefaultValue),
                    "Standard input selection defaults were not restored.")
            Return
        End If
        CheckSelectionContrast(window.Resources)
        Dim previewBorder = DirectCast(preview.Template.FindName("PreviewBorder", preview), Border)
        Dim host = DirectCast(preview.Template.FindName("PART_ContentHost", preview), ScrollViewer)
        Require(previewBorder.CornerRadius.TopLeft = 4 AndAlso BrushColor(previewBorder.BorderBrush) = BrushColor(window.Resources("InputBorderBrush")),
                "Beacon preview does not use its rounded red-toned outline.")
        WithState(preview, GetType(UIElement), "IsKeyboardFocusWithinPropertyKey", True,
            Sub() Require(BrushColor(previewBorder.BorderBrush) = accent, "Focused preview retained a blue outline."))
        WithState(preview, GetType(UIElement), "IsMouseOverPropertyKey", True,
            Sub() Require(BrushColor(previewBorder.BorderBrush) = accent, "Preview hover outline does not follow the accent."))
        Require(BrushColor(preview.SelectionBrush) = selection AndAlso BrushColor(preview.SelectionTextBrush) = BrushColor(window.Resources("TextPrimaryBrush")) AndAlso
                preview.SelectionOpacity = 1 AndAlso BrushColor(search.SelectionBrush) = selection, "Text selection retained an incompatible accent/foreground.")
        preview.Document = New FlowDocument(New Paragraph(New Run(String.Join(vbLf, Enumerable.Repeat("Readable sample text", 150)))))
        Pump(window)
        preview.SelectAll()
        Require(preview.Selection.Text.Contains("Readable sample text") AndAlso host.ScrollableHeight > 0 AndAlso preview.IsReadOnly,
                "Beacon preview template broke selection, read-only behavior or scrolling.")
        preview.ScrollToEnd()
        Pump(window)
        Require(preview.VerticalOffset > 0, "Beacon preview cannot scroll to its content.")
        preview.ScrollToHome()
        preview.Selection.Select(preview.Document.ContentStart, preview.Document.ContentStart)
    End Sub

    Private Sub CheckSelectionContrast(resources As ResourceDictionary)
        CheckContrast(DirectCast(resources("TextPrimaryBrush"), Brush), DirectCast(resources("SelectionBackgroundBrush"), Brush), "Selected primary text")
        CheckContrast(DirectCast(resources("TextSecondaryBrush"), Brush), DirectCast(resources("SelectionBackgroundBrush"), Brush), "Selected secondary text")
        CheckContrast(DirectCast(resources("AccentBrush"), Brush), DirectCast(resources("CardBackgroundBrush"), Brush), "Focus/link accent")
    End Sub

    Private Sub WithState(element As DependencyObject, declaringType As Type, fieldName As String, value As Boolean, check As Action)
        ' Exercise readonly WPF interaction-state triggers without moving the user's pointer or taking foreground focus.
        Dim field = declaringType.GetField(fieldName, BindingFlags.Static Or BindingFlags.Public Or BindingFlags.NonPublic)
        Require(field IsNot Nothing, "WPF test state key is unavailable: " & fieldName)
        Dim key = DirectCast(field.GetValue(Nothing), DependencyPropertyKey)
        Dim original = element.ReadLocalValue(key.DependencyProperty)
        Try
            element.SetValue(key, value)
            check()
        Finally
            If original Is DependencyProperty.UnsetValue Then
                element.ClearValue(key)
            Else
                element.SetValue(key, original)
            End If
        End Try
    End Sub

    Private Sub PrepareWindow(window As Window)
        window.WindowStartupLocation = WindowStartupLocation.Manual
        window.WindowState = WindowState.Normal
        window.Left = -10000
        window.Top = -10000
        window.ShowActivated = False
        window.ShowInTaskbar = False
    End Sub

    Private Sub Pump(window As Window)
        window.UpdateLayout()
        window.Dispatcher.Invoke(Sub()
                                 End Sub, DispatcherPriority.ContextIdle)
        window.UpdateLayout()
    End Sub

    Private Iterator Function Descendants(parent As DependencyObject) As IEnumerable(Of DependencyObject)
        For index = 0 To VisualTreeHelper.GetChildrenCount(parent) - 1
            Dim element = VisualTreeHelper.GetChild(parent, index)
            Yield element
            For Each nested In Descendants(element)
                Yield nested
            Next
        Next
    End Function

    Private Function BrushColor(value As Object) As Color
        Dim brush = TryCast(value, SolidColorBrush)
        Require(brush IsNot Nothing AndAlso brush.Color.A = 255 AndAlso brush.Opacity = 1, "Expected an opaque theme brush.")
        Return brush.Color
    End Function

    Private Sub CheckContrast(foreground As Brush, background As Brush, name As String)
        Dim first = Luminance(BrushColor(foreground)), second = Luminance(BrushColor(background))
        Dim ratio = (Math.Max(first, second) + 0.05) / (Math.Min(first, second) + 0.05)
        Require(ratio >= 4.5, $"{name} contrast is {ratio:F2}:1; expected at least 4.5:1.")
    End Sub

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
