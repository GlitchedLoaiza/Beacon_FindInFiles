Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Controls.Primitives
Imports System.Windows.Documents
Imports System.Windows.Media
Imports System.Windows.Threading
Imports Beacon

Module SettingsThemeChecks
    Public Sub RunReport(dark As Boolean)
        Dim owner As New SettingsWindow(New BeaconSettings(), dark)
        Dim snapshot = ReportChecks.Sample(System.IO.Path.GetTempPath())
        Dim refreshed As New ScanReportSnapshot With {.Run = New ScanRunInfo With {.State = ScanReportState.Cancelled}}
        Dim window As New ScanReportWindow(owner, snapshot, Function() refreshed, False, dark) With {
            .WindowStartupLocation = WindowStartupLocation.Manual, .Left = -10000, .Top = -10000,
            .ShowActivated = False, .ShowInTaskbar = False
        }
        Try
            window.Show()
            Pump(window)
            Require(window.FindName("Export_btn") Is Nothing AndAlso window.FindName("IncludeExcerpts_chk") Is Nothing AndAlso window.FindName("IncludeTechnical_chk") Is Nothing,
                    "Diagnostics still contains redundant export/settings controls.")
            Require(window.Title.Contains("diagnostics", StringComparison.OrdinalIgnoreCase), "Diagnostics purpose is unclear.")
            NativeCaptionChecks.Verify(window, dark)
            Dim list = DirectCast(window.FindName("Diagnostics_lst"), ListBox)
            Require(list.Items.Count = 1, "Report diagnostics were not displayed.")
            Dim row = DirectCast(list.ItemContainerGenerator.ContainerFromIndex(0), ListBoxItem)
            Dim technical = Descendants(row).OfType(Of TextBlock)().Single(Function(item) item.Name = "TechnicalDetailsText")
            Require(technical.Visibility = Visibility.Collapsed, "Technical details appeared without opt-in.")
            window.Tag = True
            Pump(window)
            Require(technical.Visibility = Visibility.Visible, "Technical details toggle did not update the report.")
            For Each label As TextBlock In Descendants(row).OfType(Of TextBlock)()
                CheckContrast(label.Foreground, list.Background, "Report issue text")
            Next
            For Each name In {"Filter_txt"}
                Dim editor = DirectCast(window.FindName(name), TextBox)
                CheckTextFits(window, editor)
            Next
            Dim filter = DirectCast(window.FindName("Filter_txt"), TextBox)
            filter.Text = "not-present"
            Pump(window)
            Require(list.Items.Count = 0 AndAlso DirectCast(window.FindName("EmptyDiagnostics_txt"), TextBlock).IsVisible, "Diagnostic filtering failed.")
            filter.Text = ""
            Dim severity = DirectCast(window.FindName("Severity_cmb"), ComboBox)
            severity.SelectedItem = "Error"
            Pump(window)
            Require(list.Items.Count = 0, "Severity filtering failed.")
            severity.SelectedItem = "All"
            Pump(window)
            Require(list.Items.Count = 1, "Clearing filters did not restore the snapshot.")
            window.Width = window.MinWidth
            window.Height = window.MinHeight
            Pump(window)
            Require(list.ActualHeight > 50, "Report controls hide the diagnostics at minimum size.")
            For Each name In {"Refresh_btn", "CopyIssues_btn", "Close_btn"}
                Dim button = DirectCast(window.FindName(name), Button)
                Dim border = DirectCast(button.Template.FindName("ButtonBorder", button), Border)
                CheckContrast(button.Foreground, border.Background, "Report button")
                Require(border.CornerRadius.TopLeft = 4, "Report button does not match Beacon styling.")
            Next
            DirectCast(window.FindName("Refresh_btn"), Button).RaiseEvent(New RoutedEventArgs(Button.ClickEvent))
            Pump(window)
            Require(list.Items.Count = 0 AndAlso DirectCast(window.FindName("Summary_txt"), TextBlock).Text.Contains("Cancelled"), "Diagnostics refresh did not capture new state.")
        Finally
            window.Close()
            owner.Close()
        End Try
    End Sub

    Public Sub Run(dark As Boolean)
        If Application.Current Is Nothing Then
            Dim application As New Application With {.ShutdownMode = ShutdownMode.OnExplicitShutdown}
        End If
        Dim settings As New BeaconSettings()
        Dim window As New SettingsWindow(settings, dark) With {
            .WindowStartupLocation = WindowStartupLocation.Manual,
            .Left = -10000, .Top = -10000, .ShowActivated = False, .ShowInTaskbar = False
        }
        Try
            window.Show()
            Pump(window)
            CheckFontSelector(window, dark)
            NativeCaptionChecks.Verify(window, dark)
            Dim theme = DirectCast(window.FindName("Theme_cmb"), ComboBox)
            Require(theme.Items.Cast(Of ComboBoxItem)().Select(Function(item) CStr(item.Content)).SequenceEqual({"Light", "Dark", "System theme", "Beacon Theme"}), "Theme choices are incorrect.")
            Require(CStr(theme.SelectedValue) = "System", "Theme should default to Windows preferences.")
            For Each choice In [Enum].GetValues(Of AppTheme)()
                Dim saved = BeaconSettingsService.Clone(New BeaconSettings With {.Theme = choice})
                Require(saved.Theme = choice, "Theme did not survive settings serialization.")
                Dim themedWindow As New SettingsWindow(saved, dark)
                Try
                    Require(CStr(DirectCast(themedWindow.FindName("Theme_cmb"), ComboBox).SelectedValue) = choice.ToString(), "Settings did not load the saved theme.")
                    DirectCast(themedWindow.FindName("RestoreDefaults_btn"), Button).RaiseEvent(New RoutedEventArgs(Button.ClickEvent))
                    Require(CStr(DirectCast(themedWindow.FindName("Theme_cmb"), ComboBox).SelectedValue) = "System", "Restore defaults did not reset the theme.")
                Finally
                    themedWindow.Close()
                End Try
            Next
            Require(BeaconSettingsService.Validate(New BeaconSettings With {.Theme = CType(999, AppTheme)}).Theme = AppTheme.System, "Invalid theme did not fall back to System.")
            Dim card = DirectCast(window.Resources("CardBackgroundBrush"), SolidColorBrush)
            Dim primary = DirectCast(window.Resources("TextPrimaryBrush"), SolidColorBrush)
            Require(card.Color = If(dark, Color.FromRgb(&H2B, &H2B, &H2B), Colors.White), "Palette differs from Beacon.")
            Dim tabs = DirectCast(window.FindName("SettingsTabs"), TabControl)
            For index = 0 To tabs.Items.Count - 1
                tabs.SelectedIndex = index
                Pump(window)
                Dim tab = DirectCast(tabs.Items(index), TabItem)
                Require(tab.Template.FindName("SelectedLine", tab) IsNot Nothing, "Selected tab indicator is missing.")
                CheckContrast(tab.Foreground, DirectCast(window.Resources("SelectionBackgroundBrush"), Brush), "Selected tab")
                Require(tabs.Template.FindName("PART_SelectedContentHost", tabs) IsNot Nothing, "Tab content host is missing.")
                Dim headings = Descendants(window).OfType(Of TextBlock)().Where(Function(text) text.Style Is window.FindResource("SectionHeader")).ToArray()
                Require(headings.Length > 0, "Section heading not rendered.")
                For Each heading In headings
                    CheckContrast(heading.Foreground, card, heading.Text)
                Next
                For Each editor As TextBox In Descendants(window).OfType(Of TextBox)().ToArray()
                    If Not editor.IsVisible Then Continue For
                    editor.ApplyTemplate()
                    CheckContrast(editor.Foreground, editor.Background, editor.Name)
                    Require(editor.Template.FindName("PART_ContentHost", editor) IsNot Nothing, "TextBox content host is missing.")
                    If editor.Name = "PART_EditableTextBox" Then
                        Require(editor.Text.Length = 0 OrElse TextFitsViewport(editor, DirectCast(editor.Template.FindName("PART_ContentHost", editor), ScrollViewer)),
                                "Editable provider text is clipped.")
                    Else
                        CheckTextFits(window, editor)
                    End If
                Next
                For Each check In Descendants(window).OfType(Of CheckBox)().ToArray()
                    CheckContrast(check.Foreground, card, check.Name)
                    Dim original = check.IsChecked
                    check.IsChecked = True
                    Pump(window)
                    Dim mark = DirectCast(check.Template.FindName("CheckMark", check), System.Windows.Shapes.Path)
                    Dim box = DirectCast(check.Template.FindName("CheckBorder", check), Border)
                    Require(mark.Visibility = Visibility.Visible, "Checked state is not visible.")
                    CheckContrast(mark.Stroke, box.Background, "Check indicator")
                    check.IsChecked = original
                Next
                For Each combo In Descendants(window).OfType(Of ComboBox)().ToArray()
                    CheckContrast(combo.Foreground, combo.Background, combo.Name)
                    Require(combo.FocusVisualStyle IsNot Nothing, "ComboBox focus cue is missing.")
                    Dim original = combo.SelectedIndex
                    combo.IsDropDownOpen = True
                    Pump(window)
                    Dim popup = DirectCast(combo.Template.FindName("PART_Popup", combo), Popup)
                    Require(popup IsNot Nothing AndAlso popup.IsOpen, "Dropdown failed to open.")
                    Dim popupBorder = DirectCast(popup.Child, Border)
                    CheckContrast(primary, popupBorder.Background, "Dropdown surface")
                    Require(popupBorder.ActualWidth >= combo.ActualWidth - 1, "Popup width does not follow the control.")
                    For itemIndex = 0 To combo.Items.Count - 1
                        combo.SelectedIndex = itemIndex
                        Pump(window)
                        Dim item = TryCast(combo.ItemContainerGenerator.ContainerFromIndex(itemIndex), ComboBoxItem)
                        Require(item IsNot Nothing, "Dropdown option was not generated.")
                        item.ApplyTemplate()
                        Dim itemBorder = DirectCast(item.Template.FindName("ItemBorder", item), Border)
                        CheckContrast(item.Foreground, itemBorder.Background, "Selected dropdown option")
                        CheckContrast(item.Foreground, popupBorder.Background, "Unselected dropdown option")
                    Next
                    combo.SelectedIndex = original
                    combo.IsDropDownOpen = False
                    combo.IsEnabled = False
                    Pump(window)
                    Require(combo.Opacity > 0 AndAlso combo.Opacity < 1, "Disabled state is not styled.")
                    combo.IsEnabled = True
                Next
            Next
            For Each name In {"Save_btn", "Cancel_btn", "RestoreDefaults_btn"}
                Dim button = DirectCast(window.FindName(name), Button)
                Dim border = DirectCast(button.Template.FindName("ButtonBorder", button), Border)
                Require(border.CornerRadius.TopLeft = 4, "Button corners differ from Beacon.")
                CheckContrast(button.Foreground, border.Background, name)
                For Each label As TextBlock In Descendants(button).OfType(Of TextBlock)()
                    CheckContrast(label.Foreground, border.Background, name & " rendered label")
                Next
                Dim hoverKey = If(name = "Save_btn", "PrimaryButtonHoverBrush", "ButtonHoverBrush")
                Dim pressedKey = If(name = "Save_btn", "PrimaryButtonPressedBrush", "ButtonPressedBrush")
                CheckContrast(button.Foreground, DirectCast(window.Resources(hoverKey), Brush), name & " hover")
                CheckContrast(button.Foreground, DirectCast(window.Resources(pressedKey), Brush), name & " pressed")
                Require(button.FocusVisualStyle IsNot Nothing, "Button focus cue is missing.")
            Next
            window.Width = window.MinWidth
            window.Height = window.MinHeight
            Pump(window)
            Require(DirectCast(tabs.Template.FindName("PART_SelectedContentHost", tabs), ContentPresenter).ActualHeight > 100,
                    "Minimum window size hides the settings content.")
            Require(settings.DefaultSearchMode = SearchMode.PlainText AndAlso window.SavedSettings Is Nothing,
                    "Theme checks changed persisted settings.")
        Finally
            window.Close()
        End Try
    End Sub

    Public Sub RunMainFields(dark As Boolean)
        Dim document = System.Xml.Linq.XDocument.Load(System.IO.Path.Combine(AppContext.BaseDirectory, "Fixtures", "MainWindow.xaml"))
        Dim presentation = System.Xml.Linq.XNamespace.Get("http://schemas.microsoft.com/winfx/2006/xaml/presentation")
        Dim xaml = System.Xml.Linq.XNamespace.Get("http://schemas.microsoft.com/winfx/2006/xaml")
        Dim element = New System.Xml.Linq.XElement(document.Root.Element(presentation + "Window.Resources").Elements(presentation + "Style").
            Single(Function(item) CStr(item.Attribute("TargetType")) = "TextBox"))
        element.SetAttributeValue("xmlns", presentation.NamespaceName)
        element.SetAttributeValue(System.Xml.Linq.XNamespace.Xmlns + "x", xaml.NamespaceName)
        Dim style = DirectCast(System.Windows.Markup.XamlReader.Parse(element.ToString()), Style)
        Dim window As New SettingsWindow(New BeaconSettings(), dark) With {
            .WindowStartupLocation = WindowStartupLocation.Manual,
            .Left = -10000, .Top = -10000, .ShowActivated = False, .ShowInTaskbar = False
        }
        Dim panel As New StackPanel With {.Margin = New Thickness(16)}
        window.Content = panel
        For Each name In {"Path_txt", "Search_txt"}
            Dim definition = document.Descendants(presentation + "TextBox").Single(Function(item) CStr(item.Attribute(xaml + "Name")) = name)
            Dim editor As New TextBox With {
                .Name = name, .Style = style,
                .IsEnabled = Not String.Equals(CStr(definition.Attribute("IsEnabled")), "False", StringComparison.OrdinalIgnoreCase),
                .Text = If(name = "Path_txt", "C:\Logs\sample", "error")
            }
            panel.Children.Add(editor)
        Next
        Try
            window.Show()
            Pump(window)
            For Each editor As TextBox In panel.Children
                Dim border = DirectCast(editor.Template.FindName("InputBorder", editor), Border)
                Require(border.CornerRadius.TopLeft = 4, "Main input is not rounded.")
                CheckContrast(editor.Foreground, border.Background, editor.Name)
                CheckTextFits(window, editor)
            Next
        Finally
            window.Close()
        End Try
    End Sub

    Public Sub RunToolbar(dark As Boolean)
        Dim document = System.Xml.Linq.XDocument.Load(System.IO.Path.Combine(AppContext.BaseDirectory, "Fixtures", "MainWindow.xaml"))
        Dim presentation = System.Xml.Linq.XNamespace.Get("http://schemas.microsoft.com/winfx/2006/xaml/presentation")
        Dim xaml = System.Xml.Linq.XNamespace.Get("http://schemas.microsoft.com/winfx/2006/xaml")
        Dim context As New System.Windows.Markup.ParserContext With {.BaseUri = New Uri("pack://application:,,,/Beacon.SafetyChecks;component/MainWindow.xaml")}
        Dim window As New SettingsWindow(New BeaconSettings(), dark) With {
            .WindowStartupLocation = WindowStartupLocation.Manual,
            .Left = -10000, .Top = -10000, .ShowActivated = False, .ShowInTaskbar = False
        }
        For Each controlType In {GetType(Button), GetType(TextBox), GetType(CheckBox)}
            Dim element = New System.Xml.Linq.XElement(document.Root.Element(presentation + "Window.Resources").Elements(presentation + "Style").
                Single(Function(item) CStr(item.Attribute("TargetType")) = controlType.Name))
            element.SetAttributeValue("xmlns", presentation.NamespaceName)
            element.SetAttributeValue(System.Xml.Linq.XNamespace.Xmlns + "x", xaml.NamespaceName)
            window.Resources(controlType) = System.Windows.Markup.XamlReader.Parse(element.ToString(), context)
        Next
        Dim markup = New System.Xml.Linq.XElement(document.Descendants(presentation + "Grid").Single(Function(item) CStr(item.Attribute(xaml + "Name")) = "MainToolbar"))
        markup.SetAttributeValue("xmlns", presentation.NamespaceName)
        markup.SetAttributeValue(System.Xml.Linq.XNamespace.Xmlns + "x", xaml.NamespaceName)
        Dim toolbar = DirectCast(System.Windows.Markup.XamlReader.Parse(markup.ToString(), context), Grid)
        window.Content = toolbar
        Try
            window.Show()
            Pump(window)
            Require(toolbar.FindName("SearchModeHelp_txt") Is Nothing, "Passive search summary is still displayed.")
            Require(toolbar.FindName("ExactMatch_chk") Is Nothing, "Duplicate whole-word checkbox is still displayed.")
            Require(toolbar.FindName("ThemeToggle_btn") Is Nothing, "Theme toggle is still on the toolbar.")
            Dim settingsButton = DirectCast(toolbar.FindName("Settings_btn"), Button)
            Dim helpButton = DirectCast(toolbar.FindName("Help_btn"), Button)
            Require(helpButton IsNot Nothing AndAlso System.Windows.Automation.AutomationProperties.GetName(helpButton) = "Help", "Accessible Help action is missing.")
            Dim icon = Descendants(settingsButton).OfType(Of TextBlock)().Single(Function(item) item.FontFamily.Source = "Segoe MDL2 Assets")
            Require(icon.Text = ChrW(&HE90F), "Settings wrench glyph is missing.")
            Dim glyphs As GlyphTypeface = Nothing
            Require(New Typeface(icon.FontFamily, FontStyles.Normal, FontWeights.Normal, FontStretches.Normal).TryGetGlyphTypeface(glyphs) AndAlso
                    glyphs.CharacterToGlyphMap.ContainsKey(&HE90F), "Installed icon font does not contain the wrench.")
            For Each width In {1000.0, 1400.0, 2200.0}
                window.Width = width
                Pump(window)
                For Each panelName In {"SourceActionsPanel", "SearchActionsPanel"}
                    Dim panel = DirectCast(toolbar.FindName(panelName), StackPanel)
                    Dim buttons = panel.Children.OfType(Of Button)().ToArray()
                    For index = 1 To buttons.Length - 1
                        Dim previous = buttons(index - 1).TranslatePoint(New Point(0, 0), panel)
                        Dim current = buttons(index).TranslatePoint(New Point(0, 0), panel)
                        Dim gap = current.X - previous.X - buttons(index - 1).ActualWidth
                        Require(Math.Abs(gap - 8) <= 1, "Action buttons no longer have consistent 8px spacing.")
                    Next
                Next
                Dim options = DirectCast(toolbar.FindName("SearchOptionsToolbar"), WrapPanel)
                Require(helpButton.TranslatePoint(New Point(), toolbar).X < settingsButton.TranslatePoint(New Point(), toolbar).X, "Help is not left of Settings.")
                Dim actions = DirectCast(toolbar.FindName("SearchActionsPanel"), StackPanel)
                Require(actions.Parent Is options.Parent AndAlso TypeOf actions.Parent Is Grid, "Search options and actions must share the same row.")
                Dim actionBounds = actions.TransformToAncestor(toolbar).TransformBounds(New Rect(actions.RenderSize))
                Dim optionBounds = options.TransformToAncestor(toolbar).TransformBounds(New Rect(options.RenderSize))
                Require(Math.Abs(actionBounds.Right - toolbar.ActualWidth) <= 1, "Search actions are not anchored to the far right.")
                Require(optionBounds.Right + 11 <= actionBounds.Left, "Search options overlap the action buttons.")
                Dim mode = DirectCast(toolbar.FindName("SearchMode_cmb"), ComboBox)
                Require(Math.Abs(mode.TranslatePoint(New Point(0, 0), toolbar).Y - actionBounds.Top) <= 2, "Actions are not aligned with the search-mode row.")
                Dim watermark = DirectCast(toolbar.FindName("ToolbarWatermark"), Image)
                Require(watermark.Source IsNot Nothing AndAlso Not watermark.IsHitTestVisible AndAlso Not watermark.Focusable,
                        "Toolbar watermark must load without intercepting input.")
                Require(watermark.Opacity > 0 AndAlso watermark.Opacity <= 0.35, "Toolbar watermark is not subtle.")
                Dim watermarkBounds = watermark.TransformToAncestor(toolbar).TransformBounds(New Rect(watermark.RenderSize))
                Require(Math.Abs(watermarkBounds.Left + watermarkBounds.Width / 2 - toolbar.ActualWidth / 2) <= 1,
                        "Toolbar watermark is not centered across the row.")
                For Each control In Descendants(toolbar).OfType(Of Control)().Where(Function(item) TypeOf item Is Button OrElse TypeOf item Is ComboBox OrElse TypeOf item Is CheckBox)
                    Dim controlBounds = control.TransformToAncestor(toolbar).TransformBounds(New Rect(control.RenderSize))
                    Require(Not watermarkBounds.IntersectsWith(controlBounds), "Watermark overlaps a toolbar control.")
                Next
                For Each button In Descendants(toolbar).OfType(Of Button)().Where(Function(item) item.Name.Length > 0)
                    Dim bounds = button.TransformToAncestor(toolbar).TransformBounds(New Rect(button.RenderSize))
                    Require(bounds.Left >= -1 AndAlso bounds.Right <= toolbar.ActualWidth + 1, "Toolbar button falls outside the window.")
                Next
            Next
        Finally
            window.Close()
        End Try
    End Sub

    Public Sub RunCompactResults(dark As Boolean)
        Dim document = System.Xml.Linq.XDocument.Load(System.IO.Path.Combine(AppContext.BaseDirectory, "Fixtures", "MainWindow.xaml"))
        Dim presentation = System.Xml.Linq.XNamespace.Get("http://schemas.microsoft.com/winfx/2006/xaml/presentation")
        Dim xaml = System.Xml.Linq.XNamespace.Get("http://schemas.microsoft.com/winfx/2006/xaml")
        Dim styleMarkup = New System.Xml.Linq.XElement(document.Root.Element(presentation + "Window.Resources").Elements(presentation + "Style").
            Single(Function(item) CStr(item.Attribute(xaml + "Key")) = "SearchResultItem"))
        styleMarkup.SetAttributeValue("xmlns", presentation.NamespaceName)
        styleMarkup.SetAttributeValue(System.Xml.Linq.XNamespace.Xmlns + "x", xaml.NamespaceName)
        Dim sidebarMarkup = New System.Xml.Linq.XElement(document.Descendants(presentation + "Grid").Single(Function(item) CStr(item.Attribute(xaml + "Name")) = "SearchResultsPane"))
        sidebarMarkup.SetAttributeValue("xmlns", presentation.NamespaceName)
        sidebarMarkup.SetAttributeValue(System.Xml.Linq.XNamespace.Xmlns + "x", xaml.NamespaceName)
        Dim resources As New System.Xml.Linq.XElement(presentation + "Grid.Resources", styleMarkup)
        sidebarMarkup.AddFirst(resources)
        Dim sidebar = DirectCast(System.Windows.Markup.XamlReader.Parse(sidebarMarkup.ToString()), Grid)
        Dim window As New SettingsWindow(New BeaconSettings(), dark) With {
            .WindowStartupLocation = WindowStartupLocation.Manual,
            .Left = -10000, .Top = -10000, .ShowActivated = False, .ShowInTaskbar = False, .Content = sidebar
        }
        Dim results = DirectCast(sidebar.FindName("Results_lst"), ListBox)
        Dim details = DirectCast(sidebar.FindName("Details_lst"), ListBox)
        Dim expander = DirectCast(sidebar.FindName("MatchDetails_exp"), Expander)
        results.ItemsSource = {New With {.DisplayName = "sample.log", .CompactCount = "1+", .LogicalPath = "C:\Logs\sample.log",
                                         .MatchSummary = "1 matching record · first matching line only", .Excerpt = "This excerpt must not appear in the file row."}}
        details.ItemsSource = {New With {.Location = "Line 4", .CountLabel = "1", .Excerpt = "Connection error"}}
        Try
            window.Show()
            Pump(window)
            Require(Not expander.IsExpanded AndAlso Not details.IsVisible, "Match details should start collapsed.")
            Require(sidebar.RowDefinitions(1).Height.IsAuto, "Collapsed details reserve unnecessary vertical space.")
            Dim row = DirectCast(results.ItemContainerGenerator.ContainerFromIndex(0), ListBoxItem)
            Dim labels = Descendants(row).OfType(Of TextBlock)().Where(Function(item) item.Text.Length > 0).ToArray()
            Require(labels.Length = 2 AndAlso labels.Any(Function(item) item.Text = "sample.log") AndAlso labels.Any(Function(item) item.Text = "1+"),
                    "File rows must show only a compact name and count.")
            Dim toggle = DirectCast(expander.Template.FindName("DetailsToggle", expander), ToggleButton)
            CheckContrast(toggle.Foreground, DirectCast(window.Resources("CardBackgroundBrush"), Brush), "Match details header")
            toggle.IsChecked = True
            Pump(window)
            Require(expander.IsExpanded AndAlso details.IsVisible AndAlso details.ActualHeight > 0, "Match details cannot be expanded.")
            Require(Descendants(expander).OfType(Of TextBlock)().Any(Function(item) item.Text.StartsWith("Select a match to jump", StringComparison.Ordinal)),
                    "Match-details purpose is not explained.")
            toggle.IsChecked = False
            Pump(window)
            Require(Not details.IsVisible, "Match details did not collapse again.")
        Finally
            window.Close()
        End Try
    End Sub

    Public Sub RunCounterBars(dark As Boolean)
        Dim document = System.Xml.Linq.XDocument.Load(System.IO.Path.Combine(AppContext.BaseDirectory, "Fixtures", "MainWindow.xaml"))
        Dim presentation = System.Xml.Linq.XNamespace.Get("http://schemas.microsoft.com/winfx/2006/xaml/presentation")
        Dim xaml = System.Xml.Linq.XNamespace.Get("http://schemas.microsoft.com/winfx/2006/xaml")
        Dim window As New SettingsWindow(New BeaconSettings(), dark) With {
            .WindowStartupLocation = WindowStartupLocation.Manual,
            .Left = -10000, .Top = -10000, .ShowActivated = False, .ShowInTaskbar = False
        }
        Dim panel As New StackPanel With {.Margin = New Thickness(16)}
        window.Content = panel
        For Each name In {"EventCounterBar", "HarCounterBar"}
            Dim element = New System.Xml.Linq.XElement(document.Descendants(presentation + "Border").Single(Function(item) CStr(item.Attribute(xaml + "Name")) = name))
            element.SetAttributeValue("xmlns", presentation.NamespaceName)
            element.SetAttributeValue(System.Xml.Linq.XNamespace.Xmlns + "x", xaml.NamespaceName)
            panel.Children.Add(DirectCast(System.Windows.Markup.XamlReader.Parse(element.ToString()), Border))
        Next
        Try
            window.Show()
            For pass = 0 To 1
                Pump(window)
                Dim expected = If(dark, Color.FromRgb(&H3A, &H3A, &H3A), Color.FromRgb(&HF0, &HF0, &HF0))
                For Each bar As Border In panel.Children
                    Require(DirectCast(bar.Background, SolidColorBrush).Color = expected, "Counter background did not follow the theme.")
                    Require(bar.ActualHeight >= 34, "Counter strip is too short.")
                    Require(bar.CornerRadius.BottomLeft = 4, "Counter corners differ from Beacon.")
                    For Each label As TextBlock In Descendants(bar).OfType(Of TextBlock)()
                        CheckContrast(label.Foreground, bar.Background, label.Name)
                        Require(label.ActualHeight >= label.DesiredSize.Height - 0.5, "Counter text is vertically clipped.")
                    Next
                Next
                dark = Not dark
                window.Resources("ButtonBackgroundBrush") = New SolidColorBrush(If(dark, Color.FromRgb(&H3A, &H3A, &H3A), Color.FromRgb(&HF0, &HF0, &HF0)))
                window.Resources("TextPrimaryBrush") = New SolidColorBrush(If(dark, Color.FromRgb(&HE0, &HE0, &HE0), Color.FromRgb(&H20, &H20, &H20)))
                window.Resources("TextSecondaryBrush") = New SolidColorBrush(If(dark, Color.FromRgb(&HB0, &HB0, &HB0), Color.FromRgb(&H66, &H66, &H66)))
            Next
        Finally
            window.Close()
        End Try
    End Sub

    Private Sub CheckFontSelector(window As SettingsWindow, dark As Boolean)
        Dim expected = Fonts.SystemFontFamilies.Select(Function(family) family.Source).
            Where(Function(name) Not String.IsNullOrWhiteSpace(name)).
            Distinct(StringComparer.OrdinalIgnoreCase).
            OrderBy(Function(name) name, StringComparer.CurrentCultureIgnoreCase).ToArray()
        Dim selector = DirectCast(window.FindName("PreviewFont_cmb"), ComboBox)
        Require(expected.Length > 0, "No installed fonts were available to test.")
        Require(selector.Items.Cast(Of String)().SequenceEqual(expected), "Font dropdown differs from the sorted installed-font list.")
        Require(Not selector.IsEditable AndAlso selector.IsTextSearchEnabled, "Font selection must be installed-only with type-to-select.")
        Require(selector.SelectedItem IsNot Nothing AndAlso expected.Contains(CStr(selector.SelectedItem)), "Font selector has no installed selection.")

        Dim saved As New BeaconSettings With {.PreviewFontFamily = expected.Last().ToUpperInvariant()}
        Dim savedWindow As New SettingsWindow(saved, dark)
        Try
            Dim savedSelector = DirectCast(savedWindow.FindName("PreviewFont_cmb"), ComboBox)
            Require(CStr(savedSelector.SelectedItem) = expected.Last(), "Saved installed font was not restored case-insensitively.")
            DirectCast(savedWindow.FindName("RestoreDefaults_btn"), Button).RaiseEvent(New RoutedEventArgs(Button.ClickEvent))
            Dim defaultFont = expected.FirstOrDefault(Function(name) String.Equals(name, BeaconSettings.CreateDefaults().PreviewFontFamily, StringComparison.OrdinalIgnoreCase))
            Require(CStr(savedSelector.SelectedItem) = If(defaultFont, expected.First()), "Restore Defaults did not select an installed default font.")
            Require(saved.PreviewFontFamily = expected.Last().ToUpperInvariant() AndAlso savedWindow.SavedSettings Is Nothing, "Selection checks changed original settings.")
        Finally
            savedWindow.Close()
        End Try

        Dim missingWindow As New SettingsWindow(New BeaconSettings With {.PreviewFontFamily = "Beacon-missing-font-" & Guid.NewGuid().ToString("N")}, dark)
        Try
            Dim fallback = DirectCast(missingWindow.FindName("PreviewFont_cmb"), ComboBox)
            Require(fallback.SelectedItem IsNot Nothing AndAlso expected.Contains(CStr(fallback.SelectedItem)), "Missing saved font did not fall back to an installed font.")
            Require(fallback.Items.Cast(Of String)().SequenceEqual(expected), "Missing saved font was added to the dropdown.")
        Finally
            missingWindow.Close()
        End Try
    End Sub

    Private Sub CheckTextFits(window As Window, editor As TextBox)
        Dim originalText = editor.Text
        Dim originalSize = editor.FontSize
        Dim host = DirectCast(editor.Template.FindName("PART_ContentHost", editor), ScrollViewer)
        Try
            editor.Text = "Agjpqy 012345"
            For Each size In {originalSize, 24.0}
                editor.FontSize = size
                Pump(window)
                Require(TextFitsViewport(editor, host), $"{editor.Name}: text is vertically clipped at size {size}.")
            Next

            editor.FontSize = originalSize
            editor.Height = editor.MinHeight
            host.Margin = editor.Padding
            Pump(window)
            Require(Not TextFitsViewport(editor, host), "Clipping check failed to detect the original duplicate-padding regression.")
        Finally
            host.ClearValue(FrameworkElement.MarginProperty)
            editor.ClearValue(FrameworkElement.HeightProperty)
            editor.FontSize = originalSize
            editor.Text = originalText
            Pump(window)
        End Try
    End Sub

    Private Function TextFitsViewport(editor As TextBox, host As ScrollViewer) As Boolean
        Dim viewport = Descendants(host).OfType(Of ScrollContentPresenter)().First()
        Dim bounds = viewport.TransformToAncestor(editor).TransformBounds(New Rect(0, 0, viewport.ActualWidth, viewport.ActualHeight))
        For Each index In {0, editor.Text.Length - 1}
            Dim glyph = editor.GetRectFromCharacterIndex(index)
            If glyph.IsEmpty OrElse glyph.Height <= 0 OrElse glyph.Top < bounds.Top - 0.5 OrElse glyph.Bottom > bounds.Bottom + 0.5 Then Return False
        Next
        Return True
    End Function

    Private Sub Pump(window As Window)
        window.UpdateLayout()
        window.Dispatcher.Invoke(Sub()
                                 End Sub, DispatcherPriority.ContextIdle)
        window.UpdateLayout()
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

    Private Sub CheckContrast(foreground As Brush, background As Brush, name As String)
        Dim fg = TryCast(foreground, SolidColorBrush)
        Dim bg = TryCast(background, SolidColorBrush)
        Require(fg IsNot Nothing AndAlso bg IsNot Nothing AndAlso fg.Color.A = 255 AndAlso bg.Color.A = 255,
                name & " does not have explicit opaque colors.")
        Dim first = Luminance(fg.Color)
        Dim second = Luminance(bg.Color)
        Dim ratio = (Math.Max(first, second) + 0.05) / (Math.Min(first, second) + 0.05)
        Require(ratio >= 4.5, $"{name} contrast is {ratio:F2}:1, below 4.5:1.")
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
