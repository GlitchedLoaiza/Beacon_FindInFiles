Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Controls.Primitives
Imports System.Windows.Media
Imports Beacon
Imports BeaconFilterControls

Module FilterControlChecks
    Public Sub Run()
        Dim leap As New DateTime(2024, 2, 29, 23, 59, 59, DateTimeKind.Utc)
        Require(UtcDateTimePicker.Adjust(leap, 0, 1).Day = 28, "Year arrow did not handle leap day.")
        Require(UtcDateTimePicker.Adjust(New DateTime(2024, 1, 31), 1, 1).Day = 29, "Month arrow did not clamp to valid day.")
        Require(UtcDateTimePicker.Adjust(leap, 5, 1) = New DateTime(2024, 3, 1), "Second arrow did not roll over the date.")
        For Each dark In {False, True}
            Dim options As New BeaconSettings With {.EvtxProvider = "Typed.Provider", .EvtxLevels = "2,3", .EvtxFromUtc = "2024-02-29 23:59:59"}
            Dim window As New SettingsWindow(options, dark) With {.Left = -10000, .Top = -10000, .WindowStartupLocation = WindowStartupLocation.Manual, .ShowActivated = False, .ShowInTaskbar = False}
            Try
                window.SetProviderSuggestions({"Zeta", "Alpha", "alpha", ""})
                window.Show()
                DirectCast(window.FindName("SettingsTabs"), TabControl).SelectedIndex = 3
                Pump(window)
                Dim provider = DirectCast(window.FindName("EvtxProvider_cmb"), ComboBox)
                Require(provider.IsEditable AndAlso provider.Text = "Typed.Provider", "Provider suggestions replaced saved text.")
                Require(provider.Items.Cast(Of String)().SequenceEqual({"", "Alpha", "Zeta"}), "Provider suggestions were not sorted/deduplicated.")
                Dim editor = DirectCast(provider.Template.FindName("PART_EditableTextBox", provider), TextBox)
                Require(editor.IsVisible, "Provider combo has no editable text surface.")
                editor.Text = "Manual.Provider"
                Require(provider.Text = "Manual.Provider", "Direct provider input did not update filter text.")
                provider.SelectedItem = "Alpha"
                Require(provider.Text = "Alpha", "Provider selection did not update text.")

                Dim picker = DirectCast(DirectCast(window.FindName("EvtxFrom_host"), ContentControl).Content, UtcDateTimePicker)
                Require(picker.Text = options.EvtxFromUtc, "Saved date did not load.")
                picker.OpenEditor()
                Pump(window)
                Dim popup = DirectCast(picker.FindName("PickerPopup"), Popup)
                Require(popup.IsOpen AndAlso popup.PopupAnimation = PopupAnimation.None, "Picker is missing or animated.")
                Dim border = DirectCast(popup.Child, Border)
                Require(DirectCast(border.Background, SolidColorBrush).Color = DirectCast(window.Resources("CardBackgroundBrush"), SolidColorBrush).Color, "Date popup did not inherit theme.")
                picker.StepComponent(5, 1)
                Require(picker.Text = options.EvtxFromUtc, "Arrow changed committed date before Apply.")
                picker.CommitEditor()
                Require(picker.Text = "2024-03-01 00:00:00", "Picker Apply lost UTC precision.")
                picker.Text = "invalid date"
                picker.OpenEditor()
                picker.CommitEditor()
                Require(popup.IsOpen AndAlso picker.Text = "invalid date", "Invalid input was silently replaced.")
                picker.Clear()
                Require(picker.Text = "" AndAlso Not popup.IsOpen, "Any time did not clear the date.")
                picker.Text = "9999-12-31 23:59:59"
                picker.OpenEditor()
                picker.StepComponent(5, 1)
                Require(DirectCast(picker.FindName("PickerError"), TextBlock).Text.Length > 0, "Date overflow was not reported.")
                popup.IsOpen = False
                picker.Text = "2024-06-01T12:34:56Z"
                picker.OpenEditor()
                Dim fields = DirectCast(picker.FindName("Components"), UniformGrid).Children.OfType(Of StackPanel)().SelectMany(Function(panel) panel.Children.OfType(Of TextBox)()).ToArray()
                fields(1).Text = "07"
                picker.CommitEditor()
                Require(picker.Text = "2024-07-01 12:34:56", "Typed component was not committed.")

                Dim levels = DirectCast(DirectCast(window.FindName("EvtxLevels_host"), ContentControl).Content, SeveritySelector)
                Require(levels.Text = "2,3", "Saved multiple severity levels changed.")
                DirectCast(levels.FindName("OpenLevels"), Button).RaiseEvent(New RoutedEventArgs(Button.ClickEvent))
                Pump(window)
                Dim levelsPopup = DirectCast(levels.FindName("LevelsPopup"), Popup)
                Require(levelsPopup.IsOpen AndAlso levelsPopup.PopupAnimation = PopupAnimation.None, "Severity dropdown did not open.")
                Dim checks = DirectCast(levels.FindName("LevelChoices"), StackPanel).Children.OfType(Of CheckBox)().ToArray()
                Require(checks(2).IsChecked.GetValueOrDefault() AndAlso checks(3).IsChecked.GetValueOrDefault(), "Saved multi-selection is not visible.")
                checks(1).IsChecked = True
                Require(levels.Text = "1,2,3", "Named severity choices did not preserve multiple values.")
                levels.Clear()
                Require(levels.Text = "" AndAlso checks.All(Function(check) Not check.IsChecked.GetValueOrDefault()), "All levels did not clear the filter.")
                levelsPopup.IsOpen = False
                DirectCast(window.FindName("RestoreDefaults_btn"), Button).RaiseEvent(New RoutedEventArgs(Button.ClickEvent))
                Require(picker.Text = "" AndAlso levels.Text = "" AndAlso provider.Text = "", "Restore defaults did not clear friendly controls.")
            Finally
                window.Close()
            End Try
        Next
    End Sub

    Private Sub Pump(window As Window)
        window.UpdateLayout()
        window.Dispatcher.Invoke(Sub()
                                 End Sub, System.Windows.Threading.DispatcherPriority.ContextIdle)
        window.UpdateLayout()
    End Sub
    Private Sub Require(condition As Boolean, message As String)
        If Not condition Then Throw New InvalidOperationException(message)
    End Sub
End Module
