Imports System.Windows
Imports System.Windows.Media

Namespace Beacon

    Partial Public Class SettingsWindow
        Inherits Window

        Public Property SavedSettings As BeaconSettings
        Private _workingSettings As BeaconSettings
        Private ReadOnly _installedFontNames As String()

        Public Sub New(settings As BeaconSettings, isDarkMode As Boolean)
            InitializeComponent()
            _workingSettings = BeaconSettingsService.Clone(settings)
            _installedFontNames = Fonts.SystemFontFamilies.Select(Function(family) family.Source).
                Where(Function(name) Not String.IsNullOrWhiteSpace(name)).
                Distinct(StringComparer.OrdinalIgnoreCase).
                OrderBy(Function(name) name, StringComparer.CurrentCultureIgnoreCase).ToArray()
            PreviewFont_cmb.ItemsSource = _installedFontNames

            SearchMode_cmb.ItemsSource = [Enum].GetValues(Of SearchMode)()
            AccessDenied_cmb.ItemsSource = [Enum].GetValues(Of AccessDeniedAction)()
            EvtxResourcePolicy_cmb.ItemsSource = [Enum].GetValues(Of EvtxResourcePolicy)()
            Diagnostics_cmb.ItemsSource = [Enum].GetValues(Of DiagnosticDetail)()

            AddHandler Save_btn.Click, AddressOf Save_btn_Click
            AddHandler Cancel_btn.Click, Sub() DialogResult = False
            AddHandler RestoreDefaults_btn.Click, AddressOf RestoreDefaults_btn_Click

            ApplyTheme(isDarkMode)
            LoadControls()
        End Sub

        Private Sub ApplyTheme(isDarkMode As Boolean)
            If Not isDarkMode Then Return
            Resources("WindowBackgroundBrush") = New SolidColorBrush(Color.FromRgb(&H20, &H20, &H20))
            Resources("CardBackgroundBrush") = New SolidColorBrush(Color.FromRgb(&H2B, &H2B, &H2B))
            Resources("CardBorderBrush") = New SolidColorBrush(Color.FromRgb(&H3F, &H3F, &H3F))
            Resources("TextPrimaryBrush") = New SolidColorBrush(Color.FromRgb(&HE0, &HE0, &HE0))
            Resources("TextSecondaryBrush") = New SolidColorBrush(Color.FromRgb(&HB0, &HB0, &HB0))
            Resources("InputBackgroundBrush") = New SolidColorBrush(Color.FromRgb(&H2B, &H2B, &H2B))
            Resources("InputBorderBrush") = New SolidColorBrush(Color.FromRgb(&H50, &H50, &H50))
            Resources("AccentBrush") = New SolidColorBrush(Color.FromRgb(&H60, &HCF, &HFF))
            Resources("ButtonBackgroundBrush") = New SolidColorBrush(Color.FromRgb(&H3A, &H3A, &H3A))
            Resources("ButtonHoverBrush") = New SolidColorBrush(Color.FromRgb(&H45, &H45, &H45))
            Resources("ButtonPressedBrush") = New SolidColorBrush(Color.FromRgb(&H50, &H50, &H50))
            Resources("SelectionBackgroundBrush") = New SolidColorBrush(Color.FromRgb(&H23, &H48, &H5B))
            Resources("CheckmarkBrush") = New SolidColorBrush(Color.FromRgb(&H20, &H20, &H20))
        End Sub

        Private Sub LoadControls()
            MaximumTotalResults_txt.Text = _workingSettings.MaximumTotalResults.ToString()
            MaximumStructuredMatches_txt.Text = _workingSettings.MaximumStructuredMatches.ToString()
            SearchMode_cmb.SelectedItem = _workingSettings.DefaultSearchMode
            SearchContents_chk.IsChecked = _workingSettings.SearchFileContents
            SearchFileNames_chk.IsChecked = _workingSettings.SearchFileNames
            SearchFullPaths_chk.IsChecked = _workingSettings.SearchFullPaths
            StopAfterFirstMatch_chk.IsChecked = _workingSettings.StopAfterFirstMatchPerFile
            ScanWorkerCount_txt.Text = _workingSettings.ScanWorkerCount.ToString()

            IncludedExtensions_txt.Text = _workingSettings.IncludedExtensions
            ExcludedDirectories_txt.Text = _workingSettings.ExcludedDirectories
            MaximumFileSize_txt.Text = _workingSettings.MaximumFileSizeMb.ToString()
            IncludeHidden_chk.IsChecked = _workingSettings.IncludeHiddenFiles
            IncludeSystem_chk.IsChecked = _workingSettings.IncludeSystemFiles
            FollowReparsePoints_chk.IsChecked = _workingSettings.FollowReparsePoints
            AccessDenied_cmb.SelectedItem = _workingSettings.AccessDeniedBehavior

            ArchiveDepth_txt.Text = _workingSettings.ArchiveNestingDepth.ToString()
            ArchiveEntries_txt.Text = _workingSettings.MaximumArchiveEntries.ToString()
            ArchiveEntrySize_txt.Text = _workingSettings.MaximumArchiveEntrySizeMb.ToString()
            ArchiveExpandedSize_txt.Text = _workingSettings.MaximumArchiveExpandedSizeMb.ToString()
            CompressionRatio_txt.Text = _workingSettings.MaximumCompressionRatio.ToString()
            ArchiveTimeout_txt.Text = _workingSettings.ArchiveProcessTimeoutSeconds.ToString()

            EvtxMatches_txt.Text = _workingSettings.EvtxMaximumMatches.ToString()
            RawXmlFallback_chk.IsChecked = _workingSettings.ShowRawXmlWhenMessageUnavailable
            EvtxResourcePolicy_cmb.SelectedItem = _workingSettings.EvtxMessageResourceBehavior
            HarMatches_txt.Text = _workingSettings.HarMaximumMatches.ToString()
            HarBodySize_txt.Text = _workingSettings.MaximumHarBodySizeMb.ToString()
            DecodeBase64_chk.IsChecked = _workingSettings.DecodeBase64HarBodies
            RedactHar_chk.IsChecked = _workingSettings.RedactSensitiveHarData

            PreviewFont_cmb.SelectedItem = _installedFontNames.FirstOrDefault(
                Function(name) String.Equals(name, _workingSettings.PreviewFontFamily, StringComparison.OrdinalIgnoreCase))
            If PreviewFont_cmb.SelectedIndex < 0 Then
                PreviewFont_cmb.SelectedItem = _installedFontNames.FirstOrDefault(
                    Function(name) String.Equals(name, BeaconSettings.CreateDefaults().PreviewFontFamily, StringComparison.OrdinalIgnoreCase))
            End If
            If PreviewFont_cmb.SelectedIndex < 0 AndAlso _installedFontNames.Length > 0 Then PreviewFont_cmb.SelectedIndex = 0
            PreviewFontSize_txt.Text = _workingSettings.PreviewFontSize.ToString()
            PreviewSize_txt.Text = _workingSettings.MaximumPreviewSizeMb.ToString()
            WordWrap_chk.IsChecked = _workingSettings.PreviewWordWrap
            PrettyPrint_chk.IsChecked = _workingSettings.PrettyPrintJsonAndXml
            SelectFirstResult_chk.IsChecked = _workingSettings.SelectFirstResult
            AbsolutePaths_chk.IsChecked = _workingSettings.DisplayAbsolutePaths
            Diagnostics_cmb.SelectedItem = _workingSettings.Diagnostics
        End Sub

        Private Sub Save_btn_Click(sender As Object, e As RoutedEventArgs)
            Try
                _workingSettings.MaximumTotalResults = ParseInteger(MaximumTotalResults_txt, "Maximum total results")
                _workingSettings.MaximumStructuredMatches = ParseInteger(MaximumStructuredMatches_txt, "Maximum structured matches")
                _workingSettings.DefaultSearchMode = DirectCast(SearchMode_cmb.SelectedItem, SearchMode)
                _workingSettings.SearchFileContents = SearchContents_chk.IsChecked.GetValueOrDefault()
                _workingSettings.SearchFileNames = SearchFileNames_chk.IsChecked.GetValueOrDefault()
                _workingSettings.SearchFullPaths = SearchFullPaths_chk.IsChecked.GetValueOrDefault()
                _workingSettings.StopAfterFirstMatchPerFile = StopAfterFirstMatch_chk.IsChecked.GetValueOrDefault()
                _workingSettings.ScanWorkerCount = ParseInteger(ScanWorkerCount_txt, "Scan workers")

                _workingSettings.IncludedExtensions = IncludedExtensions_txt.Text
                _workingSettings.ExcludedDirectories = ExcludedDirectories_txt.Text
                _workingSettings.MaximumFileSizeMb = ParseInteger(MaximumFileSize_txt, "Maximum file size")
                _workingSettings.IncludeHiddenFiles = IncludeHidden_chk.IsChecked.GetValueOrDefault()
                _workingSettings.IncludeSystemFiles = IncludeSystem_chk.IsChecked.GetValueOrDefault()
                _workingSettings.FollowReparsePoints = FollowReparsePoints_chk.IsChecked.GetValueOrDefault()
                _workingSettings.AccessDeniedBehavior = DirectCast(AccessDenied_cmb.SelectedItem, AccessDeniedAction)

                _workingSettings.ArchiveNestingDepth = ParseInteger(ArchiveDepth_txt, "Archive nesting depth")
                _workingSettings.MaximumArchiveEntries = ParseInteger(ArchiveEntries_txt, "Maximum archive entries")
                _workingSettings.MaximumArchiveEntrySizeMb = ParseInteger(ArchiveEntrySize_txt, "Maximum archive entry size")
                _workingSettings.MaximumArchiveExpandedSizeMb = ParseInteger(ArchiveExpandedSize_txt, "Maximum expanded archive size")
                _workingSettings.MaximumCompressionRatio = ParseInteger(CompressionRatio_txt, "Maximum compression ratio")
                _workingSettings.ArchiveProcessTimeoutSeconds = ParseInteger(ArchiveTimeout_txt, "Archive timeout")

                _workingSettings.EvtxMaximumMatches = ParseInteger(EvtxMatches_txt, "Maximum EVTX matches")
                _workingSettings.ShowRawXmlWhenMessageUnavailable = RawXmlFallback_chk.IsChecked.GetValueOrDefault()
                _workingSettings.EvtxMessageResourceBehavior = DirectCast(EvtxResourcePolicy_cmb.SelectedItem, EvtxResourcePolicy)
                _workingSettings.HarMaximumMatches = ParseInteger(HarMatches_txt, "Maximum HAR matches")
                _workingSettings.MaximumHarBodySizeMb = ParseInteger(HarBodySize_txt, "Maximum HAR body size")
                _workingSettings.DecodeBase64HarBodies = DecodeBase64_chk.IsChecked.GetValueOrDefault()
                _workingSettings.RedactSensitiveHarData = RedactHar_chk.IsChecked.GetValueOrDefault()

                Dim selectedFont = TryCast(PreviewFont_cmb.SelectedItem, String)
                If selectedFont Is Nothing OrElse Not _installedFontNames.Contains(selectedFont, StringComparer.Ordinal) Then
                    Throw New ArgumentException("Select a font installed on this computer.")
                End If
                _workingSettings.PreviewFontFamily = selectedFont
                _workingSettings.PreviewFontSize = ParseInteger(PreviewFontSize_txt, "Preview font size")
                _workingSettings.MaximumPreviewSizeMb = ParseInteger(PreviewSize_txt, "Maximum preview size")
                _workingSettings.PreviewWordWrap = WordWrap_chk.IsChecked.GetValueOrDefault()
                _workingSettings.PrettyPrintJsonAndXml = PrettyPrint_chk.IsChecked.GetValueOrDefault()
                _workingSettings.SelectFirstResult = SelectFirstResult_chk.IsChecked.GetValueOrDefault()
                _workingSettings.DisplayAbsolutePaths = AbsolutePaths_chk.IsChecked.GetValueOrDefault()
                _workingSettings.Diagnostics = DirectCast(Diagnostics_cmb.SelectedItem, DiagnosticDetail)

                SavedSettings = BeaconSettingsService.Validate(_workingSettings)
                BeaconSettingsService.Save(SavedSettings)
                DialogResult = True
            Catch ex As Exception
                MessageBox.Show(Me, ex.Message, "Invalid setting", MessageBoxButton.OK, MessageBoxImage.Warning)
            End Try
        End Sub

        Private Sub RestoreDefaults_btn_Click(sender As Object, e As RoutedEventArgs)
            _workingSettings = BeaconSettings.CreateDefaults()
            LoadControls()
        End Sub

        Private Shared Function ParseInteger(textBox As Controls.TextBox, settingName As String) As Integer
            Dim result As Integer
            If Not Integer.TryParse(textBox.Text, result) Then
                textBox.Focus()
                Throw New ArgumentException($"{settingName} must be a whole number.")
            End If
            Return result
        End Function
    End Class

End Namespace
