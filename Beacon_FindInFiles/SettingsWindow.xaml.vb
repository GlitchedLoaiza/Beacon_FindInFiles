Imports System.Windows
Imports System.Windows.Media
Imports System.Threading
Imports System.Threading.Tasks

Namespace Beacon

    Partial Public Class SettingsWindow
        Inherits Window

        Public Property SavedSettings As BeaconSettings
        Private _workingSettings As BeaconSettings
        Private ReadOnly _installedFontNames As String()
        Private ReadOnly _checkForUpdates As Func(Of CancellationToken, Task(Of ReleaseCheckResult))
        Private ReadOnly _updateCancellation As New CancellationTokenSource()
        Private _updateCheckRunning As Boolean
        Private _settingsClosed As Boolean
        Private ReadOnly EvtxLevels_txt As New BeaconFilterControls.SeveritySelector()
        Private ReadOnly EvtxFrom_txt As New BeaconFilterControls.UtcDateTimePicker()
        Private ReadOnly EvtxTo_txt As New BeaconFilterControls.UtcDateTimePicker()
        Private ReadOnly _harFilters As New HarFilterEditor()

        Public Sub New(settings As BeaconSettings, isDarkMode As Boolean)
            Me.New(settings, isDarkMode, Function(token) ReleaseUpdateChecker.CheckResultAsync(GetType(SettingsWindow).Assembly.GetName().Version, token))
        End Sub

        Friend Sub New(settings As BeaconSettings, isDarkMode As Boolean, checkForUpdates As Func(Of CancellationToken, Task(Of ReleaseCheckResult)))
            InitializeComponent()
            EvtxLevels_host.Content = EvtxLevels_txt
            EvtxFrom_host.Content = EvtxFrom_txt
            EvtxTo_host.Content = EvtxTo_txt
            HarFilters_host.Content = _harFilters
            _checkForUpdates = checkForUpdates
            _workingSettings = BeaconSettingsService.Clone(settings)
            _installedFontNames = Fonts.SystemFontFamilies.Select(Function(family) family.Source).
                Where(Function(name) Not String.IsNullOrWhiteSpace(name)).
                Distinct(StringComparer.OrdinalIgnoreCase).
                OrderBy(Function(name) name, StringComparer.CurrentCultureIgnoreCase).ToArray()
            PreviewFont_cmb.ItemsSource = _installedFontNames

            SearchMode_cmb.ItemsSource = [Enum].GetValues(Of SearchMode)()
            AccessDenied_cmb.ItemsSource = [Enum].GetValues(Of AccessDeniedAction)()
            EvtxResourcePolicy_cmb.ItemsSource = {EvtxResourcePolicy.OfflineOnly}
            EvtxResourceHelp_txt.Text = EvtxFilter.ResourceGuidance
            Diagnostics_cmb.ItemsSource = [Enum].GetValues(Of DiagnosticDetail)()

            AddHandler Save_btn.Click, AddressOf Save_btn_Click
            AddHandler Cancel_btn.Click, Sub() DialogResult = False
            AddHandler RestoreDefaults_btn.Click, AddressOf RestoreDefaults_btn_Click
            AddHandler CheckUpdates_btn.Click, AddressOf CheckUpdates_Click
            AddHandler RedactHar_chk.Checked, AddressOf UpdateHarPrivacyHelp
            AddHandler RedactHar_chk.Unchecked, AddressOf UpdateHarPrivacyHelp

            ApplyTheme(isDarkMode)
            LoadControls()
        End Sub

        Private Async Sub CheckUpdates_Click(sender As Object, e As RoutedEventArgs)
            If _updateCheckRunning OrElse _settingsClosed Then Return
            _updateCheckRunning = True
            CheckUpdates_btn.IsEnabled = False
            UpdateCheckStatus_txt.Text = "Checking for updates…"
            Try
                Dim result = Await _checkForUpdates(_updateCancellation.Token)
                If _settingsClosed Then Return
                Select Case result.Status
                    Case ReleaseCheckStatus.UpdateAvailable
                        Dim version = result.LatestVersion
                        Dim label = If(version.Revision > 0, version.ToString(4), version.ToString(3))
                        UpdateCheckStatus_txt.Text = $"A newer version of Beacon is available: v{label}."
                    Case ReleaseCheckStatus.UpToDate
                        UpdateCheckStatus_txt.Text = "Beacon is up to date. No updates pending."
                    Case Else
                        UpdateCheckStatus_txt.Text = "Couldn't check for updates. Check your connection and try again later."
                End Select
            Catch ex As Exception
                If Not _settingsClosed Then UpdateCheckStatus_txt.Text = "Couldn't check for updates. Please try again later."
            Finally
                _updateCheckRunning = False
                If Not _settingsClosed Then CheckUpdates_btn.IsEnabled = True
            End Try
        End Sub

        Protected Overrides Sub OnClosed(e As EventArgs)
            _settingsClosed = True
            _updateCancellation.Cancel()
            MyBase.OnClosed(e)
        End Sub

        Private Sub UpdateHarPrivacyHelp(sender As Object, e As RoutedEventArgs)
            HarPrivacyHelp_txt.Text = If(RedactHar_chk.IsChecked.GetValueOrDefault(),
                "HAR redaction: On. Recognized sensitive fields are hidden and unstructured bodies are withheld in newly captured HAR previews and exports. Searches still use original bounded data, so matching values can be hidden. Redaction is not full anonymization; review before sharing.",
                "HAR redaction: Off. Newly captured HAR previews and exports retain original matching text; body size and decoding limits still apply. Review sensitive data before sharing.") & vbCrLf & vbCrLf &
                "Save and run a new scan to apply this choice. Existing results keep the redaction setting used when they were collected."
        End Sub

        Private Sub ApplyTheme(isDarkMode As Boolean)
            NativeCaptionTheme.Apply(Me, isDarkMode)
            If isDarkMode Then
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
            End If
            BeaconThemePalette.ApplyButtons(Resources, _workingSettings.Theme = AppTheme.Beacon, isDarkMode)
        End Sub

        Private Sub LoadControls()
            Theme_cmb.SelectedValue = _workingSettings.Theme.ToString()
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
            EvtxResourcePolicy_cmb.SelectedItem = EvtxResourcePolicy.OfflineOnly
            EvtxIds_txt.Text = _workingSettings.EvtxEventIds
            EvtxProvider_cmb.Text = _workingSettings.EvtxProvider
            EvtxLevels_txt.Text = _workingSettings.EvtxLevels
            EvtxFrom_txt.Text = _workingSettings.EvtxFromUtc
            EvtxTo_txt.Text = _workingSettings.EvtxToUtc
            HarMatches_txt.Text = _workingSettings.HarMaximumMatches.ToString()
            _harFilters.LoadSettings(_workingSettings)
            HarBodySize_txt.Text = _workingSettings.MaximumHarBodySizeMb.ToString()
            DecodeBase64_chk.IsChecked = _workingSettings.DecodeBase64HarBodies
            RedactHar_chk.IsChecked = _workingSettings.RedactSensitiveHarData
            UpdateHarPrivacyHelp(Nothing, Nothing)

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
                _workingSettings.Theme = [Enum].Parse(Of AppTheme)(CStr(Theme_cmb.SelectedValue))
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
                Dim filter As New EvtxFilter(EvtxIds_txt.Text, EvtxProvider_cmb.Text, EvtxLevels_txt.Text, EvtxFrom_txt.Text, EvtxTo_txt.Text)
                _workingSettings.EvtxEventIds = EvtxIds_txt.Text.Trim()
                _workingSettings.EvtxProvider = EvtxProvider_cmb.Text.Trim()
                _workingSettings.EvtxLevels = EvtxLevels_txt.Text.Trim()
                _workingSettings.EvtxFromUtc = EvtxFrom_txt.Text.Trim()
                _workingSettings.EvtxToUtc = EvtxTo_txt.Text.Trim()
                _workingSettings.ShowRawXmlWhenMessageUnavailable = RawXmlFallback_chk.IsChecked.GetValueOrDefault()
                _workingSettings.EvtxMessageResourceBehavior = DirectCast(EvtxResourcePolicy_cmb.SelectedItem, EvtxResourcePolicy)
                _workingSettings.HarMaximumMatches = ParseInteger(HarMatches_txt, "Maximum HAR matches")
                _harFilters.SaveSettings(_workingSettings)
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

        Public Sub SetProviderSuggestions(providers As IEnumerable(Of String))
            Dim typed = EvtxProvider_cmb.Text
            EvtxProvider_cmb.ItemsSource = New String() {""}.Concat(providers.Where(Function(provider) Not String.IsNullOrWhiteSpace(provider)).
                Distinct(StringComparer.OrdinalIgnoreCase).OrderBy(Function(provider) provider, StringComparer.OrdinalIgnoreCase)).ToArray()
            EvtxProvider_cmb.Text = typed
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
