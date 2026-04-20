Imports System.IO
Imports System.IO.Compression
Imports System.Threading
Imports System.Threading.Tasks
Imports System.Collections.ObjectModel
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Documents
Imports System.Windows.Input
Imports System.Windows.Media
Imports Microsoft.Win32
Imports System.Diagnostics
Imports System.Diagnostics.Eventing.Reader
Imports WinForms = System.Windows.Forms
Imports System.Text.Json
Imports SharpCompress.Archives
Imports SharpCompress.Readers
Imports SharpCompress.Common

' ============================================================================
' Beacon: Find in Files - Advanced Log Search Utility
' ============================================================================
' This software was developed by Loaiza (luislo@microsoft)
' 
' Purpose: Searches for text patterns across multiple file types including:
'   - Plain text files (.txt, .log, .json, .xml, .csv, .html, .reg, .ini, .cfg, .config, .nfo)
'   - Windows Event Log files (.evtx) with graceful handling of missing DLLs
'   - HAR (HTTP Archive) log files
'   - Compressed archives (ZIP, 7Z, RAR, TAR, GZ, BZ2, CAB - up to 1 level nesting)
'
' Features:
'   - Real-time text highlighting in RichTextBox preview
'   - WebView2 rendering for HTML/XML/JSON files
'   - Light/Dark theme toggle matching Windows system preferences
'   - Keyboard shortcuts (F3/Shift+F3, Ctrl+Down, Ctrl+R, Enter, Esc)
'   - Cancellable async scanning with accurate progress tracking
'   - Case-sensitive/insensitive search with exact match (word boundary) support
'   - EVTX event navigation with warning messages for missing message DLLs
'   - HAR request/response preview with navigation controls
'   - Optimized performance: async UI updates, minimal LOH pressure
' ============================================================================

Namespace Beacon

    Partial Public Class MainWindow
        Inherits Window

#Region "Models"
        ' ====================================================================
        ' Data models for representing search results and event information
        ' ====================================================================

        ''' <summary>
        ''' Identifies the type of file/entry that matched the search.
        ''' Used to determine how to load and display preview content.
        ''' </summary>
        Private Enum HitKind
            DiskTextFile        ' Regular text file on disk
            ZipTextEntry        ' Text file inside a ZIP archive
            CabTextEntry        ' Text file inside a CAB archive
            EvtxFileOnDisk      ' Event log file on disk
            EvtxFromZipTemp     ' Event log extracted from ZIP to temp location
            HarFileOnDisk       ' HAR log file on disk
            HarFromZip          ' HAR log file from ZIP
            EvtxFromCabTemp     ' Event log extracted from CAB to temp location
        End Enum

        ''' <summary>
        ''' Represents a single Windows Event Log entry that matches the search term.
        ''' Contains formatted event details for display in the preview pane.
        ''' </summary>
        Private Class EventSummary
            Public Property Level As String             ' Error, Warning, Information, etc.
            Public Property Provider As String          ' Event source (e.g., "Application")
            Public Property EventId As Integer          ' Numeric event identifier
            Public Property TimeCreated As DateTime?    ' Event timestamp
            Public Property Message As String           ' Formatted event description
        End Class

        ''' <summary>
        ''' Represents a single HAR (HTTP Archive) request entry that matches the search term.
        ''' Contains formatted HTTP request/response details for display in the preview pane.
        ''' </summary>
        Private Class HarRequest
            Public Property Method As String            ' HTTP method (GET, POST, etc.)
            Public Property Url As String               ' Request URL
            Public Property StatusCode As Integer       ' Response status code
            Public Property StatusText As String        ' Response status text
            Public Property StartedDateTime As DateTime? ' Request start time
            Public Property Time As Double              ' Time in milliseconds
            Public Property RequestHeaders As String    ' Formatted request headers
            Public Property ResponseHeaders As String   ' Formatted response headers
            Public Property RequestBody As String       ' Request body/payload
            Public Property ResponseBody As String      ' Response body content
            Public Property ServerIpAddress As String   ' Server IP
        End Class

        ''' <summary>
        ''' Represents a file or EVTX log that contains at least one match.
        ''' Stores file location info and EVTX navigation state.
        ''' </summary>
        Private Class SearchHit
            Public Property DisplayName As String       ' User-friendly name shown in results list
            Public Property Kind As HitKind             ' Type of hit (text file, evtx, etc.)

            ' --- Disk file / EVTX properties ---
            Public Property FilePath As String          ' Full path to file on disk

            ' --- ZIP archive properties ---
            Public Property ZipPath As String           ' Path to ZIP file
            Public Property ZipEntryName As String      ' Entry name within ZIP

            ' --- Temp EVTX properties ---
            Public Property TempEvtxPath As String      ' Temporary file path for EVTX extracted from ZIP

            ' --- EVTX navigation state ---
            Public Property MatchingEvents As New List(Of EventSummary)  ' All matching events in this EVTX
            Public Property CurrentEventIndex As Integer = -1            ' Currently displayed event index

            ' --- HAR navigation state ---
            Public Property MatchingRequests As New List(Of HarRequest)  ' All matching requests in this HAR
            Public Property CurrentRequestIndex As Integer = -1          ' Currently displayed request index
        End Class

#End Region

#Region "Fields"
        ' ====================================================================
        ' Application state and configuration
        ' ====================================================================

        ''' <summary>Collection of all search results, bound to Results_lst</summary>
        Private ReadOnly _hits As New ObservableCollection(Of SearchHit)

        ''' <summary>Cancellation token source for aborting scan operations</summary>
        Private _scanCts As CancellationTokenSource

        ''' <summary>Flag indicating whether a scan is currently in progress</summary>
        Private _isScanning As Boolean = False

        ''' <summary>Current character position in text preview for F3 navigation</summary>
        Private _currentTextFindStart As Integer = 0

        ''' <summary>Current character position in EVTX message for F3 navigation</summary>
        Private _currentEventMessageMatchIndex As Integer = 0

        ''' <summary>Current character position in HAR content for F3 navigation</summary>
        Private _currentHarMatchIndex As Integer = 0

        ''' <summary>Root folder of current scan, used for relative path display</summary>
        Private _scanRootFolder As String = ""

        ''' <summary>Last scanned path, used to detect re-runs on same source</summary>
        Private _lastScanPath As String = ""

        ''' <summary>File extensions recognized as searchable text files</summary>
        Private ReadOnly _supportedTextExt As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase) From {
            ".txt", ".log", ".json", ".xml", ".csv", ".html", ".reg", ".ini", ".cfg", ".config", ".nfo"
        }

        ''' <summary>File extensions recognized as HAR (HTTP Archive) files</summary>
        Private ReadOnly _supportedHarExt As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase) From {
            ".har"
        }

        ''' <summary>File extensions recognized as Windows Event Logs</summary>
        Private ReadOnly _supportedEvtxExt As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase) From {
            ".evtx"
        }

        ''' <summary>File extensions recognized as archive files (for recursive scanning)</summary>
        Private ReadOnly _supportedArchiveExt As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase) From {
            ".zip", ".7z", ".rar", ".tar", ".gz", ".bz2", ".tgz", ".tbz", ".tbz2", ".tar.gz", ".tar.bz2", ".tar.xz", ".txz", ".cab"
        }

        ''' <summary>Tracks temporary EVTX files extracted from archives for cleanup</summary>
        Private ReadOnly _tempToDelete As New List(Of String)

        ' --- WebView2 state for HTML/XML preview ---
        ''' <summary>Flag indicating if WebView2 is initialized</summary>
        Private _webViewInitialized As Boolean = False

        ''' <summary>Current search term index for WebView2 F3 navigation</summary>
        Private _currentWebMatchIndex As Integer = 0

        ''' <summary>Total matches found in current WebView2 content</summary>
        Private _totalWebMatches As Integer = 0

        ''' <summary>WebView2 user data folder path (for cleanup on exit)</summary>
        Private _webView2DataFolder As String = ""

        ' --- UI throttling for "Scanning file: ..." label ---
        ''' <summary>Stopwatch for measuring time between label updates</summary>
        Private ReadOnly _fileLabelStopwatch As Stopwatch = Stopwatch.StartNew()

        ''' <summary>Last timestamp when CurrentFile_lbl was updated</summary>
        Private _lastFileLabelUpdateMs As Long = 0

        ''' <summary>Pending filename to display (queued until throttle interval passes)</summary>
        Private _pendingFileLabel As String = ""

        ''' <summary>Timer that periodically flushes pending label updates</summary>
        Private _fileLabelTimer As Threading.DispatcherTimer

        ''' <summary>Minimum milliseconds between CurrentFile_lbl updates</summary>
        Private Const FILE_LABEL_INTERVAL_MS As Integer = 120

        ' --- Elapsed time tracking ---
        ''' <summary>Stopwatch for measuring elapsed scan time</summary>
        Private ReadOnly _scanElapsedStopwatch As New Stopwatch()

        ''' <summary>Timer that updates elapsed time display every second</summary>
        Private _elapsedTimeTimer As Threading.DispatcherTimer

        ' --- File progress tracking ---
        ''' <summary>Total number of files to scan</summary>
        Private _totalFilesToScan As Integer = 0

        ''' <summary>Number of files scanned so far</summary>
        Private _filesScanned As Integer = 0

        ''' <summary>Current theme mode (True = Dark, False = Light)</summary>
        Private _isDarkMode As Boolean = False

#End Region

#Region "Init"
        ' ====================================================================
        ' Window initialization and event handler registration
        ' ====================================================================

        ''' <summary>
        ''' Constructor: Initializes UI controls, wires event handlers, and sets initial state.
        ''' </summary>
        Public Sub New()
            Debug.WriteLine("========================================")
            Debug.WriteLine("MainWindow constructor started")
            Debug.WriteLine("========================================")

            InitializeComponent()

            ' Bind results collection to ListBox
            Results_lst.ItemsSource = _hits
            Results_lst.DisplayMemberPath = NameOf(SearchHit.DisplayName)

            ' Wire button click handlers
            AddHandler BrowseFolder_btn.Click, AddressOf BrowseFolder_btn_Click
            AddHandler BrowseZip_btn.Click, AddressOf BrowseZip_btn_Click
            AddHandler Scan_btn.Click, AddressOf Scan_btn_Click
            AddHandler Reset_btn.Click, AddressOf Reset_btn_Click

            ' Wire navigation and selection handlers
            AddHandler Results_lst.SelectionChanged, AddressOf Results_lst_SelectionChanged
            AddHandler NextFile_btn.Click, AddressOf NextFile_btn_Click
            AddHandler FindNext_btn.Click, AddressOf FindNext_btn_Click
            AddHandler FindNextEvent_btn.Click, AddressOf FindNextEvent_btn_Click
            AddHandler FindPreviousEvent_btn.Click, AddressOf FindPreviousEvent_btn_Click
            AddHandler FindNextHarRequest_btn.Click, AddressOf FindNextHarRequest_btn_Click
            AddHandler FindPreviousHarRequest_btn.Click, AddressOf FindPreviousHarRequest_btn_Click

            ' Wire theme toggle handler
            AddHandler ThemeToggle_btn.Click, AddressOf ThemeToggle_btn_Click

            ' Wire input change handlers for button state updates
            AddHandler Path_txt.TextChanged, AddressOf AnyInputChanged
            AddHandler Search_txt.TextChanged, AddressOf AnyInputChanged
            AddHandler ExactMatch_chk.Checked, AddressOf AnyInputChanged
            AddHandler ExactMatch_chk.Unchecked, AddressOf AnyInputChanged

            ' Register global keyboard shortcut handler
            AddHandler Me.PreviewKeyDown, AddressOf MainWindow_PreviewKeyDown

            ' Register window closing handler for cleanup
            AddHandler Me.Closing, AddressOf MainWindow_Closing

            ' Setup UI throttling timer for filename display during scans
            _fileLabelTimer = New Threading.DispatcherTimer()
            _fileLabelTimer.Interval = TimeSpan.FromMilliseconds(FILE_LABEL_INTERVAL_MS)
            AddHandler _fileLabelTimer.Tick, AddressOf FileLabelTimer_Tick

            ' Setup elapsed time timer (updates every second)
            _elapsedTimeTimer = New Threading.DispatcherTimer()
            _elapsedTimeTimer.Interval = TimeSpan.FromSeconds(1)
            AddHandler _elapsedTimeTimer.Tick, AddressOf ElapsedTimeTimer_Tick

            ' Initialize UI to ready state
            Status("Ready")
            ShowProgress(False)
            SetCurrentFileDisplayImmediate("")
            TimeElapsed_lbl.Visibility = Visibility.Collapsed

            Scan_btn.Content = "Scan (Enter)"
            Scan_btn.IsEnabled = False

            Reset_btn.IsEnabled = False

            ' Hide navigation buttons until results are available
            NextFile_btn.Visibility = Visibility.Collapsed
            FindNext_btn.Visibility = Visibility.Collapsed
            FindNextEvent_btn.Visibility = Visibility.Collapsed
            FindNextHarRequest_btn.Visibility = Visibility.Collapsed
            FindPreviousHarRequest_btn.Visibility = Visibility.Collapsed

            ShowTextPreviewMode()
            ClearTextPreview()

            ' Initialize dark mode based on system preferences
            InitializeTheme()

            ' Initialize WebView2 after window is fully loaded (control must be in visual tree)
            AddHandler Me.Loaded, AddressOf MainWindow_Loaded
        End Sub

        ''' <summary>
        ''' Initializes theme based on Windows system preferences
        ''' </summary>
        Private Sub InitializeTheme()
            Try
                ' Detect Windows theme using Registry
                Dim isDarkModeEnabled = IsWindowsDarkModeEnabled()

                ' Apply the detected theme
                If isDarkModeEnabled Then
                    ApplyDarkTheme()
                Else
                    ApplyLightTheme()
                End If

                Debug.WriteLine($"Theme initialized: {If(_isDarkMode, "Dark", "Light")} mode")
            Catch ex As Exception
                Debug.WriteLine($"Error initializing theme, using Light mode: {ex.Message}")
                ApplyLightTheme()
            End Try
        End Sub

        ''' <summary>
        ''' Detects if Windows is using dark mode via Registry
        ''' </summary>
        Private Function IsWindowsDarkModeEnabled() As Boolean
            Try
                Using key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey("Software\Microsoft\Windows\CurrentVersion\Themes\Personalize")
                    If key IsNot Nothing Then
                        Dim value = key.GetValue("AppsUseLightTheme")
                        If value IsNot Nothing Then
                            ' Value = 0 means Dark Mode, 1 means Light Mode
                            Return CInt(value) = 0
                        End If
                    End If
                End Using
            Catch
                ' If registry read fails, default to Light mode
            End Try

            Return False
        End Function

        ''' <summary>
        ''' Handles theme toggle button click - switches between Light and Dark themes
        ''' </summary>
        Private Sub ThemeToggle_btn_Click(sender As Object, e As RoutedEventArgs)
            Try
                If _isDarkMode Then
                    ApplyLightTheme()
                Else
                    ApplyDarkTheme()
                End If

                Debug.WriteLine($"Theme manually switched to: {If(_isDarkMode, "Dark", "Light")} mode")
            Catch ex As Exception
                Debug.WriteLine($"Error switching theme: {ex.Message}")
            End Try
        End Sub

        ''' <summary>
        ''' Applies dark theme colors to the application
        ''' </summary>
        Private Sub ApplyDarkTheme()
            _isDarkMode = True

            ' Update theme button
            ThemeToggle_btn.Content = "☀️"
            ThemeToggle_btn.ToolTip = "Switch to Light theme"

            ' Apply dark theme color palette
            Resources("WindowBackgroundBrush") = New SolidColorBrush(Color.FromRgb(&H20, &H20, &H20))    ' #202020
            Resources("CardBackgroundBrush") = New SolidColorBrush(Color.FromRgb(&H2B, &H2B, &H2B))      ' #2B2B2B
            Resources("CardBorderBrush") = New SolidColorBrush(Color.FromRgb(&H3F, &H3F, &H3F))         ' #3F3F3F
            Resources("StatusBarBackgroundBrush") = New SolidColorBrush(Color.FromRgb(&H1A, &H1A, &H1A)) ' #1A1A1A
            Resources("TextPrimaryBrush") = New SolidColorBrush(Color.FromRgb(&HE0, &HE0, &HE0))        ' #E0E0E0
            Resources("TextSecondaryBrush") = New SolidColorBrush(Color.FromRgb(&HB0, &HB0, &HB0))      ' #B0B0B0
            Resources("TextTertiaryBrush") = New SolidColorBrush(Color.FromRgb(&H80, &H80, &H80))       ' #808080
            Resources("AccentBrush") = New SolidColorBrush(Color.FromRgb(&H60, &HCF, &HFF))             ' #60CFFF (Light Blue)
            Resources("ButtonBackgroundBrush") = New SolidColorBrush(Color.FromRgb(&H3A, &H3A, &H3A))   ' #3A3A3A
            Resources("ButtonHoverBrush") = New SolidColorBrush(Color.FromRgb(&H45, &H45, &H45))        ' #454545
            Resources("ButtonPressedBrush") = New SolidColorBrush(Color.FromRgb(&H50, &H50, &H50))      ' #505050
            Resources("InputBackgroundBrush") = New SolidColorBrush(Color.FromRgb(&H2B, &H2B, &H2B))    ' #2B2B2B
            Resources("InputBorderBrush") = New SolidColorBrush(Color.FromRgb(&H50, &H50, &H50))        ' #505050
            Resources("CodeBackgroundBrush") = New SolidColorBrush(Color.FromRgb(&H1E, &H1E, &H1E))     ' #1E1E1E

            ' Update event preview text colors for dark mode
            EventLevel_txt.Foreground = New SolidColorBrush(Color.FromRgb(&HFF, &H60, &H60)) ' Lighter red for dark mode
            EventId_txt.Foreground = New SolidColorBrush(Color.FromRgb(&HE0, &HE0, &HE0))    ' Light gray
            EventProvider_txt.Foreground = New SolidColorBrush(Color.FromRgb(&HE0, &HE0, &HE0)) ' Light gray
            EventTime_txt.Foreground = New SolidColorBrush(Color.FromRgb(&HB0, &HB0, &HB0))  ' Medium gray
            EventMessage_txt.Foreground = New SolidColorBrush(Color.FromRgb(&HE0, &HE0, &HE0)) ' Light gray
            EventCounter_lbl.Foreground = New SolidColorBrush(Color.FromRgb(&HB0, &HB0, &HB0)) ' Medium gray
            EventMatchCounter_lbl.Foreground = New SolidColorBrush(Color.FromRgb(&HB0, &HB0, &HB0)) ' Medium gray
        End Sub

        ''' <summary>
        ''' Applies light theme colors to the application
        ''' </summary>
        Private Sub ApplyLightTheme()
            _isDarkMode = False

            ' Update theme button
            ThemeToggle_btn.Content = "🌙"
            ThemeToggle_btn.ToolTip = "Switch to Dark theme"

            ' Apply light theme color palette (original colors)
            Resources("WindowBackgroundBrush") = New SolidColorBrush(Color.FromRgb(&HF3, &HF3, &HF3))    ' #F3F3F3
            Resources("CardBackgroundBrush") = New SolidColorBrush(Colors.White)
            Resources("CardBorderBrush") = New SolidColorBrush(Color.FromRgb(&HE0, &HE0, &HE0))         ' #E0E0E0
            Resources("StatusBarBackgroundBrush") = New SolidColorBrush(Color.FromRgb(&HE6, &HE6, &HE6)) ' #E6E6E6
            Resources("TextPrimaryBrush") = New SolidColorBrush(Color.FromRgb(&H20, &H20, &H20))        ' #202020
            Resources("TextSecondaryBrush") = New SolidColorBrush(Color.FromRgb(&H66, &H66, &H66))      ' #666666
            Resources("TextTertiaryBrush") = New SolidColorBrush(Color.FromRgb(&H88, &H88, &H88))       ' #888888
            Resources("AccentBrush") = New SolidColorBrush(Color.FromRgb(&H0, &H66, &HCC))              ' #0066CC
            Resources("ButtonBackgroundBrush") = New SolidColorBrush(Color.FromRgb(&HF0, &HF0, &HF0))   ' #F0F0F0
            Resources("ButtonHoverBrush") = New SolidColorBrush(Color.FromRgb(&HE0, &HE0, &HE0))        ' #E0E0E0
            Resources("ButtonPressedBrush") = New SolidColorBrush(Color.FromRgb(&HD0, &HD0, &HD0))      ' #D0D0D0
            Resources("InputBackgroundBrush") = New SolidColorBrush(Colors.White)
            Resources("InputBorderBrush") = New SolidColorBrush(Color.FromRgb(&HCC, &HCC, &HCC))        ' #CCCCCC
            Resources("CodeBackgroundBrush") = New SolidColorBrush(Color.FromRgb(&HFA, &HFA, &HFA))     ' #FAFAFA

            ' Restore event preview text colors for light mode
            EventLevel_txt.Foreground = New SolidColorBrush(Color.FromRgb(&HC0, &H0, &H0)) ' Original dark red
            EventId_txt.Foreground = New SolidColorBrush(Color.FromRgb(&H20, &H20, &H20))  ' Dark gray
            EventProvider_txt.Foreground = New SolidColorBrush(Color.FromRgb(&H20, &H20, &H20)) ' Dark gray
            EventTime_txt.Foreground = New SolidColorBrush(Color.FromRgb(&H66, &H66, &H66)) ' Medium gray
            EventMessage_txt.Foreground = New SolidColorBrush(Color.FromRgb(&H20, &H20, &H20)) ' Dark gray
            EventCounter_lbl.Foreground = New SolidColorBrush(Color.FromRgb(&H66, &H66, &H66)) ' Medium gray
            EventMatchCounter_lbl.Foreground = New SolidColorBrush(Color.FromRgb(&H66, &H66, &H66)) ' Medium gray
        End Sub

        ''' <summary>
        ''' Handles feedback hyperlink click - opens GitHub discussions in default browser
        ''' </summary>
        Private Sub Feedback_lnk_RequestNavigate(sender As Object, e As System.Windows.Navigation.RequestNavigateEventArgs)
            Try
                Process.Start(New ProcessStartInfo(e.Uri.AbsoluteUri) With {.UseShellExecute = True})
                e.Handled = True
            Catch ex As Exception
                MessageBox.Show($"Could not open browser: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Warning)
            End Try
        End Sub

        ''' <summary>
        ''' Called when window is fully loaded - safe time to initialize WebView2
        ''' </summary>
        Private Sub MainWindow_Loaded(sender As Object, e As RoutedEventArgs)
            Debug.WriteLine("========================================")
            Debug.WriteLine("MainWindow_Loaded event fired!")
            Debug.WriteLine("========================================")

            ' Clean up any leftover WebView2 folders from previous sessions
            CleanupOldWebView2Folders()

            ' Initialize WebView2 asynchronously - control is now in visual tree
            InitializeWebView2Async()
        End Sub

        ''' <summary>
        ''' Initializes WebView2 control asynchronously (non-blocking)
        ''' Uses default environment if already initialized, otherwise creates custom temp folder
        ''' </summary>
        Private Async Sub InitializeWebView2Async()
            Try
                Debug.WriteLine("========================================")
                Debug.WriteLine("Starting WebView2 initialization...")
                Debug.WriteLine($"WebPreview_wv2 is null: {WebPreview_wv2 Is Nothing}")

                If WebPreview_wv2 Is Nothing Then
                    Debug.WriteLine("✗ WebView2 control is null!")
                    _webViewInitialized = False
                    Return
                End If

                ' Check if WebView2 is already initialized (auto-initialized by WPF)
                If WebPreview_wv2.CoreWebView2 IsNot Nothing Then
                    Debug.WriteLine("✓ WebView2 already initialized by WPF (using default environment)")

                    ' Get the data folder path from the existing environment
                    Try
                        _webView2DataFolder = WebPreview_wv2.CoreWebView2.Environment.UserDataFolder
                        Debug.WriteLine($"Using existing data folder: {_webView2DataFolder}")
                    Catch
                        Debug.WriteLine("Could not get existing data folder path")
                    End Try

                    ' Configure settings on already-initialized WebView2
                    With WebPreview_wv2.CoreWebView2.Settings
                        .AreDefaultContextMenusEnabled = False
                        .IsScriptEnabled = True
                        .AreDevToolsEnabled = False
                        .IsWebMessageEnabled = True
                        .IsStatusBarEnabled = False
                    End With

                    _webViewInitialized = True
                    Debug.WriteLine("✓✓✓ Using existing WebView2 initialization ✓✓✓")
                    Debug.WriteLine("========================================")
                    Return
                End If

                ' Not initialized yet - try to initialize with default environment (let WPF handle location)
                Debug.WriteLine("Initializing WebView2 with default environment...")
                Await WebPreview_wv2.EnsureCoreWebView2Async(Nothing)
                Debug.WriteLine($"✓ WebView2 CoreWebView2 initialized: {WebPreview_wv2.CoreWebView2 IsNot Nothing}")

                ' Get the data folder path
                Try
                    _webView2DataFolder = WebPreview_wv2.CoreWebView2.Environment.UserDataFolder
                    Debug.WriteLine($"Data folder: {_webView2DataFolder}")
                Catch
                    Debug.WriteLine("Could not get data folder path")
                End Try

                ' Configure WebView2 settings
                With WebPreview_wv2.CoreWebView2.Settings
                    .AreDefaultContextMenusEnabled = False  ' Disable right-click menu for security
                    .IsScriptEnabled = True                  ' Enable JavaScript (required for interactive HTML)
                    .AreDevToolsEnabled = False              ' Disable F12 dev tools
                    .IsWebMessageEnabled = True              ' Allow script communication
                    .IsStatusBarEnabled = False              ' Hide status bar
                End With

                _webViewInitialized = True
                Debug.WriteLine("✓✓✓ WebView2 initialized successfully ✓✓✓")
                Debug.WriteLine("========================================")

            Catch ex As Exception
                _webViewInitialized = False
                Debug.WriteLine("========================================")
                Debug.WriteLine($"✗✗✗ WebView2 initialization FAILED ✗✗✗")
                Debug.WriteLine($"Exception Type: {ex.GetType().FullName}")
                Debug.WriteLine($"Exception Message: {ex.Message}")
                Debug.WriteLine($"Stack Trace: {ex.StackTrace}")
                Debug.WriteLine("========================================")
            End Try
        End Sub

        ''' <summary>
        ''' Ensures WebView2 is initialized before use (synchronous check with retry)
        ''' Uses default environment if already initialized, otherwise initializes with WPF default
        ''' </summary>
        Private Async Function EnsureWebView2InitializedAsync() As Task(Of Boolean)
            If _webViewInitialized Then
                Debug.WriteLine("WebView2 already initialized ✓")
                Return True
            End If

            Try
                Debug.WriteLine("========================================")
                Debug.WriteLine("WebView2 not initialized yet, initializing now...")
                Debug.WriteLine($"WebPreview_wv2 is null: {WebPreview_wv2 Is Nothing}")

                If WebPreview_wv2 Is Nothing Then
                    Debug.WriteLine("✗ WebView2 control is null!")
                    Return False
                End If

                ' Check if WebView2 is already initialized (auto-initialized by WPF)
                If WebPreview_wv2.CoreWebView2 IsNot Nothing Then
                    Debug.WriteLine("✓ WebView2 already initialized by WPF (using default environment)")

                    ' Get the data folder path from the existing environment
                    Try
                        _webView2DataFolder = WebPreview_wv2.CoreWebView2.Environment.UserDataFolder
                        Debug.WriteLine($"Using existing data folder: {_webView2DataFolder}")
                    Catch
                        Debug.WriteLine("Could not get existing data folder path")
                    End Try

                    ' Configure settings on already-initialized WebView2
                    With WebPreview_wv2.CoreWebView2.Settings
                        .AreDefaultContextMenusEnabled = False
                        .IsScriptEnabled = True
                        .AreDevToolsEnabled = False
                        .IsWebMessageEnabled = True
                        .IsStatusBarEnabled = False
                    End With

                    _webViewInitialized = True
                    Debug.WriteLine("✓✓✓ Using existing WebView2 initialization ✓✓✓")
                    Debug.WriteLine("========================================")
                    Return True
                End If

                ' Not initialized yet - initialize with default environment (let WPF handle location)
                Debug.WriteLine("Initializing WebView2 with default environment...")
                Await WebPreview_wv2.EnsureCoreWebView2Async(Nothing)
                Debug.WriteLine($"✓ WebView2 CoreWebView2 initialized: {WebPreview_wv2.CoreWebView2 IsNot Nothing}")

                ' Get the data folder path
                Try
                    _webView2DataFolder = WebPreview_wv2.CoreWebView2.Environment.UserDataFolder
                    Debug.WriteLine($"Data folder: {_webView2DataFolder}")
                Catch
                    Debug.WriteLine("Could not get data folder path")
                End Try

                ' Configure WebView2 settings
                With WebPreview_wv2.CoreWebView2.Settings
                    .AreDefaultContextMenusEnabled = False  ' Disable right-click menu for security
                    .IsScriptEnabled = True                  ' Enable JavaScript (required for interactive HTML)
                    .AreDevToolsEnabled = False              ' Disable F12 dev tools
                    .IsWebMessageEnabled = True              ' Allow script communication
                    .IsStatusBarEnabled = False              ' Hide status bar
                End With

                _webViewInitialized = True
                Debug.WriteLine("✓✓✓ WebView2 initialized on-demand successfully ✓✓✓")
                Debug.WriteLine("========================================")
                Return True
            Catch ex As Exception
                Debug.WriteLine("========================================")
                Debug.WriteLine($"✗✗✗ WebView2 initialization FAILED ✗✗✗")
                Debug.WriteLine($"Exception Type: {ex.GetType().FullName}")
                Debug.WriteLine($"Exception Message: {ex.Message}")
                Debug.WriteLine($"Stack Trace: {ex.StackTrace}")
                Debug.WriteLine("========================================")
                _webViewInitialized = False
                Return False
            End Try
        End Function

        ''' <summary>
        ''' Flag to prevent recursive closing calls
        ''' </summary>
        Private _isClosing As Boolean = False

        ''' <summary>
        ''' Handles window closing event - cleanup temp files and attempt quick WebView2 folder cleanup
        ''' Does NOT block application exit - leftover folders are cleaned on next startup
        ''' </summary>
        Private Sub MainWindow_Closing(sender As Object, e As ComponentModel.CancelEventArgs)
            ' If we're already cleaning up, allow the close to proceed
            If _isClosing Then
                Return
            End If

            ' Cancel the close event so we can finish cleanup first
            e.Cancel = True
            _isClosing = True

            Try
                ' Dispose WebView2 to release file locks
                If _webViewInitialized AndAlso WebPreview_wv2 IsNot Nothing Then
                    Try
                        Debug.WriteLine("Disposing WebView2...")

                        ' Navigate to blank page to release current page resources
                        If WebPreview_wv2.CoreWebView2 IsNot Nothing Then
                            Try
                                WebPreview_wv2.NavigateToString("about:blank")
                                WebPreview_wv2.CoreWebView2.Stop()
                            Catch
                                ' Ignore errors stopping navigation
                            End Try
                        End If

                        ' Clear source and dispose
                        WebPreview_wv2.Source = Nothing
                        WebPreview_wv2.Dispose()
                        Debug.WriteLine("✓ WebView2 disposed")

                    Catch ex As Exception
                        Debug.WriteLine($"Warning: Error disposing WebView2: {ex.Message}")
                    End Try
                End If

                ' Cleanup temp EVTX files
                CleanupTemp()

                ' QUICK attempt to delete WebView2 folder (don't block shutdown for long)
                If Not String.IsNullOrEmpty(_webView2DataFolder) AndAlso Directory.Exists(_webView2DataFolder) Then
                    Debug.WriteLine($"Attempting quick cleanup of WebView2 folder: {_webView2DataFolder}")

                    ' Single quick attempt with minimal wait
                    Try
                        System.Threading.Thread.Sleep(300)
                        Directory.Delete(_webView2DataFolder, True)
                        Debug.WriteLine($"✓ Successfully deleted WebView2 folder")
                    Catch
                        Debug.WriteLine($"⚠ WebView2 folder still in use - will be cleaned up on next app launch")
                    End Try
                End If

            Catch ex As Exception
                Debug.WriteLine($"✗ Error during cleanup: {ex.Message}")
            Finally
                ' Shut down quickly - don't keep user waiting
                Debug.WriteLine("Shutting down application...")
                Application.Current.Shutdown()
            End Try
        End Sub

        ''' <summary>
        ''' Cleans up old WebView2 folders from previous application sessions
        ''' This runs on startup when browser processes aren't running, making cleanup reliable
        ''' </summary>
        Private Sub CleanupOldWebView2Folders()
            Try
                Debug.WriteLine("========================================")
                Debug.WriteLine("Checking for old WebView2 folders to cleanup...")

                ' Look for WebView2 folders in common locations
                Dim possibleLocations As New List(Of String)

                ' Check user's AppData\Local
                Dim localAppData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData)
                If Not String.IsNullOrEmpty(localAppData) Then
                    possibleLocations.Add(Path.Combine(localAppData, "Microsoft", "Edge", "User Data"))
                    possibleLocations.Add(localAppData)
                End If

                ' Check temp folder
                possibleLocations.Add(Path.GetTempPath())

                Dim foldersDeleted = 0
                For Each location In possibleLocations
                    If Not Directory.Exists(location) Then Continue For

                    Try
                        ' Look for folders matching WebView2 pattern
                        Dim webViewFolders = Directory.GetDirectories(location, "*WebView2*", SearchOption.TopDirectoryOnly)

                        For Each folder In webViewFolders
                            ' Check if it's from our app (contains "Beacon" or is old enough to be from previous session)
                            Dim folderInfo As New DirectoryInfo(folder)
                            Dim isOldEnough = (DateTime.Now - folderInfo.LastWriteTime).TotalMinutes > 5 ' Older than 5 minutes
                            Dim isBeaconFolder = folder.Contains("Beacon", StringComparison.OrdinalIgnoreCase)

                            If isBeaconFolder OrElse isOldEnough Then
                                Try
                                    Debug.WriteLine($"Deleting old WebView2 folder: {folder}")
                                    Directory.Delete(folder, True)
                                    foldersDeleted += 1
                                    Debug.WriteLine($"✓ Deleted")
                                Catch ex As Exception
                                    Debug.WriteLine($"Could not delete {folder}: {ex.Message}")
                                End Try
                            End If
                        Next
                    Catch
                        ' Continue checking other locations
                    End Try
                Next

                If foldersDeleted > 0 Then
                    Debug.WriteLine($"✓✓✓ Cleaned up {foldersDeleted} old WebView2 folder(s)")
                Else
                    Debug.WriteLine($"No old WebView2 folders found")
                End If
                Debug.WriteLine("========================================")

            Catch ex As Exception
                Debug.WriteLine($"Error during old folder cleanup: {ex.Message}")
            End Try
        End Sub

#End Region

#Region "Keyboard"
        ' ====================================================================
        ' Global keyboard shortcut handling
        ' ====================================================================

        ''' <summary>
        ''' Handles application-wide keyboard shortcuts:
        ''' - Esc: Cancel scan
        ''' - Enter: Start scan
        ''' - F3 / Shift+F3: Find next/previous match
        ''' - Ctrl+Down: Next file
        ''' - Ctrl+R: Reset
        ''' </summary>
        Private Sub MainWindow_PreviewKeyDown(sender As Object, e As KeyEventArgs)

            ' Esc cancels active scans
            If e.Key = Key.Escape AndAlso _isScanning Then
                CancelScan()
                e.Handled = True
                Return
            End If

            ' Enter starts scan when ready (not already scanning)
            If e.Key = Key.Enter AndAlso (Not _isScanning) AndAlso Scan_btn.IsEnabled Then
                StartScan()
                e.Handled = True
                Return
            End If


            ' F3 / Shift+F3: Find next/previous match in preview
            If e.Key = Key.F3 Then

                ' EVTX preview has priority (navigate between events)
                If EventPreview_grp.Visibility = Visibility.Visible AndAlso
       FindNextEvent_btn.Visibility = Visibility.Visible Then

                    If Keyboard.Modifiers.HasFlag(ModifierKeys.Shift) Then
                        FindPreviousEvent()
                    Else
                        FindNextEvent()
                    End If

                    e.Handled = True
                    Return
                End If

                ' HAR preview (navigate between requests)
                If HarPreview_grp.Visibility = Visibility.Visible AndAlso
       FindNextHarRequest_btn.Visibility = Visibility.Visible Then

                    If Keyboard.Modifiers.HasFlag(ModifierKeys.Shift) Then
                        FindPreviousHarRequest()
                    Else
                        FindNextHarRequest()
                    End If

                    e.Handled = True
                    Return
                End If

                ' Text preview (navigate within text)
                If TextPreview_grp.Visibility = Visibility.Visible AndAlso
       FindNext_btn.Visibility = Visibility.Visible Then

                    FindNextInText()
                    e.Handled = True
                    Return
                End If

            End If


            ' Ctrl+Down: Select next file in results
            If e.Key = Key.Down AndAlso (Keyboard.Modifiers And ModifierKeys.Control) = ModifierKeys.Control Then
                If NextFile_btn.Visibility = Visibility.Visible Then
                    NextFile()
                    e.Handled = True
                    Return
                End If
            End If


            ' Ctrl+R: Reset all fields and results (when not scanning)
            If e.Key = Key.R AndAlso
   (Keyboard.Modifiers And ModifierKeys.Control) = ModifierKeys.Control AndAlso
   Not _isScanning Then

                Reset_btn_Click(Nothing, Nothing)
                e.Handled = True
                Return
            End If


        End Sub

#End Region

#Region "Validation / Enablement"
        ' ====================================================================
        ' Button state management based on input validity
        ' ====================================================================

        ''' <summary>
        ''' Triggered when Path_txt, Search_txt, or checkbox state changes.
        ''' Re-evaluates button enable states.
        ''' </summary>
        Private Sub AnyInputChanged(sender As Object, e As RoutedEventArgs)
            UpdateButtonsState()
        End Sub

        ''' <summary>
        ''' Updates button enable/disable states based on:
        ''' - Scanning state
        ''' - Path validity (folder or ZIP exists)
        ''' - Search term presence
        ''' </summary>
        Private Sub UpdateButtonsState()
            If _isScanning Then
                ' During scan: Scan button becomes "Cancel", lock most inputs
                Scan_btn.IsEnabled = True
                Reset_btn.IsEnabled = False

                BrowseFolder_btn.IsEnabled = String.IsNullOrWhiteSpace(Path_txt.Text)
                BrowseZip_btn.IsEnabled = String.IsNullOrWhiteSpace(Path_txt.Text)
                Path_txt.IsEnabled = False
                Search_txt.IsEnabled = False
                CaseSensitive_chk.IsEnabled = False
                ExactMatch_chk.IsEnabled = False
                Return
            End If

            Dim p = Path_txt.Text.Trim()
            Dim term = Search_txt.Text

            ' Scan enabled when path is valid (folder or archive) AND search term provided
            Dim pathValid As Boolean =
                Directory.Exists(p) OrElse
                (File.Exists(p) AndAlso _supportedArchiveExt.Contains(Path.GetExtension(p)))

            Dim termValid As Boolean = Not String.IsNullOrWhiteSpace(term)

            Scan_btn.IsEnabled = pathValid AndAlso termValid

            ' Reset enabled if any input is provided or results exist
            Reset_btn.IsEnabled =
                (Not String.IsNullOrWhiteSpace(Path_txt.Text)) OrElse
                (Not String.IsNullOrWhiteSpace(Search_txt.Text)) OrElse
                (_hits.Count > 0)

            ' Browse buttons only enabled when path is empty (prevent overwriting typed paths)
            BrowseFolder_btn.IsEnabled = String.IsNullOrWhiteSpace(Path_txt.Text)
            BrowseZip_btn.IsEnabled = String.IsNullOrWhiteSpace(Path_txt.Text)
            Path_txt.IsEnabled = False
            Search_txt.IsEnabled = True
            CaseSensitive_chk.IsEnabled = True
            ExactMatch_chk.IsEnabled = True
        End Sub

#End Region

#Region "Browse (Folder/Zip) - Two Button Approach"

        ''' <summary>
        ''' Prompts user to select a folder to scan recursively
        ''' Tries multiple folder picker methods for maximum compatibility
        ''' </summary>
        Private Sub BrowseFolder_btn_Click(sender As Object, e As RoutedEventArgs)
            Dim selectedFolder As String = TryPickFolderBest()
            If Not String.IsNullOrWhiteSpace(selectedFolder) Then
                Path_txt.Text = selectedFolder
            End If
        End Sub

        ''' <summary>
        ''' Prompts user to select an archive file to scan
        ''' </summary>
        Private Sub BrowseZip_btn_Click(sender As Object, e As RoutedEventArgs)
            Dim ofd As New OpenFileDialog() With {
                .Filter = "Archive files|*.zip;*.7z;*.rar;*.tar;*.gz;*.bz2;*.tgz;*.tbz;*.tbz2;*.tar.gz;*.tar.bz2;*.tar.xz;*.txz;*.cab|ZIP files (*.zip)|*.zip|7-Zip files (*.7z)|*.7z|RAR files (*.rar)|*.rar|TAR files|*.tar;*.tar.gz;*.tar.bz2;*.tar.xz;*.tgz;*.tbz;*.tbz2;*.txz|GZIP files (*.gz)|*.gz|BZIP2 files (*.bz2)|*.bz2|CAB files (*.cab)|*.cab",
                .Multiselect = False,
                .Title = "Select archive to scan"
            }
            If ofd.ShowDialog() = True Then
                Path_txt.Text = ofd.FileName
            End If
        End Sub

        ''' <summary>
        ''' Attempts multiple folder picker methods for maximum compatibility
        ''' Tries: (1) WindowsAPICodePack, (2) WinForms FolderBrowserDialog, (3) InputBox manual entry
        ''' </summary>
        ''' <returns>Selected folder path, or Nothing if cancelled</returns>
        Private Function TryPickFolderBest() As String

            ' 1) Try Windows API Code Pack CommonOpenFileDialog (best UX, requires NuGet package)
            Try
                Dim t = Type.GetType("Microsoft.WindowsAPICodePack.Dialogs.CommonOpenFileDialog, Microsoft.WindowsAPICodePack.Shell")
                If t IsNot Nothing Then
                    Dim dlg = Activator.CreateInstance(t)
                    ' Use reflection to configure dialog without direct reference
                    t.GetProperty("IsFolderPicker")?.SetValue(dlg, True)
                    t.GetProperty("Title")?.SetValue(dlg, "Select folder containing logs")

                    Dim result = t.GetMethod("ShowDialog", Type.EmptyTypes).Invoke(dlg, Nothing)
                    Dim okEnum = [Enum].Parse(Type.GetType("Microsoft.WindowsAPICodePack.Dialogs.CommonFileDialogResult, Microsoft.WindowsAPICodePack.Shell"), "Ok")

                    If result.Equals(okEnum) Then
                        Return CStr(t.GetProperty("FileName")?.GetValue(dlg))
                    Else
                        Return Nothing ' USER CANCELLED → STOP
                    End If
                End If
            Catch
                ' Package not available, try next method
            End Try

            ' 2) Use WinForms FolderBrowserDialog (reliable, basic UX)
            Try
                Using dialog As New WinForms.FolderBrowserDialog()
                    dialog.Description = "Select folder containing logs"
                    dialog.ShowNewFolderButton = True

                    ' Set UseDescriptionForTitle if available (Windows Vista+)
                    Try
                        dialog.UseDescriptionForTitle = True
                    Catch
                        ' Property not available on older systems
                    End Try

                    If dialog.ShowDialog() = WinForms.DialogResult.OK Then
                        Return dialog.SelectedPath
                    Else
                        Return Nothing ' USER CANCELLED → STOP
                    End If
                End Using
            Catch ex As Exception
                ' WinForms not available or error occurred, try last resort
            End Try

            ' 3) Fallback: InputBox for manual path entry (no confusing file picker)
            Try
                Dim path = Microsoft.VisualBasic.Interaction.InputBox(
                    "Enter folder path manually (advanced folder pickers not available):",
                    "Select Folder",
                    "")

                If Not String.IsNullOrWhiteSpace(path) AndAlso Directory.Exists(path) Then
                    Return path
                ElseIf Not String.IsNullOrWhiteSpace(path) Then
                    MessageBox.Show("The specified folder does not exist.", "Invalid Path", MessageBoxButton.OK, MessageBoxImage.Warning)
                End If
            Catch
            End Try

            Return Nothing
        End Function


#End Region

#Region "Scan/Cancel"

        ''' <summary>
        ''' Toggles between starting a new scan and cancelling an active scan
        ''' </summary>
        Private Sub Scan_btn_Click(sender As Object, e As RoutedEventArgs)
            If _isScanning Then
                CancelScan()
            Else
                StartScan()
            End If
        End Sub

        ''' <summary>
        ''' Requests cancellation of the active scan operation
        ''' </summary>
        Private Sub CancelScan()
            Try
                If _scanCts IsNot Nothing Then _scanCts.Cancel()
                Status("Cancelling...")
            Catch
            End Try
        End Sub

        ''' <summary>
        ''' Initiates a new scan operation with async task orchestration
        ''' Clears previous results, configures UI, and spawns background scan task
        ''' </summary>
        Private Sub StartScan()
            If _isScanning Then Return
            If Not Scan_btn.IsEnabled Then Return

            CleanupTemp()

            ' Detect if user is re-running same search (for status message)
            Dim isRerun As Boolean = (_lastScanPath = Path_txt.Text)
            _lastScanPath = Path_txt.Text

            ' Create new cancellation token source for this scan
            If _scanCts IsNot Nothing Then
                _scanCts.Cancel()
                _scanCts.Dispose()
            End If
            _scanCts = New CancellationTokenSource()

            ' Reset all scan state
            _hits.Clear()
            Results_lst.SelectedIndex = -1
            _currentTextFindStart = 0

            NextFile_btn.Visibility = Visibility.Collapsed
            FindNext_btn.Visibility = Visibility.Collapsed
            FindNextEvent_btn.Visibility = Visibility.Collapsed
            FindNextHarRequest_btn.Visibility = Visibility.Collapsed
            FindPreviousHarRequest_btn.Visibility = Visibility.Collapsed

            ShowTextPreviewMode()
            ClearTextPreview()

            _isScanning = True
            Scan_btn.Content = "Cancel (Esc)"

            Status(If(isRerun, "Re-running search on same source...", "Scanning..."))
            ShowProgress(True)

            ' Start label throttle timer and reset timestamps
            _pendingFileLabel = ""
            _lastFileLabelUpdateMs = 0
            If Not _fileLabelTimer.IsEnabled Then _fileLabelTimer.Start()

            ' Start elapsed time tracking
            _scanElapsedStopwatch.Restart()
            TimeElapsed_lbl.Visibility = Visibility.Visible
            TimeElapsed_lbl.Text = "Time Elapsed: 00:00:00"
            If Not _elapsedTimeTimer.IsEnabled Then _elapsedTimeTimer.Start()

            ' Reset file progress tracking
            _totalFilesToScan = 0
            _filesScanned = 0

            UpdateButtonsState()

            ' Capture scan parameters for background task
            Dim p = Path_txt.Text.Trim()
            Dim term = Search_txt.Text
            Dim cs = CaseSensitive_chk.IsChecked.GetValueOrDefault(False)
            Dim exactMatch = ExactMatch_chk.IsChecked.GetValueOrDefault(False)
            Dim ct = _scanCts.Token

            ' Capture scan root folder for relative display
            ' (e.g., show "logs\app.log" instead of "C:\Users\...\logs\app.log")
            _scanRootFolder = If(Directory.Exists(p), p, "")

            ' Count total files before scanning (for progress bar)
            Task.Run(Sub()
                         Try
                             If Directory.Exists(p) Then
                                 _totalFilesToScan = CountFilesInFolder(p)
                             ElseIf File.Exists(p) AndAlso _supportedArchiveExt.Contains(Path.GetExtension(p)) Then
                                 _totalFilesToScan = CountFilesInArchive(p)
                             End If
                         Catch ex As Exception
                             _totalFilesToScan = 0
                         End Try
                     End Sub)

            ' Spawn background scan task to keep UI responsive
            Task.Run(Async Function()
                         Dim wasCancelled As Boolean = False

                         Try
                             ' Route to appropriate scan method based on source type
                             If Directory.Exists(p) Then
                                 Await ScanFolderAsync(p, term, cs, exactMatch, ct)
                             ElseIf File.Exists(p) AndAlso _supportedArchiveExt.Contains(Path.GetExtension(p)) Then
                                 ' CAB files use 7-Zip extraction
                                 If Path.GetExtension(p).Equals(".cab", StringComparison.OrdinalIgnoreCase) Then
                                     Await ScanCabArchiveAsync(p, term, cs, exactMatch, ct, depth:=0)
                                 Else
                                     Await ScanArchiveAsync(p, term, cs, exactMatch, ct, depth:=0)
                                 End If
                             End If
                         Catch ex As OperationCanceledException
                             wasCancelled = True
                         Catch ex As TaskCanceledException
                             wasCancelled = True
                         Catch ex As Exception
                             Dispatcher.Invoke(Sub() Status("Error: " & ex.Message))
                         Finally
                             ' Allow pending UI updates to complete before showing final status
                             Task.Delay(50).Wait()

                             Dispatcher.Invoke(Sub()
                                                   _isScanning = False
                                                   ShowProgress(False)
                                                   Scan_btn.Content = "Scan (Enter)"

                                                   ' Stop timers and clear labels
                                                   If _fileLabelTimer.IsEnabled Then _fileLabelTimer.Stop()
                                                   If _elapsedTimeTimer.IsEnabled Then _elapsedTimeTimer.Stop()
                                                   _scanElapsedStopwatch.Stop()
                                                   SetCurrentFileDisplayImmediate("")

                                                   If wasCancelled OrElse (ct.IsCancellationRequested) Then
                                                       Status("Scan cancelled")
                                                   Else
                                                       If _hits.Count > 0 Then
                                                           Status($"Scan complete - {_hits.Count} result(s)")
                                                           NextFile_btn.Visibility = Visibility.Visible
                                                           Results_lst.SelectedIndex = 0
                                                       Else
                                                           Status("Scan complete - No matches")
                                                       End If
                                                   End If

                                                   UpdateButtonsState()

                                                   If wasCancelled OrElse ct.IsCancellationRequested Then
                                                       Search_txt.Focus()
                                                       Keyboard.Focus(Search_txt)
                                                   End If

                                               End Sub)
                         End Try
                     End Function)
        End Sub

#End Region

#Region "Reset"

        ''' <summary>
        ''' Resets application to initial state: clears inputs, results, and stops any active scan
        ''' </summary>
        Private Sub Reset_btn_Click(sender As Object, e As RoutedEventArgs)

            ' Stop any active scan first
            If _isScanning Then
                CancelScan()
            End If

            ' Clear all input fields
            Path_txt.Text = ""
            Search_txt.Text = ""
            CaseSensitive_chk.IsChecked = False
            ExactMatch_chk.IsChecked = False

            ' Clear results and selection
            _hits.Clear()
            Results_lst.SelectedIndex = -1

            ' Reset preview pane
            ClearTextPreview()
            ShowTextPreviewMode()

            ' Clear WebView2 content
            ClearWebPreview()
            _currentWebMatchIndex = 0
            _totalWebMatches = 0

            NextFile_btn.Visibility = Visibility.Collapsed
            FindNext_btn.Visibility = Visibility.Collapsed
            FindNextEvent_btn.Visibility = Visibility.Collapsed

            ShowProgress(False)
            If _fileLabelTimer.IsEnabled Then _fileLabelTimer.Stop()
            If _elapsedTimeTimer.IsEnabled Then _elapsedTimeTimer.Stop()
            _scanElapsedStopwatch.Reset()
            TimeElapsed_lbl.Visibility = Visibility.Collapsed
            TimeElapsed_lbl.Text = "Time Elapsed: 00:00:00"
            SetCurrentFileDisplayImmediate("")

            _currentTextFindStart = 0
            _scanRootFolder = ""

            Status("Ready")

            _isScanning = False
            Scan_btn.Content = "Scan (Enter)"
            Scan_btn.IsEnabled = False
            _lastScanPath = ""
            BrowseFolder_btn.IsEnabled = True
            BrowseZip_btn.IsEnabled = True

            CleanupTemp()

            UpdateButtonsState()
        End Sub

#End Region

#Region "Relative Display Name"

        ''' <summary>
        ''' Converts absolute file path to relative path for cleaner display
        ''' Falls back to filename only if path is not under root folder
        ''' </summary>
        ''' <param name="rootFolder">Base folder for relative path calculation</param>
        ''' <param name="fullPath">Absolute path to file</param>
        ''' <returns>Relative path (e.g., "logs\app.log") or filename if not under root</returns>
        Private Function MakeRelativeDisplay(rootFolder As String, fullPath As String) As String
            Try
                If String.IsNullOrWhiteSpace(rootFolder) Then
                    Return Path.GetFileName(fullPath)
                End If

                Dim root = Path.GetFullPath(rootFolder.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar) & Path.DirectorySeparatorChar)
                Dim full = Path.GetFullPath(fullPath)

                If Not full.StartsWith(root, StringComparison.OrdinalIgnoreCase) Then
                    Return Path.GetFileName(fullPath)
                End If

                Return full.Substring(root.Length)
            Catch
                Return Path.GetFileName(fullPath)
            End Try
        End Function

#End Region

#Region "Scan Helpers"

        ''' <summary>
        ''' Recursively scans a folder for searchable files and archives
        ''' Supports: Text files, EVTX logs, HAR files, and nested archives (ZIP, 7Z, RAR, TAR, GZ, CAB)
        ''' Progress tracking: Increments file counter AFTER archive scanning to avoid double-counting
        ''' </summary>
        ''' <param name="folder">Root folder path to scan recursively</param>
        ''' <param name="term">Search term to find</param>
        ''' <param name="caseSensitive">Whether to perform case-sensitive search</param>
        ''' <param name="exactMatch">Whether to use word boundary matching</param>
        ''' <param name="ct">Cancellation token for aborting scan</param>
        Private Async Function ScanFolderAsync(folder As String, term As String, caseSensitive As Boolean, exactMatch As Boolean, ct As CancellationToken) As Task
            ' Enumerate all files recursively
            For Each f In Directory.EnumerateFiles(folder, "*.*", SearchOption.AllDirectories)
                If ct.IsCancellationRequested Then Exit For

                ' Show relative path (more useful than just file name)
                Dim relScanning = MakeRelativeDisplay(folder, f)
                ThrottledSetCurrentFileDisplay(relScanning)

                Dim ext = Path.GetExtension(f)

                ' Handle nested archives recursively
                ' Note: Archive files are counted individually when their contents are scanned,
                ' so we don't increment _filesScanned here to avoid double-counting
                If _supportedArchiveExt.Contains(ext) Then
                    ' CAB files use 7-Zip extraction
                    If ext.Equals(".cab", StringComparison.OrdinalIgnoreCase) Then
                        Await ScanCabArchiveAsync(f, term, caseSensitive, exactMatch, ct, depth:=0)
                    Else
                        Await ScanArchiveAsync(f, term, caseSensitive, exactMatch, ct, depth:=0)
                    End If
                    Continue For
                End If

                ' Increment file counter and update progress
                ' (only for non-archive files since archive contents are counted separately)
                _filesScanned += 1
                UpdateScanProgress()

                If _supportedTextExt.Contains(ext) Then
                    Dim containsTerm As Boolean

                    ' For HTML/XML/JSON files, search only visible/data content (not tags/markup)
                    If ext = ".html" OrElse ext = ".xml" OrElse ext = ".json" Then
                        containsTerm = Await HtmlXmlJsonFileContainsAsync(f, term, caseSensitive, exactMatch, ct)
                    Else
                        ' For plain text files, search entire content
                        containsTerm = Await TextFileContainsAsync(f, term, caseSensitive, exactMatch, ct)
                    End If

                    If containsTerm Then
                        Dim relName = MakeRelativeDisplay(folder, f)
                        AddHit(New SearchHit With {
                            .DisplayName = relName,
                            .Kind = HitKind.DiskTextFile,
                            .FilePath = f
                        })
                    End If

                ElseIf _supportedHarExt.Contains(ext) Then
                    Dim har = Await HarCollectMatchesAsync(f, term, caseSensitive, exactMatch, ct)
                    If har IsNot Nothing Then
                        har.DisplayName = MakeRelativeDisplay(folder, f)
                        AddHit(har)
                    End If

                ElseIf _supportedEvtxExt.Contains(ext) Then
                    Dim ev = Await EvtxCollectMatchesAsync(f, term, caseSensitive, exactMatch, ct)
                    If ev IsNot Nothing Then
                        ev.DisplayName = MakeRelativeDisplay(folder, f)
                        AddHit(ev)
                    End If
                End If
            Next
        End Function

        ''' <summary>
        ''' Scans all entries in an archive (ZIP, 7Z, RAR, TAR, etc.) for text and EVTX files
        ''' Extracts EVTX files to temp storage for EventLogReader access
        ''' Uses SharpCompress for multi-format archive support (ZIP, 7Z, RAR, TAR, GZ, BZ2)
        ''' Supports one level of nested archives (e.g., ZIP inside ZIP)
        ''' </summary>
        ''' <param name="depth">Current nesting depth (0 = top-level archive, 1 = nested archive)</param>
        ''' <param name="parentArchiveName">Display name chain from parent archives (e.g., "diagnostic.zip »")</param>
        Private Async Function ScanArchiveAsync(archivePath As String, term As String, caseSensitive As Boolean, exactMatch As Boolean, ct As CancellationToken, Optional depth As Integer = 0, Optional parentArchiveName As String = "") As Task
            Try
                Debug.WriteLine($"=== Starting scan of {Path.GetFileName(archivePath)} (depth {depth}) ===")

                Using archive = OpenArchive(archivePath)
                    Debug.WriteLine($"Archive opened successfully, {archive.Entries.Count()} total entries")
                    For Each entry In archive.Entries.Where(Function(e) Not e.IsDirectory)
                        Await ProcessArchiveEntry(entry, archivePath, term, caseSensitive, exactMatch, ct, depth, parentArchiveName)
                        If ct.IsCancellationRequested Then Exit For
                    Next
                End Using
                Debug.WriteLine($"Archive scan completed successfully for {Path.GetFileName(archivePath)}")
            Catch ex As Exception
                ' Archive might be corrupted or unsupported format, log and continue
                Debug.WriteLine($"ERROR scanning archive {Path.GetFileName(archivePath)}: {ex.Message}")
                Debug.WriteLine($"Stack trace: {ex.StackTrace}")
                Dispatcher.Invoke(Sub() Status($"Error reading archive {Path.GetFileName(archivePath)}: {ex.Message}"))
            End Try
        End Function

        ''' <summary>
        ''' Processes a single archive entry (Archive API version)
        ''' Supports nested archives up to 1 level deep
        ''' </summary>
        ''' <param name="depth">Current nesting depth (0 = top-level archive, 1 = nested archive)</param>
        ''' <param name="parentArchiveName">Display name chain from parent archives (e.g., "diagnostic.zip »")</param>
        Private Async Function ProcessArchiveEntry(entry As IArchiveEntry, archivePath As String, term As String, caseSensitive As Boolean, exactMatch As Boolean, ct As CancellationToken, Optional depth As Integer = 0, Optional parentArchiveName As String = "") As Task
            If String.IsNullOrEmpty(entry.Key) Then Return

            ' Show "archive.ext | path/to/file.ext" format
            ThrottledSetCurrentFileDisplay($"{Path.GetFileName(archivePath)} | {entry.Key}")

            Dim ext = Path.GetExtension(entry.Key)

            ' Handle nested archives (only if depth allows - max 1 level deep)
            ' Note: Nested archive entries are counted when their contents are scanned,
            ' so we don't increment _filesScanned here to avoid double-counting
            If _supportedArchiveExt.Contains(ext) AndAlso depth < 1 Then
                Debug.WriteLine($"Found nested archive: {entry.Key} at depth {depth}, extracting and scanning...")

                ' Extract nested archive to temp
                Dim tempNestedArchive = ExtractArchiveEntryToTemp(entry, archivePath)

                ' Build archive chain for nested display (e.g., "diagnostic.zip » reports.cab")
                Dim nestedArchiveChain = parentArchiveName & Path.GetFileName(archivePath) & " » "

                Try
                    ' Scan the nested archive with increased depth and archive chain
                    If ext.Equals(".cab", StringComparison.OrdinalIgnoreCase) Then
                        Await ScanCabArchiveAsync(tempNestedArchive, term, caseSensitive, exactMatch, ct, depth:=depth + 1, parentArchiveName:=nestedArchiveChain)
                    Else
                        Await ScanArchiveAsync(tempNestedArchive, term, caseSensitive, exactMatch, ct, depth:=depth + 1, parentArchiveName:=nestedArchiveChain)
                    End If
                Finally
                    ' Clean up extracted nested archive
                    SafeDelete(tempNestedArchive)
                End Try

                Return
            End If

            ' Increment file counter and update progress
            ' (only for non-archive entries since archive contents are counted separately)
            _filesScanned += 1
            UpdateScanProgress()

            If _supportedTextExt.Contains(ext) Then
                Using entryStream = entry.OpenEntryStream()
                    Dim containsTerm As Boolean

                    ' For HTML/XML/JSON files, search only visible/data content
                    If ext = ".html" OrElse ext = ".xml" OrElse ext = ".json" Then
                        containsTerm = Await HtmlXmlJsonStreamContainsAsync(entryStream, term, caseSensitive, exactMatch, ct)
                    Else
                        ' For plain text files, search entire content
                        containsTerm = Await StreamContainsAsync(entryStream, term, caseSensitive, exactMatch, ct)
                    End If

                    If containsTerm Then
                        ' Build display name with parent archive chain (e.g., "diagnostic.zip » reports.cab | data.log")
                        Dim displayName = If(String.IsNullOrEmpty(parentArchiveName),
                                            $"{Path.GetFileName(archivePath)} | {entry.Key}",
                                            $"{parentArchiveName}{Path.GetFileName(archivePath)} | {entry.Key}")

                        AddHit(New SearchHit With {
                               .DisplayName = displayName,
                               .Kind = HitKind.ZipTextEntry,
                               .ZipPath = archivePath,
                               .ZipEntryName = entry.Key
                        })
                    End If
                End Using

            ElseIf _supportedEvtxExt.Contains(ext) Then
                ' EventLogReader requires file path, extract to temp
                Dim tempEvtx = ExtractArchiveEntryToTemp(entry, archivePath)
                Dim ev = Await EvtxCollectMatchesAsync(tempEvtx, term, caseSensitive, exactMatch, ct)

                If ev IsNot Nothing Then
                    ' Build display name with parent archive chain
                    ev.DisplayName = If(String.IsNullOrEmpty(parentArchiveName),
                                       $"{Path.GetFileName(archivePath)} | {entry.Key}",
                                       $"{parentArchiveName}{Path.GetFileName(archivePath)} | {entry.Key}")
                    ev.Kind = HitKind.EvtxFromZipTemp
                    ev.ZipPath = archivePath
                    ev.ZipEntryName = entry.Key
                    ev.TempEvtxPath = tempEvtx
                    AddHit(ev)
                Else
                    SafeDelete(tempEvtx)
                End If
            End If
        End Function

        ''' <summary>
        ''' Scans CAB archive using 7-Zip command line tool (simple and reliable)
        ''' Extracts to temp directory, scans files, then cleans up
        ''' Supports one level of nested archives (e.g., ZIP inside CAB)
        ''' </summary>
        ''' <param name="depth">Current nesting depth (0 = top-level archive, 1 = nested archive)</param>
        ''' <param name="parentArchiveName">Display name chain from parent archives (e.g., "diagnostic.zip »")</param>
        Private Async Function ScanCabArchiveAsync(cabPath As String, term As String, caseSensitive As Boolean, exactMatch As Boolean, ct As CancellationToken, Optional depth As Integer = 0, Optional parentArchiveName As String = "") As Task
            Await Task.Run(Sub()
                               Try
                                   Debug.WriteLine($"=== Starting CAB scan using 7-Zip: {Path.GetFileName(cabPath)} (depth {depth}) ===")

                                   ' Extract CAB to temp directory using 7-Zip
                                   Dim tempExtractDir = Path.Combine(Path.GetTempPath(), "BeaconCabExtract_" & Guid.NewGuid().ToString("N"))
                                   Directory.CreateDirectory(tempExtractDir)

                                   Try
                                       ' Extract using 7-Zip command line
                                       Dim success = SevenZipHelper.ExtractCab(cabPath, tempExtractDir)

                                       If Not success Then
                                           Debug.WriteLine($"CAB extraction failed for {Path.GetFileName(cabPath)}")
                                           Dispatcher.Invoke(Sub() Status($"Error extracting CAB {Path.GetFileName(cabPath)}"))
                                           Return
                                       End If

                                       Debug.WriteLine($"CAB extracted to: {tempExtractDir}")

                                       ' Scan extracted files
                                       Dim extractedFiles = Directory.GetFiles(tempExtractDir, "*.*", SearchOption.AllDirectories)
                                       Debug.WriteLine($"CAB contains {extractedFiles.Length} file(s)")

                                       For Each filePath In extractedFiles
                                           If ct.IsCancellationRequested Then Exit For

                                           Dim relativePath = filePath.Substring(tempExtractDir.Length + 1)
                                           ThrottledSetCurrentFileDisplay($"{Path.GetFileName(cabPath)} | {relativePath}")

                                           Dim ext = Path.GetExtension(filePath)

                                           ' Handle nested archives (only if depth allows - max 1 level deep)
                                           ' Note: Nested archive entries are counted when their contents are scanned,
                                           ' so we don't increment _filesScanned here to avoid double-counting
                                           If _supportedArchiveExt.Contains(ext) AndAlso depth < 1 Then
                                               Debug.WriteLine($"Found nested archive in CAB: {relativePath} at depth {depth}, scanning...")

                                               ' Build archive chain for nested display (e.g., "diagnostic.cab » logs.zip")
                                               Dim nestedArchiveChain = parentArchiveName & Path.GetFileName(cabPath) & " » "

                                               ' Scan the nested archive with increased depth and archive chain
                                               If ext.Equals(".cab", StringComparison.OrdinalIgnoreCase) Then
                                                   ScanCabArchiveAsync(filePath, term, caseSensitive, exactMatch, ct, depth:=depth + 1, parentArchiveName:=nestedArchiveChain).Wait()
                                               Else
                                                   ScanArchiveAsync(filePath, term, caseSensitive, exactMatch, ct, depth:=depth + 1, parentArchiveName:=nestedArchiveChain).Wait()
                                               End If

                                               Continue For
                                           End If

                                           ' Increment file counter and update progress
                                           ' (only for non-archive files since archive contents are counted separately)
                                           _filesScanned += 1
                                           UpdateScanProgress()

                                           If _supportedTextExt.Contains(ext) Then
                                               Dim containsTerm As Boolean

                                               If ext = ".html" OrElse ext = ".xml" OrElse ext = ".json" Then
                                                   containsTerm = HtmlXmlJsonFileContainsAsync(filePath, term, caseSensitive, exactMatch, ct).Result
                                               Else
                                                   containsTerm = TextFileContainsAsync(filePath, term, caseSensitive, exactMatch, ct).Result
                                               End If

                                               If containsTerm Then
                                                   ' Build display name with parent archive chain (e.g., "diagnostic.zip » reports.cab | data.log")
                                                   Dim displayName = If(String.IsNullOrEmpty(parentArchiveName),
                                                                       $"{Path.GetFileName(cabPath)} | {relativePath}",
                                                                       $"{parentArchiveName}{Path.GetFileName(cabPath)} | {relativePath}")

                                                   AddHit(New SearchHit With {
                                                          .DisplayName = displayName,
                                                          .Kind = HitKind.CabTextEntry,
                                                          .ZipPath = cabPath,
                                                          .ZipEntryName = relativePath
                                                   })
                                               End If

                                           ElseIf _supportedEvtxExt.Contains(ext) Then
                                               Debug.WriteLine($"[CAB] Processing EVTX file: {relativePath}")
                                               Dim ev = EvtxCollectMatchesAsync(filePath, term, caseSensitive, exactMatch, ct).Result

                                               If ev IsNot Nothing Then
                                                   Debug.WriteLine($"[CAB] ✓ EVTX SearchHit returned with {ev.MatchingEvents.Count} event(s)")
                                                   ' Build display name with parent archive chain
                                                   ev.DisplayName = If(String.IsNullOrEmpty(parentArchiveName),
                                                                      $"{Path.GetFileName(cabPath)} | {relativePath}",
                                                                      $"{parentArchiveName}{Path.GetFileName(cabPath)} | {relativePath}")
                                                   ev.Kind = HitKind.EvtxFromCabTemp
                                                   ev.ZipPath = cabPath
                                                   ev.ZipEntryName = relativePath
                                                   ev.TempEvtxPath = filePath
                                                   Debug.WriteLine($"[CAB] Calling AddHit for EVTX: {ev.DisplayName}")
                                                   AddHit(ev)
                                                   _tempToDelete.Add(filePath) ' Track for cleanup
                                                   Debug.WriteLine($"[CAB] ✓✓✓ AddHit completed successfully")
                                               Else
                                                   Debug.WriteLine($"[CAB] ✗ EVTX SearchHit is Nothing - no matching events found")
                                               End If
                                           End If
                                       Next

                                       Debug.WriteLine($"CAB scan completed")

                                   Finally
                                       ' Cleanup temp directory (except EVTX files we're tracking)
                                       Try
                                           For Each file In Directory.GetFiles(tempExtractDir, "*.*", SearchOption.AllDirectories)
                                               If Not _tempToDelete.Contains(file) Then
                                                   SafeDelete(file)
                                               End If
                                           Next
                                           ' Try to remove directory (will fail if EVTX files still there, which is fine)
                                           Directory.Delete(tempExtractDir, True)
                                       Catch
                                           ' Ignore cleanup errors
                                       End Try
                                   End Try

                               Catch ex As Exception
                                   Debug.WriteLine($"ERROR scanning CAB {Path.GetFileName(cabPath)}: {ex.Message}")
                                   Debug.WriteLine($"Stack trace: {ex.StackTrace}")
                                   Dispatcher.Invoke(Sub() Status($"Error reading CAB {Path.GetFileName(cabPath)}: {ex.Message}"))
                               End Try
                           End Sub, ct)
        End Function

        ''' <summary>
        ''' Thread-safe addition of search result to UI-bound collection
        ''' Marshals to UI thread asynchronously via Dispatcher (non-blocking)
        ''' </summary>
        Private Sub AddHit(hit As SearchHit)
            Dispatcher.BeginInvoke(Sub() _hits.Add(hit))
        End Sub

#End Region

#Region "Text Search"

        ''' <summary>
        ''' Checks if file contains search term using line-by-line streaming
        ''' Uses FileShare.ReadWrite to access locked files (e.g., active logs)
        ''' </summary>
        Private Async Function TextFileContainsAsync(filePath As String, term As String, caseSensitive As Boolean, exactMatch As Boolean, ct As CancellationToken) As Task(Of Boolean)
            ' Empty search term should not match anything
            If String.IsNullOrWhiteSpace(term) Then Return False

            Dim comparison = If(caseSensitive, StringComparison.Ordinal, StringComparison.OrdinalIgnoreCase)

            Using fs As New FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)
                Using sr As New StreamReader(fs, detectEncodingFromByteOrderMarks:=True)
                    ' Scan line-by-line for memory efficiency
                    While True
                        If ct.IsCancellationRequested Then Return False
                        Dim line = Await sr.ReadLineAsync()
                        If line Is Nothing Then Exit While

                        If exactMatch Then
                            ' Exact match: use word boundaries
                            If ContainsExactMatch(line, term, caseSensitive) Then Return True
                        Else
                            ' Partial match: simple contains
                            If line.IndexOf(term, comparison) >= 0 Then Return True
                        End If
                    End While
                End Using
            End Using

            Return False
        End Function

        ''' <summary>
        ''' Checks if stream contains search term (used for ZIP entry scanning)
        ''' Leaves stream open for caller to manage (leaveOpen:=True)
        ''' </summary>
        Private Async Function StreamContainsAsync(stream As Stream, term As String, caseSensitive As Boolean, exactMatch As Boolean, ct As CancellationToken) As Task(Of Boolean)
            ' Empty search term should not match anything
            If String.IsNullOrWhiteSpace(term) Then Return False

            Dim comparison = If(caseSensitive, StringComparison.Ordinal, StringComparison.OrdinalIgnoreCase)

            Using sr As New StreamReader(stream, detectEncodingFromByteOrderMarks:=True, bufferSize:=4096, leaveOpen:=True)
                While True
                    If ct.IsCancellationRequested Then Return False
                    Dim line = Await sr.ReadLineAsync()
                    If line Is Nothing Then Exit While

                    If exactMatch Then
                        ' Exact match: use word boundaries
                        If ContainsExactMatch(line, term, caseSensitive) Then Return True
                    Else
                        ' Partial match: simple contains
                        If line.IndexOf(term, comparison) >= 0 Then Return True
                    End If
                End While
            End Using

            Return False
        End Function

        ''' <summary>
        ''' Checks if HTML/XML/JSON file contains search term in visible/data content only (strips tags/markup)
        ''' </summary>
        Private Async Function HtmlXmlJsonFileContainsAsync(filePath As String, term As String, caseSensitive As Boolean, exactMatch As Boolean, ct As CancellationToken) As Task(Of Boolean)
            If String.IsNullOrWhiteSpace(term) Then Return False

            Try
                Using fs As New FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)
                    Using sr As New StreamReader(fs, detectEncodingFromByteOrderMarks:=True)
                        Dim content = Await sr.ReadToEndAsync()
                        Dim visibleText = ExtractVisibleText(content)

                        If exactMatch Then
                            Return ContainsExactMatch(visibleText, term, caseSensitive)
                        Else
                            Dim comparison = If(caseSensitive, StringComparison.Ordinal, StringComparison.OrdinalIgnoreCase)
                            Return visibleText.IndexOf(term, comparison) >= 0
                        End If
                    End Using
                End Using
            Catch
                Return False
            End Try
        End Function

        ''' <summary>
        ''' Checks if HTML/XML/JSON stream contains search term in visible/data content only
        ''' </summary>
        Private Async Function HtmlXmlJsonStreamContainsAsync(stream As Stream, term As String, caseSensitive As Boolean, exactMatch As Boolean, ct As CancellationToken) As Task(Of Boolean)
            If String.IsNullOrWhiteSpace(term) Then Return False

            Try
                Using sr As New StreamReader(stream, detectEncodingFromByteOrderMarks:=True, bufferSize:=4096, leaveOpen:=True)
                    Dim content = Await sr.ReadToEndAsync()
                    Dim visibleText = ExtractVisibleText(content)

                    If exactMatch Then
                        Return ContainsExactMatch(visibleText, term, caseSensitive)
                    Else
                        Dim comparison = If(caseSensitive, StringComparison.Ordinal, StringComparison.OrdinalIgnoreCase)
                        Return visibleText.IndexOf(term, comparison) >= 0
                    End If
                End Using
            Catch
                Return False
            End Try
        End Function

        ''' <summary>
        ''' Extracts visible text from HTML/XML/JSON by removing tags and markup
        ''' </summary>
        Private Function ExtractVisibleText(content As String) As String
            If String.IsNullOrEmpty(content) Then Return ""

            ' Remove script and style tags and their content
            content = System.Text.RegularExpressions.Regex.Replace(content, "<script[^>]*>.*?</script>", "", System.Text.RegularExpressions.RegexOptions.Singleline Or System.Text.RegularExpressions.RegexOptions.IgnoreCase)
            content = System.Text.RegularExpressions.Regex.Replace(content, "<style[^>]*>.*?</style>", "", System.Text.RegularExpressions.RegexOptions.Singleline Or System.Text.RegularExpressions.RegexOptions.IgnoreCase)

            ' Remove all HTML/XML tags
            content = System.Text.RegularExpressions.Regex.Replace(content, "<[^>]+>", " ")

            ' Decode HTML entities (e.g., &amp; -> &, &lt; -> <)
            content = System.Net.WebUtility.HtmlDecode(content)

            ' Normalize whitespace
            content = System.Text.RegularExpressions.Regex.Replace(content, "\s+", " ")

            Return content.Trim()
        End Function

        ''' <summary>
        ''' Checks if text contains an exact match of the search term (whole word match)
        ''' Uses word boundaries to ensure the term is not part of a larger word
        ''' </summary>
        ''' <param name="text">Text to search in</param>
        ''' <param name="term">Search term to find</param>
        ''' <param name="caseSensitive">Whether to perform case-sensitive matching</param>
        ''' <returns>True if exact match found, False otherwise</returns>
        Private Function ContainsExactMatch(text As String, term As String, caseSensitive As Boolean) As Boolean
            If String.IsNullOrEmpty(text) OrElse String.IsNullOrEmpty(term) Then Return False

            Try
                ' Escape regex special characters in the search term
                Dim escapedTerm = System.Text.RegularExpressions.Regex.Escape(term)

                ' Build regex pattern with word boundaries (\b)
                ' \b matches position between word and non-word character
                Dim pattern = $"\b{escapedTerm}\b"

                ' Set regex options based on case sensitivity
                Dim options = System.Text.RegularExpressions.RegexOptions.None
                If Not caseSensitive Then
                    options = System.Text.RegularExpressions.RegexOptions.IgnoreCase
                End If

                ' Perform regex match
                Dim regex As New System.Text.RegularExpressions.Regex(pattern, options)
                Return regex.IsMatch(text)
            Catch
                ' If regex fails (e.g., invalid pattern), fall back to simple contains
                Dim comparison = If(caseSensitive, StringComparison.Ordinal, StringComparison.OrdinalIgnoreCase)
                Return text.IndexOf(term, comparison) >= 0
            End Try
        End Function

#End Region

#Region "EVTX Search (Cancellation-safe)"

        ''' <summary>
        ''' Scans EVTX file for events containing search term with graceful DLL handling
        ''' OPTIMIZATION: Pre-filters using XML representation before expensive FormatDescription() call
        ''' This yields 2-10x performance improvement for EVTX scanning
        ''' RESILIENCE: Safely handles missing message DLLs (LevelDisplayName, ProviderName)
        ''' Limits to 300 matches per file to prevent memory issues
        ''' </summary>
        ''' <param name="evtxPath">Path to the EVTX file</param>
        ''' <param name="term">Search term to find</param>
        ''' <param name="caseSensitive">Whether to perform case-sensitive search</param>
        ''' <param name="exactMatch">Whether to use word boundary matching</param>
        ''' <param name="ct">Cancellation token for aborting scan</param>
        ''' <returns>SearchHit containing matching events, or Nothing if no matches</returns>
        Private Async Function EvtxCollectMatchesAsync(evtxPath As String,
                                                     term As String,
                                                     caseSensitive As Boolean,
                                                     exactMatch As Boolean,
                                                     ct As CancellationToken) As Task(Of SearchHit)

            Try
                Return Await Task.Run(Function()

                                          Dim comparison = If(caseSensitive, StringComparison.Ordinal, StringComparison.OrdinalIgnoreCase)

                                          Dim hit As New SearchHit With {
                                              .Kind = HitKind.EvtxFileOnDisk,
                                              .FilePath = evtxPath,
                                              .DisplayName = evtxPath
                                          }

                                          ' Read EVTX events with graceful cancellation handling
                                          Using rdr As New EventLogReader(evtxPath, PathType.FilePath)
                                              While True
                                                  ' Do NOT throw - exit cleanly to avoid VS "user-unhandled"
                                                  If ct.IsCancellationRequested Then Return Nothing

                                                  Dim rec = rdr.ReadEvent()
                                                  If rec Is Nothing Then Exit While

                                                  Using rec
                                                      ' OPTIMIZATION: Check XML representation first (much faster than FormatDescription)
                                                      ' This allows us to skip expensive formatting for events that don't contain the search term
                                                      Dim xmlString As String = Nothing
                                                      Try
                                                          xmlString = rec.ToXml()
                                                      Catch
                                                          ' If XML conversion fails, skip this event
                                                          Continue While
                                                      End Try

                                                      ' Quick pre-filter: if search term not in XML, it won't be in formatted message either
                                                      ' Note: This is a fast substring check to avoid expensive FormatDescription() calls
                                                      If String.IsNullOrEmpty(xmlString) OrElse xmlString.IndexOf(term, comparison) < 0 Then
                                                          Continue While
                                                      End If

                                                      ' At this point, we KNOW the XML contains the search term (pre-filter passed)
                                                      ' Now try to get the formatted message
                                                      Dim msg As String = ""
                                                      Dim msgFailed As Boolean = False
                                                      Dim useFallback As Boolean = False
                                                      Dim hasMatch As Boolean = False

                                                      Try
                                                          msg = rec.FormatDescription()
                                                          ' FormatDescription can return Nothing or empty string when DLLs are missing
                                                          If String.IsNullOrEmpty(msg) Then
                                                              msgFailed = True
                                                              useFallback = True
                                                          Else
                                                              ' Check if the formatted message contains the search term
                                                              If exactMatch Then
                                                                  hasMatch = ContainsExactMatch(msg, term, caseSensitive)
                                                              Else
                                                                  hasMatch = msg.IndexOf(term, comparison) >= 0
                                                              End If

                                                              ' If formatted message doesn't contain the term, fall back to XML
                                                              If Not hasMatch Then
                                                                  useFallback = True
                                                              End If
                                                          End If
                                                      Catch ex As Exception
                                                          ' Exception thrown when message DLLs are completely missing
                                                          msgFailed = True
                                                          useFallback = True
                                                      End Try

                                                      ' Use fallback message with raw XML when:
                                                      ' 1. Formatting failed (DLL missing), OR
                                                      ' 2. Formatted message doesn't contain the search term (term only in raw XML)
                                                      If useFallback Then
                                                          ' Build fallback message with XML
                                                          If msgFailed Then
                                                              msg = "⚠️ Warning: Message unavailable - required DLL not found." & vbCrLf &
                                                                    "Event data may not render correctly. Showing raw XML data below:" & vbCrLf & vbCrLf &
                                                                    xmlString
                                                          Else
                                                              msg = "ℹ️ Note: Search term found in raw event data (XML), not in formatted message." & vbCrLf &
                                                                    "Showing raw XML data below:" & vbCrLf & vbCrLf &
                                                                        xmlString
                                                          End If

                                                          ' IMPORTANT: Re-validate match against the complete fallback message (including XML)
                                                          ' This ensures exact match rules are respected
                                                          If exactMatch Then
                                                              hasMatch = ContainsExactMatch(msg, term, caseSensitive)
                                                          Else
                                                              hasMatch = msg.IndexOf(term, comparison) >= 0
                                                          End If
                                                      End If

                                                      ' Only add if it actually matches (respects exact match rules)
                                                      If hasMatch Then
                                                          ' Safely get properties that may throw when DLLs are missing
                                                          Dim eventLevel As String = "Unknown"
                                                          Dim eventProvider As String = "Unknown"

                                                          Try
                                                              eventLevel = rec.LevelDisplayName
                                                              If String.IsNullOrEmpty(eventLevel) Then eventLevel = "Unknown"
                                                          Catch
                                                              ' LevelDisplayName throws when DLL is missing
                                                              eventLevel = "Unknown"
                                                          End Try

                                                          Try
                                                              eventProvider = rec.ProviderName
                                                              If String.IsNullOrEmpty(eventProvider) Then eventProvider = "Unknown"
                                                          Catch
                                                              ' ProviderName may also throw
                                                              eventProvider = "Unknown"
                                                          End Try

                                                          hit.MatchingEvents.Add(New EventSummary With {
                                                               .Level = eventLevel,
                                                               .Provider = eventProvider,
                                                               .EventId = rec.Id,
                                                               .TimeCreated = rec.TimeCreated,
                                                               .Message = msg
                                                               })

                                                          ' Limit matches per file to prevent memory exhaustion
                                                          If hit.MatchingEvents.Count >= 300 Then
                                                              Exit While
                                                          End If
                                                      End If
                                                  End Using
                                              End While
                                          End Using

                                          If hit.MatchingEvents.Count > 0 Then
                                              Return hit
                                          End If

                                          Return Nothing

                                      End Function)
            Catch ex As OperationCanceledException
                Return Nothing
            Catch ex As TaskCanceledException
                Return Nothing
            Catch ex As Exception
                Return Nothing
            End Try
        End Function

#End Region

#Region "HAR Search (Cancellation-safe)"

        ''' <summary>
        ''' Scans HAR file for HTTP requests containing search term
        ''' Searches in URL, headers, request/response bodies
        ''' Limits to 300 matches per file to prevent memory issues
        ''' </summary>
        Private Async Function HarCollectMatchesAsync(harPath As String,
                                                       term As String,
                                                       caseSensitive As Boolean,
                                                       exactMatch As Boolean,
                                                       ct As CancellationToken) As Task(Of SearchHit)
            Try
                Using fs As New FileStream(harPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)
                    Return Await HarCollectMatchesFromStreamAsync(fs, term, caseSensitive, exactMatch, ct, harPath)
                End Using
            Catch ex As Exception
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' Scans HAR stream for HTTP requests containing search term
        ''' Collects all matching requests into a single SearchHit
        ''' </summary>
        Private Async Function HarCollectMatchesFromStreamAsync(stream As Stream,
                                                                 term As String,
                                                                 caseSensitive As Boolean,
                                                                 exactMatch As Boolean,
                                                                 ct As CancellationToken,
                                                                 Optional filePath As String = Nothing) As Task(Of SearchHit)
            Try
                Return Await Task.Run(Async Function()
                                          Dim comparison = If(caseSensitive, StringComparison.Ordinal, StringComparison.OrdinalIgnoreCase)

                                          Dim hit As New SearchHit With {
                                              .Kind = HitKind.HarFileOnDisk,
                                              .FilePath = filePath,
                                              .DisplayName = filePath
                                          }

                                          ' Parse HAR JSON
                                          Using reader As New StreamReader(stream, detectEncodingFromByteOrderMarks:=True, bufferSize:=4096, leaveOpen:=True)
                                              Dim jsonContent = Await reader.ReadToEndAsync()

                                              Using doc = JsonDocument.Parse(jsonContent)
                                                  Dim root = doc.RootElement

                                                  ' HAR structure: { "log": { "entries": [...] } }
                                                  If Not root.TryGetProperty("log", Nothing) Then Return Nothing
                                                  Dim log = root.GetProperty("log")

                                                  If Not log.TryGetProperty("entries", Nothing) Then Return Nothing
                                                  Dim entries = log.GetProperty("entries")

                                                  Const MAX_MATCHES As Integer = 300

                                                  For Each entryElement In entries.EnumerateArray()
                                                      If ct.IsCancellationRequested Then Return Nothing
                                                      If hit.MatchingRequests.Count >= MAX_MATCHES Then Exit For

                                                      Dim harReq = ParseHarEntry(entryElement)
                                                      If harReq Is Nothing Then Continue For

                                                      ' Check if any field matches the search term
                                                      Dim matchFound As Boolean = False

                                                      ' HTTP Method
                                                      If harReq.Method IsNot Nothing Then
                                                          If exactMatch Then
                                                              If ContainsExactMatch(harReq.Method, term, caseSensitive) Then matchFound = True
                                                          Else
                                                              If harReq.Method.IndexOf(term, comparison) >= 0 Then matchFound = True
                                                          End If
                                                      End If

                                                      ' URL
                                                      If Not matchFound AndAlso harReq.Url IsNot Nothing Then
                                                          If exactMatch Then
                                                              If ContainsExactMatch(harReq.Url, term, caseSensitive) Then matchFound = True
                                                          Else
                                                              If harReq.Url.IndexOf(term, comparison) >= 0 Then matchFound = True
                                                          End If
                                                      End If

                                                      ' Status Code (as string)
                                                      If Not matchFound AndAlso harReq.StatusCode > 0 Then
                                                          Dim statusStr = harReq.StatusCode.ToString()
                                                          If exactMatch Then
                                                              If ContainsExactMatch(statusStr, term, caseSensitive) Then matchFound = True
                                                          Else
                                                              If statusStr.IndexOf(term, comparison) >= 0 Then matchFound = True
                                                          End If
                                                      End If

                                                      ' Status Text
                                                      If Not matchFound AndAlso harReq.StatusText IsNot Nothing Then
                                                          If exactMatch Then
                                                              If ContainsExactMatch(harReq.StatusText, term, caseSensitive) Then matchFound = True
                                                          Else
                                                              If harReq.StatusText.IndexOf(term, comparison) >= 0 Then matchFound = True
                                                          End If
                                                      End If

                                                      ' Request Headers
                                                      If Not matchFound AndAlso harReq.RequestHeaders IsNot Nothing Then
                                                          If exactMatch Then
                                                              If ContainsExactMatch(harReq.RequestHeaders, term, caseSensitive) Then matchFound = True
                                                          Else
                                                              If harReq.RequestHeaders.IndexOf(term, comparison) >= 0 Then matchFound = True
                                                          End If
                                                      End If

                                                      ' Response Headers
                                                      If Not matchFound AndAlso harReq.ResponseHeaders IsNot Nothing Then
                                                          If exactMatch Then
                                                              If ContainsExactMatch(harReq.ResponseHeaders, term, caseSensitive) Then matchFound = True
                                                          Else
                                                              If harReq.ResponseHeaders.IndexOf(term, comparison) >= 0 Then matchFound = True
                                                          End If
                                                      End If

                                                      ' Request Body
                                                      If Not matchFound AndAlso harReq.RequestBody IsNot Nothing Then
                                                          If exactMatch Then
                                                              If ContainsExactMatch(harReq.RequestBody, term, caseSensitive) Then matchFound = True
                                                          Else
                                                              If harReq.RequestBody.IndexOf(term, comparison) >= 0 Then matchFound = True
                                                          End If
                                                      End If

                                                      ' Response Body
                                                      If Not matchFound AndAlso harReq.ResponseBody IsNot Nothing Then
                                                          If exactMatch Then
                                                              If ContainsExactMatch(harReq.ResponseBody, term, caseSensitive) Then matchFound = True
                                                          Else
                                                              If harReq.ResponseBody.IndexOf(term, comparison) >= 0 Then matchFound = True
                                                          End If
                                                      End If

                                                      If matchFound Then
                                                          hit.MatchingRequests.Add(harReq)
                                                      End If
                                                  Next
                                              End Using
                                          End Using

                                          If hit.MatchingRequests.Count > 0 Then Return hit
                                          Return Nothing

                                      End Function)
            Catch ex As OperationCanceledException
                Return Nothing
            Catch ex As TaskCanceledException
                Return Nothing
            Catch ex As Exception
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' Parses a single HAR entry into a HarRequest object
        ''' </summary>
        Private Function ParseHarEntry(entryElement As JsonElement) As HarRequest
            Try
                Dim req As New HarRequest()

                ' Parse request
                If entryElement.TryGetProperty("request", Nothing) Then
                    Dim request = entryElement.GetProperty("request")
                    req.Method = GetJsonString(request, "method")
                    req.Url = GetJsonString(request, "url")

                    ' Parse request headers
                    If request.TryGetProperty("headers", Nothing) Then
                        req.RequestHeaders = FormatHeaders(request.GetProperty("headers"))
                    End If

                    ' Parse request body
                    If request.TryGetProperty("postData", Nothing) Then
                        Dim postData = request.GetProperty("postData")
                        req.RequestBody = GetJsonString(postData, "text")
                    End If
                End If

                ' Parse response
                If entryElement.TryGetProperty("response", Nothing) Then
                    Dim response = entryElement.GetProperty("response")
                    req.StatusCode = GetJsonInt(response, "status")
                    req.StatusText = GetJsonString(response, "statusText")

                    ' Parse response headers
                    If response.TryGetProperty("headers", Nothing) Then
                        req.ResponseHeaders = FormatHeaders(response.GetProperty("headers"))
                    End If

                    ' Parse response body
                    If response.TryGetProperty("content", Nothing) Then
                        Dim content = response.GetProperty("content")
                        req.ResponseBody = GetJsonString(content, "text")
                    End If
                End If

                ' Parse timing
                req.Time = GetJsonDouble(entryElement, "time")

                ' Parse start time
                Dim startedDateTime = GetJsonString(entryElement, "startedDateTime")
                If Not String.IsNullOrEmpty(startedDateTime) Then
                    Dim dt As DateTime
                    If DateTime.TryParse(startedDateTime, dt) Then
                        req.StartedDateTime = dt
                    End If
                End If

                ' Parse server IP
                req.ServerIpAddress = GetJsonString(entryElement, "serverIPAddress")

                Return req
            Catch
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' Safely gets a string property from JSON element
        ''' </summary>
        Private Function GetJsonString(element As JsonElement, propertyName As String) As String
            Try
                If element.TryGetProperty(propertyName, Nothing) Then
                    Dim prop = element.GetProperty(propertyName)
                    If prop.ValueKind = JsonValueKind.String Then
                        Return prop.GetString()
                    End If
                End If
            Catch
            End Try
            Return Nothing
        End Function

        ''' <summary>
        ''' Safely gets an integer property from JSON element
        ''' Handles both numeric and string representations
        ''' </summary>
        Private Function GetJsonInt(element As JsonElement, propertyName As String) As Integer
            Try
                If element.TryGetProperty(propertyName, Nothing) Then
                    Dim prop = element.GetProperty(propertyName)

                    ' Try as number first
                    If prop.ValueKind = JsonValueKind.Number Then
                        Return prop.GetInt32()
                    End If

                    ' Some HAR files might store status as string "404" instead of number 404
                    If prop.ValueKind = JsonValueKind.String Then
                        Dim strValue = prop.GetString()
                        Dim intValue As Integer
                        If Integer.TryParse(strValue, intValue) Then
                            Return intValue
                        End If
                    End If
                End If
            Catch
            End Try
            Return 0
        End Function

        ''' <summary>
        ''' Safely gets a double property from JSON element
        ''' </summary>
        Private Function GetJsonDouble(element As JsonElement, propertyName As String) As Double
            Try
                If element.TryGetProperty(propertyName, Nothing) Then
                    Dim prop = element.GetProperty(propertyName)
                    If prop.ValueKind = JsonValueKind.Number Then
                        Return prop.GetDouble()
                    End If
                End If
            Catch
            End Try
            Return 0
        End Function

        ''' <summary>
        ''' Formats HAR headers array into readable string
        ''' </summary>
        Private Function FormatHeaders(headersArray As JsonElement) As String
            Try
                Dim sb As New System.Text.StringBuilder()
                For Each header In headersArray.EnumerateArray()
                    Dim name = GetJsonString(header, "name")
                    Dim value = GetJsonString(header, "value")
                    If Not String.IsNullOrEmpty(name) Then
                        sb.AppendLine($"{name}: {value}")
                    End If
                Next
                Return sb.ToString()
            Catch
                Return ""
            End Try
        End Function

#End Region

#Region "Preview Selection"

        ''' <summary>
        ''' Handles result selection changes - loads appropriate preview (text or EVTX) and sets up navigation
        ''' </summary>
        Private Sub Results_lst_SelectionChanged(sender As Object, e As SelectionChangedEventArgs)
            Dim hit = TryCast(Results_lst.SelectedItem, SearchHit)
            If hit Is Nothing Then Return

            FindNext_btn.Visibility = Visibility.Collapsed
            FindNextEvent_btn.Visibility = Visibility.Collapsed
            FindPreviousEvent_btn.Visibility = Visibility.Collapsed
            FindNextHarRequest_btn.Visibility = Visibility.Collapsed
            FindPreviousHarRequest_btn.Visibility = Visibility.Collapsed
            _currentTextFindStart = 0
            _currentEventMessageMatchIndex = 0
            _currentHarMatchIndex = 0
            _currentWebMatchIndex = 0
            _totalWebMatches = 0

            ' Load preview based on hit type
            Select Case hit.Kind
                Case HitKind.DiskTextFile
                    ' ROUTE BASED ON EXTENSION
                    Dim ext = Path.GetExtension(hit.FilePath).ToLowerInvariant()

                    If ext = ".html" OrElse ext = ".xml" OrElse ext = ".json" Then
                        ' Use WebView2 engine for HTML/XML/JSON (will initialize if needed)
                        ShowWebPreviewMode()
                        LoadWebContentFromDisk(hit.FilePath, ext)
                        FindNext_btn.Visibility = Visibility.Visible
                    Else
                        ' Use RichTextBox engine for plain text
                        ShowTextPreviewMode()
                        LoadTextFromDisk(hit.FilePath)
                        FindNext_btn.Visibility = Visibility.Visible
                    End If

                Case HitKind.ZipTextEntry
                    Dim ext = Path.GetExtension(hit.ZipEntryName).ToLowerInvariant()

                    If ext = ".html" OrElse ext = ".xml" OrElse ext = ".json" Then
                        ' Use WebView2 engine for HTML/XML/JSON (will initialize if needed)
                        ShowWebPreviewMode()
                        LoadWebContentFromArchive(hit.ZipPath, hit.ZipEntryName, ext)
                        FindNext_btn.Visibility = Visibility.Visible
                    Else
                        ' Use RichTextBox engine for plain text
                        ShowTextPreviewMode()
                        LoadTextFromArchive(hit.ZipPath, hit.ZipEntryName)
                        FindNext_btn.Visibility = Visibility.Visible
                    End If

                Case HitKind.CabTextEntry
                    Dim ext = Path.GetExtension(hit.ZipEntryName).ToLowerInvariant()

                    If ext = ".html" OrElse ext = ".xml" OrElse ext = ".json" Then
                        ' Use WebView2 engine for HTML/XML/JSON (will initialize if needed)
                        ShowWebPreviewMode()
                        LoadWebContentFromCab(hit.ZipPath, hit.ZipEntryName, ext)
                        FindNext_btn.Visibility = Visibility.Visible
                    Else
                        ' Use RichTextBox engine for plain text
                        ShowTextPreviewMode()
                        LoadTextFromCab(hit.ZipPath, hit.ZipEntryName)
                        FindNext_btn.Visibility = Visibility.Visible
                    End If

                Case HitKind.EvtxFileOnDisk, HitKind.EvtxFromZipTemp, HitKind.EvtxFromCabTemp
                    ShowEventPreviewMode()
                    hit.CurrentEventIndex = -1
                    RenderFirstEvent(hit)
                    FindNextEvent_btn.Visibility = Visibility.Visible
                    FindPreviousEvent_btn.Visibility = Visibility.Visible

                Case HitKind.HarFileOnDisk, HitKind.HarFromZip
                    ShowHarPreviewMode()
                    hit.CurrentRequestIndex = -1
                    RenderFirstHarRequest(hit)
                    FindNextHarRequest_btn.Visibility = Visibility.Visible
                    FindPreviousHarRequest_btn.Visibility = Visibility.Visible
            End Select

            ' IMPORTANT: If scanning, do NOT change status to Ready when user clicks results
            If Not _isScanning Then
                Status("Ready")
            End If
        End Sub

        ''' <summary>Loads text file from disk into preview pane</summary>
        Private Sub LoadTextFromDisk(path As String)
            Dim text = File.ReadAllText(path)
            SetTextPreview(text)
        End Sub

        ''' <summary>
        ''' Extracts text entry from archive and loads into preview pane
        ''' Supports ZIP, 7Z, RAR, TAR, and other formats via SharpCompress
        ''' </summary>
        Private Sub LoadTextFromArchive(archivePath As String, entryName As String)
            Try
                Using archive = OpenArchive(archivePath)
                    For Each entry In archive.Entries
                        If entry.Key = entryName Then
                            Dim stream As Stream = entry.OpenEntryStream()
                            Using stream
                                Using sr As New StreamReader(stream, detectEncodingFromByteOrderMarks:=True)
                                    SetTextPreview(sr.ReadToEnd())
                                End Using
                            End Using
                            Return
                        End If
                    Next
                    ' Entry not found
                    SetTextPreview("[Entry not found]")
                End Using
            Catch ex As Exception
                SetTextPreview($"[Error loading from archive: {ex.Message}]")
            End Try
        End Sub

        ''' <summary>
        ''' Extracts text entry from CAB archive and loads into preview pane
        ''' Uses 7-Zip command line for extraction
        ''' </summary>
        Private Sub LoadTextFromCab(cabPath As String, entryName As String)
            Dim tempExtractDir As String = Nothing
            Try
                ' Extract entire CAB to temp directory using 7-Zip
                tempExtractDir = Path.Combine(Path.GetTempPath(), "BeaconCabPreview_" & Guid.NewGuid().ToString("N"))

                Dim success = SevenZipHelper.ExtractCab(cabPath, tempExtractDir)
                If Not success Then
                    SetTextPreview($"[Error: Failed to extract CAB file]")
                    Return
                End If

                ' Find the extracted file
                Dim extractedFile = Path.Combine(tempExtractDir, entryName)
                If File.Exists(extractedFile) Then
                    Dim text = File.ReadAllText(extractedFile)
                    SetTextPreview(text)
                Else
                    SetTextPreview($"[Error: File '{entryName}' not found in extracted CAB contents]")
                End If
            Catch ex As Exception
                SetTextPreview($"[Error loading CAB preview: {ex.Message}]")
            Finally
                ' Cleanup temp directory
                If tempExtractDir IsNot Nothing AndAlso Directory.Exists(tempExtractDir) Then
                    Try
                        Directory.Delete(tempExtractDir, True)
                    Catch
                        ' Ignore cleanup errors
                    End Try
                End If
            End Try
        End Sub

        ''' <summary>
        ''' Sets preview text and automatically highlights first search term occurrence
        ''' Uses Dispatcher.BeginInvoke to ensure document is fully rendered before highlighting
        ''' </summary>
        Private Sub SetTextPreview(text As String)

            TextPreview_rtb.Document = New FlowDocument(
        New Paragraph(New Run(text))
    )

            _currentTextFindStart = 0

            Dispatcher.BeginInvoke(
        Sub()
            FindNextInText()
        End Sub,
        System.Windows.Threading.DispatcherPriority.Loaded)

        End Sub


        Private Sub ClearTextPreview()
            TextPreview_rtb.Document.Blocks.Clear()
        End Sub

        ''' <summary>
        ''' Clears WebView2 content to prevent black screen or lingering content
        ''' </summary>
        Private Sub ClearWebPreview()
            If _webViewInitialized Then
                Try
                    WebPreview_wv2.NavigateToString("<html><body style='margin:0;padding:0;background:white;'></body></html>")
                Catch
                    ' Ignore errors clearing WebView2 preview
                End Try
            End If
        End Sub

        ''' <summary>
        ''' Loads HTML or XML file from disk into WebView2 with search highlighting
        ''' </summary>
        Private Async Sub LoadWebContentFromDisk(filePath As String, extension As String)
            ' Ensure WebView2 is initialized (will wait if needed)
            Dim isReady = Await EnsureWebView2InitializedAsync()

            If Not isReady Then
                ' Fallback to text preview if WebView2 not available
                ShowTextPreviewMode()
                LoadTextFromDisk(filePath)
                Status("HTML/XML shown as text - WebView2 unavailable")
                Return
            End If

            Try
                ' Read file content and wrap with theme-aware CSS
                Dim content = File.ReadAllText(filePath)
                Await LoadWebContentAsync(content, extension)
            Catch ex As Exception
                Debug.WriteLine($"Error loading content: {ex.Message}")
                ' On error, show error message in WebView
                WebPreview_wv2.NavigateToString($"<html><body><h3>Error loading file:</h3><pre>{System.Security.SecurityElement.Escape(ex.Message)}</pre></body></html>")
            End Try
        End Sub

        ''' <summary>
        ''' Loads HTML or XML entry from archive into WebView2 with search highlighting
        ''' Supports ZIP, 7Z, RAR, TAR, and other formats via SharpCompress
        ''' </summary>
        Private Async Sub LoadWebContentFromArchive(archivePath As String, entryName As String, extension As String)
            ' Ensure WebView2 is initialized (will wait if needed)
            Dim isReady = Await EnsureWebView2InitializedAsync()

            If Not isReady Then
                ' Fallback to text preview if WebView2 not available
                ShowTextPreviewMode()
                LoadTextFromArchive(archivePath, entryName)
                Status("HTML/XML shown as text - WebView2 unavailable")
                Return
            End If

            Try
                Using archive = OpenArchive(archivePath)
                    For Each entry In archive.Entries
                        If entry.Key = entryName Then
                            Dim stream As Stream = entry.OpenEntryStream()
                            Using stream
                                Using sr As New StreamReader(stream, detectEncodingFromByteOrderMarks:=True)
                                    Dim content = sr.ReadToEnd()
                                    Await LoadWebContentAsync(content, extension)
                                End Using
                            End Using
                            Return
                        End If
                    Next
                    ' Entry not found
                    WebPreview_wv2.NavigateToString("<html><body><h3>Error: Entry not found in archive</h3></body></html>")
                End Using
            Catch ex As Exception
                WebPreview_wv2.NavigateToString($"<html><body><h3>Error loading from archive:</h3><pre>{System.Security.SecurityElement.Escape(ex.Message)}</pre></body></html>")
            End Try
        End Sub

        ''' <summary>
        ''' Loads HTML/XML/JSON entry from CAB archive into WebView2 with search highlighting
        ''' Uses 7-Zip command line for extraction
        ''' </summary>
        Private Async Sub LoadWebContentFromCab(cabPath As String, entryName As String, extension As String)
            ' Ensure WebView2 is initialized (will wait if needed)
            Dim isReady = Await EnsureWebView2InitializedAsync()

            If Not isReady Then
                ' Fallback to text preview if WebView2 not available
                ShowTextPreviewMode()
                LoadTextFromCab(cabPath, entryName)
                Status("HTML/XML shown as text - WebView2 unavailable")
                Return
            End If

            Dim tempExtractDir As String = Nothing
            Try
                ' Extract entire CAB to temp directory using 7-Zip
                tempExtractDir = Path.Combine(Path.GetTempPath(), "BeaconCabPreview_" & Guid.NewGuid().ToString("N"))

                Dim success = SevenZipHelper.ExtractCab(cabPath, tempExtractDir)
                If Not success Then
                    WebPreview_wv2.NavigateToString($"<html><body><h3>Error: Failed to extract CAB file</h3></body></html>")
                    Return
                End If

                ' Find the extracted file
                Dim extractedFile = Path.Combine(tempExtractDir, entryName)
                If File.Exists(extractedFile) Then
                    Dim content = File.ReadAllText(extractedFile)
                    Await LoadWebContentAsync(content, extension)
                Else
                    WebPreview_wv2.NavigateToString($"<html><body><h3>Error: File '{System.Security.SecurityElement.Escape(entryName)}' not found in extracted CAB contents</h3></body></html>")
                End If
            Catch ex As Exception
                WebPreview_wv2.NavigateToString($"<html><body><h3>Error loading CAB preview:</h3><pre>{System.Security.SecurityElement.Escape(ex.Message)}</pre></body></html>")
            Finally
                ' Cleanup temp directory
                If tempExtractDir IsNot Nothing AndAlso Directory.Exists(tempExtractDir) Then
                    Try
                        Directory.Delete(tempExtractDir, True)
                    Catch
                        ' Ignore cleanup errors
                    End Try
                End If
            End Try
        End Sub

        ''' <summary>
        ''' Core method to load and highlight HTML/XML/JSON content in WebView2
        ''' </summary>
        Private Async Function LoadWebContentAsync(content As String, extension As String) As Task
            Dim htmlToRender As String

            If extension = ".xml" Then
                ' Wrap XML in styled HTML viewer
                htmlToRender = CreateXmlViewerHtml(content)
            ElseIf extension = ".json" Then
                ' Wrap JSON in styled HTML viewer with syntax highlighting
                htmlToRender = CreateJsonViewerHtml(content)
            Else
                ' HTML content - inject theme-aware CSS to ensure good contrast
                htmlToRender = InjectThemeAwareCSS(content)
            End If

            ' Navigate to content
            WebPreview_wv2.NavigateToString(htmlToRender)

            ' Wait for navigation to complete before highlighting
            Await Task.Delay(300) ' Small delay to ensure DOM is ready

            ' Apply search highlighting
            Await HighlightSearchInWebViewAsync()
        End Function

        ''' <summary>
        ''' Injects theme-aware CSS into HTML content to ensure good contrast in both light and dark modes
        ''' </summary>
        Private Function InjectThemeAwareCSS(htmlContent As String) As String
            ' Determine colors based on current theme
            Dim bgColor = If(_isDarkMode, "#1E1E1E", "#FFFFFF")
            Dim textColor = If(_isDarkMode, "#E0E0E0", "#202020")
            Dim linkColor = If(_isDarkMode, "#60CFFF", "#0066CC")
            Dim borderColor = If(_isDarkMode, "#3F3F3F", "#CCCCCC")

            ' CSS to inject - uses !important to override existing styles
            Dim themeCSS = $"
<style id='beacon-theme-override'>
    /* Force readable colors for better contrast */
    body {{
        background-color: {bgColor} !important;
        color: {textColor} !important;
    }}

    /* Ensure all text elements have good contrast */
    div, span, p, td, th, li, dt, dd, label, h1, h2, h3, h4, h5, h6 {{
        color: {textColor} !important;
    }}

    /* Style tables for better visibility */
    table {{
        border-color: {borderColor} !important;
    }}

    td, th {{
        border-color: {borderColor} !important;
        background-color: transparent !important;
    }}

    /* Style links */
    a {{
        color: {linkColor} !important;
    }}

    /* Ensure input fields are visible */
    input, textarea, select {{
        background-color: {If(_isDarkMode, "#2B2B2B", "#FFFFFF")} !important;
        color: {textColor} !important;
        border-color: {borderColor} !important;
    }}

    /* Style code blocks */
    pre, code {{
        background-color: {If(_isDarkMode, "#252525", "#F5F5F5")} !important;
        color: {textColor} !important;
        border-color: {borderColor} !important;
    }}

    /* Search highlighting (maintain high visibility) */
    mark.search-highlight {{
        background-color: #FFFF00 !important;
        color: #000 !important;
        font-weight: bold;
        padding: 2px 0;
    }}

    mark.current-highlight {{
        background-color: #FF9500 !important;
        color: #000 !important;
        font-weight: bold;
        padding: 2px 0;
    }}
</style>
"

            ' Try to inject CSS into <head> if it exists, otherwise prepend to content
            Dim headEndIndex = htmlContent.IndexOf("</head>", StringComparison.OrdinalIgnoreCase)

            If headEndIndex > 0 Then
                ' Insert before </head>
                Return htmlContent.Insert(headEndIndex, themeCSS)
            Else
                ' Check if <html> tag exists
                Dim htmlStartIndex = htmlContent.IndexOf("<html", StringComparison.OrdinalIgnoreCase)

                If htmlStartIndex >= 0 Then
                    ' Find the end of <html> tag
                    Dim htmlTagEndIndex = htmlContent.IndexOf(">", htmlStartIndex)

                    If htmlTagEndIndex > 0 Then
                        ' Insert <head> with CSS after <html>
                        Dim headSection = $"<head><meta charset='utf-8'>{themeCSS}</head>"
                        Return htmlContent.Insert(htmlTagEndIndex + 1, headSection)
                    End If
                End If

                ' No proper HTML structure, wrap entire content
                Return $"<!DOCTYPE html><html><head><meta charset='utf-8'>{themeCSS}</head><body>{htmlContent}</body></html>"
            End If
        End Function

        ''' <summary>
        ''' Creates HTML wrapper for XML content with syntax highlighting and search support
        ''' Respects current theme (dark/light mode)
        ''' </summary>
        Private Function CreateXmlViewerHtml(xmlContent As String) As String
            ' Escape XML for safe display in HTML
            Dim escapedXml = System.Security.SecurityElement.Escape(xmlContent)

            ' Adapt styling based on current theme
            ' Theme-aware colors
            Dim bgColor = If(_isDarkMode, "#1E1E1E", "white")
            Dim textColor = If(_isDarkMode, "#D4D4D4", "#000")

            Return $"<!DOCTYPE html>
<html>
<head>
    <meta charset='utf-8'>
    <style>
        body {{
            margin: 0;
            padding: 12px;
            font-family: Consolas, 'Courier New', monospace;
            font-size: 13px;
            background: {bgColor};
            color: {textColor};
        }}
        .xml-content {{
            white-space: pre-wrap;
            word-wrap: break-word;
            line-height: 1.5;
        }}
        mark.search-highlight {{
            background-color: #FFFF00;
            color: #000;
            font-weight: bold;
            padding: 2px 0;
        }}
        mark.current-highlight {{
            background-color: #FF9500;
            color: #000;
            font-weight: bold;
            padding: 2px 0;
        }}
    </style>
</head>
<body>
    <div class='xml-content' id='content'>{escapedXml}</div>
</body>
</html>"
        End Function

        ''' <summary>
        ''' Creates HTML wrapper for JSON content with syntax highlighting and search support
        ''' </summary>
        Private Function CreateJsonViewerHtml(jsonContent As String) As String
            ' Pretty-print JSON if it appears to be minified
            Dim formattedJson As String
            Try
                ' Try to parse and format JSON
                Dim jsonObj = System.Text.Json.JsonDocument.Parse(jsonContent)
                Dim options As New System.Text.Json.JsonSerializerOptions With {
                    .WriteIndented = True
                }
                formattedJson = System.Text.Json.JsonSerializer.Serialize(jsonObj, options)
            Catch
                ' If parsing fails, use original content
                formattedJson = jsonContent
            End Try

            ' Escape for HTML display
            Dim escapedJson = System.Security.SecurityElement.Escape(formattedJson)

            ' Apply basic syntax highlighting using HTML/CSS with theme-aware colors
            If _isDarkMode Then
                ' Dark mode: Use VS Code dark theme colors
                escapedJson = System.Text.RegularExpressions.Regex.Replace(escapedJson, """([^""]+)""\s*:", "<span style='color:#9CDCFE'>""$1""</span>:") ' Keys (light blue)
                escapedJson = System.Text.RegularExpressions.Regex.Replace(escapedJson, """([^""]+)""", "<span style='color:#CE9178'>""$1""</span>") ' String values (orange)
                escapedJson = System.Text.RegularExpressions.Regex.Replace(escapedJson, "\b(\d+\.?\d*)\b", "<span style='color:#B5CEA8'>$1</span>") ' Numbers (light green)
                escapedJson = System.Text.RegularExpressions.Regex.Replace(escapedJson, "\b(true|false|null)\b", "<span style='color:#569CD6'>$1</span>") ' Booleans/null (blue)
            Else
                ' Light mode: Use darker colors for better contrast
                escapedJson = System.Text.RegularExpressions.Regex.Replace(escapedJson, """([^""]+)""\s*:", "<span style='color:#0451A5'>""$1""</span>:") ' Keys (blue)
                escapedJson = System.Text.RegularExpressions.Regex.Replace(escapedJson, """([^""]+)""", "<span style='color:#A31515'>""$1""</span>") ' String values (red)
                escapedJson = System.Text.RegularExpressions.Regex.Replace(escapedJson, "\b(\d+\.?\d*)\b", "<span style='color:#098658'>$1</span>") ' Numbers (green)
                escapedJson = System.Text.RegularExpressions.Regex.Replace(escapedJson, "\b(true|false|null)\b", "<span style='color:#0000FF'>$1</span>") ' Booleans/null (blue)
            End If

            ' Adapt styling based on current theme
            Dim bgColor = If(_isDarkMode, "#1E1E1E", "white")
            Dim textColor = If(_isDarkMode, "#D4D4D4", "#000")

            Return $"<!DOCTYPE html>
<html>
<head>
    <meta charset='utf-8'>
    <style>
        body {{
            margin: 0;
            padding: 12px;
            font-family: Consolas, 'Courier New', monospace;
            font-size: 13px;
            background: {bgColor};
            color: {textColor};
        }}
        .json-content {{
            white-space: pre-wrap;
            word-wrap: break-word;
            line-height: 1.5;
        }}
        mark.search-highlight {{
            background-color: #FFFF00;
            color: #000;
            font-weight: bold;
            padding: 2px 0;
        }}
        mark.current-highlight {{
            background-color: #FF9500;
            color: #000;
            font-weight: bold;
            padding: 2px 0;
        }}
    </style>
</head>
<body>
    <div class='json-content' id='content'>{escapedJson}</div>
</body>
</html>"
        End Function

        ''' <summary>
        ''' Injects JavaScript to highlight all search term occurrences in WebView2
        ''' </summary>
        Private Async Function HighlightSearchInWebViewAsync() As Task
            Dim searchTerm = Search_txt.Text
            If String.IsNullOrWhiteSpace(searchTerm) OrElse Not _webViewInitialized Then
                _totalWebMatches = 0
                Return
            End If

            Dim caseSensitive = CaseSensitive_chk.IsChecked.GetValueOrDefault(False)

            ' Escape JavaScript string and regex special characters
            Dim escapedTerm = searchTerm.Replace("\", "\\").Replace("'", "\'").Replace(vbLf, "\n").Replace(vbCr, "\r")
            Dim regexFlags = If(caseSensitive, "g", "gi")

            ' Build regex pattern, escaping special regex characters
            Dim regexEscapedTerm = System.Text.RegularExpressions.Regex.Escape(searchTerm)

            ' JavaScript to highlight all matches and count them
            Dim script = $"
(function() {{
    // Remove existing highlights
    document.querySelectorAll('mark.search-highlight, mark.current-highlight').forEach(function(el) {{
        var parent = el.parentNode;
        parent.replaceChild(document.createTextNode(el.textContent), el);
        parent.normalize();
    }});

    var body = document.body;
    var searchTerm = '{escapedTerm}';
    var regex = new RegExp('({regexEscapedTerm})', '{regexFlags}');
    var matchCount = 0;

    // Recursive function to traverse and highlight text nodes
    function highlightInNode(node) {{
        if (node.nodeType === Node.TEXT_NODE) {{
            var text = node.textContent;
            var matches = text.match(regex);

            if (matches && matches.length > 0) {{
                var fragment = document.createDocumentFragment();
                var lastIndex = 0;
                var tempText = text;

                tempText.replace(regex, function(match, ...args) {{
                    var offset = args[args.length - 2]; // offset is second to last argument

                    // Add text before match
                    if (offset > lastIndex) {{
                        fragment.appendChild(document.createTextNode(text.substring(lastIndex, offset)));
                    }}

                    // Add highlighted match
                    var mark = document.createElement('mark');
                    mark.className = 'search-highlight';
                    mark.setAttribute('data-match-index', matchCount);
                    mark.textContent = match;
                    fragment.appendChild(mark);
                    matchCount++;

                    lastIndex = offset + match.length;
                    return match;
                }});

                // Add remaining text
                if (lastIndex < text.length) {{
                    fragment.appendChild(document.createTextNode(text.substring(lastIndex)));
                }}

                node.parentNode.replaceChild(fragment, node);
            }}
        }} else if (node.nodeType === Node.ELEMENT_NODE && node.tagName !== 'SCRIPT' && node.tagName !== 'STYLE' && node.tagName !== 'MARK') {{
            Array.from(node.childNodes).forEach(highlightInNode);
        }}
    }}

    highlightInNode(body);
    return matchCount;
}})();
"

            Try
                Dim result = Await WebPreview_wv2.ExecuteScriptAsync(script)
                ' Parse match count from result
                If Integer.TryParse(result, _totalWebMatches) Then
                    _currentWebMatchIndex = 0
                    If _totalWebMatches > 0 Then
                        ' Highlight first match
                        Await HighlightWebMatchAtIndexAsync(0)
                    End If
                End If
            Catch ex As Exception
                _totalWebMatches = 0
            End Try
        End Function

        ''' <summary>
        ''' Highlights specific match index and scrolls it into view
        ''' </summary>
        Private Async Function HighlightWebMatchAtIndexAsync(index As Integer) As Task
            If Not _webViewInitialized OrElse _totalWebMatches = 0 Then Return

            Dim script = $"
(function() {{
    var marks = document.querySelectorAll('mark.search-highlight, mark.current-highlight');

    // Remove current highlight from all marks
    marks.forEach(function(mark) {{
        mark.className = 'search-highlight';
    }});

    // Highlight the current match
    if (marks.length > {index}) {{
        var currentMark = marks[{index}];
        currentMark.className = 'current-highlight';
        currentMark.scrollIntoView({{ behavior: 'smooth', block: 'center' }});
        return true;
    }}
    return false;
}})();
"

            Try
                Await WebPreview_wv2.ExecuteScriptAsync(script)
            Catch
            End Try
        End Function

        ''' <summary>
        ''' Injects CSS to override HTML page colors for better visibility in dark mode
        ''' Only applies when in dark mode to prevent breaking light-themed HTML
        ''' </summary>
        Private Async Function InjectThemeOverrideCssAsync() As Task
            If Not _webViewInitialized Then Return

            ' Only inject dark mode overrides when in dark mode
            If Not _isDarkMode Then Return

            Dim script = "
(function() {
    // Inject CSS to override page colors for dark mode readability
    var style = document.createElement('style');
    style.id = 'beacon-dark-mode-override';
    style.textContent = `
        /* Force readable colors in dark mode */
        body {
            background-color: #1E1E1E !important;
            color: #D4D4D4 !important;
        }

        /* Override common text elements */
        p, div, span, td, th, li, a, h1, h2, h3, h4, h5, h6 {
            color: #D4D4D4 !important;
        }

        /* Override links for visibility */
        a {
            color: #60CFFF !important;
        }

        /* Override tables */
        table {
            border-color: #3F3F3F !important;
        }

        /* Keep search highlights visible */
        mark.search-highlight {
            background-color: #FFFF00 !important;
            color: #000 !important;
        }

        mark.current-highlight {
            background-color: #FF9500 !important;
            color: #000 !important;
        }
    `;

    // Remove existing override if present
    var existing = document.getElementById('beacon-dark-mode-override');
    if (existing) {
        existing.remove();
    }

    document.head.appendChild(style);
    return true;
})();
"

            Try
                Await WebPreview_wv2.ExecuteScriptAsync(script)
                Debug.WriteLine("✓ Dark mode CSS override injected")
            Catch ex As Exception
                Debug.WriteLine($"Failed to inject dark mode CSS: {ex.Message}")
            End Try
        End Function

#End Region

#Region "Find Next (Text)"

        ''' <summary>Navigate to previous event in EVTX preview</summary>
        Private Sub FindPreviousEvent_btn_Click(sender As Object, e As RoutedEventArgs)
            FindPreviousEvent()
        End Sub

        Private Sub FindNext_btn_Click(sender As Object, e As RoutedEventArgs)
            FindNextInText()
        End Sub

        ''' <summary>
        ''' Finds and highlights next occurrence of search term in text preview or WebView2
        ''' Wraps to beginning if no more matches found
        ''' </summary>
        Private Sub FindNextInText()
            ' Check if we're in WebView2 mode
            If WebPreview_grp.Visibility = Visibility.Visible AndAlso _webViewInitialized Then
                FindNextInWebView()
                Return
            End If

            ' Original text preview logic
            Dim term = Search_txt.Text
            If String.IsNullOrWhiteSpace(term) Then Return

            Dim full = New TextRange(TextPreview_rtb.Document.ContentStart, TextPreview_rtb.Document.ContentEnd).Text
            If String.IsNullOrEmpty(full) Then Return

            Dim cs = CaseSensitive_chk.IsChecked.GetValueOrDefault(False)
            Dim comparison = If(cs, StringComparison.Ordinal, StringComparison.OrdinalIgnoreCase)

            Dim idx = full.IndexOf(term, _currentTextFindStart, comparison)
            Dim wrapped As Boolean = False

            ' If not found from current position, try from beginning (wrap-around)
            If idx < 0 Then
                idx = full.IndexOf(term, 0, comparison)
                If idx < 0 Then
                    If Not _isScanning Then Status("No matches in preview")
                    Return
                End If
                wrapped = True
            End If

            SelectInRichTextBox(idx, term.Length)
            _currentTextFindStart = idx + term.Length

            If Not _isScanning Then
                Status(If(wrapped, "Wrapped to top", "Ready"))
            End If
        End Sub

        ''' <summary>
        ''' Navigates to next match in WebView2 content
        ''' </summary>
        Private Async Sub FindNextInWebView()
            If _totalWebMatches = 0 Then
                If Not _isScanning Then Status("No matches in preview")
                Return
            End If

            _currentWebMatchIndex += 1

            Dim wrapped As Boolean = False
            If _currentWebMatchIndex >= _totalWebMatches Then
                _currentWebMatchIndex = 0
                wrapped = True
            End If

            Await HighlightWebMatchAtIndexAsync(_currentWebMatchIndex)

            If Not _isScanning Then
                Status(If(wrapped, $"Wrapped to top (match {_currentWebMatchIndex + 1} of {_totalWebMatches})", $"Match {_currentWebMatchIndex + 1} of {_totalWebMatches}"))
            End If
        End Sub

        ''' <summary>
        ''' Selects text range in RichTextBox and scrolls into view
        ''' Uses TextPointer navigation for character-based positioning
        ''' </summary>
        Private Sub SelectInRichTextBox(startIndex As Integer, length As Integer)

            Dim startPtr = GetTextPointerAt(TextPreview_rtb.Document.ContentStart, startIndex)
            Dim endPtr = GetTextPointerAt(startPtr, length)

            If startPtr Is Nothing OrElse endPtr Is Nothing Then Return

            ' Normalize pointers to valid insertion positions (avoids selection errors)
            startPtr = startPtr.GetInsertionPosition(LogicalDirection.Forward)
            endPtr = endPtr.GetInsertionPosition(LogicalDirection.Backward)

            TextPreview_rtb.BeginChange()

            ' Collapse selection first (workaround for RichTextBox quirks)
            TextPreview_rtb.Selection.Select(startPtr, startPtr)
            TextPreview_rtb.CaretPosition = startPtr

            ' Apply actual selection
            TextPreview_rtb.Selection.Select(startPtr, endPtr)
            TextPreview_rtb.CaretPosition = endPtr

            TextPreview_rtb.EndChange()

            TextPreview_rtb.Focus()

            ' Re-apply selection at Render priority to ensure visibility
            Dispatcher.BeginInvoke(
        Sub()
            TextPreview_rtb.Selection.Select(startPtr, endPtr)
        End Sub,
        System.Windows.Threading.DispatcherPriority.Render)

            Dim p = startPtr.Paragraph
            If p IsNot Nothing Then p.BringIntoView()

        End Sub

        ''' <summary>
        ''' Navigates TextPointer by character offset, accounting for non-text content
        ''' </summary>
        Private Function GetTextPointerAt(start As TextPointer, charOffset As Integer) As TextPointer
            Dim nav = start
            Dim count = 0

            While nav IsNot Nothing
                If nav.GetPointerContext(LogicalDirection.Forward) = TextPointerContext.Text Then
                    Dim textRun = nav.GetTextInRun(LogicalDirection.Forward)
                    If count + textRun.Length >= charOffset Then
                        Return nav.GetPositionAtOffset(charOffset - count)
                    End If
                    count += textRun.Length
                End If

                Dim nextPtr = nav.GetNextContextPosition(LogicalDirection.Forward)
                If nextPtr Is Nothing Then Exit While
                nav = nextPtr
            End While

            Return TextPreview_rtb.Document.ContentEnd
        End Function

#End Region

#Region "EVTX Preview + Find Next Event"

        ''' <summary>
        ''' Navigates to previous occurrence of search term in EVTX events
        ''' Searches within current event first, then moves to previous event
        ''' </summary>
        Private Sub FindPreviousEvent()
            Dim hit = TryCast(Results_lst.SelectedItem, SearchHit)
            If hit Is Nothing OrElse hit.MatchingEvents.Count = 0 Then Return

            Dim currentEvent = hit.MatchingEvents(hit.CurrentEventIndex)
            Dim term = Search_txt.Text
            Dim cs = CaseSensitive_chk.IsChecked.GetValueOrDefault(False)
            Dim comparison = If(cs, StringComparison.Ordinal, StringComparison.OrdinalIgnoreCase)

            ' Try to find previous match within current event message
            If _currentEventMessageMatchIndex > term.Length Then
                ' Search backwards from before the current match
                Dim searchUpTo = _currentEventMessageMatchIndex - term.Length - 1
                Dim prevMatchIndex = currentEvent.Message.LastIndexOf(term, searchUpTo, searchUpTo + 1, comparison)

                If prevMatchIndex >= 0 Then
                    _currentEventMessageMatchIndex = prevMatchIndex + term.Length
                    RenderEventWithHighlight(currentEvent, prevMatchIndex, term.Length)
                    UpdateEventCounter(hit)
                    If Not _isScanning Then Status("Ready")
                    Return
                End If
            End If

            ' No previous match in current event, move to previous event
            hit.CurrentEventIndex -= 1
            If hit.CurrentEventIndex < 0 Then
                hit.CurrentEventIndex = hit.MatchingEvents.Count - 1
                If Not _isScanning Then Status("Wrapped to last event")
            Else
                If Not _isScanning Then Status("Ready")
            End If

            ' Start at the end of the new event for reverse search
            Dim newEvent = hit.MatchingEvents(hit.CurrentEventIndex)
            _currentEventMessageMatchIndex = newEvent.Message.Length
            RenderEvent(newEvent)
            UpdateEventCounter(hit)
        End Sub

        Private Sub FindNextEvent_btn_Click(sender As Object, e As RoutedEventArgs)
            FindNextEvent()
        End Sub

        ''' <summary>
        ''' Renders first event in list and highlights first search term occurrence
        ''' </summary>
        Private Sub RenderFirstEvent(hit As SearchHit)
            If hit.MatchingEvents.Count = 0 Then
                EventLevel_txt.Text = "[No matches]"
                EventId_txt.Text = ""
                EventProvider_txt.Text = ""
                EventTime_txt.Text = ""
                EventMessage_txt.Inlines.Clear()
                EventMatchCounter_lbl.Text = "0 match(es) in this event"
                Return
            End If

            hit.CurrentEventIndex = 0
            _currentEventMessageMatchIndex = 0
            RenderEvent(hit.MatchingEvents(0))
            UpdateEventCounter(hit)
        End Sub

        ''' <summary>
        ''' Navigates to next occurrence of search term in EVTX events
        ''' Searches within current event first, then moves to next event
        ''' Wraps to first event when reaching end
        ''' </summary>
        Private Sub FindNextEvent()
            Dim hit = TryCast(Results_lst.SelectedItem, SearchHit)
            If hit Is Nothing OrElse hit.MatchingEvents.Count = 0 Then Return

            Dim currentEvent = hit.MatchingEvents(hit.CurrentEventIndex)
            Dim term = Search_txt.Text
            Dim cs = CaseSensitive_chk.IsChecked.GetValueOrDefault(False)
            Dim comparison = If(cs, StringComparison.Ordinal, StringComparison.OrdinalIgnoreCase)

            ' Try to find next match within the current event message
            Dim nextMatchIndex = currentEvent.Message.IndexOf(term, _currentEventMessageMatchIndex, comparison)

            If nextMatchIndex >= 0 Then
                ' Found another match in the same event
                _currentEventMessageMatchIndex = nextMatchIndex + term.Length
                RenderEventWithHighlight(currentEvent, nextMatchIndex, term.Length)
                UpdateEventCounter(hit)
                If Not _isScanning Then Status("Ready")
                Return
            End If

            ' No more matches in current event, move to next event
            hit.CurrentEventIndex += 1
            If hit.CurrentEventIndex >= hit.MatchingEvents.Count Then
                hit.CurrentEventIndex = 0
                If Not _isScanning Then Status("Wrapped to first event")
            Else
                If Not _isScanning Then Status("Ready")
            End If

            _currentEventMessageMatchIndex = 0
            RenderEvent(hit.MatchingEvents(hit.CurrentEventIndex))
            UpdateEventCounter(hit)
        End Sub

        ''' <summary>
        ''' Renders event details and highlights first search term occurrence in message
        ''' </summary>
        Private Sub RenderEvent(ev As EventSummary)
            Dim lvl = If(String.IsNullOrWhiteSpace(ev.Level), "Information", ev.Level)
            EventLevel_txt.Text = $"[{lvl}]"
            EventId_txt.Text = $"Event ID {ev.EventId}"
            EventProvider_txt.Text = ev.Provider
            EventTime_txt.Text = If(ev.TimeCreated.HasValue, ev.TimeCreated.Value.ToString("yyyy-MM-dd HH:mm:ss"), "")

            ' Highlight the first occurrence of the search term
            Dim term = Search_txt.Text
            If String.IsNullOrWhiteSpace(term) Then
                EventMessage_txt.Inlines.Clear()
                EventMessage_txt.Inlines.Add(New Run(ev.Message))
                _currentEventMessageMatchIndex = 0
                EventMatchCounter_lbl.Text = "0 match(es) in this event"
                Return
            End If

            Dim cs = CaseSensitive_chk.IsChecked.GetValueOrDefault(False)
            Dim comparison = If(cs, StringComparison.Ordinal, StringComparison.OrdinalIgnoreCase)

            ' Count total matches in this event message
            Dim totalMatches = CountMatchesInString(ev.Message, term, comparison)
            EventMatchCounter_lbl.Text = $"{totalMatches} match(es) in this event"

            Dim firstMatchIndex = ev.Message.IndexOf(term, comparison)

            If firstMatchIndex >= 0 Then
                _currentEventMessageMatchIndex = firstMatchIndex + term.Length
                RenderEventWithHighlight(ev, firstMatchIndex, term.Length)
            Else
                EventMessage_txt.Inlines.Clear()
                EventMessage_txt.Inlines.Add(New Run(ev.Message))
                _currentEventMessageMatchIndex = 0
            End If
        End Sub

        ''' <summary>
        ''' Renders event message with inline highlighting of specific occurrence
        ''' Uses TextBlock.Inlines with colored Run for highlighted text
        ''' </summary>
        Private Sub RenderEventWithHighlight(ev As EventSummary, highlightStart As Integer, highlightLength As Integer)
            EventMessage_txt.Inlines.Clear()

            Dim message = ev.Message
            If String.IsNullOrEmpty(message) Then Return

            ' Ensure indices are valid
            If highlightStart < 0 OrElse highlightStart >= message.Length Then
                EventMessage_txt.Inlines.Add(New Run(message))
                Return
            End If

            Dim highlightEnd = Math.Min(highlightStart + highlightLength, message.Length)

            ' Text before highlight
            If highlightStart > 0 Then
                EventMessage_txt.Inlines.Add(New Run(message.Substring(0, highlightStart)))
            End If

            ' Highlighted text
            Dim highlightedRun As New Run(message.Substring(highlightStart, highlightEnd - highlightStart)) With {
                .Background = New SolidColorBrush(Color.FromRgb(&HFF, &HFF, &H0)),
                .Foreground = New SolidColorBrush(Colors.Black)
            }
            EventMessage_txt.Inlines.Add(highlightedRun)

            ' Text after highlight
            If highlightEnd < message.Length Then
                EventMessage_txt.Inlines.Add(New Run(message.Substring(highlightEnd)))
            End If
        End Sub

        ''' <summary>
        ''' Counts how many times a search term appears in a string
        ''' </summary>
        Private Function CountMatchesInString(text As String, term As String, comparison As StringComparison) As Integer
            If String.IsNullOrEmpty(text) OrElse String.IsNullOrEmpty(term) Then Return 0

            Dim count = 0
            Dim index = 0

            While index < text.Length
                index = text.IndexOf(term, index, comparison)
                If index < 0 Then Exit While
                count += 1
                index += term.Length
            End While

            Return count
        End Function

        Private Sub UpdateEventCounter(hit As SearchHit)
            EventCounter_lbl.Text =
        $"Viewing {hit.CurrentEventIndex + 1} of {hit.MatchingEvents.Count} event(s)"
        End Sub

#End Region

#Region "HAR Preview + Find Next Request"

        ''' <summary>Button handler for Previous HAR Request navigation</summary>
        Private Sub FindPreviousHarRequest_btn_Click(sender As Object, e As RoutedEventArgs)
            FindPreviousHarRequest()
        End Sub

        ''' <summary>Button handler for Next HAR Request navigation</summary>
        Private Sub FindNextHarRequest_btn_Click(sender As Object, e As RoutedEventArgs)
            FindNextHarRequest()
        End Sub

        ''' <summary>
        ''' Navigates to previous occurrence of search term in HAR requests
        ''' Searches within current request first, then moves to previous request
        ''' </summary>
        Private Sub FindPreviousHarRequest()
            Dim hit = TryCast(Results_lst.SelectedItem, SearchHit)
            If hit Is Nothing OrElse hit.MatchingRequests.Count = 0 Then Return

            hit.CurrentRequestIndex -= 1
            If hit.CurrentRequestIndex < 0 Then
                hit.CurrentRequestIndex = hit.MatchingRequests.Count - 1
                If Not _isScanning Then Status("Wrapped to last request")
            Else
                If Not _isScanning Then Status("Ready")
            End If

            RenderHarRequest(hit)
        End Sub

        ''' <summary>
        ''' Navigates to next occurrence of search term in HAR requests
        ''' Wraps to first request when reaching end
        ''' </summary>
        Private Sub FindNextHarRequest()
            Dim hit = TryCast(Results_lst.SelectedItem, SearchHit)
            If hit Is Nothing OrElse hit.MatchingRequests.Count = 0 Then Return

            hit.CurrentRequestIndex += 1
            If hit.CurrentRequestIndex >= hit.MatchingRequests.Count Then
                hit.CurrentRequestIndex = 0
                If Not _isScanning Then Status("Wrapped to first request")
            Else
                If Not _isScanning Then Status("Ready")
            End If

            RenderHarRequest(hit)
        End Sub

        ''' <summary>
        ''' Renders the first HAR request with highlighting
        ''' </summary>
        Private Sub RenderFirstHarRequest(hit As SearchHit)
            If hit.MatchingRequests.Count = 0 Then Return

            hit.CurrentRequestIndex = 0
            RenderHarRequest(hit)
        End Sub

        ''' <summary>
        ''' Renders the current HAR request with highlighting
        ''' </summary>
        Private Sub RenderHarRequest(hit As SearchHit)
            If hit.CurrentRequestIndex < 0 OrElse hit.CurrentRequestIndex >= hit.MatchingRequests.Count Then Return

            Dim req = hit.MatchingRequests(hit.CurrentRequestIndex)
            Dim term = Search_txt.Text
            Dim cs = CaseSensitive_chk.IsChecked.GetValueOrDefault(False)
            Dim comparison = If(cs, StringComparison.Ordinal, StringComparison.OrdinalIgnoreCase)

            ' Highlight HTTP Method
            HighlightTextInBlock(HarMethod_txt, req.Method, term, comparison)

            ' Highlight URL
            HighlightTextInBlock(HarUrl_txt, req.Url, term, comparison)

            ' Highlight Status Code and Text
            Dim statusText = $"{req.StatusCode} {req.StatusText}"
            HighlightTextInBlock(HarStatus_txt, statusText, term, comparison)

            HarTime_txt.Text = If(req.StartedDateTime.HasValue, req.StartedDateTime.Value.ToString("yyyy-MM-dd HH:mm:ss.fff"), "N/A")
            HarDuration_txt.Text = $"{req.Time:F2} ms"
            HarServerIp_txt.Text = If(String.IsNullOrEmpty(req.ServerIpAddress), "N/A", req.ServerIpAddress)

            ' Highlight Request Headers
            HighlightTextInBlock(HarRequestHeaders_txt,
                                If(String.IsNullOrEmpty(req.RequestHeaders), "(no headers)", req.RequestHeaders),
                                term, comparison)

            ' Highlight Response Headers
            HighlightTextInBlock(HarResponseHeaders_txt,
                                If(String.IsNullOrEmpty(req.ResponseHeaders), "(no headers)", req.ResponseHeaders),
                                term, comparison)

            ' Highlight Request Body/Payload
            HighlightTextInBlock(HarRequestBody_txt,
                                If(String.IsNullOrEmpty(req.RequestBody), "(no request payload)", req.RequestBody),
                                term, comparison)

            ' Highlight Response Body/Payload
            HighlightTextInBlock(HarResponseBody_txt,
                                If(String.IsNullOrEmpty(req.ResponseBody), "(no response payload)", req.ResponseBody),
                                term, comparison)

            ' Update counter
            HarCounter_lbl.Text = $"Viewing {hit.CurrentRequestIndex + 1} of {hit.MatchingRequests.Count} request(s)"

            ' Count total matches in this request
            Dim totalMatches = 0
            If req.Method IsNot Nothing Then totalMatches += CountMatchesInString(req.Method, term, comparison)
            If req.Url IsNot Nothing Then totalMatches += CountMatchesInString(req.Url, term, comparison)
            totalMatches += CountMatchesInString(req.StatusCode.ToString(), term, comparison)
            If req.StatusText IsNot Nothing Then totalMatches += CountMatchesInString(req.StatusText, term, comparison)
            If req.RequestHeaders IsNot Nothing Then totalMatches += CountMatchesInString(req.RequestHeaders, term, comparison)
            If req.ResponseHeaders IsNot Nothing Then totalMatches += CountMatchesInString(req.ResponseHeaders, term, comparison)
            If req.RequestBody IsNot Nothing Then totalMatches += CountMatchesInString(req.RequestBody, term, comparison)
            If req.ResponseBody IsNot Nothing Then totalMatches += CountMatchesInString(req.ResponseBody, term, comparison)

            HarMatchCounter_lbl.Text = $"{totalMatches} match(es) in this request"
        End Sub

        ''' <summary>
        ''' Highlights all occurrences of search term in a TextBlock using Inlines
        ''' </summary>
        Private Sub HighlightTextInBlock(textBlock As TextBlock, text As String, term As String, comparison As StringComparison)
            textBlock.Inlines.Clear()

            If String.IsNullOrEmpty(text) Then
                textBlock.Inlines.Add(New Run(""))
                Return
            End If

            If String.IsNullOrWhiteSpace(term) Then
                textBlock.Inlines.Add(New Run(text))
                Return
            End If

            Dim currentIndex = 0
            While currentIndex < text.Length
                Dim matchIndex = text.IndexOf(term, currentIndex, comparison)

                If matchIndex < 0 Then
                    ' No more matches, add remaining text
                    textBlock.Inlines.Add(New Run(text.Substring(currentIndex)))
                    Exit While
                End If

                ' Add text before match
                If matchIndex > currentIndex Then
                    textBlock.Inlines.Add(New Run(text.Substring(currentIndex, matchIndex - currentIndex)))
                End If

                ' Add highlighted match
                Dim highlightedRun As New Run(text.Substring(matchIndex, term.Length)) With {
                    .Background = New SolidColorBrush(Color.FromRgb(&HFF, &HFF, &H0)),
                    .Foreground = New SolidColorBrush(Colors.Black),
                    .FontWeight = FontWeights.Bold
                }
                textBlock.Inlines.Add(highlightedRun)

                currentIndex = matchIndex + term.Length
            End While
        End Sub

#End Region

#Region "Next File"

        ''' <summary>Button handler for Next File navigation</summary>
        Private Sub NextFile_btn_Click(sender As Object, e As RoutedEventArgs)
            NextFile()
        End Sub

        ''' <summary>
        ''' Navigates to next result file in list (Ctrl+Down shortcut)
        ''' </summary>
        Private Sub NextFile()
            If _hits.Count = 0 Then Return

            If Results_lst.SelectedIndex < 0 Then
                Results_lst.SelectedIndex = 0
                Return
            End If

            If Results_lst.SelectedIndex < _hits.Count - 1 Then
                Results_lst.SelectedIndex += 1
            Else
                If Not _isScanning Then Status("No more files")
            End If
        End Sub

#End Region

#Region "Preview Mode Switch"

        ''' <summary>Shows text preview pane, hides EVTX and Web previews</summary>
        Private Sub ShowTextPreviewMode()
            TextPreview_grp.Visibility = Visibility.Visible
            EventPreview_grp.Visibility = Visibility.Collapsed
            HarPreview_grp.Visibility = Visibility.Collapsed
            WebPreview_grp.Visibility = Visibility.Collapsed
            EventCounter_lbl.Visibility = Visibility.Collapsed

            ' Clear WebView2 to prevent black screen
            ClearWebPreview()
        End Sub

        ''' <summary>Shows EVTX preview pane, hides text and Web previews</summary>
        Private Sub ShowEventPreviewMode()
            TextPreview_grp.Visibility = Visibility.Collapsed
            EventPreview_grp.Visibility = Visibility.Visible
            HarPreview_grp.Visibility = Visibility.Collapsed
            WebPreview_grp.Visibility = Visibility.Collapsed
            EventCounter_lbl.Visibility = Visibility.Visible

            ' Clear WebView2 to prevent black screen
            ClearWebPreview()
        End Sub

        ''' <summary>Shows web preview pane (HTML/XML), hides text and EVTX previews</summary>
        Private Sub ShowWebPreviewMode()
            TextPreview_grp.Visibility = Visibility.Collapsed
            EventPreview_grp.Visibility = Visibility.Collapsed
            WebPreview_grp.Visibility = Visibility.Visible
            EventCounter_lbl.Visibility = Visibility.Collapsed
        End Sub

        Private Sub ShowHarPreviewMode()
            TextPreview_grp.Visibility = Visibility.Collapsed
            EventPreview_grp.Visibility = Visibility.Collapsed
            HarPreview_grp.Visibility = Visibility.Visible
            WebPreview_grp.Visibility = Visibility.Collapsed
            EventCounter_lbl.Visibility = Visibility.Collapsed

            ' Clear WebView2 to prevent black screen
            ClearWebPreview()
        End Sub

#End Region

#Region "Status / Progress"

        ''' <summary>Updates status label with message</summary>
        Private Sub Status(msg As String)
            Status_lbl.Text = "Status: " & msg
        End Sub

        Private Sub ShowProgress(scanning As Boolean)
            ScanProgress_pb.Visibility = If(scanning, Visibility.Visible, Visibility.Collapsed)
        End Sub

#End Region

#Region "Current file label (throttled)"

        ''' <summary>
        ''' Updates CurrentFile_lbl with throttling to prevent excessive UI updates during fast file scanning
        ''' Only updates if FILE_LABEL_INTERVAL_MS (120ms) has elapsed since last update
        ''' </summary>
        Private Sub ThrottledSetCurrentFileDisplay(text As String)
            _pendingFileLabel = text
            ' Note: UpdateScanProgress now handles the actual label update with file counter
        End Sub

        ''' <summary>
        ''' Timer tick handler that flushes pending CurrentFile_lbl updates
        ''' Ensures label stays current even if throttle blocks immediate updates
        ''' </summary>
        Private Sub FileLabelTimer_Tick(sender As Object, e As EventArgs)
            ' UpdateScanProgress now handles all file label updates
        End Sub

        Private Sub SetCurrentFileDisplayImmediate(text As String)
            CurrentFile_lbl.Text = text
        End Sub

        ''' <summary>
        ''' Timer tick handler that updates elapsed time display every second
        ''' </summary>
        Private Sub ElapsedTimeTimer_Tick(sender As Object, e As EventArgs)
            If Not _isScanning Then Return
            Dim elapsed = _scanElapsedStopwatch.Elapsed
            TimeElapsed_lbl.Text = $"Time Elapsed: {elapsed:hh\:mm\:ss}"
        End Sub

        ''' <summary>
        ''' Updates progress bar and file counter based on files scanned
        ''' </summary>
        Private Sub UpdateScanProgress()
            Dispatcher.BeginInvoke(Sub()
                                       If _totalFilesToScan > 0 Then
                                           Dim percentage = (_filesScanned / _totalFilesToScan) * 100
                                           ScanProgress_pb.Value = Math.Min(percentage, 100)
                                       End If

                                       ' Update file label with progress counter
                                       If Not String.IsNullOrEmpty(_pendingFileLabel) Then
                                           If _totalFilesToScan > 0 Then
                                               CurrentFile_lbl.Text = $"Scanning {_filesScanned} of {_totalFilesToScan}: {_pendingFileLabel}"
                                           Else
                                               CurrentFile_lbl.Text = $"Scanning file {_filesScanned}: {_pendingFileLabel}"
                                           End If
                                       End If
                                   End Sub)
        End Sub

        ''' <summary>
        ''' Counts total files in a folder recursively (for progress calculation)
        ''' Archives are counted as their contents, not as single files
        ''' </summary>
        Private Function CountFilesInFolder(folder As String) As Integer
            Try
                Dim totalCount As Integer = 0

                For Each filePath In Directory.EnumerateFiles(folder, "*.*", SearchOption.AllDirectories)
                    Dim ext = Path.GetExtension(filePath)

                    ' If it's an archive, count its contents instead of counting it as 1 file
                    If _supportedArchiveExt.Contains(ext) Then
                        Try
                            Dim archiveCount = CountFilesInArchive(filePath)
                            totalCount += archiveCount
                        Catch ex As Exception
                            ' If we can't count archive contents, count it as 1 file
                            totalCount += 1
                        End Try
                    Else
                        ' Regular file - count as 1
                        totalCount += 1
                    End If
                Next

                Return totalCount
            Catch ex As Exception
                Return 0
            End Try
        End Function

        ''' <summary>
        ''' Opens an archive using the appropriate SharpCompress type based on file extension
        ''' Supports ZIP, 7Z, RAR, TAR, GZ, BZ2 formats
        ''' </summary>
        Private Function OpenArchive(archivePath As String) As IArchive
            Try
                ' Try generic ArchiveFactory which auto-detects format
                Return SharpCompress.Archives.ArchiveFactory.Open(archivePath)
            Catch ex As Exception
                ' Log the error for debugging
                Debug.WriteLine($"Failed to open archive {Path.GetFileName(archivePath)}: {ex.Message}")
                Throw New Exception($"Unsupported or corrupted archive format: {Path.GetFileName(archivePath)}", ex)
            End Try
        End Function

        ''' <summary>
        ''' Counts total entries in an archive file (for progress calculation)
        ''' Supports ZIP, 7Z, RAR, TAR, CAB and other formats
        ''' Recursively counts nested archives up to 1 level deep
        ''' </summary>
        Private Function CountFilesInArchive(archivePath As String, Optional depth As Integer = 0) As Integer
            Try
                Dim totalCount As Integer = 0

                ' CAB files use 7-Zip listing
                If Path.GetExtension(archivePath).Equals(".cab", StringComparison.OrdinalIgnoreCase) Then
                    Try
                        Dim files = SevenZipHelper.ListCabFiles(archivePath)

                        ' Count each file, recursively counting nested archives
                        For Each fileName In files
                            Dim ext = Path.GetExtension(fileName)

                            ' If nested archive and depth allows (max 1 level)
                            If _supportedArchiveExt.Contains(ext) AndAlso depth < 1 Then
                                ' We'd need to extract it to count contents, but that's expensive
                                ' For now, estimate nested archives as 5 files each
                                totalCount += 5
                            Else
                                totalCount += 1
                            End If
                        Next

                        Return totalCount
                    Catch ex As Exception
                        Debug.WriteLine($"Failed to count CAB files using 7-Zip: {ex.Message}")
                        Return 0
                    End Try
                End If

                ' Other archives use SharpCompress
                Using archive = OpenArchive(archivePath)
                    For Each entry In archive.Entries
                        If entry.IsDirectory OrElse String.IsNullOrEmpty(entry.Key) Then Continue For

                        Dim ext = Path.GetExtension(entry.Key)

                        ' If nested archive and depth allows (max 1 level)
                        If _supportedArchiveExt.Contains(ext) AndAlso depth < 1 Then
                            ' Extract nested archive to temp and count its contents
                            Try
                                Dim tempNested = ExtractArchiveEntryToTemp(entry, archivePath)
                                Try
                                    totalCount += CountFilesInArchive(tempNested, depth + 1)
                                Finally
                                    SafeDelete(tempNested)
                                End Try
                            Catch
                                ' If extraction fails, estimate as 5 files
                                totalCount += 5
                            End Try
                        Else
                            totalCount += 1
                        End If
                    Next

                    Return totalCount
                End Using
            Catch ex As Exception
                Debug.WriteLine($"Failed to count files in archive {Path.GetFileName(archivePath)}: {ex.Message}")
                Return 0
            End Try
        End Function

#End Region

#Region "Temp EVTX Helpers"

        ''' <summary>
        ''' Extracts ZIP entry to temporary file for EventLogReader access
        ''' EventLogReader requires file path, cannot read from stream
        ''' </summary>
        Private Function ExtractEntryToTemp(entry As ZipArchiveEntry) As String
            Dim tempDir = Path.Combine(Path.GetTempPath(), "BeaconFindInFiles")
            Directory.CreateDirectory(tempDir)

            Dim tempFile = Path.Combine(tempDir, Guid.NewGuid().ToString("N") & "_" & Path.GetFileName(entry.FullName))

            Using input = entry.Open()
                Using output As New FileStream(tempFile, FileMode.Create, FileAccess.Write, FileShare.None)
                    input.CopyTo(output)
                End Using
            End Using

            ' Track for cleanup on Reset or app close
            _tempToDelete.Add(tempFile)
            Return tempFile
        End Function

        ''' <summary>
        ''' Extracts archive entry to temporary file for EventLogReader access (SharpCompress version)
        ''' EventLogReader requires file path, cannot read from stream
        ''' </summary>
        Private Function ExtractArchiveEntryToTemp(entry As IArchiveEntry, archivePath As String) As String
            Dim tempDir = Path.Combine(Path.GetTempPath(), "BeaconFindInFiles")
            Directory.CreateDirectory(tempDir)

            Dim tempFile = Path.Combine(tempDir, Guid.NewGuid().ToString("N") & "_" & Path.GetFileName(entry.Key))

            Using input = entry.OpenEntryStream()
                Using output As New FileStream(tempFile, FileMode.Create, FileAccess.Write, FileShare.None)
                    input.CopyTo(output)
                End Using
            End Using

            ' Track for cleanup on Reset or app close
            _tempToDelete.Add(tempFile)
            Return tempFile
        End Function

        ''' <summary>
        ''' Deletes all temporary EVTX files extracted from archives
        ''' </summary>
        Private Sub CleanupTemp()
            For Each f In _tempToDelete.ToList()
                SafeDelete(f)
            Next
            _tempToDelete.Clear()
        End Sub

        Private Sub SafeDelete(path As String)
            Try
                If File.Exists(path) Then File.Delete(path)
            Catch
            End Try
        End Sub

        ''' <summary>
        ''' Cleanup on application close: delete temp files and cancel any active scans
        ''' </summary>
        Protected Overrides Sub OnClosed(e As EventArgs)
            CleanupTemp()
            If _scanCts IsNot Nothing Then
                _scanCts.Cancel()
                _scanCts.Dispose()
            End If
            MyBase.OnClosed(e)
        End Sub

#End Region

    End Class

End Namespace
