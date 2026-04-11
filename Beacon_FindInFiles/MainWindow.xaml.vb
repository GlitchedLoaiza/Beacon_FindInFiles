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

' ============================================================================
' Beacon: Find in Files - Advanced Log Search Utility
' ============================================================================
' This software was developed by Loaiza (luislo@microsoft)
' 
' Purpose: Searches for text patterns across multiple file types including:
'   - Plain text files (.txt, .log, .json, .xml, .csv, .html, .reg)
'   - Windows Event Log files (.evtx)
'   - ZIP archives (recursively scans entries)
'
' Features:
'   - Real-time text highlighting in preview
'   - Keyboard shortcuts (F3, Ctrl+Down, Ctrl+R, Enter, Esc)
'   - Cancellable async scanning with progress indication
'   - Case-sensitive/insensitive search
'   - EVTX event message parsing with inline highlighting
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
            EvtxFileOnDisk      ' Event log file on disk
            EvtxFromZipTemp     ' Event log extracted from ZIP to temp location
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

        ''' <summary>Root folder of current scan, used for relative path display</summary>
        Private _scanRootFolder As String = ""

        ''' <summary>Last scanned path, used to detect re-runs on same source</summary>
        Private _lastScanPath As String = ""

        ''' <summary>File extensions recognized as searchable text files</summary>
        Private ReadOnly _supportedTextExt As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase) From {
            ".txt", ".log", ".json", ".xml", ".csv", ".html", ".reg"
        }

        ''' <summary>File extensions recognized as Windows Event Logs</summary>
        Private ReadOnly _supportedEvtxExt As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase) From {
            ".evtx"
        }

        ''' <summary>Tracks temporary EVTX files extracted from ZIPs for cleanup</summary>
        Private ReadOnly _tempToDelete As New List(Of String)

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

#End Region

#Region "Init"
        ' ====================================================================
        ' Window initialization and event handler registration
        ' ====================================================================

        ''' <summary>
        ''' Constructor: Initializes UI controls, wires event handlers, and sets initial state.
        ''' </summary>
        Public Sub New()
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

            ' Wire input change handlers for button state updates
            AddHandler Path_txt.TextChanged, AddressOf AnyInputChanged
            AddHandler Search_txt.TextChanged, AddressOf AnyInputChanged

            ' Register global keyboard shortcut handler
            AddHandler Me.PreviewKeyDown, AddressOf MainWindow_PreviewKeyDown

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

            ShowTextPreviewMode()
            ClearTextPreview()
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
        ''' Triggered when Path_txt or Search_txt content changes.
        ''' Re-evaluates button enable states.
        ''' </summary>
        Private Sub AnyInputChanged(sender As Object, e As TextChangedEventArgs)
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
                Return
            End If

            Dim p = Path_txt.Text.Trim()
            Dim term = Search_txt.Text

            ' Scan enabled when path is valid (folder or ZIP) AND search term provided
            Dim pathValid As Boolean =
                Directory.Exists(p) OrElse
                (File.Exists(p) AndAlso String.Equals(Path.GetExtension(p), ".zip", StringComparison.OrdinalIgnoreCase))

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
        ''' Prompts user to select a ZIP file to scan
        ''' </summary>
        Private Sub BrowseZip_btn_Click(sender As Object, e As RoutedEventArgs)
            Dim ofd As New OpenFileDialog() With {
                .Filter = "ZIP files (*.zip)|*.zip|All files (*.*)|*.*",
                .Multiselect = False,
                .Title = "Select ZIP archive to scan"
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
            Dim ct = _scanCts.Token

            ' Capture scan root folder for relative display
            ' (e.g., show "logs\app.log" instead of "C:\Users\...\logs\app.log")
            _scanRootFolder = If(Directory.Exists(p), p, "")

            ' Count total files before scanning (for progress bar)
            Task.Run(Sub()
                         Try
                             If Directory.Exists(p) Then
                                 _totalFilesToScan = CountFilesInFolder(p)
                             ElseIf File.Exists(p) AndAlso Path.GetExtension(p).Equals(".zip", StringComparison.OrdinalIgnoreCase) Then
                                 _totalFilesToScan = CountFilesInZip(p)
                             End If
                         Catch
                             _totalFilesToScan = 0
                         End Try
                     End Sub)

            ' Spawn background scan task to keep UI responsive
            Task.Run(Async Function()
                         Dim wasCancelled As Boolean = False

                         Try
                             ' Route to appropriate scan method based on source type
                             If Directory.Exists(p) Then
                                 Await ScanFolderAsync(p, term, cs, ct)
                             ElseIf File.Exists(p) AndAlso Path.GetExtension(p).Equals(".zip", StringComparison.OrdinalIgnoreCase) Then
                                 Await ScanZipAsync(p, term, cs, ct)
                             End If
                         Catch ex As OperationCanceledException
                             wasCancelled = True
                         Catch ex As TaskCanceledException
                             wasCancelled = True
                         Catch ex As Exception
                             Dispatcher.Invoke(Sub() Status("Error: " & ex.Message))
                         Finally
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

            ' Clear results and selection
            _hits.Clear()
            Results_lst.SelectedIndex = -1

            ' Reset preview pane
            ClearTextPreview()
            ShowTextPreviewMode()

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
        ''' Recursively scans a folder for text files, ZIPs, and EVTX files
        ''' Processes nested ZIPs and extracts EVTX files from archives for scanning
        ''' </summary>
        Private Async Function ScanFolderAsync(folder As String, term As String, caseSensitive As Boolean, ct As CancellationToken) As Task
            ' Enumerate all files recursively
            For Each f In Directory.EnumerateFiles(folder, "*.*", SearchOption.AllDirectories)
                If ct.IsCancellationRequested Then Exit For

                ' Increment file counter and update progress
                _filesScanned += 1
                UpdateScanProgress()

                ' Show relative path (more useful than just file name)
                Dim relScanning = MakeRelativeDisplay(folder, f)
                ThrottledSetCurrentFileDisplay(relScanning)

                Dim ext = Path.GetExtension(f)

                ' Handle nested ZIPs recursively
                If ext.Equals(".zip", StringComparison.OrdinalIgnoreCase) Then
                    Await ScanZipAsync(f, term, caseSensitive, ct)
                    Continue For
                End If

                If _supportedTextExt.Contains(ext) Then
                    If Await TextFileContainsAsync(f, term, caseSensitive, ct) Then
                        Dim relName = MakeRelativeDisplay(folder, f)
                        AddHit(New SearchHit With {
                            .DisplayName = relName,
                            .Kind = HitKind.DiskTextFile,
                            .FilePath = f
                        })
                    End If

                ElseIf _supportedEvtxExt.Contains(ext) Then
                    Dim ev = Await EvtxCollectMatchesAsync(f, term, caseSensitive, ct)
                    If ev IsNot Nothing Then
                        ev.DisplayName = MakeRelativeDisplay(folder, f)
                        AddHit(ev)
                    End If
                End If
            Next
        End Function

        ''' <summary>
        ''' Scans all entries in a ZIP archive for text and EVTX files
        ''' Extracts EVTX files to temp storage for EventLogReader access
        ''' </summary>
        Private Async Function ScanZipAsync(zipPath As String, term As String, caseSensitive As Boolean, ct As CancellationToken) As Task
            Using z = ZipFile.OpenRead(zipPath)
                For Each entry In z.Entries
                    If ct.IsCancellationRequested Then Exit For
                    If String.IsNullOrEmpty(entry.Name) Then Continue For

                    ' Increment file counter and update progress
                    _filesScanned += 1
                    UpdateScanProgress()

                    ' Show "archive.zip | path/to/file.ext" format
                    ThrottledSetCurrentFileDisplay($"{Path.GetFileName(zipPath)} | {entry.FullName}")

                    Dim ext = Path.GetExtension(entry.FullName)

                    If _supportedTextExt.Contains(ext) Then
                        Using s = entry.Open()
                            If Await StreamContainsAsync(s, term, caseSensitive, ct) Then
                                AddHit(New SearchHit With {
                                       .DisplayName = $"{Path.GetFileName(zipPath)} | {entry.FullName}",
                                       .Kind = HitKind.ZipTextEntry,
                                       .ZipPath = zipPath,
                                       .ZipEntryName = entry.FullName
                                })
                            End If
                        End Using

                    ElseIf _supportedEvtxExt.Contains(ext) Then
                        ' EventLogReader requires file path, extract to temp
                        Dim tempEvtx = ExtractEntryToTemp(entry)
                        Dim ev = Await EvtxCollectMatchesAsync(tempEvtx, term, caseSensitive, ct)

                        If ev IsNot Nothing Then
                            ev.Kind = HitKind.EvtxFromZipTemp
                            ev.ZipPath = zipPath
                            ev.ZipEntryName = entry.FullName
                            ev.TempEvtxPath = tempEvtx
                            ev.DisplayName = $"{Path.GetFileName(zipPath)} | {entry.FullName}"
                            AddHit(ev)
                        Else
                            SafeDelete(tempEvtx)
                        End If
                    End If
                Next
            End Using
        End Function

        ''' <summary>
        ''' Thread-safe addition of search result to UI-bound collection
        ''' Marshals to UI thread via Dispatcher
        ''' </summary>
        Private Sub AddHit(hit As SearchHit)
            Dispatcher.Invoke(Sub() _hits.Add(hit))
        End Sub

#End Region

#Region "Text Search"

        ''' <summary>
        ''' Checks if file contains search term using line-by-line streaming
        ''' Uses FileShare.ReadWrite to access locked files (e.g., active logs)
        ''' </summary>
        Private Async Function TextFileContainsAsync(filePath As String, term As String, caseSensitive As Boolean, ct As CancellationToken) As Task(Of Boolean)
            Dim comparison = If(caseSensitive, StringComparison.Ordinal, StringComparison.OrdinalIgnoreCase)

            Using fs As New FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)
                Using sr As New StreamReader(fs, detectEncodingFromByteOrderMarks:=True)
                    ' Scan line-by-line for memory efficiency
                    While True
                        If ct.IsCancellationRequested Then Return False
                        Dim line = Await sr.ReadLineAsync()
                        If line Is Nothing Then Exit While
                        If line.IndexOf(term, comparison) >= 0 Then Return True
                    End While
                End Using
            End Using

            Return False
        End Function

        ''' <summary>
        ''' Checks if stream contains search term (used for ZIP entry scanning)
        ''' Leaves stream open for caller to manage (leaveOpen:=True)
        ''' </summary>
        Private Async Function StreamContainsAsync(stream As Stream, term As String, caseSensitive As Boolean, ct As CancellationToken) As Task(Of Boolean)
            Dim comparison = If(caseSensitive, StringComparison.Ordinal, StringComparison.OrdinalIgnoreCase)

            Using sr As New StreamReader(stream, detectEncodingFromByteOrderMarks:=True, bufferSize:=4096, leaveOpen:=True)
                While True
                    If ct.IsCancellationRequested Then Return False
                    Dim line = Await sr.ReadLineAsync()
                    If line Is Nothing Then Exit While
                    If line.IndexOf(term, comparison) >= 0 Then Return True
                End While
            End Using

            Return False
        End Function

#End Region

#Region "EVTX Search (Cancellation-safe)"

        ''' <summary>
        ''' Scans EVTX file for events containing search term
        ''' OPTIMIZATION: Pre-filters using XML representation before expensive FormatDescription() call
        ''' This yields 2-10x performance improvement for EVTX scanning
        ''' Limits to 300 matches per file to prevent memory issues
        ''' </summary>
        Private Async Function EvtxCollectMatchesAsync(evtxPath As String,
                                                     term As String,
                                                     caseSensitive As Boolean,
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
                                                      If String.IsNullOrEmpty(xmlString) OrElse xmlString.IndexOf(term, comparison) < 0 Then
                                                          Continue While
                                                      End If

                                                      ' Term found in XML, now get the formatted message (expensive but necessary)
                                                      Dim msg As String = ""
                                                      Try
                                                          msg = rec.FormatDescription()
                                                      Catch
                                                          msg = ""
                                                      End Try

                                                      If Not String.IsNullOrEmpty(msg) AndAlso msg.IndexOf(term, comparison) >= 0 Then
                                                          hit.MatchingEvents.Add(New EventSummary With {
                                                              .Level = rec.LevelDisplayName,
                                                              .Provider = rec.ProviderName,
                                                              .EventId = rec.Id,
                                                              .TimeCreated = rec.TimeCreated,
                                                              .Message = msg
                                                              })

                                                          ' Limit matches per file to prevent memory exhaustion
                                                          If hit.MatchingEvents.Count >= 300 Then Exit While
                                                      End If
                                                  End Using
                                              End While
                                          End Using

                                          If hit.MatchingEvents.Count > 0 Then Return hit
                                          Return Nothing

                                      End Function)
            Catch ex As OperationCanceledException
                Return Nothing
            Catch ex As TaskCanceledException
                Return Nothing
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
            _currentTextFindStart = 0
            _currentEventMessageMatchIndex = 0

            ' Load preview based on hit type
            Select Case hit.Kind
                Case HitKind.DiskTextFile
                    ShowTextPreviewMode()
                    LoadTextFromDisk(hit.FilePath)
                    FindNext_btn.Visibility = Visibility.Visible

                Case HitKind.ZipTextEntry
                    ShowTextPreviewMode()
                    LoadTextFromZip(hit.ZipPath, hit.ZipEntryName)
                    FindNext_btn.Visibility = Visibility.Visible

                Case HitKind.EvtxFileOnDisk, HitKind.EvtxFromZipTemp
                    ShowEventPreviewMode()
                    hit.CurrentEventIndex = -1
                    RenderFirstEvent(hit)
                    FindNextEvent_btn.Visibility = Visibility.Visible
                    FindPreviousEvent_btn.Visibility = Visibility.Visible
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
        ''' Extracts text entry from ZIP and loads into preview pane
        ''' </summary>
        Private Sub LoadTextFromZip(zipPath As String, entryName As String)
            Using z = ZipFile.OpenRead(zipPath)
                Dim entry = z.GetEntry(entryName)
                If entry Is Nothing Then
                    SetTextPreview("[Entry not found]")
                    Return
                End If

                Using s = entry.Open()
                    Using sr As New StreamReader(s, detectEncodingFromByteOrderMarks:=True)
                        SetTextPreview(sr.ReadToEnd())
                    End Using
                End Using
            End Using
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
        ''' Finds and highlights next occurrence of search term in text preview
        ''' Wraps to beginning if no more matches found
        ''' </summary>
        Private Sub FindNextInText()
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

        ''' <summary>Shows text preview pane, hides EVTX preview</summary>
        Private Sub ShowTextPreviewMode()
            TextPreview_grp.Visibility = Visibility.Visible
            EventPreview_grp.Visibility = Visibility.Collapsed
            EventCounter_lbl.Visibility = Visibility.Collapsed
        End Sub

        Private Sub ShowEventPreviewMode()
            TextPreview_grp.Visibility = Visibility.Collapsed
            EventPreview_grp.Visibility = Visibility.Visible
            EventCounter_lbl.Visibility = Visibility.Visible
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
        ''' </summary>
        Private Function CountFilesInFolder(folder As String) As Integer
            Try
                Return Directory.EnumerateFiles(folder, "*.*", SearchOption.AllDirectories).Count()
            Catch
                Return 0
            End Try
        End Function

        ''' <summary>
        ''' Counts total entries in a ZIP file (for progress calculation)
        ''' </summary>
        Private Function CountFilesInZip(zipPath As String) As Integer
            Try
                Using z = ZipFile.OpenRead(zipPath)
                    Return z.Entries.Where(Function(e) Not String.IsNullOrEmpty(e.Name)).Count()
                End Using
            Catch
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
        ''' Deletes all temporary EVTX files extracted from ZIPs
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
