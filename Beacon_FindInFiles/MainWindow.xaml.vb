Imports System.IO
Imports System.IO.Compression
Imports System.Threading
Imports System.Threading.Tasks
Imports System.Collections.ObjectModel
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Documents
Imports System.Windows.Input
Imports Microsoft.Win32
Imports System.Diagnostics
Imports System.Diagnostics.Eventing.Reader
'This software was developed by Loaiza
'luislo@microsoft
Namespace Beacon

    Partial Public Class MainWindow
        Inherits Window

#Region "Models"

        Private Enum HitKind
            DiskTextFile
            ZipTextEntry
            EvtxFileOnDisk
            EvtxFromZipTemp
        End Enum

        Private Class EventSummary
            Public Property Level As String
            Public Property Provider As String
            Public Property EventId As Integer
            Public Property TimeCreated As DateTime?
            Public Property Message As String
        End Class

        Private Class SearchHit
            Public Property DisplayName As String
            Public Property Kind As HitKind

            ' Disk file / EVTX
            Public Property FilePath As String

            ' Zip
            Public Property ZipPath As String
            Public Property ZipEntryName As String

            ' Temp EVTX extracted from ZIP
            Public Property TempEvtxPath As String

            ' EVTX navigation
            Public Property MatchingEvents As New List(Of EventSummary)
            Public Property CurrentEventIndex As Integer = -1
        End Class

#End Region

#Region "Fields"

        Private ReadOnly _hits As New ObservableCollection(Of SearchHit)

        Private _scanCts As CancellationTokenSource
        Private _isScanning As Boolean = False

        Private _currentTextFindStart As Integer = 0

        ' Root folder used for relative display during folder scans
        Private _scanRootFolder As String = ""
        Private _lastScanPath As String = ""
        Private ReadOnly _supportedTextExt As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase) From {
            ".txt", ".log", ".json", ".xml", ".csv", ".html", ".reg"
        }

        Private ReadOnly _supportedEvtxExt As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase) From {
            ".evtx"
        }

        Private ReadOnly _tempToDelete As New List(Of String)

        ' ---- Throttle for CurrentFile_lbl updates ----
        Private ReadOnly _fileLabelStopwatch As Stopwatch = Stopwatch.StartNew()
        Private _lastFileLabelUpdateMs As Long = 0
        Private _pendingFileLabel As String = ""
        Private _fileLabelTimer As Threading.DispatcherTimer

        Private Const FILE_LABEL_INTERVAL_MS As Integer = 120 ' small throttle

#End Region

#Region "Init"

        Public Sub New()
            InitializeComponent()

            Results_lst.ItemsSource = _hits
            Results_lst.DisplayMemberPath = NameOf(SearchHit.DisplayName)

            ' Wire handlers
            AddHandler Browse_btn.Click, AddressOf Browse_btn_Click
            AddHandler Scan_btn.Click, AddressOf Scan_btn_Click
            AddHandler Reset_btn.Click, AddressOf Reset_btn_Click

            AddHandler Results_lst.SelectionChanged, AddressOf Results_lst_SelectionChanged
            AddHandler NextFile_btn.Click, AddressOf NextFile_btn_Click
            AddHandler FindNext_btn.Click, AddressOf FindNext_btn_Click
            AddHandler FindNextEvent_btn.Click, AddressOf FindNextEvent_btn_Click
            AddHandler FindPreviousEvent_btn.Click, AddressOf FindPreviousEvent_btn_Click

            AddHandler Path_txt.TextChanged, AddressOf AnyInputChanged
            AddHandler Search_txt.TextChanged, AddressOf AnyInputChanged

            ' Keyboard shortcuts (global)
            AddHandler Me.PreviewKeyDown, AddressOf MainWindow_PreviewKeyDown

            ' Setup throttled label timer (UI thread)
            _fileLabelTimer = New Threading.DispatcherTimer()
            _fileLabelTimer.Interval = TimeSpan.FromMilliseconds(FILE_LABEL_INTERVAL_MS)
            AddHandler _fileLabelTimer.Tick, AddressOf FileLabelTimer_Tick

            ' Initial UI state
            Status("Ready")
            ShowProgress(False)
            SetCurrentFileDisplayImmediate("")

            Scan_btn.Content = "Scan (Enter)"
            Scan_btn.IsEnabled = False

            Reset_btn.IsEnabled = False

            NextFile_btn.Visibility = Visibility.Collapsed
            FindNext_btn.Visibility = Visibility.Collapsed
            FindNextEvent_btn.Visibility = Visibility.Collapsed

            ShowTextPreviewMode()
            ClearTextPreview()
        End Sub

#End Region

#Region "Keyboard"

        Private Sub MainWindow_PreviewKeyDown(sender As Object, e As KeyEventArgs)

            ' Cancel with Esc if scanning
            If e.Key = Key.Escape AndAlso _isScanning Then
                CancelScan()
                e.Handled = True
                Return
            End If

            ' Enter = Scan when enabled and NOT scanning
            If e.Key = Key.Enter AndAlso (Not _isScanning) AndAlso Scan_btn.IsEnabled Then
                StartScan()
                e.Handled = True
                Return
            End If


            ' F3 / Shift+F3 = Find Next / Previous (text or event)
            If e.Key = Key.F3 Then

                ' EVTX preview has priority
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

                ' Text preview
                If TextPreview_grp.Visibility = Visibility.Visible AndAlso
       FindNext_btn.Visibility = Visibility.Visible Then

                    FindNextInText()
                    e.Handled = True
                    Return
                End If

            End If


            ' Ctrl + Down = Next File
            If e.Key = Key.Down AndAlso (Keyboard.Modifiers And ModifierKeys.Control) = ModifierKeys.Control Then
                If NextFile_btn.Visibility = Visibility.Visible Then
                    NextFile()
                    e.Handled = True
                    Return
                End If
            End If


            ' Ctrl + R = Reset (when not scanning)
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

        Private Sub AnyInputChanged(sender As Object, e As TextChangedEventArgs)
            UpdateButtonsState()
        End Sub

        Private Sub UpdateButtonsState()
            If _isScanning Then
                ' During scanning, keep Scan enabled (as Cancel) and lock others
                Scan_btn.IsEnabled = True
                Reset_btn.IsEnabled = False

                Browse_btn.IsEnabled = String.IsNullOrWhiteSpace(Path_txt.Text)
                Path_txt.IsEnabled = False
                Search_txt.IsEnabled = False
                CaseSensitive_chk.IsEnabled = False
                Return
            End If

            Dim p = Path_txt.Text.Trim()
            Dim term = Search_txt.Text

            Dim pathValid As Boolean =
                Directory.Exists(p) OrElse
                (File.Exists(p) AndAlso String.Equals(Path.GetExtension(p), ".zip", StringComparison.OrdinalIgnoreCase))

            Dim termValid As Boolean = Not String.IsNullOrWhiteSpace(term)

            Scan_btn.IsEnabled = pathValid AndAlso termValid

            Reset_btn.IsEnabled =
                (Not String.IsNullOrWhiteSpace(Path_txt.Text)) OrElse
                (Not String.IsNullOrWhiteSpace(Search_txt.Text)) OrElse
                (_hits.Count > 0)

            Browse_btn.IsEnabled = String.IsNullOrWhiteSpace(Path_txt.Text)
            Path_txt.IsEnabled = False
            Search_txt.IsEnabled = True
            CaseSensitive_chk.IsEnabled = True
        End Sub

#End Region

#Region "Browse (Folder/Zip) - Fixed"

        Private Sub Browse_btn_Click(sender As Object, e As RoutedEventArgs)

            Dim choice = MessageBox.Show("Select a folder? (No = select a ZIP file)",
                                         "Browse",
                                         MessageBoxButton.YesNoCancel,
                                         MessageBoxImage.Question)

            If choice = MessageBoxResult.Cancel Then Return

            If choice = MessageBoxResult.No Then
                Dim ofd As New OpenFileDialog() With {
                    .Filter = "ZIP files (*.zip)|*.zip|All files (*.*)|*.*",
                    .Multiselect = False
                }
                If ofd.ShowDialog() = True Then
                    Path_txt.Text = ofd.FileName
                End If
                Return
            End If

            Dim selectedFolder As String = TryPickFolderBest()
            If Not String.IsNullOrWhiteSpace(selectedFolder) Then
                Path_txt.Text = selectedFolder
            End If
        End Sub


        Private Function TryPickFolderBest() As String

            ' 1) Try Windows API Code Pack CommonOpenFileDialog
            Try
                Dim t = Type.GetType("Microsoft.WindowsAPICodePack.Dialogs.CommonOpenFileDialog, Microsoft.WindowsAPICodePack.Shell")
                If t IsNot Nothing Then
                    Dim dlg = Activator.CreateInstance(t)
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
            End Try

            ' 2) Try WinForms FolderBrowserDialog
            Try
                Dim fbdType = Type.GetType("System.Windows.Forms.FolderBrowserDialog, System.Windows.Forms")
                If fbdType IsNot Nothing Then
                    Dim dlg = Activator.CreateInstance(fbdType)
                    fbdType.GetProperty("Description")?.SetValue(dlg, "Select folder containing logs")
                    fbdType.GetProperty("UseDescriptionForTitle")?.SetValue(dlg, True)

                    Dim showRes = fbdType.GetMethod("ShowDialog", Type.EmptyTypes).Invoke(dlg, Nothing)
                    If showRes.ToString().EndsWith("OK", StringComparison.OrdinalIgnoreCase) Then
                        Return CStr(fbdType.GetProperty("SelectedPath")?.GetValue(dlg))
                    Else
                        Return Nothing ' USER CANCELLED → STOP
                    End If
                End If
            Catch
            End Try

            ' 3) Fallback OpenFileDialog trick
            Try
                Dim dlg As New OpenFileDialog() With {
            .CheckFileExists = False,
            .CheckPathExists = True,
            .ValidateNames = False,
            .FileName = "Select Folder"
        }

                If dlg.ShowDialog() = True Then
                    Return System.IO.Path.GetDirectoryName(dlg.FileName)
                Else
                    Return Nothing ' USER CANCELLED → STOP
                End If
            Catch
            End Try

            Return Nothing
        End Function


#End Region

#Region "Scan/Cancel"

        Private Sub Scan_btn_Click(sender As Object, e As RoutedEventArgs)
            If _isScanning Then
                CancelScan()
            Else
                StartScan()
            End If
        End Sub

        Private Sub CancelScan()
            Try
                If _scanCts IsNot Nothing Then _scanCts.Cancel()
                Status("Cancelling...")
            Catch
            End Try
        End Sub

        Private Sub StartScan()
            If _isScanning Then Return
            If Not Scan_btn.IsEnabled Then Return

            CleanupTemp()

            Dim isRerun As Boolean = (_lastScanPath = Path_txt.Text)
            _lastScanPath = Path_txt.Text

            If _scanCts IsNot Nothing Then
                _scanCts.Cancel()
                _scanCts.Dispose()
            End If
            _scanCts = New CancellationTokenSource()

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

            UpdateButtonsState()

            Dim p = Path_txt.Text.Trim()
            Dim term = Search_txt.Text
            Dim cs = CaseSensitive_chk.IsChecked.GetValueOrDefault(False)
            Dim ct = _scanCts.Token

            ' Capture scan root folder for relative display
            _scanRootFolder = If(Directory.Exists(p), p, "")

            Task.Run(Async Function()
                         Dim wasCancelled As Boolean = False

                         Try
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

                                                   ' Stop timer and clear label
                                                   If _fileLabelTimer.IsEnabled Then _fileLabelTimer.Stop()
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

        Private Sub Reset_btn_Click(sender As Object, e As RoutedEventArgs)

            If _isScanning Then
                CancelScan()
            End If

            Path_txt.Text = ""
            Search_txt.Text = ""
            CaseSensitive_chk.IsChecked = False

            _hits.Clear()
            Results_lst.SelectedIndex = -1

            ClearTextPreview()
            ShowTextPreviewMode()

            NextFile_btn.Visibility = Visibility.Collapsed
            FindNext_btn.Visibility = Visibility.Collapsed
            FindNextEvent_btn.Visibility = Visibility.Collapsed

            ShowProgress(False)
            If _fileLabelTimer.IsEnabled Then _fileLabelTimer.Stop()
            SetCurrentFileDisplayImmediate("")

            _currentTextFindStart = 0
            _scanRootFolder = ""

            Status("Ready")

            _isScanning = False
            Scan_btn.Content = "Scan (Enter)"
            Scan_btn.IsEnabled = False
            _lastScanPath = ""
            Browse_btn.IsEnabled = True

            CleanupTemp()

            UpdateButtonsState()
        End Sub

#End Region

#Region "Relative Display Name"

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

        Private Async Function ScanFolderAsync(folder As String, term As String, caseSensitive As Boolean, ct As CancellationToken) As Task
            For Each f In Directory.EnumerateFiles(folder, "*.*", SearchOption.AllDirectories)
                If ct.IsCancellationRequested Then Exit For

                ' Show relative path (more useful than just file name)
                Dim relScanning = MakeRelativeDisplay(folder, f)
                ThrottledSetCurrentFileDisplay(relScanning)

                Dim ext = Path.GetExtension(f)

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

        Private Async Function ScanZipAsync(zipPath As String, term As String, caseSensitive As Boolean, ct As CancellationToken) As Task
            Using z = ZipFile.OpenRead(zipPath)
                For Each entry In z.Entries
                    If ct.IsCancellationRequested Then Exit For
                    If String.IsNullOrEmpty(entry.Name) Then Continue For

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

        Private Sub AddHit(hit As SearchHit)
            Dispatcher.Invoke(Sub() _hits.Add(hit))
        End Sub

#End Region

#Region "Text Search"

        Private Async Function TextFileContainsAsync(filePath As String, term As String, caseSensitive As Boolean, ct As CancellationToken) As Task(Of Boolean)
            Dim comparison = If(caseSensitive, StringComparison.Ordinal, StringComparison.OrdinalIgnoreCase)

            Using fs As New FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)
                Using sr As New StreamReader(fs, detectEncodingFromByteOrderMarks:=True)
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

                                          Using rdr As New EventLogReader(evtxPath, PathType.FilePath)
                                              While True
                                                  ' Do NOT throw - exit cleanly to avoid VS "user-unhandled"
                                                  If ct.IsCancellationRequested Then Return Nothing

                                                  Dim rec = rdr.ReadEvent()
                                                  If rec Is Nothing Then Exit While

                                                  Using rec
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

        Private Sub Results_lst_SelectionChanged(sender As Object, e As SelectionChangedEventArgs)
            Dim hit = TryCast(Results_lst.SelectedItem, SearchHit)
            If hit Is Nothing Then Return

            FindNext_btn.Visibility = Visibility.Collapsed
            FindNextEvent_btn.Visibility = Visibility.Collapsed
            FindPreviousEvent_btn.Visibility = Visibility.Collapsed
            _currentTextFindStart = 0

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

        Private Sub LoadTextFromDisk(path As String)
            Dim text = File.ReadAllText(path)
            SetTextPreview(text)
        End Sub

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

        Private Sub FindPreviousEvent_btn_Click(sender As Object, e As RoutedEventArgs)
            FindPreviousEvent()
        End Sub

        Private Sub FindNext_btn_Click(sender As Object, e As RoutedEventArgs)
            FindNextInText()
        End Sub

        Private Sub FindNextInText()
            Dim term = Search_txt.Text
            If String.IsNullOrWhiteSpace(term) Then Return

            Dim full = New TextRange(TextPreview_rtb.Document.ContentStart, TextPreview_rtb.Document.ContentEnd).Text
            If String.IsNullOrEmpty(full) Then Return

            Dim cs = CaseSensitive_chk.IsChecked.GetValueOrDefault(False)
            Dim comparison = If(cs, StringComparison.Ordinal, StringComparison.OrdinalIgnoreCase)

            Dim idx = full.IndexOf(term, _currentTextFindStart, comparison)
            Dim wrapped As Boolean = False

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


        Private Sub SelectInRichTextBox(startIndex As Integer, length As Integer)

            Dim startPtr = GetTextPointerAt(TextPreview_rtb.Document.ContentStart, startIndex)
            Dim endPtr = GetTextPointerAt(startPtr, length)

            If startPtr Is Nothing OrElse endPtr Is Nothing Then Return

            ' Normalize pointer into valid insertion positions
            startPtr = startPtr.GetInsertionPosition(LogicalDirection.Forward)
            endPtr = endPtr.GetInsertionPosition(LogicalDirection.Backward)

            TextPreview_rtb.BeginChange()

            ' Collapse selection first
            TextPreview_rtb.Selection.Select(startPtr, startPtr)
            TextPreview_rtb.CaretPosition = startPtr

            ' Apply actual selection
            TextPreview_rtb.Selection.Select(startPtr, endPtr)
            TextPreview_rtb.CaretPosition = endPtr

            TextPreview_rtb.EndChange()

            TextPreview_rtb.Focus()

            Dispatcher.BeginInvoke(
        Sub()
            TextPreview_rtb.Selection.Select(startPtr, endPtr)
        End Sub,
        System.Windows.Threading.DispatcherPriority.Render)

            Dim p = startPtr.Paragraph
            If p IsNot Nothing Then p.BringIntoView()

        End Sub


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

        Private Sub FindPreviousEvent()
            Dim hit = TryCast(Results_lst.SelectedItem, SearchHit)
            If hit Is Nothing OrElse hit.MatchingEvents.Count = 0 Then Return

            hit.CurrentEventIndex -= 1
            If hit.CurrentEventIndex < 0 Then
                hit.CurrentEventIndex = hit.MatchingEvents.Count - 1
            End If

            RenderEvent(hit.MatchingEvents(hit.CurrentEventIndex))
            UpdateEventCounter(hit)
        End Sub

        Private Sub FindNextEvent_btn_Click(sender As Object, e As RoutedEventArgs)
            FindNextEvent()
        End Sub

        Private Sub RenderFirstEvent(hit As SearchHit)
            If hit.MatchingEvents.Count = 0 Then
                EventLevel_txt.Text = "[No matches]"
                EventId_txt.Text = ""
                EventProvider_txt.Text = ""
                EventTime_txt.Text = ""
                EventMessage_txt.Text = ""
                Return
            End If

            hit.CurrentEventIndex = 0
            RenderEvent(hit.MatchingEvents(0))
            UpdateEventCounter(hit)
        End Sub

        Private Sub FindNextEvent()
            Dim hit = TryCast(Results_lst.SelectedItem, SearchHit)
            If hit Is Nothing OrElse hit.MatchingEvents.Count = 0 Then Return

            hit.CurrentEventIndex += 1
            If hit.CurrentEventIndex >= hit.MatchingEvents.Count Then
                hit.CurrentEventIndex = 0
                If Not _isScanning Then Status("Wrapped to first event")
            Else
                If Not _isScanning Then Status("Ready")
            End If

            RenderEvent(hit.MatchingEvents(hit.CurrentEventIndex))
            UpdateEventCounter(hit)
        End Sub

        Private Sub RenderEvent(ev As EventSummary)
            Dim lvl = If(String.IsNullOrWhiteSpace(ev.Level), "Information", ev.Level)
            EventLevel_txt.Text = $"[{lvl}]"
            EventId_txt.Text = $"Event ID {ev.EventId}"
            EventProvider_txt.Text = ev.Provider
            EventTime_txt.Text = If(ev.TimeCreated.HasValue, ev.TimeCreated.Value.ToString("yyyy-MM-dd HH:mm:ss"), "")
            EventMessage_txt.Text = ev.Message
        End Sub

        Private Sub UpdateEventCounter(hit As SearchHit)
            EventCounter_lbl.Text =
        $"Viewing {hit.CurrentEventIndex + 1} of {hit.MatchingEvents.Count} event(s)"
        End Sub

#End Region

#Region "Next File"

        Private Sub NextFile_btn_Click(sender As Object, e As RoutedEventArgs)
            NextFile()
        End Sub

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

        Private Sub Status(msg As String)
            Status_lbl.Text = "Status: " & msg
        End Sub

        Private Sub ShowProgress(scanning As Boolean)
            ScanProgress_pb.Visibility = If(scanning, Visibility.Visible, Visibility.Collapsed)
        End Sub

#End Region

#Region "Current file label (throttled)"

        ' Called from background scan loops
        Private Sub ThrottledSetCurrentFileDisplay(text As String)
            _pendingFileLabel = text

            Dim nowMs = _fileLabelStopwatch.ElapsedMilliseconds
            If (nowMs - _lastFileLabelUpdateMs) >= FILE_LABEL_INTERVAL_MS Then
                _lastFileLabelUpdateMs = nowMs
                Dispatcher.BeginInvoke(Sub() CurrentFile_lbl.Text = _pendingFileLabel)
            End If
        End Sub

        ' Timer flush (UI thread)
        Private Sub FileLabelTimer_Tick(sender As Object, e As EventArgs)
            If Not _isScanning Then Return
            Dim nowMs = _fileLabelStopwatch.ElapsedMilliseconds
            If (nowMs - _lastFileLabelUpdateMs) >= FILE_LABEL_INTERVAL_MS Then
                _lastFileLabelUpdateMs = nowMs
                CurrentFile_lbl.Text = _pendingFileLabel
            End If
        End Sub

        Private Sub SetCurrentFileDisplayImmediate(text As String)
            CurrentFile_lbl.Text = text
        End Sub

#End Region

#Region "Temp EVTX Helpers"

        Private Function ExtractEntryToTemp(entry As ZipArchiveEntry) As String
            Dim tempDir = Path.Combine(Path.GetTempPath(), "BeaconFindInFiles")
            Directory.CreateDirectory(tempDir)

            Dim tempFile = Path.Combine(tempDir, Guid.NewGuid().ToString("N") & "_" & Path.GetFileName(entry.FullName))

            Using input = entry.Open()
                Using output As New FileStream(tempFile, FileMode.Create, FileAccess.Write, FileShare.None)
                    input.CopyTo(output)
                End Using
            End Using

            _tempToDelete.Add(tempFile)
            Return tempFile
        End Function

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
