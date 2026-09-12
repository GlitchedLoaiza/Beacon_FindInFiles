Imports System.IO
Imports System.Linq
Imports System.Text
Imports System.Text.Json
Imports System.Threading
Imports System.Threading.Tasks
Imports Microsoft.VisualBasic.FileIO
Imports Beacon

Module ReportChecks
    Public Function Sample(root As String) As ScanReportSnapshot
        Dim report As New ScanReportSnapshot With {
            .ApplicationVersion = "test", .FilesScanned = 7, .EstimatedTotalFiles = 10, .ElapsedMilliseconds = 1250,
            .Run = New ScanRunInfo With {.SourceRoot = root, .QueryText = " =SUM(1,2)", .Mode = SearchMode.PlainText,
                .Options = New BeaconSettings(), .StartedUtc = DateTimeOffset.UtcNow.AddSeconds(-2),
                .CompletedUtc = DateTimeOffset.UtcNow, .State = ScanReportState.Completed}
        }
        report.Results.Add(New ReportFile With {.DisplayName = "source.log", .SourcePath = Path.Combine(root, "source.log"), .FileType = "LOG",
            .PartialReason = "first matching line only", .Matches = New List(Of ReportMatch) From {
                New ReportMatch With {.Location = "Line 4", .LineNumber = 4, .VisibleMatches = 2,
                    .Excerpt = "PRIVATE_EXCERPT <script>alert('x')</script>, ""quoted""" & vbCrLf & "第二行"}
            }})
        report.Diagnostics.Items.Add(New ScanDiagnostic With {.TimestampUtc = DateTimeOffset.UtcNow, .Severity = "Warning", .Category = "Access denied",
            .Stage = "File / archive", .SourcePath = Path.Combine(root, "locked.evtx"), .Message = "Denied <img src=x onerror=alert(1)>",
            .ExceptionType = "System.UnauthorizedAccessException", .ErrorCode = "0x80070005", .TechnicalDetails = "PRIVATE_STACK_TRACE", .Occurrences = 2})
        report.Diagnostics.TotalNotifications = 2
        Return report
    End Function

    Public Sub DiagnosticBounds()
        Dim store As New ScanDiagnosticStore(2)
        Parallel.For(0, 1000, Sub(index) store.Record("same.log", New IOException("same failure"), "Scan", False))
        Dim snapshot = store.Snapshot()
        Ensure(snapshot.Items.Count = 1 AndAlso snapshot.Items(0).Occurrences = 1000 AndAlso snapshot.TotalNotifications = 1000, "Concurrent diagnostic deduplication failed.")
        snapshot.Items(0).Message = "mutated"
        Ensure(store.Snapshot().Items(0).Message = "same failure", "Diagnostics snapshot aliases live state.")
        store.Record("other.log", New IOException("other failure"), "Scan", False)
        store.Record("third.log", New IOException("third failure"), "Scan", False)
        Ensure(store.Count = 2 AndAlso store.OmittedCount = 1, "Diagnostic capacity was not enforced.")
        store.Clear()
        store.Record(New String("p"c, 5000), New IOException(New String("m"c, 4000)), "Scan", False)
        snapshot = store.Snapshot()
        Ensure(snapshot.Items(0).TextTruncated AndAlso snapshot.Items(0).SourcePath.Length <= 4096 AndAlso snapshot.Items(0).Message.Length <= 2000,
               "Diagnostic strings were not bounded/labeled.")
    End Sub

    Public Sub SnapshotAndCopy()
        Dim report = Sample(Path.GetTempPath())
        Dim exported = report.ForExport(False, False)
        Ensure(Not exported.IsComplete AndAlso exported.Results(0).Matches(0).Excerpt Is Nothing AndAlso exported.Diagnostics.Items(0).TechnicalDetails Is Nothing,
               "Export projection leaked optional data or hid incomplete coverage.")
        report.Run.QueryText = "changed"
        report.Run.Options.MaximumTotalResults = 1
        report.Results(0).Matches(0).Location = "changed"
        report.Diagnostics.Items(0).Occurrences = 9
        Ensure(exported.Run.QueryText <> "changed" AndAlso exported.Run.Options.MaximumTotalResults = 10000 AndAlso
               exported.Results(0).Matches(0).Location = "Line 4" AndAlso exported.Diagnostics.Items(0).Occurrences = 2, "Report snapshot is not detached.")
        Ensure(Not ScanReportWriter.CopyResult(report.Results(0)).Contains("PRIVATE_EXCERPT"), "Copy result included excerpts without opt-in.")
        Ensure(ScanReportWriter.CopyPaths(report) = report.Results(0).SourcePath, "Copy paths did not retain logical source paths.")
        Ensure(Not ScanReportWriter.CopyDiagnostics(report, False).Contains("PRIVATE_STACK_TRACE"), "Copy diagnostics leaked technical data.")
        For Each state In {ScanReportState.Ready, ScanReportState.Running, ScanReportState.Cancelled, ScanReportState.Failed, ScanReportState.ResultLimitReached}
            Dim incomplete As New ScanReportSnapshot With {.Run = New ScanRunInfo With {.State = state}}
            Ensure(Not incomplete.IsComplete, "An unfinished scan was labeled complete.")
        Next
        Dim complete As New ScanReportSnapshot With {.Run = New ScanRunInfo With {.State = ScanReportState.Completed}}
        Ensure(complete.IsComplete, "A complete empty search was not recognized.")
        complete.Diagnostics.OmittedNotifications = 1
        Ensure(Not complete.IsComplete, "Omitted diagnostics were hidden.")
    End Sub

    Public Sub FormatsAndAtomicWrites()
        Dim root = Path.Combine(Path.GetTempPath(), "BeaconReportTests-" & Guid.NewGuid().ToString("N"))
        Directory.CreateDirectory(root)
        Try
            Dim report = Sample(root)
            File.WriteAllText(report.Results(0).SourcePath, "source sentinel")
            Dim json = Path.Combine(root, "report.json")
            ScanReportWriter.SaveAsync(report, json, ScanReportFormat.Json, False, False, CancellationToken.None).GetAwaiter().GetResult()
            Dim text = File.ReadAllText(json)
            Ensure(Not text.Contains("PRIVATE_EXCERPT") AndAlso Not text.Contains("PRIVATE_STACK_TRACE"), "Default JSON export leaked optional data.")
            Using parsed = JsonDocument.Parse(text)
                Ensure(parsed.RootElement.GetProperty("Run").GetProperty("QueryText").GetString() = report.Run.QueryText, "JSON query was not preserved.")
                Ensure(Not parsed.RootElement.GetProperty("IsComplete").GetBoolean(), "JSON lost partial status.")
            End Using
            ScanReportWriter.SaveAsync(report, json, ScanReportFormat.Json, True, True, CancellationToken.None).GetAwaiter().GetResult()
            Using parsed = JsonDocument.Parse(File.ReadAllText(json))
                Ensure(parsed.RootElement.GetProperty("Results")(0).GetProperty("Matches")(0).GetProperty("Excerpt").GetString() = report.Results(0).Matches(0).Excerpt,
                       "JSON excerpt failed to round-trip.")
            End Using

            Dim csv = Path.Combine(root, "report.csv")
            ScanReportWriter.SaveAsync(report, csv, ScanReportFormat.Csv, True, False, CancellationToken.None).GetAwaiter().GetResult()
            Dim rows As New List(Of String())()
            Using parser As New TextFieldParser(csv, Encoding.UTF8)
                parser.SetDelimiters(",")
                parser.HasFieldsEnclosedInQuotes = True
                parser.TrimWhiteSpace = False
                While Not parser.EndOfData
                    rows.Add(parser.ReadFields())
                End While
            End Using
            Ensure(rows.All(Function(row) row.Length = 16), "CSV rows are not rectangular.")
            Ensure(rows.Single(Function(row) row(0) = "match")(4) = report.Results(0).Matches(0).Excerpt, "CSV quoting/newlines/Unicode failed to round-trip.")
            Ensure(rows.Single(Function(row) row(0) = "metadata" AndAlso row(14) = "Query")(15).StartsWith("'"), "CSV query was not formula-protected.")
            For Each dangerous In {"=1+1", " +SUM(1)", "-1+1", "@formula", vbTab & "data", "＝1+1"}
                Ensure(ScanReportWriter.CsvField(dangerous).StartsWith("""'"), "CSV formula protection missed an input.")
            Next

            Dim html = Path.Combine(root, "report.html")
            ScanReportWriter.SaveAsync(report, html, ScanReportFormat.Html, True, False, CancellationToken.None).GetAwaiter().GetResult()
            text = File.ReadAllText(html)
            Ensure(text.Contains("&lt;script&gt;") AndAlso text.Contains("&lt;img") AndAlso Not text.Contains("<script>") AndAlso Not text.Contains("<img src=x"), "HTML export contains active untrusted markup.")
            Ensure(text.Contains("Content-Security-Policy") AndAlso Not text.Contains("PRIVATE_STACK_TRACE"), "HTML policy or technical-data omission failed.")
            Using logo = GetType(ScanReportWriter).Assembly.GetManifestResourceStream("Beacon.ReportLogo.png")
                Ensure(logo IsNot Nothing, "Export logo was not embedded in the assembly.")
                Using bytes As New MemoryStream()
                    logo.CopyTo(bytes)
                    Dim expected = "data:image/png;base64," & Convert.ToBase64String(bytes.ToArray())
                    Dim images = System.Text.RegularExpressions.Regex.Matches(text, "src='(data:image/png;base64,[^']+)'")
                    Ensure(images.Count = 2 AndAlso images.Cast(Of System.Text.RegularExpressions.Match)().All(Function(image) image.Groups(1).Value = expected),
                           "HTML branding does not embed the actual Beacon logo twice.")
                End Using
            End Using
            Ensure(text.Contains("img-src data:") AndAlso text.Contains("default-src 'none'"), "Report image policy is not restricted to embedded data.")

            Using cancelled As New CancellationTokenSource()
                cancelled.Cancel()
                Expect(Of OperationCanceledException)(Sub() ScanReportWriter.SaveAsync(report, json, ScanReportFormat.Json, False, False, cancelled.Token).GetAwaiter().GetResult())
            End Using
            Ensure(File.ReadAllText(json).Contains("PRIVATE_EXCERPT"), "Cancelled export replaced the previous report.")
            Dim lockedReport = Path.Combine(root, "locked.json")
            File.WriteAllText(lockedReport, "existing report sentinel")
            Using locked As New FileStream(lockedReport, FileMode.Open, FileAccess.ReadWrite, FileShare.None)
                Dim rejected As Boolean = False
                Try
                    ScanReportWriter.SaveAsync(report, lockedReport, ScanReportFormat.Json, False, False, CancellationToken.None).GetAwaiter().GetResult()
                Catch ex As Exception When TypeOf ex Is IOException OrElse TypeOf ex Is UnauthorizedAccessException
                    rejected = True
                End Try
                Ensure(rejected, "Locked report replacement unexpectedly succeeded.")
            End Using
            Ensure(File.ReadAllText(lockedReport) = "existing report sentinel", "Failed replacement damaged the existing report.")
            Expect(Of IOException)(Sub() ScanReportWriter.SaveAsync(report, report.Results(0).SourcePath, ScanReportFormat.Json, False, False, CancellationToken.None).GetAwaiter().GetResult())
            Ensure(File.ReadAllText(report.Results(0).SourcePath) = "source sentinel", "Export overwrote its scanned source.")
            Ensure(Not Directory.EnumerateFiles(root, ".beacon-report-*").Any(), "Export left a temporary file behind.")
        Finally
            Directory.Delete(root, True)
        End Try
    End Sub

    Public Sub MainWindowSnapshot()
        Dim window As New MainWindow()
        Dim flags = Reflection.BindingFlags.Instance Or Reflection.BindingFlags.NonPublic
        Dim type = GetType(MainWindow)
        window.RemoveHandler(System.Windows.FrameworkElement.LoadedEvent,
                             type.GetMethod("MainWindow_Loaded", flags).CreateDelegate(GetType(System.Windows.RoutedEventHandler), window))
        RemoveHandler window.Closing, DirectCast(type.GetMethod("MainWindow_Closing", flags).CreateDelegate(GetType(System.ComponentModel.CancelEventHandler), window), System.ComponentModel.CancelEventHandler)
        Dim run As New ScanRunInfo With {.SourceRoot = Path.GetTempPath(), .QueryText = "original query", .Mode = SearchMode.AllTerms,
                                        .Options = New BeaconSettings(), .State = ScanReportState.Running, .StartedUtc = DateTimeOffset.UtcNow}
        type.GetField("_scanRun", flags).SetValue(window, run)
        Dim hitType = type.GetNestedType("SearchHit", Reflection.BindingFlags.NonPublic)
        Dim hit = Activator.CreateInstance(hitType, True)
        hitType.GetProperty("DisplayName").SetValue(hit, "source.log")
        hitType.GetProperty("LogicalPath").SetValue(hit, Path.Combine(Path.GetTempPath(), "source.log"))
        DirectCast(hitType.GetProperty("Details").GetValue(hit), List(Of SearchDetail)).Add(New SearchDetail With {.Location = "File name", .IsMetadata = True, .VisibleMatches = 1, .Excerpt = "source.log"})
        Try
            type.GetMethod("AddHit", flags).Invoke(window, {hit})
            Dim preferences = DirectCast(type.GetField("_settings", flags).GetValue(window), BeaconSettings)
            For Each theme In {AppTheme.Light, AppTheme.Dark, AppTheme.System}
                preferences.Theme = theme
                type.GetMethod("ApplySettings", flags).Invoke(window, Nothing)
                Dim expectedDark = theme = AppTheme.Dark OrElse (theme = AppTheme.System AndAlso CBool(type.GetMethod("IsWindowsDarkModeEnabled", flags).Invoke(window, Nothing)))
                Ensure(CBool(type.GetField("_isDarkMode", flags).GetValue(window)) = expectedDark, "Main window did not apply the saved theme.")
                type.GetMethod("SystemThemeChanged", flags).Invoke(window, {Nothing, New Microsoft.Win32.UserPreferenceChangedEventArgs(Microsoft.Win32.UserPreferenceCategory.General)})
                window.Dispatcher.Invoke(Sub()
                                         End Sub, System.Windows.Threading.DispatcherPriority.ContextIdle)
                Ensure(CBool(type.GetField("_isDarkMode", flags).GetValue(window)) = expectedDark, "System notification overrode the selected theme.")
            Next
            Dim temporary = Path.Combine(Path.GetTempPath(), "temporary-inner.zip")
            Dim logical = Path.Combine(Path.GetTempPath(), "original.zip") & " | inner.zip"
            type.GetMethod("RememberSourcePath", flags).Invoke(window, {temporary, logical})
            type.GetMethod("RecordFileSystemIssue", flags).Invoke(window, New Object() {temporary & " | bad.json", New InvalidDataException("limit reached")})
            DirectCast(window.FindName("Search_txt"), System.Windows.Controls.TextBox).Text = "next query"
            Dim snapshot = DirectCast(type.GetMethod("BuildReportSnapshot", flags).Invoke(window, Nothing), ScanReportSnapshot)
            Ensure(snapshot.Run.QueryText = "original query" AndAlso snapshot.Run.Mode = SearchMode.AllTerms AndAlso Not snapshot.IsComplete,
                   "Main-window report used edited controls or hid the running state.")
            Ensure(snapshot.Diagnostics.Items.Single().SourcePath = logical & " | bad.json", "Diagnostic source exposed a temporary archive path.")
            run.QueryText = "mutated query"
            run.Options.MaximumTotalResults = 1
            Ensure(snapshot.Run.QueryText = "original query" AndAlso snapshot.Run.Options.MaximumTotalResults = 10000, "Main-window snapshot is not isolated.")
            DirectCast(window.FindName("Results_lst"), System.Windows.Controls.ListBox).SelectedItem = hit
            Ensure(window.FindName("CopyResult_btn") Is Nothing, "The removed Copy result button is still present.")
            Dim exportButton = DirectCast(window.FindName("ExportResults_btn"), System.Windows.Controls.Button)
            Ensure(exportButton.IsEnabled AndAlso CStr(exportButton.Content) = "Export…", "Export action has the wrong label or state.")
            window.WindowStartupLocation = System.Windows.WindowStartupLocation.Manual
            window.WindowState = System.Windows.WindowState.Normal
            window.Left = -10000
            window.Top = -10000
            window.ShowActivated = False
            window.ShowInTaskbar = False
            window.Show()
            Dim notification = DirectCast(window.FindName("UpdateNotification"), System.Windows.Controls.Border)
            Ensure(notification.Visibility = System.Windows.Visibility.Collapsed, "Update notification appeared without a newer release.")
            type.GetMethod("ShowUpdateNotification", flags).Invoke(window, {New Version(2, 10, 0, 0)})
            Ensure(notification.Visibility = System.Windows.Visibility.Visible AndAlso
                   DirectCast(window.FindName("UpdateVersion_txt"), System.Windows.Controls.TextBlock).Text.Contains("2.10.0"), "New-release notification is missing.")
            DirectCast(window.FindName("DismissUpdate_btn"), System.Windows.Controls.Button).RaiseEvent(New System.Windows.RoutedEventArgs(System.Windows.Controls.Button.ClickEvent))
            Ensure(window.FindName("TestUpdate_btn") Is Nothing, "Dummy update button must not ship in the application.")
            Ensure(type.GetMethod("TestUpdateNotification", flags) Is Nothing, "Dummy notification handler must not ship in the application.")
            type.GetMethod("ShowUpdateNotification", flags).Invoke(window, {New Version(2, 10, 0, 0)})
            Ensure(notification.Visibility = System.Windows.Visibility.Visible AndAlso
                   DirectCast(window.FindName("UpdateVersion_txt"), System.Windows.Controls.TextBlock).Text = "Version 2.10.0", "Notification contains incorrect version or dummy text.")
            For Each name In {"NextFile_btn", "FindPreviousEvent_btn", "FindNextEvent_btn"}
                DirectCast(window.FindName(name), System.Windows.Controls.Button).Visibility = System.Windows.Visibility.Visible
            Next
            For Each width In {1000.0, 1400.0, 2200.0}
                window.Width = width
                Dim bar = DirectCast(window.FindName("ResultActionsBar"), System.Windows.Controls.Grid)
                Dim layout = DirectCast(bar.Parent, System.Windows.Controls.Grid)
                For Each leftWidth In {320.0, 460.0}
                layout.ColumnDefinitions(0).Width = New System.Windows.GridLength(leftWidth)
                window.UpdateLayout()
                window.Dispatcher.Invoke(Sub()
                                         End Sub, System.Windows.Threading.DispatcherPriority.ContextIdle)
                Dim preview = DirectCast(DirectCast(window.FindName("TextPreview_grp"), System.Windows.FrameworkElement).Parent, System.Windows.FrameworkElement)
                Dim previewLeft = preview.TranslatePoint(New System.Windows.Point(), window).X
                For Each action In {"ViewRelease_btn", "DismissUpdate_btn"}
                    Dim button = DirectCast(window.FindName(action), System.Windows.Controls.Button)
                    Dim bounds = button.TransformToAncestor(notification).TransformBounds(New System.Windows.Rect(button.RenderSize))
                    Ensure(bounds.Left >= 0 AndAlso bounds.Right <= notification.ActualWidth, "Update notification actions overflow.")
                Next
                Ensure(Math.Abs(bar.TranslatePoint(New System.Windows.Point(), window).X - previewLeft) <= 1,
                       "Bottom actions do not follow the preview-pane divider.")
                Dim sidebar = DirectCast(window.FindName("SearchResultsPane"), System.Windows.Controls.Grid)
                Dim splitter = layout.Children.OfType(Of System.Windows.Controls.GridSplitter)().Single()
                Ensure(System.Windows.Controls.Grid.GetRowSpan(sidebar) = 2 AndAlso System.Windows.Controls.Grid.GetRowSpan(splitter) = 2, "Left pane or divider does not span the action row.")
                Dim sidebarBounds = sidebar.TransformToAncestor(layout).TransformBounds(New System.Windows.Rect(sidebar.RenderSize))
                Ensure(Math.Abs(sidebarBounds.Bottom - (layout.ActualHeight - sidebar.Margin.Bottom)) <= 1, "Left pane does not reach the bottom content edge.")
                Dim details = DirectCast(window.FindName("MatchDetails_exp"), System.Windows.Controls.Expander)
                details.IsExpanded = True
                window.UpdateLayout()
                Dim detailsBounds = details.TransformToAncestor(layout).TransformBounds(New System.Windows.Rect(details.RenderSize))
                Ensure(Math.Abs(detailsBounds.Bottom - sidebarBounds.Bottom) <= 1, "Expanded Match details does not use the extended left pane.")
                details.IsExpanded = False
                For Each name In {"ExportResults_btn", "CopyPaths_btn", "Diagnostics_btn", "NextFile_btn", "FindPreviousEvent_btn", "FindNextEvent_btn"}
                    Dim button = DirectCast(window.FindName(name), System.Windows.Controls.Button)
                    Dim bounds = button.TransformToAncestor(bar).TransformBounds(New System.Windows.Rect(button.RenderSize))
                    Ensure(bounds.Left >= -1 AndAlso bounds.Right <= bar.ActualWidth + 1, "Reporting/navigation controls overflow the bottom bar.")
                Next
                Next
            Next
            DirectCast(window.FindName("DismissUpdate_btn"), System.Windows.Controls.Button).RaiseEvent(New System.Windows.RoutedEventArgs(System.Windows.Controls.Button.ClickEvent))
            PumpNotification(window, 400)
            Ensure(notification.Visibility = System.Windows.Visibility.Collapsed, "Update notification could not be dismissed.")
            Dim timer = DirectCast(type.GetProperty("_updateNotificationTimer", flags).GetValue(window), System.Windows.Threading.DispatcherTimer)
            Ensure(timer.Interval <= TimeSpan.FromSeconds(1) AndAlso Not timer.IsEnabled, "Countdown timer or manual-dismiss cleanup is incorrect.")
            type.GetMethod("ShowUpdateNotification", flags).Invoke(window, {New Version(2, 10, 0, 0)})
            window.UpdateLayout()
            Dim popup = DirectCast(window.FindName("UpdatePopup"), System.Windows.Controls.Primitives.Popup)
            Dim host = DirectCast(window.FindName("UpdatePopupHost"), System.Windows.Controls.Grid)
            Dim message = DirectCast(window.FindName("UpdateMessage_txt"), System.Windows.Controls.TextBlock)
            Ensure(popup.IsOpen AndAlso timer.IsEnabled AndAlso message.Text = "A newer version of Beacon is available (5s)", "Popup or initial countdown is incorrect.")
            Ensure(host.ClipToBounds AndAlso popup.Placement = System.Windows.Controls.Primitives.PlacementMode.AbsolutePoint,
                   "Notification still uses window-relative placement.")
            PumpNotification(window, 150)
            Dim initialPosition = host.PointToScreen(New System.Windows.Point())
            Dim screenBottomRight = host.PointToScreen(New System.Windows.Point(host.ActualWidth, host.ActualHeight))
            Dim expected = DirectCast(type.GetMethod("NotificationDesktopPosition", Reflection.BindingFlags.NonPublic Or Reflection.BindingFlags.Static).
                Invoke(Nothing, {CInt(Math.Round(screenBottomRight.X - initialPosition.X)), CInt(Math.Round(screenBottomRight.Y - initialPosition.Y))}), System.Windows.Point)
            Ensure(Math.Abs(initialPosition.X - expected.X) <= 2 AndAlso Math.Abs(initialPosition.Y - expected.Y) <= 2,
                   "Actual popup screen coordinates do not match the desktop-taskbar target.")
            For Each width In {1000.0, 1400.0, 2200.0}
                window.Width = width
                window.Left += 40
                window.Top += 25
                window.UpdateLayout()
                PumpNotification(window, 150)
                Dim actual = host.PointToScreen(New System.Windows.Point())
                Ensure(Math.Abs(actual.X - initialPosition.X) <= 2 AndAlso Math.Abs(actual.Y - initialPosition.Y) <= 2,
                       "Notification moved away from the desktop taskbar when Beacon moved or resized.")
            Next
            type.GetMethod("ShowUpdateNotification", flags).Invoke(window, {New Version(2, 10, 0, 0)})
            PumpNotification(window, 1200)
            Ensure(message.Text.Contains("(4s)"), "Countdown did not decrease.")
            type.GetMethod("ShowUpdateNotification", flags).Invoke(window, {New Version(2, 10, 0, 0)})
            Ensure(message.Text.Contains("(5s)"), "Repeated notification did not restart countdown.")
            Dim elapsed = System.Diagnostics.Stopwatch.StartNew()
            While popup.IsOpen AndAlso elapsed.Elapsed < TimeSpan.FromSeconds(7)
                window.Dispatcher.Invoke(Sub()
                                         End Sub, System.Windows.Threading.DispatcherPriority.ContextIdle)
                System.Threading.Thread.Sleep(20)
            End While
            Ensure(elapsed.Elapsed >= TimeSpan.FromSeconds(4.5) AndAlso notification.Visibility = System.Windows.Visibility.Collapsed AndAlso Not timer.IsEnabled,
                   "Notification did not automatically dismiss after five seconds.")
        Finally
            window.Close()
        End Try
    End Sub

    Private Sub PumpNotification(window As MainWindow, milliseconds As Integer)
        Dim elapsed = System.Diagnostics.Stopwatch.StartNew()
        While elapsed.ElapsedMilliseconds < milliseconds
            window.Dispatcher.Invoke(Sub()
                                     End Sub, System.Windows.Threading.DispatcherPriority.ContextIdle)
            System.Threading.Thread.Sleep(10)
        End While
    End Sub

    Private Sub Ensure(condition As Boolean, message As String)
        If Not condition Then Throw New InvalidOperationException(message)
    End Sub
    Private Sub Expect(Of T As Exception)(action As Action)
        Try
            action()
        Catch ex As T
            Return
        End Try
        Throw New InvalidOperationException("Expected " & GetType(T).Name)
    End Sub
End Module
