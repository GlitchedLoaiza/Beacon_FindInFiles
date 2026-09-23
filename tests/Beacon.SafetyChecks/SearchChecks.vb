Imports System.Collections
Imports System.IO
Imports System.IO.Compression
Imports System.Reflection
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Threading
Imports System.Threading.Tasks
Imports System.Windows.Controls
Imports System.Windows.Threading
Imports Beacon

Module SearchChecks
    Public Sub Queries()
        Ensure(New SearchQuery("a.b", SearchMode.PlainText, False).IsMatch("A.B"), "Literal case handling failed.")
        Ensure(Not New SearchQuery("a.b", SearchMode.PlainText, False).IsMatch("axb"), "Literal syntax became regex.")
        Ensure(Not New SearchQuery("Error", SearchMode.PlainText, True).IsMatch("error"), "Case-sensitive matching failed.")
        Dim whole As New SearchQuery("cat", SearchMode.ExactWord, False)
        Ensure(whole.IsMatch("a cat!") AndAlso Not whole.IsMatch("catalog") AndAlso Not whole.IsMatch("猫cat猫"), "Unicode whole-word matching failed.")
        Dim pattern As New SearchQuery("(?<=code=)\d+", SearchMode.RegularExpression, False)
        Dim spans = pattern.FindHighlights("code=123 and code=45")
        Ensure(spans.Count = 2 AndAlso spans(0).Start = 5 AndAlso spans(0).Length = 3, "Regex ranges failed.")
        Ensure(New SearchQuery("error timeout", SearchMode.AnyTerm, False).IsMatch("TIMEOUT"), "Any-term search failed.")
        Dim all As New SearchQuery("""connection failed"" timeout", SearchMode.AllTerms, False)
        Ensure(all.IsMatch("TIMEOUT: connection failed") AndAlso Not all.IsMatch("connection failed"), "Quoted all-term search failed.")
        Ensure(New SearchQuery("error err", SearchMode.AnyTerm, False).FindHighlights("error").Count = 1, "Overlapping highlights were duplicated.")
        Expect(Of ArgumentException)(Sub()
                                         Dim invalid As New SearchQuery("[", SearchMode.RegularExpression, False)
                                     End Sub)
        Expect(Of ArgumentException)(Sub()
                                         Dim invalid As New SearchQuery("""unclosed", SearchMode.AllTerms, False)
                                     End Sub)
        Expect(Of ArgumentException)(Sub()
                                         Dim invalid As New SearchQuery(" ", SearchMode.PlainText, False)
                                     End Sub)
        Dim expensive As New SearchQuery("^(a+)+$", SearchMode.RegularExpression, False)
        Expect(Of RegexMatchTimeoutException)(Sub() expensive.IsMatch(New String("a"c, 10000) & "!"))
        Ensure(New SearchQuery("(?=a)", SearchMode.RegularExpression, False).FindHighlights(New String("a"c, 10000), 10).Count = 0, "Zero-width highlights were emitted.")
        Using cancelled As New CancellationTokenSource()
            cancelled.Cancel()
            Expect(Of OperationCanceledException)(Sub() pattern.FindHighlights("code=1", token:=cancelled.Token))
        End Using
    End Sub

    Public Sub TextRecords()
        Dim query As New SearchQuery("error timeout", SearchMode.AllTerms, False)
        Using stream As New MemoryStream(Encoding.UTF8.GetBytes("error" & vbCrLf & "timeout" & vbCrLf & "error timeout"))
            Dim result = SearchTextCollector.CollectAsync(stream, query, 10, False, False, CancellationToken.None).GetAwaiter().GetResult()
            Ensure(result.Details.Count = 1 AndAlso result.Details(0).LineNumber = 3, "All terms crossed text-record boundaries.")
            Ensure(result.Details(0).VisibleMatches = 2 AndAlso result.PartialReason = "", "Detailed text counts failed.")
        End Using
        For Each firstOnly In {False, True}
            Using stream As New MemoryStream(Encoding.UTF8.GetBytes("error" & vbLf & "error" & vbLf & "error"))
                Dim result = SearchTextCollector.CollectAsync(stream, New SearchQuery("error", SearchMode.PlainText, False), 2, firstOnly, False, CancellationToken.None).GetAwaiter().GetResult()
                Ensure(result.Details.Count = If(firstOnly, 1, 2) AndAlso result.PartialReason.Length > 0, "Partial-result limits were not labeled.")
            End Using
        Next
        Using stream As New MemoryStream(Encoding.UTF8.GetBytes("<script>error timeout</script><p>safe</p>"))
            Ensure(SearchTextCollector.CollectAsync(stream, query, 10, False, True, CancellationToken.None).GetAwaiter().GetResult().Details.Count = 0, "Script contents were searched as visible text.")
        End Using
        Using stream As New MemoryStream(Encoding.UTF8.GetBytes("<p>error</p><p>timeout</p>"))
            Dim result = SearchTextCollector.CollectAsync(stream, query, 10, False, True, CancellationToken.None).GetAwaiter().GetResult()
            Ensure(result.Details.Count = 1 AndAlso result.Details(0).Location = "Document text", "Document search semantics failed.")
        End Using
    End Sub

    Public Sub Pipeline()
        Dim root = Path.Combine(Path.GetTempPath(), "BeaconSearchTests-" & Guid.NewGuid().ToString("N"))
        Directory.CreateDirectory(root)
        Try
            File.WriteAllText(Path.Combine(root, "one.log"), "error" & vbLf & "timeout" & vbLf & "error timeout")
            Dim options As New BeaconSettings With {.StopAfterFirstMatchPerFile = False}
            CheckPipeline(root, "error timeout", SearchMode.AllTerms, options,
                          Sub(hits)
                              Ensure(hits.Count = 1 AndAlso Details(hits(0)).Single().LineNumber = 3, "Disk pipeline ignored all-term record scope.")
                          End Sub)
            File.WriteAllBytes(Path.Combine(root, "broken.evtx"), {CByte(1), CByte(2)})
            options = New BeaconSettings With {.SearchFileContents = False, .SearchFileNames = True}
            CheckPipeline(root, "broken.evtx", SearchMode.PlainText, options,
                          Sub(hits)
                              Ensure(hits.Count = 1 AndAlso Details(hits(0)).All(Function(detail) detail.IsMetadata), "Metadata-only EVTX attempted parsing.")
                          End Sub)
            options.SearchFileNames = False
            options.SearchFullPaths = True
            options.DisplayAbsolutePaths = True
            CheckPipeline(root, "broken.evtx", SearchMode.PlainText, options,
                          Sub(hits) Ensure(hits.Count = 1 AndAlso CStr(PropertyValue(hits(0), "DisplayName")).StartsWith(root), "Absolute-path search/display failed."))
            File.Delete(Path.Combine(root, "broken.evtx"))

            Dim har = Path.Combine(root, "sample.har")
            File.WriteAllText(har, "{""log"":{""entries"":[{""request"":{""method"":""GET"",""url"":""https://example.invalid/"",""headers"":[]},""response"":{""status"":503,""statusText"":""Unavailable"",""headers"":[],""content"":{""text"":""forbidden""}}}]}}")
            options = New BeaconSettings With {.StopAfterFirstMatchPerFile = False}
            CheckPipeline(root, "GET forbidden", SearchMode.AllTerms, options,
                          Sub(hits) Ensure(hits.Count = 1 AndAlso Details(hits(0))(0).RecordIndex = 0, "HAR fields were not matched within one request."))

            Dim zipPath = Path.Combine(root, "outer.zip")
            Using inner As New MemoryStream()
                Using archive As New ZipArchive(inner, ZipArchiveMode.Create, True)
                    Using writer As New StreamWriter(archive.CreateEntry("target.log").Open())
                        writer.Write("error timeout")
                    End Using
                End Using
                Using archive = ZipFile.Open(zipPath, ZipArchiveMode.Create)
                    Using output = archive.CreateEntry("inner.zip").Open()
                        output.Write(inner.ToArray())
                    End Using
                End Using
            End Using
            CheckPipeline(zipPath, "error timeout", SearchMode.AllTerms, options,
                          Sub(hits)
                              Ensure(hits.Count = 1, "Nested archive content was not searched.")
                              Ensure(CStr(PropertyValue(hits(0), "LogicalPath")) = zipPath & " | inner.zip | target.log", "Temporary names leaked into logical paths.")
                          End Sub)
            CheckPipeline(zipPath, "error", SearchMode.PlainText, New BeaconSettings With {.ArchiveNestingDepth = 0},
                          Sub(hits) Ensure(hits.Count = 0, "Depth-zero scan entered a nested archive."), expectedCount:=0)
            Using archive = ZipFile.Open(zipPath, ZipArchiveMode.Update)
                Using writer As New StreamWriter(archive.CreateEntry("second.log").Open())
                    writer.Write("unmatched text")
                End Using
                Using writer As New StreamWriter(archive.CreateEntry(".git/ignored.log").Open())
                    writer.Write("error")
                End Using
                archive.CreateEntry("empty-directory/")
            End Using
            CheckPipeline(zipPath, "error", SearchMode.PlainText, New BeaconSettings With {.ArchiveNestingDepth = 1},
                          Sub(hits) Ensure(hits.Count = 1, "Archive exclusions changed results."), expectedCount:=2)
            CheckPipeline(zipPath, "error", SearchMode.PlainText, New BeaconSettings With {.ArchiveNestingDepth = 0},
                          Sub(hits) Ensure(hits.Count = 0, "Depth-zero scan entered a nested archive."), expectedCount:=1)
            options = New BeaconSettings With {.SearchFileContents = False, .SearchFullPaths = True}
            CheckPipeline(zipPath, "outer.zip target.log", SearchMode.AllTerms, options,
                          Sub(hits) Ensure(hits.Count = 1 AndAlso Details(hits(0))(0).IsMetadata, "Logical archive-path search failed."))
            options = New BeaconSettings With {.StopAfterFirstMatchPerFile = False, .MaximumStructuredMatches = 1, .MaximumTotalResults = 1}
            CheckPipeline(root, "error", SearchMode.PlainText, options,
                          Sub(hits) Ensure(hits.Count = 1 AndAlso Details(hits(0)).Count = 1 AndAlso CStr(PropertyValue(hits(0), "PartialReason")).Contains("limit"), "Pipeline result limits failed."))
        Finally
            Directory.Delete(root, True)
        End Try
    End Sub

    Public Sub NavigationAndCab()
        Dim root = Path.Combine(Path.GetTempPath(), "BeaconNavigationTests-" & Guid.NewGuid().ToString("N"))
        Directory.CreateDirectory(root)
        Try
            Dim sourceFile = Path.Combine(root, "sample.log")
            File.WriteAllText(sourceFile, "header" & vbLf & "nothing" & vbLf & "error code=7" & vbLf & "error code=99")
            Dim options As New BeaconSettings With {.StopAfterFirstMatchPerFile = False}
            CheckPipeline(sourceFile, "code=\d+", SearchMode.RegularExpression, options,
                          Sub(hits) Ensure(hits.Count = 1 AndAlso Details(hits(0)).Count = 2, "Text details were not collected."),
                          Sub(window, hits)
                              Dim results = DirectCast(window.FindName("Results_lst"), ListBox)
                              results.SelectedItem = hits(0)
                              DirectCast(window.FindName("MatchDetails_exp"), Expander).IsExpanded = True
                              DirectCast(window.FindName("Details_lst"), ListBox).SelectedItem = Details(hits(0))(1)
                              Pump(window)
                              Ensure(DirectCast(window.FindName("TextPreview_rtb"), RichTextBox).Selection.Text = "code=99", "Detail selection did not navigate to the matching line/range.")
                              Dim flags = BindingFlags.Instance Or BindingFlags.NonPublic
                              Dim events = DirectCast(PropertyValue(hits(0), "MatchingEvents"), IList)
                              Dim eventType = GetType(MainWindow).GetNestedType("EventSummary", BindingFlags.NonPublic)
                              For Each message In {"code=7 then code=12", "code=99 then code=345"}
                                  Dim record = Activator.CreateInstance(eventType, True)
                                  eventType.GetProperty("Message").SetValue(record, message)
                                  eventType.GetProperty("Provider").SetValue(record, "Synthetic test provider")
                                  eventType.GetProperty("EventId").SetValue(record, 42)
                                  events.Add(record)
                              Next
                              Details(hits(0)).Clear()
                              Details(hits(0)).Add(New SearchDetail With {.RecordIndex = 0, .Location = "Event 1"})
                              Details(hits(0)).Add(New SearchDetail With {.RecordIndex = 1, .Location = "Event 2"})
                              Dim list = DirectCast(window.FindName("Details_lst"), ListBox)
                              list.Items.Refresh()
                              list.SelectedItem = Details(hits(0))(1)
                              Pump(window)
                              Ensure(CInt(PropertyValue(hits(0), "CurrentEventIndex")) = 1, "Event detail navigation selected the wrong record.")
                              GetType(MainWindow).GetMethod("NavigateEventMatch", flags).Invoke(window, New Object() {True})
                              Ensure(CInt(GetType(MainWindow).GetField("_currentEventMessageMatchIndex", flags).GetValue(window)) = "code=99 then code=345".Length,
                                     "Forward event navigation ignored variable-length regex matches.")
                              GetType(MainWindow).GetMethod("NavigateEventMatch", flags).Invoke(window, New Object() {False})
                              Ensure(CInt(GetType(MainWindow).GetField("_currentEventMessageMatchIndex", flags).GetValue(window)) = "code=99".Length,
                                     "Reverse event navigation selected the wrong match.")
                          End Sub)

            Dim har = Path.Combine(root, "multi.har")
            File.WriteAllText(har, "{""log"":{""entries"":[{""request"":{""method"":""GET"",""url"":""https://example.invalid/one"",""headers"":[]},""response"":{""status"":500,""headers"":[],""content"":{""text"":""error""}}},{""request"":{""method"":""GET"",""url"":""https://example.invalid/two"",""headers"":[]},""response"":{""status"":503,""headers"":[],""content"":{""text"":""error""}}}]}}")
            CheckPipeline(har, "error", SearchMode.PlainText, options,
                          Sub(hits) Ensure(Details(hits(0)).Count = 2, "HAR request collection failed."),
                          Sub(window, hits)
                              DirectCast(window.FindName("Results_lst"), ListBox).SelectedItem = hits(0)
                              DirectCast(window.FindName("Details_lst"), ListBox).SelectedItem = Details(hits(0))(1)
                              Pump(window)
                              Dim urlBlock = DirectCast(window.FindName("HarUrl_txt"), TextBlock)
                              Dim renderedUrl = New System.Windows.Documents.TextRange(urlBlock.ContentStart, urlBlock.ContentEnd).Text
                              Ensure(CInt(PropertyValue(hits(0), "CurrentRequestIndex")) = 1 AndAlso
                                     renderedUrl.TrimEnd().EndsWith("/two"),
                                     $"HAR detail did not navigate: index={PropertyValue(hits(0), "CurrentRequestIndex")}, rendered URL={renderedUrl}.")
                          End Sub)
            options.StopAfterFirstMatchPerFile = True
            CheckPipeline(har, "error", SearchMode.PlainText, options,
                          Sub(hits) Ensure(Details(hits(0)).Count = 1 AndAlso CStr(PropertyValue(hits(0), "PartialReason")).Contains("first"), "HAR first-match policy failed."))

            Dim cab = Path.Combine(root, "sample.cab")
            Dim start As New System.Diagnostics.ProcessStartInfo(Path.Combine(Environment.SystemDirectory, "makecab.exe")) With {.UseShellExecute = False, .CreateNoWindow = True}
            start.ArgumentList.Add(sourceFile)
            start.ArgumentList.Add(cab)
            Using child = System.Diagnostics.Process.Start(start)
                If Not child.WaitForExit(15000) Then
                    child.Kill(True)
                    Throw New TimeoutException("Test CAB creation timed out.")
                End If
                Ensure(child.ExitCode = 0, "Test CAB creation failed.")
            End Using
            options.StopAfterFirstMatchPerFile = False
            CheckPipeline(cab, "code=\d+", SearchMode.RegularExpression, options,
                          Sub(hits) Ensure(hits.Count = 1 AndAlso Details(hits(0)).Count = 2 AndAlso
                                           CStr(PropertyValue(hits(0), "LogicalPath")) = cab & " | sample.log", "CAB content or logical paths failed."))
            Dim ddf = Path.Combine(root, "multiple.ddf")
            Dim multiCab = Path.Combine(root, "multiple.cab")
            File.WriteAllLines(ddf, {".OPTION EXPLICIT", ".Set Cabinet=on", ".Set Compress=on", ".Set CabinetNameTemplate=multiple.cab",
                                    ".Set DiskDirectoryTemplate=""" & root & """", ".Set InfFileName=""" & Path.Combine(root, "multiple.inf") & """",
                                    ".Set RptFileName=""" & Path.Combine(root, "multiple.rpt") & """",
                                    """" & sourceFile & """ one.log", """" & sourceFile & """ two.log", """" & sourceFile & """ three.log"})
            start.ArgumentList.Clear()
            start.ArgumentList.Add("/F")
            start.ArgumentList.Add(ddf)
            Using child = System.Diagnostics.Process.Start(start)
                If Not child.WaitForExit(15000) Then
                    child.Kill(True)
                    Throw New TimeoutException("Multi-file CAB creation timed out.")
                End If
                Ensure(child.ExitCode = 0, "Multi-file CAB creation failed.")
            End Using
            CheckPipeline(multiCab, "code=", SearchMode.PlainText, options,
                          Sub(hits) Ensure(hits.Count = 3, "Multi-file CAB search missed files."), expectedCount:=3)
            Dim wrappedCab = Path.Combine(root, "wrapped.zip")
            Using archive = ZipFile.Open(wrappedCab, ZipArchiveMode.Create)
                archive.CreateEntryFromFile(multiCab, "nested.cab")
            End Using
            CheckPipeline(wrappedCab, "code=", SearchMode.PlainText, options,
                          Sub(hits) Ensure(hits.Count = 3, "Nested CAB search missed files."), expectedCount:=3)
            CheckPipeline(wrappedCab, "code=", SearchMode.PlainText, New BeaconSettings With {.ArchiveNestingDepth = 0},
                          Sub(hits) Ensure(hits.Count = 0, "Nested CAB bypassed the depth limit."), expectedCount:=0)
        Finally
            Directory.Delete(root, True)
        End Try
    End Sub

    Private Sub Pump(window As System.Windows.Window)
        If TypeOf window Is MainWindow Then PreviewTestHelpers.WaitForPreview(DirectCast(window, MainWindow))
        window.UpdateLayout()
        window.Dispatcher.Invoke(Sub()
                                 End Sub, DispatcherPriority.ContextIdle)
        window.UpdateLayout()
    End Sub

    Friend Sub CheckPipeline(root As String, text As String, mode As SearchMode, settings As BeaconSettings, check As Action(Of List(Of Object)),
                              Optional inspect As Action(Of MainWindow, List(Of Object)) = Nothing, Optional expectedCount As Integer? = Nothing)
        Dim window As New MainWindow()
        Dim flags = BindingFlags.Instance Or BindingFlags.NonPublic
        Dim type = GetType(MainWindow)
        window.RemoveHandler(System.Windows.FrameworkElement.LoadedEvent,
                             type.GetMethod("MainWindow_Loaded", flags).CreateDelegate(GetType(System.Windows.RoutedEventHandler), window))
        RemoveHandler window.Closing, DirectCast(type.GetMethod("MainWindow_Closing", flags).CreateDelegate(GetType(System.ComponentModel.CancelEventHandler), window), System.ComponentModel.CancelEventHandler)
        Using source As New CancellationTokenSource(TimeSpan.FromSeconds(20))
            type.GetField("_settings", flags).SetValue(window, settings)
            type.GetMethod("ApplySettings", flags).Invoke(window, Nothing)
            type.GetField("_searchOptions", flags).SetValue(window, settings)
            type.GetField("_activeQuery", flags).SetValue(window, New SearchQuery(text, mode, False))
            type.GetField("_scanRootFolder", flags).SetValue(window, If(Directory.Exists(root), root, ""))
            Try
                Dim work = Task.Run(Async Function()
                                        Using cancelled As New CancellationTokenSource()
                                            cancelled.Cancel()
                                            Dim cancelledWork = DirectCast(type.GetMethod("CountSourceFilesAsync", flags).Invoke(window, {root, CType(cancelled.Token, Object)}), Task)
                                            Try
                                                Await cancelledWork
                                                Throw New InvalidOperationException("Counting swallowed cancellation.")
                                            Catch ex As OperationCanceledException
                                            End Try
                                            Ensure(Not CBool(type.GetField("_isCountingFiles", flags).GetValue(window)), "Cancelled counting left the phase active.")
                                        End Using
                                        Await DirectCast(type.GetMethod("CountSourceFilesAsync", flags).Invoke(window, {root, CType(source.Token, Object)}), Task)
                                        Ensure(CInt(type.GetField("_filesScanned", flags).GetValue(window)) = 0, "Counting incremented the scanned counter.")
                                        Ensure(Not DirectCast(type.GetField("_hits", flags).GetValue(window), IEnumerable).Cast(Of Object)().Any(), "Counting published search results.")
                                        Ensure(Not DirectCast(type.GetField("_tempDirectories", flags).GetValue(window), IEnumerable).Cast(Of Object)().Any(), "Counting retained extracted directories.")
                                        If expectedCount.HasValue Then Ensure(CInt(type.GetField("_totalFilesToScan", flags).GetValue(window)) = expectedCount.Value, "Full nested count is incorrect for " & root & ".")
                                        Dim task = DirectCast(type.GetMethod("SearchSourceAsync", flags).Invoke(window, {root, CType(source.Token, Object)}), Task)
                                        Await task
                                        If Not CBool(type.GetField("_resultLimitReached", flags).GetValue(window)) Then
                                            Ensure(CInt(type.GetField("_filesScanned", flags).GetValue(window)) = CInt(type.GetField("_totalFilesToScan", flags).GetValue(window)), "Count and search disagree on eligible files.")
                                        End If
                                    End Function)
                While Not work.IsCompleted
                    window.Dispatcher.Invoke(Sub()
                                             End Sub, DispatcherPriority.ContextIdle)
                    Thread.Sleep(5)
                End While
                work.GetAwaiter().GetResult()
                Dim hits = DirectCast(type.GetField("_hits", flags).GetValue(window), IEnumerable).Cast(Of Object)().ToList()
                Try
                    check(hits)
                    If inspect IsNot Nothing Then
                        window.WindowStartupLocation = System.Windows.WindowStartupLocation.Manual
                        window.WindowState = System.Windows.WindowState.Normal
                        window.Left = -10000
                        window.Top = -10000
                        window.ShowActivated = False
                        window.ShowInTaskbar = False
                        window.Show()
                        Pump(window)
                        inspect(window, hits)
                    End If
                Catch ex As Exception
                    Dim issues = DirectCast(type.GetField("_diagnostics", flags).GetValue(window), ScanDiagnosticStore).Snapshot().Items.
                        Select(Function(item) item.SourcePath & ": " & item.Message)
                    Throw New InvalidOperationException(ex.Message & " Reported scan issues: " & String.Join("; ", issues), ex)
                End Try
                DirectCast(window.FindName("Search_txt"), TextBox).Text = "a different future query"
                Dim actual = DirectCast(type.GetMethod("PreviewSpans", flags).Invoke(window, {text}), List(Of SearchSpan))
                Ensure(actual.Count = New SearchQuery(text, mode, False).FindHighlights(text).Count, "Editing input changed the scan snapshot.")
            Finally
                type.GetMethod("CleanupTemp", flags).Invoke(window, Nothing)
                window.Close()
            End Try
        End Using
    End Sub

    Private Function PropertyValue(hit As Object, name As String) As Object
        Return hit.GetType().GetProperty(name).GetValue(hit)
    End Function
    Private Function Details(hit As Object) As List(Of SearchDetail)
        Return DirectCast(PropertyValue(hit, "Details"), List(Of SearchDetail))
    End Function
    Private Sub Ensure(value As Boolean, message As String)
        If Not value Then Throw New InvalidOperationException(message)
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
