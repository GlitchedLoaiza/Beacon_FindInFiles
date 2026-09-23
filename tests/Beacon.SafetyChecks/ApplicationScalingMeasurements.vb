Imports System.Collections
Imports System.Diagnostics
Imports System.IO
Imports System.Reflection
Imports System.Text.Json
Imports System.Threading
Imports System.Threading.Tasks
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Threading
Imports Beacon

Friend Module ApplicationScalingMeasurements
    Private Const Flags As BindingFlags = BindingFlags.Instance Or BindingFlags.NonPublic
    Private Const Marker As String = "BEACON_PRODUCTION_SCALING_MATCH"

    Public Function Run(corpus As String, evidence As String) As Integer
        Console.WriteLine("APPLICATION measurement starting: " & DateTimeOffset.UtcNow.ToString("O"))
        Dim window As MainWindow = Nothing
        Try
            Using manifest = JsonDocument.Parse(File.ReadAllText(Path.Combine(corpus, "beacon-production-corpus.json")))
                Dim files = manifest.RootElement.GetProperty("Files").GetInt32()
                If files < Environment.ProcessorCount Then Throw New InvalidDataException("Insufficient independent sources for this machine.")
                Directory.CreateDirectory(evidence)
                Dim application = If(System.Windows.Application.Current, New System.Windows.Application With {.ShutdownMode = ShutdownMode.OnExplicitShutdown})
                Console.WriteLine("APPLICATION creating measurement window")
                window = New MainWindow()
                Console.WriteLine("APPLICATION window constructed")
                window.ApplicationExitForChecks = Sub() window.Dispatcher.BeginInvoke(New Action(AddressOf window.Close), DispatcherPriority.Background)
                window.RemoveHandler(FrameworkElement.LoadedEvent,
                    GetType(MainWindow).GetMethod("MainWindow_Loaded", Flags).CreateDelegate(GetType(RoutedEventHandler), window))
                Dim options As New BeaconSettings With {
                    .IncludedExtensions = ".log;.json", .StopAfterFirstMatchPerFile = False,
                    .SearchFileNames = False, .SearchFullPaths = False, .SearchFileContents = True,
                    .MaximumStructuredMatches = 100000, .MaximumTotalResults = 100000,
                    .MaximumFileSizeMb = 1024, .MaximumArchiveEntrySizeMb = 1024,
                    .MaximumArchiveExpandedSizeMb = 2048, .MaximumCompressionRatio = 100, .SelectFirstResult = False}
                GetType(MainWindow).GetField("_settings", Flags).SetValue(window, options)
                GetType(MainWindow).GetMethod("ApplySettings", Flags).Invoke(window, Nothing)
                window.WindowStartupLocation = WindowStartupLocation.Manual
                window.WindowState = WindowState.Normal
                window.Width = 1000
                window.Height = 700
                window.Left = -10000
                window.Top = -10000
                window.ShowActivated = False
                window.ShowInTaskbar = False
                window.Show()
                Pump(window)
                Console.WriteLine("APPLICATION window ready")
                Dim workloads As (Name As String, Source As String, Files As Integer, Preview As Boolean)() = {
                    ("PlainFiles", Path.Combine(corpus, "plain"), files, True),
                    ("CompressedZips", Path.Combine(corpus, "zip"), files, True),
                    ("SmallFiles", Path.Combine(corpus, "small"), 1024, True),
                    ("DenseResults", Path.Combine(corpus, "dense"), files, True),
                    ("WholeDocuments", Path.Combine(corpus, "documents"), files, False),
                    ("SingleArchive", Path.Combine(corpus, "zip", "00000.zip"), 1, True)}
                For Each workload In workloads
                    For repetition As Integer = 0 To 3
                        For Each workers In If(repetition Mod 2 = 0, New Integer() {1, 0}, New Integer() {0, 1})
                            window.ScanPolicyForChecks = If(workers = 0, Nothing, New SearchConcurrencyPolicy(1))
                            GetType(MainWindow).GetField("_summaryPreferred", Flags).SetValue(window, False)
                            DirectCast(window.FindName("Path_txt"), TextBox).Text = workload.Source
                            DirectCast(window.FindName("Search_txt"), TextBox).Text = Marker
                            Console.WriteLine($"APPLICATION BEGIN {workload.Name} workers={workers} repetition={repetition}")
                            Dim allocated = GC.GetTotalAllocatedBytes(True)
                            Using process = System.Diagnostics.Process.GetCurrentProcess()
                                Dim cpu = process.TotalProcessorTime
                                Dim peak = process.WorkingSet64
                                Dim clock = Stopwatch.StartNew()
                                GetType(MainWindow).GetMethod("StartScan", Flags).Invoke(window, Nothing)
                                Dim task = DirectCast(GetType(MainWindow).GetField("_scanTask", Flags).GetValue(window), Task)
                                If task Is Nothing Then Throw New InvalidOperationException("The application did not start the scan.")
                                Dim countObservedMs As Double? = Nothing
                                Dim maxDispatcherMs As Double = 0
                                Dim sampledAt As Long = -20
                                While Not task.IsCompleted
                                    Dim started = Stopwatch.GetTimestamp()
                                    Pump(window)
                                    maxDispatcherMs = Math.Max(maxDispatcherMs, Stopwatch.GetElapsedTime(started).TotalMilliseconds)
                                    If Not countObservedMs.HasValue AndAlso Not DirectCast(window.FindName("ScanProgress_pb"), ProgressBar).IsIndeterminate Then
                                        countObservedMs = clock.Elapsed.TotalMilliseconds
                                    End If
                                    If clock.ElapsedMilliseconds - sampledAt >= 20 Then
                                        process.Refresh()
                                        peak = Math.Max(peak, process.WorkingSet64)
                                        sampledAt = clock.ElapsedMilliseconds
                                    End If
                                    If clock.Elapsed > TimeSpan.FromMinutes(10) Then Throw New TimeoutException("Application measurement exceeded its bound.")
                                    Thread.Sleep(1)
                                End While
                                task.GetAwaiter().GetResult()
                                Dim scanMs = clock.Elapsed.TotalMilliseconds
                                Dim scanAllocations = GC.GetTotalAllocatedBytes(True) - allocated
                                process.Refresh()
                                Dim cpuMs = (process.TotalProcessorTime - cpu).TotalMilliseconds
                                peak = Math.Max(peak, process.WorkingSet64)
                                Console.WriteLine($"APPLICATION SEARCH COMPLETE {workload.Name}: {scanMs:F1} ms")
                                Dim results = DirectCast(window.FindName("Results_lst"), ListBox)
                                Dim scanned = CInt(GetType(MainWindow).GetField("_filesScanned", Flags).GetValue(window))
                                Dim counted = CInt(GetType(MainWindow).GetField("_totalFilesToScan", Flags).GetValue(window))
                                Dim runInfo = DirectCast(GetType(MainWindow).GetField("_scanRun", Flags).GetValue(window), ScanRunInfo)
                                Dim diagnostics = DirectCast(GetType(MainWindow).GetField("_diagnostics", Flags).GetValue(window), ScanDiagnosticStore)
                                If scanned <> workload.Files OrElse counted <> workload.Files OrElse results.Items.Count <> workload.Files OrElse
                                   runInfo.State <> ScanReportState.Completed OrElse diagnostics.Count <> 0 Then
                                    Throw New InvalidOperationException($"Application scan differs from expected coverage: {workload.Name}; state={runInfo.State}; counted={counted}; scanned={scanned}; results={results.Items.Count}; issues={diagnostics.Count}.")
                                End If
                                Dim runs = DirectCast(GetType(MainWindow).GetField("_sourceSearches", Flags).GetValue(window), IEnumerable).Cast(Of Object)()
                                Dim coordinated = runs.OfType(Of SearchRunCoordinator)().Single()
                                Dim fullMs As Double? = Nothing
                                Dim summaryMs As Double? = Nothing
                                If workload.Preview Then
                                    Console.WriteLine("APPLICATION FULL PREVIEW BEGIN")
                                    clock.Restart()
                                    results.SelectedIndex = 0
                                    PreviewTestHelpers.WaitForPreview(window)
                                    fullMs = clock.Elapsed.TotalMilliseconds
                                    Console.WriteLine($"APPLICATION FULL PREVIEW COMPLETE: {fullMs:F1} ms")
                                    Console.WriteLine("APPLICATION SUMMARY BEGIN")
                                    clock.Restart()
                                    DirectCast(window.FindName("TextPreviewMode_cmb"), ComboBox).SelectedIndex = 1
                                    PreviewTestHelpers.WaitForPreview(window)
                                    summaryMs = clock.Elapsed.TotalMilliseconds
                                    If Not DirectCast(window.FindName("SummaryPreview_lst"), ListBox).IsVisible Then Throw New InvalidOperationException("Summary interaction did not complete.")
                                End If
                                Dim record = New With {
                                    .Workload = workload.Name, .Workers = workers, .Mode = If(workers = 0, "Automatic", "ForcedOne"),
                                    .Repetition = repetition, .Warmup = repetition = 0, .LogicalProcessors = Environment.ProcessorCount,
                                    .Files = scanned, .PeakActiveJobs = coordinated.PeakWorkers, .ScanMs = scanMs,
                                    .CountPhaseObservedMs = countObservedMs, .CpuMs = cpuMs, .AllocatedBytes = scanAllocations,
                                    .SampledPeakWorkingSetBytes = peak, .MaxDispatcherRoundTripMs = maxDispatcherMs,
                                    .FullPreviewMs = fullMs, .SummarySwitchMs = summaryMs}
                                File.AppendAllText(Path.Combine(evidence, "application.jsonl"), JsonSerializer.Serialize(record) & Environment.NewLine)
                                Console.WriteLine($"APPLICATION {workload.Name} workers={workers} repetition={repetition}: count+scan+publication={scanMs:F1} ms; jobs={coordinated.PeakWorkers}; full={fullMs:F1} ms; summary={summaryMs:F1} ms.")
                            End Using
                        Next
                    Next
                Next
            End Using
            Return 0
        Catch ex As Exception
            Console.Error.WriteLine(ex)
            Return 1
        Finally
            If window IsNot Nothing AndAlso window.IsVisible Then
                Dim cancellation = DirectCast(GetType(MainWindow).GetField("_scanCts", Flags).GetValue(window), CancellationTokenSource)
                cancellation?.Cancel()
                window.Close()
                Dim timeout = Stopwatch.StartNew()
                While window.IsVisible AndAlso timeout.Elapsed < TimeSpan.FromMinutes(1)
                    Pump(window)
                    Thread.Sleep(1)
                End While
            End If
        End Try
    End Function

    Private Sub Pump(window As MainWindow)
        window.Dispatcher.Invoke(Sub()
                                 End Sub, DispatcherPriority.ContextIdle)
    End Sub
End Module
