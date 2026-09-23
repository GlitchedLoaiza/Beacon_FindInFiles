Imports System.Diagnostics
Imports System.IO
Imports System.Text
Imports System.Threading
Imports Beacon

Module Program
    Private _passed As Integer
    Private _failed As Integer

    <STAThread>
    Function Main(args As String()) As Integer
        If args.Length = 3 AndAlso args(0) = "--measure-production-app" Then Return ApplicationScalingMeasurements.Run(args(1), args(2))
        If args.Length = 3 AndAlso args(0) = "--single-instance-check" Then Return SingleInstanceChecks.RunChildMode(args)
        If args.Length > 0 AndAlso (args(0) = "l" OrElse args(0) = "e") Then Return FakeReader(args)
        If args.Length <> 1 OrElse Not File.Exists(args(0)) Then
            Console.Error.WriteLine("Usage: dotnet run --project tests/Beacon.SafetyChecks -- <path to bundled 7za.exe>")
            Return 2
        End If

        Check("XML indentation and declaration", Sub()
                                                      Dim xml = "<?xml version=""1.0""?><root><child>error</child></root>"
                                                      Dim rendered = PreviewFormatting.FormatXml(xml, True, 1024)
                                                      Require(rendered.Contains(Environment.NewLine & "  <child>error</child>"), "XML was not indented.")
                                                      Require(rendered.StartsWith("<?xml version=""1.0""?>", StringComparison.Ordinal), "Declaration lost.")
                                                  End Sub)
        Check("XML formatting can be disabled", Sub()
                                                    Dim xml = "<root><child/></root>"
                                                    Require(PreviewFormatting.FormatXml(xml, False, 1024) = xml, "Disabled formatting changed XML.")
                                                End Sub)
        Check("Malformed XML remains readable", Sub()
                                                    Dim xml = "<root><unclosed>"
                                                    Require(PreviewFormatting.FormatXml(xml, True, 1024) = xml, "Malformed input changed.")
                                                End Sub)
        Check("DTD and oversized XML remain unexpanded", Sub()
                                                            Dim xml = "<!DOCTYPE root [<!ENTITY test 'expanded'>]><root>&test;</root>"
                                                            Require(PreviewFormatting.FormatXml(xml, True, 1024) = xml, "DTD was processed.")
                                                            xml = "<root>" & New String("x"c, 100) & "</root>"
                                                            Require(PreviewFormatting.FormatXml(xml, True, 20) = xml, "Oversized XML was parsed.")
                                                        End Sub)
        Check("HTML uses a low-priority light canvas", Sub()
                                                          Require(PreviewFormatting.HtmlCanvasStyle.Contains("@layer"), "Author precedence is missing.")
                                                          Require(PreviewFormatting.HtmlCanvasStyle.Contains("background-color: white"), "Missing light canvas.")
                                                          Require(Not PreviewFormatting.HtmlCanvasStyle.Contains("!important"), "Canvas overrides author styling.")
                                                      End Sub)
        Check("Unsafe archive paths rejected", Sub()
                                                   For Each name In {"../escape", "dir\..\escape", "C:\escape", "\\server\share", "/absolute", "file:stream", "NUL.txt", "dir/COM1", "file. ", "bad?name"}
                                                       Reject(Of InvalidDataException)(Sub() ArchiveSafety.ValidateEntryName(name))
                                                   Next
                                                   ArchiveSafety.ValidateEntryName("./logs/report.xml")
                                                   ArchiveSafety.ValidateEntryName("logs with spaces/report.xml")
                                               End Sub)
        Check("Encrypted entries and links rejected", Sub()
                                                         Reject(Of InvalidDataException)(Sub() NewBudget().Register("a.txt", 1, 1, True))
                                                         Reject(Of InvalidDataException)(Sub() NewBudget().Register("a.txt", 1, 1, False, "outside"))
                                                     End Sub)
        Check("Entry count limit enforced", Sub()
                                                Dim settings As New BeaconSettings With {.MaximumArchiveEntries = 1}
                                                Dim budget As New ArchiveReadBudget(settings, 1000)
                                                budget.Register("a", 1, 1, False)
                                                Reject(Of InvalidDataException)(Sub() budget.Register("b", 1, 1, False))
                                            End Sub)
        Check("Entry size and ratio limits enforced", Sub()
                                                         Dim settings As New BeaconSettings With {.MaximumArchiveEntrySizeMb = 1, .MaximumCompressionRatio = 2}
                                                         Dim budget As New ArchiveReadBudget(settings, 1024 * 1024)
                                                         Reject(Of InvalidDataException)(Sub() budget.Register("large", 1024 * 1024 + 1, 0, False))
                                                         Reject(Of InvalidDataException)(Sub() NewBudget(settings).Register("ratio", 100, 1, False))
                                                     End Sub)
        Check("Aggregate actual byte budget enforced", Sub()
                                                          Dim budget As New ArchiveReadBudget(New BeaconSettings With {.MaximumCompressionRatio = 2}, 4)
                                                          budget.Consume(5)
                                                          Reject(Of InvalidDataException)(Sub() budget.Consume(4))
                                                      End Sub)
        Check("Bounded stream allows exact limit and EOF", Sub()
                                                              Using input As New BoundedReadStream(New MemoryStream({CByte(1), CByte(2)}), 2, CancellationToken.None)
                                                                  Using output As New MemoryStream()
                                                                      input.CopyTo(output)
                                                                      Require(output.Length = 2, "Exact limit failed.")
                                                                  End Using
                                                              End Using
                                                          End Sub)
        Check("Bounded stream detects lying metadata", Sub()
                                                          Dim reported As Boolean
                                                          Using input As New BoundedReadStream(New MemoryStream({CByte(1), CByte(2), CByte(3)}), 2, CancellationToken.None,
                                                                                               onError:=Sub(ex) reported = True)
                                                              Reject(Of InvalidDataException)(Sub() input.CopyTo(Stream.Null))
                                                          End Using
                                                          Require(reported, "Limit failure was not reported.")
                                                      End Sub)
        Check("Async reads enforce byte limits", Sub()
                                                     Using input As New BoundedReadStream(New MemoryStream(New Byte(31) {}), 2, CancellationToken.None)
                                                         Reject(Of InvalidDataException)(Sub() input.CopyToAsync(Stream.Null).GetAwaiter().GetResult())
                                                     End Using
                                                 End Sub)
        Check("Read cancellation honored", Sub()
                                               Using source As New CancellationTokenSource()
                                                   source.Cancel()
                                                   Using input As New BoundedReadStream(New MemoryStream({CByte(1)}), 1, source.Token)
                                                       Reject(Of OperationCanceledException)(Sub() input.ReadByte())
                                                   End Using
                                               End Using
                                           End Sub)
        Check("Filesystem filters and size reporting", AddressOf TestTraversal)
        Check("Followed junction cycle terminates", AddressOf TestJunction)
        Check("Traversal cancellation honored", Sub()
                                                    Using source As New CancellationTokenSource()
                                                        source.Cancel()
                                                        Reject(Of OperationCanceledException)(Sub() FileSystemTraversal.EnumerateFiles(Path.GetTempPath(), source.Token, False, Nothing, Nothing).ToArray())
                                                    End Using
                                                End Sub)
        Check("CAB extraction with spaces in filename", Sub() TestCab(args(0)))
        Check("Corrupt CAB fails safely", Sub()
                                              WithFixture(Sub(root)
                                                              Dim cab = Path.Combine(root, "corrupt.cab")
                                                              File.WriteAllText(cab, "not a cabinet")
                                                              Reject(Of InvalidDataException)(Sub() CabExtractor.ExtractAsync(args(0), cab, Path.Combine(root, "out"), New BeaconSettings(), CancellationToken.None).GetAwaiter().GetResult())
                                                          End Sub)
                                          End Sub)
        Check("CAB output cannot exceed advertised size", Sub() TestFakeReader("overflow", False))
        Check("CAB paths rejected before extraction", Sub() TestFakeReader("unsafe", False))
        Check("CAB timeout stops stalled reader", Sub() TestFakeReader("hang", False))
        Check("CAB cancellation stops active reader", Sub() TestFakeReader("hang", True))
        Check("Settings light theme templates and contrast", Sub() SettingsThemeChecks.Run(False))
        Check("Settings dark theme templates and contrast", Sub() SettingsThemeChecks.Run(True))
        Check("Main Path/Search fields in light theme", Sub() SettingsThemeChecks.RunMainFields(False))
        Check("Main Path/Search fields in dark theme", Sub() SettingsThemeChecks.RunMainFields(True))
        Check("Preview counters from light to dark", Sub() SettingsThemeChecks.RunCounterBars(False))
        Check("Preview counters from dark to light", Sub() SettingsThemeChecks.RunCounterBars(True))
        Check("Compact toolbar and wrench in light theme", Sub() SettingsThemeChecks.RunToolbar(False))
        Check("Compact toolbar and wrench in dark theme", Sub() SettingsThemeChecks.RunToolbar(True))
        Check("Compact results and optional details in light theme", Sub() SettingsThemeChecks.RunCompactResults(False))
        Check("Compact results and optional details in dark theme", Sub() SettingsThemeChecks.RunCompactResults(True))
        Check("Splash scene, motion preferences and cleanup", AddressOf SplashChecks.SceneAndMotion)
        Check("Splash work-area scaling", AddressOf SplashChecks.WorkAreaSizing)
        Check("Search modes, validation and regex safety", AddressOf SearchChecks.Queries)
        Check("Bidirectional text cursor, Unicode and wraparound", AddressOf TextNavigationChecks.Run)
        Check("Captured summary parity, overlapping windows and anchors", AddressOf TextSummaryChecks.Projection)
        Check("Summary preview, reverse navigation, source independence and session mode", AddressOf TextSummaryChecks.Preview)
        Check("Text match counter navigation and theme transitions", AddressOf TextCounterChecks.NavigationAndThemes)
        Check("Text match counter loading, limits, errors and clearing", AddressOf TextCounterChecks.LoadingLimitsAndErrors)
        Check("Text match counter zero-width anchors and metadata", AddressOf TextCounterChecks.RecordAnchorsAndMetadata)
        Check("Reset and new scans clear all text, event and HAR navigation", AddressOf PreviewResetChecks.Run)
        Check("Text preview direction, newline and Unicode selection", AddressOf PreviewAcceptanceChecks.Directions)
        Check("ZIP, nested and CAB summaries without source reads", AddressOf PreviewAcceptanceChecks.ArchivedSummary)
        Check("Summary metadata, preview bounds and zero-width anchors", AddressOf PreviewAcceptanceChecks.LimitsAndMetadata)
        Check("Obsolete preview readers cannot replace current content", AddressOf PreviewAcceptanceChecks.StaleLoads)
        Check("Summary themes, minimum layout and bounded virtualization", AddressOf PreviewPresentationChecks.Run)
        Check("Bounded matching text records", AddressOf SearchChecks.TextRecords)
        Check("Disk, HAR and archive search pipeline", AddressOf SearchChecks.Pipeline)
        Check("SharpCompress archive formats and byte-exact rereads", Sub() ArchiveChecks.Formats(args(0)))
        Check("Compressed TAR preview, nesting and logical paths", Sub() ArchiveChecks.PreviewAndNested(args(0)))
        Check("ZIP hidden and system attributes", AddressOf ArchiveChecks.ZipAttributes)
        Check("Archive encryption, paths, links and checksums", Sub() ArchiveChecks.RejectionAndChecksums(args(0)))
        Check("Compressed TAR limits, cancellation and cleanup", Sub() ArchiveChecks.LimitsAndCleanup(args(0)))
        Check("RAR and RAR5 solid payload hashes", AddressOf ArchiveChecks.RarPayloads)
        Check("Malicious archive directories and preview limits", AddressOf ArchiveChecks.DirectoryTraversalAndPreviewLimits)
        Check("ZIP checksum streams and corrupt payload rejection", AddressOf ZipChecksumChecks.Run)
        Check("Readable bounded previews and Unicode boundaries", AddressOf PreviewLimitChecks.Run)
        Check("Selected archive labels, nested paths and previews", AddressOf ArchiveDisplayChecks.Run)
        Check("Text, event, HAR navigation and CAB search", AddressOf SearchChecks.NavigationAndCab)
        Check("WebView2 query ranges and stale captures", AddressOf WebSearchChecks.Run)
        Check("Bounded concurrent diagnostics", AddressOf ReportChecks.DiagnosticBounds)
        Check("Report snapshots and safe copy formatting", AddressOf ReportChecks.SnapshotAndCopy)
        Check("CSV/JSON/HTML export and atomic writes", AddressOf ReportChecks.FormatsAndAtomicWrites)
        Check("Scan report filtering and light theme", Sub() SettingsThemeChecks.RunReport(False))
        Check("Scan report filtering and dark theme", Sub() SettingsThemeChecks.RunReport(True))
        Check("Main-window report snapshot and action layout", AddressOf ReportChecks.MainWindowSnapshot)
        Check("Five-line context and readable HTML reports", AddressOf ContextReportChecks.Run)
        Check("Startup stable-release checks and silent failures", AddressOf ReleaseUpdateChecks.Run)
        Check("Settings inline update results and cancellation", AddressOf ReleaseUpdateChecks.SettingsResults)
        Check("EVTX filters, XML tools and offline guidance", AddressOf EvtxChecks.Run)
        Check("Native caption themes preserve Windows frames", AddressOf NativeCaptionChecks.Transitions)
        Check("Single-instance ownership, activation and crash recovery", AddressOf SingleInstanceChecks.Run)
        Check("Editable UTC picker, severity and provider dropdowns", AddressOf FilterControlChecks.Run)
        Check("HAR filters, bounded decoding and redacted exports", AddressOf HarChecks.Run)
        Check("Headless HAR service isolation and failure semantics", AddressOf StructuredSearchChecks.Har)
        Check("Headless EVTX service isolation and XML semantics", AddressOf StructuredSearchChecks.Events)
        Check("Beginner help and multiple-record defaults", AddressOf HelpChecks.Run)
        Check("Headless preview service limits, sources and cleanup", AddressOf PreviewContentChecks.Run)
        Check("Headless source traversal, snapshots and lifetime", AddressOf SourceSearchChecks.Run)
        Check("Automatic processor and memory policy", AddressOf MulticoreChecks.Policy)
        Check("Heavy-work bounds, reserved head lane and cancellation", AddressOf MulticoreChecks.Resources)
        Check("Ordered multicore results and completed progress", AddressOf MulticoreChecks.Parity)
        Check("Ordered result caps, callback faults and cancellation", AddressOf MulticoreChecks.Termination)
        Check("Parallel nested preview lifetime and cancelled-result cleanup", AddressOf MulticoreChecks.Lifetime)
        Check("Ordered archive floods, capped prefixes and cancelled backpressure", AddressOf MulticoreStressChecks.Backpressure)
        Check("Mixed-source parallel parity, redaction, solid archives and resource fallback", AddressOf MulticoreStressChecks.MixedSources)
        Check("Once-only optional welcome tour and splash version", AddressOf TourChecks.Run)
        Check("Help tour replay preserves existing, absent and blocked marker state", AddressOf TourReplayChecks.MarkerIsolation)
        Check("Tour replay and Summary remain independent of active multicore search", AddressOf TourReplayChecks.ActiveSearch)
        Check("Optimized line-context order and isolation", AddressOf OptimizationChecks.TextContextParity)
        Check("Beacon Theme system backgrounds and logo-red buttons", AddressOf BeaconThemeChecks.Run)
        Console.WriteLine($"{_passed} passed; {_failed} failed.")
        Return If(_failed = 0, 0, 1)
    End Function

    Private Function NewBudget(Optional settings As BeaconSettings = Nothing) As ArchiveReadBudget
        Return New ArchiveReadBudget(If(settings, New BeaconSettings()), 1024 * 1024)
    End Function

    Private Function FakeReader(args As String()) As Integer
        Dim archive = args(Array.IndexOf(args, "--") + 1)
        Dim mode = File.ReadAllText(archive)
        If args(0) = "l" Then
            Console.WriteLine("Type = Cab")
            Console.WriteLine("----------")
            Console.WriteLine("Path = " & If(mode = "unsafe", "../escape.txt", "payload.txt"))
            Console.WriteLine("Size = 2")
            Console.WriteLine()
        ElseIf mode = "hang" Then
            Thread.Sleep(30000)
        Else
            Using output = Console.OpenStandardOutput()
                output.Write({CByte(1), CByte(2), CByte(3)}, 0, 3)
            End Using
        End If
        Return 0
    End Function

    Private Sub TestFakeReader(mode As String, cancel As Boolean)
        WithFixture(Sub(root)
                        Dim cab = Path.Combine(root, "fake.cab")
                        File.WriteAllText(cab, mode)
                        Dim outputDirectory = Path.Combine(root, "output")
                        Dim settings As New BeaconSettings With {.ArchiveProcessTimeoutSeconds = If(cancel, 10, 1)}
                        Dim timer = Stopwatch.StartNew()
                        Using cancellation As New CancellationTokenSource()
                            If cancel Then cancellation.CancelAfter(500)
                            Dim extract As Action = Sub() CabExtractor.ExtractAsync(Environment.ProcessPath, cab, outputDirectory, settings, cancellation.Token).GetAwaiter().GetResult()
                            If mode = "hang" Then
                                Reject(Of OperationCanceledException)(extract)
                                Require(timer.Elapsed < TimeSpan.FromSeconds(8), "Reader did not stop promptly.")
                            Else
                                Reject(Of InvalidDataException)(extract)
                            End If
                        End Using
                        If mode = "unsafe" Then Require(Not Directory.Exists(outputDirectory), "Unsafe metadata produced output.")
                        If mode = "overflow" Then Require(New FileInfo(Path.Combine(outputDirectory, "payload.txt")).Length <= 2, "Oversized output was written.")
                    End Sub)
    End Sub

    Private Sub TestTraversal()
        WithFixture(Sub(root)
                        Directory.CreateDirectory(Path.Combine(root, ".git"))
                        File.WriteAllText(Path.Combine(root, ".git", "excluded.txt"), "error")
                        File.WriteAllText(Path.Combine(root, "included.txt"), "error")
                        Dim hidden = Path.Combine(root, "hidden.txt")
                        File.WriteAllText(hidden, "error")
                        File.SetAttributes(hidden, FileAttributes.Hidden)
                        Using large = File.Create(Path.Combine(root, "large.txt"))
                            large.SetLength(1024 * 1024 + 1)
                        End Using
                        Dim issues As New List(Of String)()
                        Dim settings As New BeaconSettings With {.MaximumFileSizeMb = 1, .IncludeHiddenFiles = False}
                        Dim files = FileSystemTraversal.EnumerateFiles(root, CancellationToken.None, False, Nothing,
                                                                      Sub(p, ex) issues.Add(p), settings).ToArray()
                        Require(files.Length = 1 AndAlso Path.GetFileName(files(0)) = "included.txt", "Filesystem filters were not applied.")
                        Require(issues.Count = 1, "Size skip was not reported exactly once.")
                    End Sub)
    End Sub

    Private Sub TestJunction()
        WithFixture(Sub(root)
                        File.WriteAllText(Path.Combine(root, "one.txt"), "error")
                        Dim link = Path.Combine(root, "loop")
                        Dim start As New ProcessStartInfo(Path.Combine(Environment.SystemDirectory, "cmd.exe")) With {.UseShellExecute = False, .CreateNoWindow = True}
                        start.Arguments = $"/c mklink /J ""{link}"" ""{root}"""

                        RunProcess(start)
                        Try
                            Using source As New CancellationTokenSource(TimeSpan.FromSeconds(5))
                                Dim files = FileSystemTraversal.EnumerateFiles(root, source.Token, True, Nothing,
                                                                              Sub(p, ex) Throw New IOException(p, ex), New BeaconSettings With {.FollowReparsePoints = True}).ToArray()
                                Require(files.Length = 1, "Junction caused duplicate traversal.")
                            End Using
                        Finally
                            Directory.Delete(link)
                        End Try
                    End Sub)
    End Sub

    Private Sub TestCab(executable As String)
        WithFixture(Sub(root)
                        Dim original = Path.Combine(root, "sample payload.txt")
                        Dim bytes(4095) As Byte
                        Dim random As New Random(1234)
                        random.NextBytes(bytes)
                        File.WriteAllBytes(original, bytes)
                        Dim cab = Path.Combine(root, "sample.cab")
                        Dim start As New ProcessStartInfo(Path.Combine(Environment.SystemDirectory, "makecab.exe")) With {.UseShellExecute = False, .CreateNoWindow = True}
                        start.ArgumentList.Add(original)
                        start.ArgumentList.Add(cab)
                        RunProcess(start)
                        Dim destination = Path.Combine(root, "out")
                        CabExtractor.ExtractAsync(executable, cab, destination, New BeaconSettings(), CancellationToken.None).GetAwaiter().GetResult()
                        Require(File.ReadAllBytes(Path.Combine(destination, Path.GetFileName(original))).SequenceEqual(bytes), "CAB payload differs.")
                        Using cancelled As New CancellationTokenSource()
                            cancelled.Cancel()
                            Reject(Of OperationCanceledException)(Sub() CabExtractor.ExtractAsync(executable, cab, Path.Combine(root, "cancelled"), New BeaconSettings(), cancelled.Token).GetAwaiter().GetResult())
                        End Using
                    End Sub)
    End Sub

    Private Sub RunProcess(start As ProcessStartInfo)
        Using child = Process.Start(start)
            If child Is Nothing Then Throw New IOException("Test process did not start.")
            If Not child.WaitForExit(15000) Then
                child.Kill(True)
                Throw New TimeoutException("Test process timed out.")
            End If
            Require(child.ExitCode = 0, $"Test process failed: {start.FileName} ({child.ExitCode}).")
        End Using
    End Sub

    Private Sub WithFixture(action As Action(Of String))
        Dim root = Path.Combine(Path.GetTempPath(), "BeaconSafetyTests-" & Guid.NewGuid().ToString("N"))
        Directory.CreateDirectory(root)
        Try
            action(root)
        Finally
            Directory.Delete(root, True)
        End Try
    End Sub

    Private Sub Check(name As String, action As Action)
        Try
            action()
            _passed += 1
            Console.WriteLine("PASS " & name)
        Catch ex As Exception
            _failed += 1
            Console.Error.WriteLine($"FAIL {name}: {ex}")
        End Try
    End Sub

    Private Sub Require(condition As Boolean, message As String)
        If Not condition Then Throw New InvalidOperationException(message)
    End Sub

    Private Sub Reject(Of T As Exception)(action As Action)
        Try
            action()
        Catch ex As T
            Return
        End Try
        Throw New InvalidOperationException($"Expected {GetType(T).Name}.")
    End Sub
End Module
