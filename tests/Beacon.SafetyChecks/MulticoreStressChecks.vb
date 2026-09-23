Imports System.IO
Imports System.IO.Compression
Imports System.Text.Json
Imports System.Threading
Imports System.Threading.Tasks
Imports Beacon

Module MulticoreStressChecks
    Private ReadOnly Extensions As String() = {".zip", ".cab", ".rar", ".gz", ".tar", ".7z"}

    Public Sub Backpressure()
        Dim root = Path.Combine(Path.GetTempPath(), "BeaconBackpressure-" & Guid.NewGuid().ToString("N"))
        Directory.CreateDirectory(root)
        Try
            For archiveIndex = 0 To 2
                Using archive = ZipFile.Open(Path.Combine(root, $"{archiveIndex:D2}.zip"), ZipArchiveMode.Create)
                    For entryIndex = 0 To 99
                        Using writer As New StreamWriter(archive.CreateEntry($"{entryIndex:D3}.log").Open())
                            writer.Write($"error {archiveIndex}:{entryIndex}")
                        End Using
                    Next
                End Using
            Next
            Dim query As New SearchQuery("error", SearchMode.PlainText, False)
            Dim options As New BeaconSettings()
            Dim expected = Sequential(root, query, options)
            For Each limit In {1, 105, 10000}
                Dim settings = BeaconSettingsService.Clone(options)
                settings.MaximumTotalResults = limit
                Dim flood As New TaskCompletionSource(Of Boolean)(TaskCreationOptions.RunContinuationsAsynchronously)
                Dim output As New List(Of SourceSearchResult)()
                Using timeout As New CancellationTokenSource(TimeSpan.FromSeconds(20))
                    Using run As New SearchRunCoordinator(query, settings, Extensions, Policy(3), AddressOf output.Add)
                        run.BeforeSourceAsync = Async Function(path, ordinal, token)
                                                    If ordinal = 0 Then Await flood.Task.WaitAsync(token).ConfigureAwait(False)
                                                End Function
                        run.EventQueuedForChecks = Sub(ordinal, kind)
                                                       If ordinal = 1 AndAlso kind = SourceSearchEventKind.Result Then flood.TrySetResult(True)
                                                   End Sub
                        run.RunAsync(root, timeout.Token).GetAwaiter().GetResult()
                        Require(Snapshot(output) = Snapshot(expected.Take(limit)), "Backpressure changed the ordered capped prefix.")
                        Require(run.FilesProcessed = Math.Min(limit, 300) AndAlso run.ResultLimitReached = (limit <= 300),
                                "Archive containers were double-counted or capped completion was wrong.")
                    End Using
                End Using
            Next
            For pass = 1 To 3
                Dim started As New TaskCompletionSource(Of Boolean)(TaskCreationOptions.RunContinuationsAsynchronously)
                Using cancelled As New CancellationTokenSource()
                    Using run As New SearchRunCoordinator(query, options, Extensions, Policy(3))
                        run.BeforeSourceAsync = Async Function(path, ordinal, token)
                                                    If ordinal = 0 Then Await started.Task.WaitAsync(token).ConfigureAwait(False)
                                                End Function
                        run.EventQueuedForChecks = Sub(ordinal, kind)
                                                       If ordinal = 1 AndAlso kind = SourceSearchEventKind.Result Then
                                                           cancelled.Cancel()
                                                           started.TrySetResult(True)
                                                       End If
                                                   End Sub
                        Try
                            run.RunAsync(root, cancelled.Token).WaitAsync(TimeSpan.FromSeconds(20)).GetAwaiter().GetResult()
                            Throw New Exception("Cancellation under output pressure was swallowed.")
                        Catch ex As OperationCanceledException
                        End Try
                    End Using
                End Using
            Next
        Finally
            Directory.Delete(root, True)
        End Try
    End Sub

    Public Sub MixedSources()
        Dim root = Path.Combine(Path.GetTempPath(), "BeaconMixedParallel-" & Guid.NewGuid().ToString("N"))
        Directory.CreateDirectory(root)
        Try
            File.WriteAllText(Path.Combine(root, "plain.log"), "before" & vbLf & "error" & vbLf & "after")
            File.WriteAllText(Path.Combine(root, "none.log"), "no match")
            File.WriteAllText(Path.Combine(root, "view.json"), "{""message"":""error""}")
            File.WriteAllText(Path.Combine(root, "view.xml"), "<root><message>error</message></root>")
            File.WriteAllText(Path.Combine(root, "view.html"), "<script>hidden</script><p>error</p>")
            File.WriteAllText(Path.Combine(root, "broken.har"), "not json")
            File.WriteAllText(Path.Combine(root, "broken.evtx"), "not an event log")
            File.WriteAllText(Path.Combine(root, "broken.zip"), "not an archive")
            Dim har = "{""log"":{""entries"":[{""request"":{""method"":""GET"",""url"":""https://example.invalid/?token=PRIVATE_SECRET""},""response"":{""status"":503,""content"":{""mimeType"":""text/plain"",""text"":""error PRIVATE_SECRET""}}}]}}"
            File.WriteAllText(Path.Combine(root, "sample.har"), har)
            For Each name In {"Rar.solid.rar", "Rar5.solid.rar"}
                File.Copy(Path.Combine(AppContext.BaseDirectory, "Fixtures", "Archives", name), Path.Combine(root, name))
            Next
            Using archive = ZipFile.Open(Path.Combine(root, "unsafe.zip"), ZipArchiveMode.Create)
                Using writer As New StreamWriter(archive.CreateEntry("../error.log").Open())
                    writer.Write("error")
                End Using
            End Using
            Dim options As New BeaconSettings With {.IncludedExtensions = ".log;.json;.xml;.html;.txt", .SearchFileNames = True, .RedactSensitiveHarData = True}
            For Each term In {"error", "PRIVATE_SECRET", "тест.txt", "absent-marker"}
                Dim query As New SearchQuery(term, SearchMode.PlainText, False)
                Dim issues As New List(Of String)()
                Dim expected = Sequential(root, query, options, issues)
                For Each workers In {1, 4, 14}
                    Dim actual As New List(Of SourceSearchResult)()
                    Dim parallelIssues As New List(Of String)()
                    Using timeout As New CancellationTokenSource(TimeSpan.FromSeconds(30))
                        Using run As New SearchRunCoordinator(query, options, Extensions, Policy(workers), AddressOf actual.Add,
                            report:=Sub(path, ex, stage, severity) parallelIssues.Add(Issue(path, ex, stage, severity)))
                            run.RunAsync(root, timeout.Token).GetAwaiter().GetResult()
                        End Using
                    End Using
                    Require(Snapshot(actual) = Snapshot(expected), $"Mixed-source result parity failed for {term}/{workers}.")
                    Require(parallelIssues.SequenceEqual(issues), $"Committed diagnostics changed order for {term}/{workers}.")
                    If term = "PRIVATE_SECRET" Then
                        Require(actual.Where(Function(result) result.Kind = SourceResultKind.DiskHar).All(
                            Function(result) Not JsonSerializer.Serialize(result).Contains("PRIVATE_SECRET")), "Parallel HAR presentation exposed a redacted value.")
                    End If
                Next
            Next
            Dim fallback As New SearchConcurrencyPolicy(processors:=Function() 32, availableMemory:=Function() 0)
            Using run As New SearchRunCoordinator(New SearchQuery("error", SearchMode.PlainText, False), options, Extensions, fallback)
                run.RunAsync(root, CancellationToken.None).WaitAsync(TimeSpan.FromSeconds(30)).GetAwaiter().GetResult()
                Require(run.PeakWorkers = 1, "Unknown/low resources did not fall back to one worker.")
            End Using
        Finally
            Directory.Delete(root, True)
        End Try
    End Sub

    Private Function Sequential(root As String, query As SearchQuery, options As BeaconSettings, Optional issues As List(Of String) = Nothing) As List(Of SourceSearchResult)
        Dim result As New List(Of SourceSearchResult)()
        Using service As New SourceSearchService(query, options, Extensions, AddressOf result.Add,
            report:=Sub(path, ex, stage, severity) issues?.Add(Issue(path, ex, stage, severity)))
            service.RunAsync(root, False, CancellationToken.None).GetAwaiter().GetResult()
        End Using
        Return result
    End Function

    Private Function Snapshot(results As IEnumerable(Of SourceSearchResult)) As String
        Return JsonSerializer.Serialize(results.Select(Function(result) New With {.Path = result.LogicalPath, .Kind = result.Kind,
            .Partial = result.PartialReason, .Details = result.Details, .Events = result.Events, .Requests = result.Requests}))
    End Function

    Private Function Issue(path As String, ex As Exception, stage As String, severity As String) As String
        Return path & " | " & stage & " | " & severity & " | " & ex.GetType().FullName
    End Function

    Private Function Policy(workers As Integer) As SearchConcurrencyPolicy
        Return New SearchConcurrencyPolicy(workers, availableMemory:=Function() 64L * 1024 * 1024 * 1024)
    End Function

    Private Sub Require(value As Boolean, message As String)
        If Not value Then Throw New InvalidOperationException(message)
    End Sub
End Module
