Imports System.IO
Imports System.Diagnostics
Imports System.Reflection
Imports System.Text
Imports System.Threading
Imports Beacon

Module StructuredSearchChecks
    Private Const HarEntry As String = "{""request"":{""method"":""GET"",""url"":""https://example.invalid/""},""response"":{""status"":503,""content"":{""mimeType"":""application/json"",""text"":""{\""token\"":\""PRIVATE_SECRET\"",\""safe\"":\""error\""}""}}}"

    Public Sub Har()
        Dim options As New BeaconSettings With {.StopAfterFirstMatchPerFile = False, .MaximumStructuredMatches = 2, .HarMaximumMatches = 2, .RedactSensitiveHarData = True}
        Dim service As New HarSearchService(New SearchQuery("PRIVATE_SECRET", SearchMode.PlainText, False), options)
        options.RedactSensitiveHarData = False
        options.HarStatusCodes = "200"
        Dim result As New StructuredSearchResult(Of HarRecord)()
        Dim issues As New List(Of String)()
        Using stream As New MemoryStream(Encoding.UTF8.GetBytes("{""log"":{""entries"":[{}," & HarEntry & "," & HarEntry & "," & HarEntry & "]}}"))
            service.CollectAsync(stream, result, CancellationToken.None, Sub(ex, stage, severity) issues.Add(stage)).GetAwaiter().GetResult()
            Require(stream.CanRead, "HAR service disposed the caller-owned stream.")
        End Using
        Require(result.Records.Count = 2 AndAlso result.Details.Count = 2, "HAR record limits or settings snapshot failed.")
        Require(result.Details.Select(Function(detail) detail.RecordIndex).SequenceEqual({0, 1}), "HAR detail indexes are not aligned.")
        Require(result.Details(0).Location.StartsWith("Request 2"), "Malformed entries changed source ordinals.")
        Require(result.PartialReason.Contains("request limit") AndAlso result.PartialReason.Contains("malformed"), "HAR partial coverage was lost.")
        Require(issues.SequenceEqual({"HAR parsing"}), "HAR diagnostics changed.")
        Require(result.Records.All(Function(record) record.RedactionEnabled AndAlso Not record.SearchText().Contains("PRIVATE_SECRET")), "HAR service did not snapshot redaction settings.")
        Require(result.Details.All(Function(detail) Not detail.Excerpt.Contains("PRIVATE_SECRET") AndAlso detail.VisibleMatches = 0), "Sensitive context was exposed.")
        Dim second As New StructuredSearchResult(Of HarRecord)()
        Using stream As New MemoryStream(Encoding.UTF8.GetBytes("{""log"":{""entries"":[" & HarEntry & ",{}]}}"))
            Reject(Of IOException)(Sub() service.CollectAsync(stream, second, CancellationToken.None,
                Sub(ex, stage, severity)
                    Throw New IOException("Simulated later reporting failure")
                End Sub).GetAwaiter().GetResult())
        End Using
        Require(second.Records.Count = 1 AndAlso second.Details.Count = 1, "Later failure discarded prior HAR results.")
        Using cancelled As New CancellationTokenSource(), stream As New MemoryStream()
            cancelled.Cancel()
            Reject(Of OperationCanceledException)(Sub() service.CollectAsync(stream, New StructuredSearchResult(Of HarRecord)(), cancelled.Token).GetAwaiter().GetResult())
        End Using
        Using stream As New MemoryStream(Encoding.UTF8.GetBytes("{}"))
            Reject(Of InvalidDataException)(Sub() service.CollectAsync(stream, New StructuredSearchResult(Of HarRecord)(), CancellationToken.None).GetAwaiter().GetResult())
        End Using
    End Sub

    Public Sub Events()
        Dim options As New BeaconSettings With {.StopAfterFirstMatchPerFile = False, .EvtxEventIds = "100", .EvtxMaximumMatches = 2, .ShowRawXmlWhenMessageUnavailable = False}
        Dim service As New EvtxSearchService(New SearchQuery("error", SearchMode.PlainText, False), options)
        options.EvtxEventIds = "999"
        options.EvtxMaximumMatches = 1
        Dim result As New StructuredSearchResult(Of EventRecordSummary)()
        Dim source As New EventRecordSummary With {.EventId = 100, .Provider = "provider", .Message = "error", .RawXml = "<Event>safe</Event>"}
        Require(Not service.CollectRecord(source, result, CancellationToken.None), "EVTX settings snapshot failed.")
        source.Message = "changed"
        Require(result.Records(0).Message = "error", "Captured EVTX record aliases its source.")
        Require(service.CollectRecord(New EventRecordSummary With {.EventId = 100, .Message = "", .RawXml = "<Event>error</Event>"}, result, CancellationToken.None), "EVTX limit was not reached.")
        Require(result.Details(1).ContextKind = "Event XML lines" AndAlso result.Records(1).MessageUnavailable, "XML matching/fallback changed.")
        Require(result.Details.Select(Function(detail) detail.RecordIndex).SequenceEqual({0, 1}) AndAlso result.PartialReason = "event limit reached", "EVTX indexes/partial state changed.")
        Dim unmatched As New StructuredSearchResult(Of EventRecordSummary)()
        service.CollectRecord(New EventRecordSummary With {.EventId = 999, .Message = "error"}, unmatched, CancellationToken.None)
        Require(unmatched.Records.Count = 0, "EVTX filter was not preserved.")
        Dim fallback As New EvtxSearchService(New SearchQuery("provider", SearchMode.PlainText, False), New BeaconSettings With {.ShowRawXmlWhenMessageUnavailable = False})
        Dim missing As New StructuredSearchResult(Of EventRecordSummary)()
        fallback.CollectRecord(New EventRecordSummary With {.Provider = "provider", .RawXml = "<Event>xml payload</Event>"}, missing, CancellationToken.None)
        Require(Not missing.Records(0).Message.Contains("xml payload") AndAlso missing.Records(0).RawXml.Contains("xml payload"), "Explicit XML/fallback preference changed.")
        Using cancelled As New CancellationTokenSource()
            cancelled.Cancel()
            Reject(Of OperationCanceledException)(Sub() service.CollectRecord(source, result, cancelled.Token))
            Reject(Of OperationCanceledException)(Sub() service.Collect("does-not-exist.evtx", result, cancelled.Token))
        End Using
        Require(result.Records.Count = 2, "Cancellation changed accumulated EVTX records.")
    End Sub

    Public Sub EvtxIsolation()
        Dim previousExe = EvtxWorkerHost.WorkerExecutableOverride
        Dim previousRender = EvtxWorkerHost.RenderingBudget
        Dim previousXml = EvtxWorkerHost.XmlFallbackBudget
        Dim previousMode = Environment.GetEnvironmentVariable("BEACON_EVTX_TEST_WORKER_MODE")
        EvtxWorkerHost.WorkerExecutableOverride = Environment.ProcessPath
        EvtxWorkerHost.RenderingBudget = Nothing
        EvtxWorkerHost.XmlFallbackBudget = TimeSpan.FromMilliseconds(750)
        Try
            WithEvtxFixture(Sub(path)
                                Require(Not EvtxWorkerHost.EffectiveRenderingBudgetForTest(New BeaconSettings With {.EvtxDeepSearch = True}).HasValue, "Unlimited Deep Search should not retain the fixed rendering deadline.")
                                Require(EvtxWorkerHost.EffectiveRenderingBudgetForTest(New BeaconSettings With {.EvtxDeepSearch = True, .EvtxLimitDeepSearchBeforeXml = True, .EvtxDeepSearchTimeoutSeconds = 42}).Value = TimeSpan.FromSeconds(42), "Configured Deep Search timeout was not used.")
                                Require(EvtxWorkerHost.EffectiveRenderingBudgetForTest(New BeaconSettings()).Value = TimeSpan.FromSeconds(180), "Normal EVTX search should keep its fixed render safety budget.")

                                Environment.SetEnvironmentVariable("BEACON_EVTX_TEST_WORKER_MODE", "success")
                                Dim query As New SearchQuery("success", SearchMode.PlainText, False)
                                Dim isolated = Collect(path, query, New BeaconSettings With {.EvtxDeepSearch = True})
                                Dim direct As New StructuredSearchResult(Of EventRecordSummary)()
                                Dim service As New EvtxSearchService(query, New BeaconSettings With {.EvtxDeepSearch = True})
                                service.CollectRecord(New EventRecordSummary With {.EventSequence = 0, .Provider = "Provider", .Level = "Information", .EventId = 100,
                                    .LevelNumber = 4, .Message = "rendered success message", .RawXml = "<Event><System><EventID>100</EventID><Provider Name=""Provider""/><Level>4</Level></System></Event>",
                                    .TimeCreated = New DateTime(2026, 1, 2, 3, 4, 5, DateTimeKind.Utc)}, direct, CancellationToken.None)
                                Require(isolated.Records.Count = 1 AndAlso isolated.PartialReason = "", "Successful worker scan was marked partial.")
                                Require(isolated.Details(0).Excerpt = direct.Details(0).Excerpt AndAlso isolated.Records(0).RawXml.Contains("EventID"), "Worker result differs from direct collection or lost XML preview.")

                                Environment.SetEnvironmentVariable("BEACON_EVTX_TEST_WORKER_MODE", "inspect-fast-path")
                                Dim beforeFast = FastModeDirectories()
                                Dim fast = Collect(path, New SearchQuery("fast-path-ok", SearchMode.PlainText, False), New BeaconSettings())
                                Require(fast.Records.Count = 1 AndAlso fast.PartialReason.Contains("EVTX normal search") AndAlso NoNewFastModeDirectories(beforeFast), "Default normal search did not use a private EVTX-only copy or report coverage.")

                                Environment.SetEnvironmentVariable("BEACON_EVTX_TEST_WORKER_MODE", "inspect-original-path")
                                Dim full = Collect(path, New SearchQuery("original-path-ok", SearchMode.PlainText, False), New BeaconSettings With {.EvtxDeepSearch = True})
                                Require(full.Records.Count = 1 AndAlso full.PartialReason = "", "Full EVTX mode stopped using the original file and sibling resources.")

                                Dim copyFailure As New StructuredSearchResult(Of EventRecordSummary)()
                                Dim copyIssues As New List(Of String)()
                                Dim copyService As New EvtxSearchService(New SearchQuery("original-path-ok", SearchMode.PlainText, False), New BeaconSettings With {.EvtxFastMode = True})
                                copyService.Collect(path, copyFailure, CancellationToken.None, Sub(ex, stage, severity) copyIssues.Add(stage), fastCopyLimitBytes:=1)
                                Require(copyFailure.Records.Count = 1 AndAlso copyFailure.PartialReason = "", "Full-resource fallback was incorrectly labeled as reduced coverage.")
                                Require(copyIssues.Contains("EVTX normal search unavailable") AndAlso Not copyIssues.Contains("EVTX normal search coverage"), "Copy failure reported that archived resources were skipped.")
                                Require(NoNewFastModeDirectories(beforeFast), "Failed fast copy leaked its temporary directory.")

                                Environment.SetEnvironmentVariable("BEACON_EVTX_TEST_WORKER_MODE", "inspect-fast-path")
                                Dim sourceHits As New List(Of SourceSearchResult)()
                                Using sourceScan As New SourceSearchService(New SearchQuery("fast-path-ok", SearchMode.PlainText, False),
                                    New BeaconSettings With {.EvtxFastMode = True}, Array.Empty(Of String)(), Sub(result) sourceHits.Add(result))
                                    sourceScan.RunAsync(path, False, CancellationToken.None).GetAwaiter().GetResult()
                                End Using
                                Require(sourceHits.Count = 1 AndAlso sourceHits(0).PhysicalPath = path AndAlso sourceHits(0).LogicalPath = path AndAlso Not sourceHits(0).LogicalPath.Contains("BeaconEvtxFast_"), "Fast mode leaked temporary paths into source results.")

                                Dim archivePath = IO.Path.Combine(IO.Path.GetDirectoryName(path), "events.zip")
                                Using archive = System.IO.Compression.ZipFile.Open(archivePath, System.IO.Compression.ZipArchiveMode.Create)
                                    Dim entry = archive.CreateEntry("sample.evtx")
                                    Using input = File.OpenRead(path), output = entry.Open()
                                        input.CopyTo(output)
                                    End Using
                                End Using
                                Dim archiveHits As New List(Of SourceSearchResult)()
                                Using archiveScan As New SourceSearchService(New SearchQuery("fast-path-ok", SearchMode.PlainText, False),
                                    New BeaconSettings With {.EvtxDeepSearch = True}, {".zip"}, Sub(result) archiveHits.Add(result))
                                    archiveScan.RunAsync(archivePath, False, CancellationToken.None).GetAwaiter().GetResult()
                                End Using
                                Require(archiveHits.Count = 1 AndAlso archiveHits(0).Kind = SourceResultKind.ArchiveEvent AndAlso archiveHits(0).PartialReason.Contains("EVTX normal search"), "Archive EVTX entries must force normal search even when Deep Search is saved.")

                                Dim cabPath = IO.Path.Combine(IO.Path.GetDirectoryName(path), "events.cab")
                                Dim makeCab As New ProcessStartInfo(IO.Path.Combine(Environment.SystemDirectory, "makecab.exe")) With {.UseShellExecute = False, .CreateNoWindow = True}
                                makeCab.ArgumentList.Add(path)
                                makeCab.ArgumentList.Add(cabPath)
                                Using child = Process.Start(makeCab)
                                    If Not child.WaitForExit(15000) Then
                                        child.Kill(True)
                                        Throw New TimeoutException("EVTX CAB fixture creation timed out.")
                                    End If
                                    Require(child.ExitCode = 0, "EVTX CAB fixture creation failed.")
                                End Using
                                Dim wrappedCab = IO.Path.Combine(IO.Path.GetDirectoryName(path), "wrapped-cab.zip")
                                Using archive = System.IO.Compression.ZipFile.Open(wrappedCab, System.IO.Compression.ZipArchiveMode.Create)
                                    Using input = File.OpenRead(cabPath), output = archive.CreateEntry("events.cab").Open()
                                        input.CopyTo(output)
                                    End Using
                                End Using
                                For Each source In {cabPath, wrappedCab}
                                    Dim cabHits As New List(Of SourceSearchResult)()
                                    Dim folderPreference As New BeaconSettings With {.EvtxDeepSearch = True}
                                    Using cabScan As New SourceSearchService(New SearchQuery("fast-path-ok", SearchMode.PlainText, False),
                                        folderPreference, {".cab", ".zip"}, Sub(result) cabHits.Add(result))
                                        cabScan.RunAsync(source, False, CancellationToken.None).GetAwaiter().GetResult()
                                    End Using
                                    Require(cabHits.Count = 1 AndAlso cabHits(0).PartialReason.Contains("EVTX normal search"), "CAB EVTX entries inherited folder Deep Search.")
                                    Require(folderPreference.EvtxDeepSearch, "Archive scanning changed the saved folder preference.")
                                Next

                                Dim entryDll = Assembly.GetEntryAssembly().Location
                                If File.Exists(entryDll) AndAlso entryDll.EndsWith(".dll", StringComparison.OrdinalIgnoreCase) Then
                                    EvtxWorkerHost.WorkerExecutableOverride = entryDll
                                    Environment.SetEnvironmentVariable("BEACON_EVTX_TEST_WORKER_MODE", "success")
                                    Dim dotnetHosted = Collect(path, query, New BeaconSettings With {.EvtxDeepSearch = True})
                                    Require(dotnetHosted.Records.Count = 1 AndAlso dotnetHosted.PartialReason = "", "Worker DLL launch via dotnet did not complete cleanly.")
                                    EvtxWorkerHost.WorkerExecutableOverride = Environment.ProcessPath
                                End If

                                Environment.SetEnvironmentVariable("BEACON_EVTX_TEST_WORKER_MODE", "slow-success")
                                Dim unlimited = Collect(path, New SearchQuery("slow-success", SearchMode.PlainText, False), New BeaconSettings With {.EvtxDeepSearch = True})
                                Require(unlimited.Records.Count = 1 AndAlso unlimited.PartialReason = "", "Unlimited Deep Search should not fall back on the old short render budget.")

                                Environment.SetEnvironmentVariable("BEACON_EVTX_TEST_WORKER_MODE", "timeout-zero")
                                Dim configuredTimeout = Collect(path, New SearchQuery("missing", SearchMode.PlainText, False), New BeaconSettings With {.EvtxDeepSearch = True, .EvtxLimitDeepSearchBeforeXml = True, .EvtxDeepSearchTimeoutSeconds = 1})
                                Require(configuredTimeout.PartialReason.Contains("message rendering timed out after 1 second"), "Configured Deep Search timeout label was not reported.")

                                EvtxWorkerHost.RenderingBudget = TimeSpan.FromMilliseconds(750)
                                Environment.SetEnvironmentVariable("BEACON_EVTX_TEST_WORKER_MODE", "timeout-fallback")
                                Dim mixed = Collect(path, New SearchQuery("message-only-before-timeout xml-after-timeout", SearchMode.AnyTerm, False), New BeaconSettings With {.EvtxDeepSearch = True})
                                Require(mixed.Records.Count = 2, "Timeout fallback did not retain rendered and XML-only hits.")
                                Require(mixed.Records(0).Message.Contains("message-only-before-timeout") AndAlso mixed.Records(1).MessageUnavailable AndAlso mixed.Records(1).RawXml.Contains("xml-after-timeout"), "Timeout fallback coverage was mislabeled or discarded.")
                                Require(mixed.PartialReason.Contains("message rendering timed out") AndAlso mixed.Details.Select(Function(item) item.RecordIndex).SequenceEqual({0, 1}), "Timeout partial reason or record indexes are wrong.")

                                Environment.SetEnvironmentVariable("BEACON_EVTX_TEST_WORKER_MODE", "timeout-fallback-limit")
                                Dim limited = Collect(path, New SearchQuery("timeout-limit", SearchMode.PlainText, False), New BeaconSettings With {.EvtxDeepSearch = True, .MaximumStructuredMatches = 2, .EvtxMaximumMatches = 2})
                                Require(limited.Records.Count = 2 AndAlso limited.PartialReason.Contains("message rendering timed out") AndAlso limited.PartialReason.Contains("event limit reached"), "Fallback event limit overwrote timeout coverage.")

                                Environment.SetEnvironmentVariable("BEACON_EVTX_TEST_WORKER_MODE", "duplicate-fallback")
                                Dim deduped = Collect(path, New SearchQuery("dup", SearchMode.PlainText, False), New BeaconSettings With {.EvtxDeepSearch = True, .MaximumStructuredMatches = 3, .EvtxMaximumMatches = 3})
                                Require(deduped.Records.Select(Function(item) item.EventSequence).SequenceEqual({0L, 1L}), "Fallback duplicated a rendered event or lost sequence order.")
                                Require(deduped.Details.Select(Function(item) item.RecordIndex).SequenceEqual({0, 1}), "Fallback duplicate handling broke detail indexes.")

                                Environment.SetEnvironmentVariable("BEACON_EVTX_TEST_WORKER_MODE", "replay-quota")
                                Dim replayed = Collect(path, New SearchQuery("replay-quota", SearchMode.PlainText, False), New BeaconSettings With {.EvtxDeepSearch = True, .MaximumStructuredMatches = 3, .EvtxMaximumMatches = 3})
                                Require(replayed.Records.Select(Function(item) item.EventSequence).SequenceEqual({0L, 1L, 2L}), "Replayed hits consumed the fallback quota and hid later matches.")
                                Require(replayed.PartialReason.Contains("message rendering timed out") AndAlso replayed.PartialReason.Contains("event limit reached"), "Replay limit lost incomplete coverage.")

                                Environment.SetEnvironmentVariable("BEACON_EVTX_TEST_WORKER_MODE", "regex-timeout")
                                Reject(Of System.Text.RegularExpressions.RegexMatchTimeoutException)(Sub() Collect(path, New SearchQuery("a+", SearchMode.RegularExpression, False), New BeaconSettings With {.EvtxDeepSearch = True}))

                                Environment.SetEnvironmentVariable("BEACON_EVTX_TEST_WORKER_MODE", "timeout-zero")
                                Dim zero = Collect(path, New SearchQuery("missing", SearchMode.PlainText, False), New BeaconSettings With {.EvtxDeepSearch = True})
                                Require(zero.Records.Count = 0 AndAlso zero.PartialReason.Contains("message rendering timed out"), "Zero-hit timeout was reported as a complete scan.")

                                For Each mode In {"crash", "malformed", "invalid-json", "incomplete", "backward", "after-complete", "oversized-unterminated"}
                                    Environment.SetEnvironmentVariable("BEACON_EVTX_TEST_WORKER_MODE", mode)
                                    Dim issues As New List(Of String)()
                                    Dim failed = Collect(path, New SearchQuery("anything", SearchMode.PlainText, False), New BeaconSettings With {.EvtxDeepSearch = True}, issues)
                                    Require(failed.PartialReason.Contains("worker failed") AndAlso issues.Count > 0, "Worker " & mode & " was not diagnosed as incomplete.")
                                Next

                                Environment.SetEnvironmentVariable("BEACON_EVTX_TEST_WORKER_MODE", "limit-stop")
                                Dim limitedStop = Collect(path, New SearchQuery("limit-stop", SearchMode.PlainText, False), New BeaconSettings With {.EvtxDeepSearch = True, .MaximumStructuredMatches = 1, .EvtxMaximumMatches = 1})
                                Require(limitedStop.Records.Count = 1 AndAlso limitedStop.PartialReason.Contains("event limit reached"), "Worker limit stop did not complete after the requested match count.")
                                Environment.SetEnvironmentVariable("BEACON_EVTX_TEST_WORKER_MODE", "first-stop")
                                Dim firstStop = Collect(path, New SearchQuery("first-stop", SearchMode.PlainText, False), New BeaconSettings With {.EvtxDeepSearch = True, .StopAfterFirstMatchPerFile = True})
                                Require(firstStop.Records.Count = 1 AndAlso firstStop.PartialReason.Contains("first matching event only"), "Worker first-match stop did not complete immediately.")

                                Environment.SetEnvironmentVariable("BEACON_EVTX_TEST_WORKER_MODE", "hang")
                                Using cancel As New CancellationTokenSource(200)
                                    Dim timer = Stopwatch.StartNew()
                                    Reject(Of OperationCanceledException)(Sub() Collect(path, New SearchQuery("anything", SearchMode.PlainText, False), New BeaconSettings With {.EvtxDeepSearch = True}, Nothing, cancel.Token))
                                    Require(timer.Elapsed < TimeSpan.FromSeconds(5), "Cancellation did not contain a hung render worker.")
                                End Using

                                Environment.SetEnvironmentVariable("BEACON_EVTX_TEST_WORKER_MODE", "hang")
                                beforeFast = FastModeDirectories()
                                Using cancel As New CancellationTokenSource(200)
                                    Reject(Of OperationCanceledException)(Sub() Collect(path, New SearchQuery("anything", SearchMode.PlainText, False), New BeaconSettings With {.EvtxFastMode = True}, Nothing, cancel.Token))
                                End Using
                                Require(NoNewFastModeDirectories(beforeFast), "Cancellation leaked an EVTX fast-mode copy.")

                                Environment.SetEnvironmentVariable("BEACON_EVTX_TEST_WORKER_MODE", "fallback-hang")
                                Using cancel As New CancellationTokenSource(350)
                                    Dim timer = Stopwatch.StartNew()
                                    Reject(Of OperationCanceledException)(Sub() Collect(path, New SearchQuery("anything", SearchMode.PlainText, False), New BeaconSettings With {.EvtxDeepSearch = True}, Nothing, cancel.Token))
                                    Require(timer.Elapsed < TimeSpan.FromSeconds(5), "Cancellation did not contain a hung XML fallback worker.")
                                End Using
                            End Sub)
        Finally
            EvtxWorkerHost.WorkerExecutableOverride = previousExe
            EvtxWorkerHost.RenderingBudget = previousRender
            EvtxWorkerHost.XmlFallbackBudget = previousXml
            Environment.SetEnvironmentVariable("BEACON_EVTX_TEST_WORKER_MODE", previousMode)
        End Try
    End Sub

    Private Function Collect(path As String, query As SearchQuery, options As BeaconSettings,
                             Optional issues As List(Of String) = Nothing,
                             Optional token As CancellationToken = Nothing) As StructuredSearchResult(Of EventRecordSummary)
        Dim result As New StructuredSearchResult(Of EventRecordSummary)()
        Dim service As New EvtxSearchService(query, options)
        service.Collect(path, result, token, Sub(ex, stage, severity)
                                                If issues IsNot Nothing Then issues.Add(stage & ":" & ex.GetType().Name)
                                            End Sub)
        Return result
    End Function

    Private Sub WithEvtxFixture(action As Action(Of String))
        Dim root = IO.Path.Combine(IO.Path.GetTempPath(), "BeaconEvtxIsolation-" & Guid.NewGuid().ToString("N"))
        Directory.CreateDirectory(root)
        Dim path = IO.Path.Combine(root, "sample.evtx")
        File.WriteAllText(path, "placeholder")
        Directory.CreateDirectory(IO.Path.Combine(root, "LocaleMetaData"))
        Try
            action(path)
        Finally
            Directory.Delete(root, True)
        End Try
    End Sub

    Private Sub Reject(Of T As Exception)(action As Action)
        Try
            action()
        Catch ex As T
            Return
        End Try
        Throw New InvalidOperationException("Expected " & GetType(T).Name)
    End Sub

    Private Function FastModeDirectories() As HashSet(Of String)
        Return Directory.GetDirectories(IO.Path.GetTempPath(), "BeaconEvtxFast_*").ToHashSet(StringComparer.OrdinalIgnoreCase)
    End Function

    Private Function NoNewFastModeDirectories(before As HashSet(Of String)) As Boolean
        Return Directory.GetDirectories(IO.Path.GetTempPath(), "BeaconEvtxFast_*").All(Function(path) before.Contains(path))
    End Function
    Private Sub Require(condition As Boolean, message As String)
        If Not condition Then Throw New InvalidOperationException(message)
    End Sub
End Module
