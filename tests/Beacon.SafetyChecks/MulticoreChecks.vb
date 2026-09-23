Imports System.IO
Imports System.IO.Compression
Imports System.Text.Json
Imports System.Threading
Imports System.Threading.Tasks
Imports Beacon

Module MulticoreChecks
    Private ReadOnly Extensions As String() = {".zip", ".cab", ".7z", ".tar", ".gz"}

    Public Sub Policy()
        For Each count In {1, 2, 4, 8, 14, 20, 32, 128}
            Dim policy As New SearchConcurrencyPolicy(processors:=Function() count, availableMemory:=Function() 64L * 1024 * 1024 * 1024)
            Require(policy.MaximumWorkers = count AndAlso policy.AdmissionLimit() = count, "Processor policy imposed an artificial ceiling.")
        Next
        Require(New SearchConcurrencyPolicy(processors:=Function() 0, availableMemory:=Function() 0).AdmissionLimit() = 1, "Unknown capacity must make progress with one worker.")
        Dim low As New SearchConcurrencyPolicy(processors:=Function() 32, availableMemory:=Function() SearchConcurrencyPolicy.ReserveBytes)
        Require(low.AdmissionLimit() = 1 AndAlso Not low.CanStartHeavy(1024), "Memory pressure did not reduce admission.")
        For Each saved In {0, 1, 64}
            Dim options As New BeaconSettings With {.ScanWorkerCount = saved}
            Using run As New SearchRunCoordinator(New SearchQuery("error", SearchMode.PlainText, False), options, Extensions)
                Require(run.WorkerCeiling = Math.Max(1, Environment.ProcessorCount), "A legacy preference capped automatic scheduling.")
                Require(BeaconSettingsService.Clone(options).ScanWorkerCount = saved, "Scheduling rewrote saved settings.")
            End Using
        Next
    End Sub

    Public Sub Resources()
        Dim head As Long = 0
        Using gate As New SearchResourceGate(TestPolicy(14), Function() Volatile.Read(head))
            Dim background = gate.EnterAsync(1, 1024, CancellationToken.None).GetAwaiter().GetResult()
            Using cancellation As New CancellationTokenSource()
                Dim waiting = gate.EnterAsync(2, 1024, cancellation.Token)
                Require(Not waiting.IsCompleted, "Speculative heavy work was not bounded.")
                Using reserved = gate.EnterAsync(0, Long.MaxValue, CancellationToken.None).WaitAsync(TimeSpan.FromSeconds(2)).GetAwaiter().GetResult()
                    Require(reserved Is Nothing, "The ordered head lost its reserved lane.")
                End Using
                cancellation.Cancel()
                Try
                    waiting.GetAwaiter().GetResult()
                    Throw New Exception("Heavy admission ignored cancellation.")
                Catch ex As OperationCanceledException
                End Try
            End Using
            background.Dispose()
            background.Dispose()
            Using nextLease = gate.EnterAsync(3, 1024, CancellationToken.None).WaitAsync(TimeSpan.FromSeconds(2)).GetAwaiter().GetResult()
                Require(nextLease IsNot Nothing, "A cancelled waiter leaked its permit.")
            End Using
        End Using
        Dim acquired As Integer = 0
        Dim execution As New SourceSearchExecution(Function(item, token) Task.CompletedTask,
            Function(bytes, token)
                acquired += 1
                Return Task.FromResult(Of IDisposable)(New SearchResourceLease(Sub() acquired -= 1))
            End Function)
        Using outer = execution.EnterHeavyAsync(1024, CancellationToken.None).GetAwaiter().GetResult()
            Using inner = execution.EnterHeavyAsync(1024, CancellationToken.None).GetAwaiter().GetResult()
                Require(acquired = 1, "Nested archive/document admission was not reentrant.")
            End Using
            Require(acquired = 1, "A nested lease released its parent's permit.")
        End Using
        Require(acquired = 0, "The heavy-work permit was not released.")
        execution.Temporary.Dispose()
    End Sub

    Public Sub Parity()
        Dim root = Path.Combine(Path.GetTempPath(), "BeaconParallel-" & Guid.NewGuid().ToString("N"))
        Directory.CreateDirectory(root)
        Try
            For index = 0 To 15
                File.WriteAllText(Path.Combine(root, $"{index:D2}.log"), $"before{vbCrLf}code={index} then code={index + 100}{vbCrLf}after")
            Next
            Using zip = ZipFile.Open(Path.Combine(root, "outer.zip"), ZipArchiveMode.Create)
                Using writer As New StreamWriter(zip.CreateEntry("inside.log").Open())
                    writer.Write("before" & vbLf & "code=42" & vbLf & "after")
                End Using
            End Using
            Dim query As New SearchQuery("code=\d+", SearchMode.RegularExpression, False)
            Dim options As New BeaconSettings With {.StopAfterFirstMatchPerFile = False}
            Dim baseline As New List(Of SourceSearchResult)()
            Using service As New SourceSearchService(query, options, Extensions, AddressOf baseline.Add)
                service.RunAsync(root, False, CancellationToken.None).GetAwaiter().GetResult()
            End Using
            Dim expected = Snapshot(baseline)
            For Each count In {1, 2, 4, 14}
                Dim actual As New List(Of SourceSearchResult)()
                Dim progress As Integer = 0
                Dim entered As Integer = 0
                Dim ready As New TaskCompletionSource(Of Boolean)(TaskCreationOptions.RunContinuationsAsynchronously)
                Using timeout As New CancellationTokenSource(TimeSpan.FromSeconds(30))
                    Using run As New SearchRunCoordinator(query, options, Extensions, TestPolicy(count), AddressOf actual.Add,
                        Sub(path, completed)
                            If completed Then progress += 1
                        End Sub)
                        run.BeforeSourceAsync = Async Function(path, ordinal, token)
                                                    If ordinal < count Then
                                                        If Interlocked.Increment(entered) = count Then ready.TrySetResult(True)
                                                        Await ready.Task.WaitAsync(token).ConfigureAwait(False)
                                                    End If
                                                End Function
                        run.RunAsync(root, timeout.Token).GetAwaiter().GetResult()
                        Require(Snapshot(actual) = expected, $"Ordered result/context parity failed with {count} workers.")
                        Require(run.FilesProcessed = 17 AndAlso progress = 17 AndAlso run.MatchingFiles = 17 AndAlso Not run.ResultLimitReached,
                                $"Counts for {count} workers: completed={run.FilesProcessed}, progress={progress}, matches={run.MatchingFiles}, limit={run.ResultLimitReached}.")
                        Require(run.PeakWorkers = count, "Forced worker capacity was not exercised.")
                    End Using
                End Using
            Next
        Finally
            Directory.Delete(root, True)
        End Try
    End Sub

    Public Sub Termination()
        Dim root = Path.Combine(Path.GetTempPath(), "BeaconTermination-" & Guid.NewGuid().ToString("N"))
        Directory.CreateDirectory(root)
        Try
            For index = 0 To 3
                File.WriteAllText(Path.Combine(root, $"{index:D2}.log"), "error")
            Next
            Dim query As New SearchQuery("error", SearchMode.PlainText, False)
            Dim options As New BeaconSettings With {.MaximumTotalResults = 1}
            Dim expected As New List(Of SourceSearchResult)()
            Using serial As New SourceSearchService(query, options, Extensions, AddressOf expected.Add)
                serial.RunAsync(root, False, CancellationToken.None).GetAwaiter().GetResult()
            End Using
            Dim laterStarted As New TaskCompletionSource(Of Boolean)(TaskCreationOptions.RunContinuationsAsynchronously)
            Dim actual As New List(Of SourceSearchResult)()
            Using timeout As New CancellationTokenSource(TimeSpan.FromSeconds(20))
                Using run As New SearchRunCoordinator(query, options, Extensions, TestPolicy(4), AddressOf actual.Add)
                    run.BeforeSourceAsync = Async Function(path, ordinal, token)
                                                If ordinal = 0 Then
                                                    Await laterStarted.Task.WaitAsync(token).ConfigureAwait(False)
                                                Else
                                                    laterStarted.TrySetResult(True)
                                                    Throw New InvalidDataException("Speculative failure beyond the result cap.")
                                                End If
                                            End Function
                    run.RunAsync(root, timeout.Token).GetAwaiter().GetResult()
                    Require(Snapshot(actual) = Snapshot(expected) AndAlso run.ResultLimitReached AndAlso run.FilesProcessed = 1,
                            "A later fault changed the ordered cap or completed-file count.")
                End Using
            End Using
            Using timeout As New CancellationTokenSource(TimeSpan.FromSeconds(20))
                Using run As New SearchRunCoordinator(query, New BeaconSettings(), Extensions, TestPolicy(4),
                    publishAsync:=Function(result, token) Task.FromException(New InvalidOperationException("Publisher failed")))
                    Try
                        run.RunAsync(root, timeout.Token).GetAwaiter().GetResult()
                        Throw New Exception("Callback failure was swallowed.")
                    Catch ex As InvalidOperationException When ex.Message = "Publisher failed"
                    End Try
                End Using
            End Using
            Using cancelled As New CancellationTokenSource()
                Using run As New SearchRunCoordinator(query, New BeaconSettings(), Extensions, TestPolicy(4))
                    run.BeforeSourceAsync = Async Function(path, ordinal, token)
                                                cancelled.Cancel()
                                                Await Task.Delay(Timeout.Infinite, token).ConfigureAwait(False)
                                            End Function
                    Try
                        run.RunAsync(root, cancelled.Token).WaitAsync(TimeSpan.FromSeconds(20)).GetAwaiter().GetResult()
                        Throw New Exception("Cancelled search completed successfully.")
                    Catch ex As OperationCanceledException
                    End Try
                End Using
            End Using
        Finally
            Directory.Delete(root, True)
        End Try
    End Sub

    Public Sub Lifetime()
        Dim root = Path.Combine(Path.GetTempPath(), "BeaconParallelLifetime-" & Guid.NewGuid().ToString("N"))
        Directory.CreateDirectory(root)
        Try
            Dim outer = Path.Combine(root, "outer.zip")
            Using data As New MemoryStream()
                Using inner As New ZipArchive(data, ZipArchiveMode.Create, True)
                    Using writer As New StreamWriter(inner.CreateEntry("nested.log").Open())
                        writer.Write("nested error café")
                    End Using
                End Using
                Using archive = ZipFile.Open(outer, ZipArchiveMode.Create)
                    Using output = archive.CreateEntry("inner.zip").Open()
                        output.Write(data.ToArray())
                    End Using
                End Using
            End Using
            For Each cancelAfterPublish In {False, True}
                Dim captured As SourceSearchResult = Nothing
                Using cancellation As New CancellationTokenSource(TimeSpan.FromSeconds(20))
                    Using run As New SearchRunCoordinator(New SearchQuery("error", SearchMode.PlainText, False), New BeaconSettings(), Extensions,
                        TestPolicy(4), Sub(result)
                                           captured = result
                                           If cancelAfterPublish Then cancellation.Cancel()
                                       End Sub)
                        Try
                            run.RunAsync(outer, cancellation.Token).GetAwaiter().GetResult()
                        Catch ex As OperationCanceledException When cancelAfterPublish
                        End Try
                        Require(captured IsNot Nothing AndAlso File.Exists(captured.ArchivePath), "Accepted nested sources were disposed before preview.")
                        Require(New PreviewContentService(New BeaconSettings()).ReadArchive(captured.ArchivePath, captured.EntryName).Text = "nested error café",
                                "Retained source could not be previewed after completion/cancellation.")
                    End Using
                End Using
                Require(Not File.Exists(captured.ArchivePath), "Releasing the run leaked nested temporary sources.")
            Next
        Finally
            Directory.Delete(root, True)
        End Try
    End Sub

    Private Function TestPolicy(count As Integer) As SearchConcurrencyPolicy
        Return New SearchConcurrencyPolicy(count, availableMemory:=Function() 64L * 1024 * 1024 * 1024)
    End Function

    Private Function Snapshot(results As IEnumerable(Of SourceSearchResult)) As String
        Return JsonSerializer.Serialize(results.Select(Function(result) New With {
            .Path = result.LogicalPath, .Kind = result.Kind, .Partial = result.PartialReason, .Details = result.Details}))
    End Function

    Private Sub Require(condition As Boolean, message As String)
        If Not condition Then Throw New InvalidOperationException(message)
    End Sub
End Module
