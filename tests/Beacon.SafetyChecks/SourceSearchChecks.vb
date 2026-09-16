Imports System.IO
Imports System.IO.Compression
Imports System.Text
Imports System.Threading
Imports Beacon

Module SourceSearchChecks
    Private ReadOnly Extensions As String() = {".zip", ".7z", ".rar", ".tar", ".gz", ".bz2", ".xz", ".tgz", ".tbz", ".tbz2", ".txz", ".cab"}

    Public Sub Run()
        Dim root = Path.Combine(Path.GetTempPath(), "BeaconSourceTests-" & Guid.NewGuid().ToString("N"))
        Directory.CreateDirectory(root)
        Try
            File.WriteAllText(Path.Combine(root, "one.log"), "error" & vbLf & "error again")
            Directory.CreateDirectory(Path.Combine(root, ".git"))
            File.WriteAllText(Path.Combine(root, ".git", "ignored.log"), "error")
            Dim outer = Path.Combine(root, "outer.zip")
            Using inner As New MemoryStream()
                Using zip As New ZipArchive(inner, ZipArchiveMode.Create, True)
                    Using writer As New StreamWriter(zip.CreateEntry("nested.log").Open())
                        writer.Write("nested error café")
                    End Using
                End Using
                Using zip = ZipFile.Open(outer, ZipArchiveMode.Create)
                    Using output = zip.CreateEntry("inner.zip").Open()
                        output.Write(inner.ToArray())
                    End Using
                End Using
            End Using
            Dim before = Directory.GetDirectories(Path.GetTempPath(), "BeaconSearch_*").ToHashSet(StringComparer.OrdinalIgnoreCase)
            Dim query As New SearchQuery("error", SearchMode.PlainText, False)
            Dim options As New BeaconSettings With {.StopAfterFirstMatchPerFile = False}
            Dim published As New List(Of SourceSearchResult)()
            Dim progress As Integer = 0
            Using counter As New SourceSearchService(query, options, Extensions, Sub(result) published.Add(result),
                Sub(logical, processed)
                    If processed Then progress += 1
                End Sub)
                counter.RunAsync(root, True, CancellationToken.None).GetAwaiter().GetResult()
                Require(counter.FilesProcessed = 2 AndAlso progress = 2 AndAlso counter.MatchingFiles = 0 AndAlso published.Count = 0, "Count pass did not share filters or published search results.")
            End Using
            Require(NoNewTemporaryDirectories(before), "Count pass leaked temporary nested archives.")
            Dim captured As SourceSearchResult = Nothing
            Dim search As New SourceSearchService(query, options, Extensions, Sub(result) published.Add(result))
            Try
                options.ArchiveNestingDepth = 0
                options.SearchFileContents = False
                options.IncludedExtensions = ""
                search.RunAsync(root, False, CancellationToken.None).GetAwaiter().GetResult()
                Require(search.FilesProcessed = 2 AndAlso search.MatchingFiles = 2 AndAlso published.Count = 2, "Mutable caller settings changed the source search.")
                captured = published.Single(Function(result) result.Kind = SourceResultKind.ArchiveText)
                Require(captured.LogicalPath = outer & " | inner.zip | nested.log", "Service exposed temporary names in logical paths.")
                Require(File.Exists(captured.ArchivePath), "Nested archive disappeared before preview.")
                Dim preview As New PreviewContentService(New BeaconSettings())
                Require(preview.ReadArchive(captured.ArchivePath, captured.EntryName).Text = "nested error café", "Service-owned nested source cannot be previewed.")
                Require(published.Single(Function(result) result.Kind = SourceResultKind.DiskText).Details.Count = 2, "Multiple text records were lost.")
                Reject(Of InvalidOperationException)(Sub() search.RunAsync(root, False, CancellationToken.None).GetAwaiter().GetResult())
            Finally
                search.Dispose()
                search.Dispose()
            End Try
            Require(Not File.Exists(captured.ArchivePath) AndAlso NoNewTemporaryDirectories(before), "Search disposal leaked nested sources.")
            Reject(Of ObjectDisposedException)(Sub() search.RunAsync(root, False, CancellationToken.None).GetAwaiter().GetResult())
            published.Clear()
            Using limited As New SourceSearchService(query, New BeaconSettings With {.MaximumTotalResults = 1}, Extensions, Sub(result) published.Add(result))
                limited.RunAsync(root, False, CancellationToken.None).GetAwaiter().GetResult()
                Require(limited.ResultLimitReached AndAlso limited.MatchingFiles = 1 AndAlso published.Count = 1, "Headless result cap failed.")
            End Using
            Using cancelled As New CancellationTokenSource()
                Using scan As New SourceSearchService(query, New BeaconSettings(), Extensions, progress:=Sub(logical, processed)
                                                                                                           If processed Then cancelled.Cancel()
                                                                                                       End Sub)
                    Reject(Of OperationCanceledException)(Sub() scan.RunAsync(outer, False, cancelled.Token).GetAwaiter().GetResult())
                End Using
            End Using
            Require(NoNewTemporaryDirectories(before), "Cancellation leaked nested extraction files.")
            Dim invalid = Path.Combine(root, "name-only.evtx")
            File.WriteAllText(invalid, "Not a valid EVTX file")
            published.Clear()
            Using scan As New SourceSearchService(New SearchQuery("name-only", SearchMode.PlainText, False),
                New BeaconSettings With {.SearchFileContents = False, .SearchFileNames = True}, Extensions, Sub(result) published.Add(result))
                scan.RunAsync(invalid, False, CancellationToken.None).GetAwaiter().GetResult()
                Require(published.Count = 1 AndAlso published(0).Details.All(Function(detail) detail.IsMetadata) AndAlso published(0).Events.Count = 0, "Metadata search parsed invalid EVTX content.")
            End Using
        Finally
            Directory.Delete(root, True)
        End Try
    End Sub

    Private Function NoNewTemporaryDirectories(before As HashSet(Of String)) As Boolean
        Return Directory.GetDirectories(Path.GetTempPath(), "BeaconSearch_*").All(Function(path) before.Contains(path))
    End Function
    Private Sub Reject(Of T As Exception)(action As Action)
        Try
            action()
        Catch ex As T
            Return
        End Try
        Throw New InvalidOperationException("Expected " & GetType(T).Name)
    End Sub
    Private Sub Require(condition As Boolean, message As String)
        If Not condition Then Throw New InvalidOperationException(message)
    End Sub
End Module
