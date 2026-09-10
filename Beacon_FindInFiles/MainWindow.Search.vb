Imports System.IO
Imports System.Threading
Imports System.Threading.Tasks
Imports System.Diagnostics.Eventing.Reader
Imports System.Text.Json
Imports SharpCompress.Archives

Namespace Beacon
    Partial Public Class MainWindow
        Private Function NewSourceHit(physicalPath As String, logicalPath As String, kind As HitKind,
                                      Optional archivePath As String = Nothing, Optional entryName As String = Nothing) As SearchHit
            RememberSourcePath(physicalPath, logicalPath)
            Dim hit As New SearchHit With {.FilePath = physicalPath, .LogicalPath = logicalPath, .Kind = kind,
                .ZipPath = archivePath, .ZipEntryName = entryName, .DisplayName = DisplaySourcePath(logicalPath)}
            Dim leaf = If(entryName, physicalPath)
            If _searchOptions.SearchFileNames Then
                Dim detail = SearchDetail.FromRecord(_activeQuery, Path.GetFileName(leaf), "File name", metadata:=True)
                If detail IsNot Nothing Then hit.Details.Add(detail)
            End If
            If _searchOptions.SearchFullPaths Then
                Dim detail = SearchDetail.FromRecord(_activeQuery, logicalPath, "Full path", metadata:=True)
                If detail IsNot Nothing Then hit.Details.Add(detail)
            End If
            Return hit
        End Function

        Private Function DisplaySourcePath(logicalPath As String) As String
            If _searchOptions.DisplayAbsolutePaths Then Return logicalPath
            Dim separator = logicalPath.IndexOf(" | ", StringComparison.Ordinal)
            Dim root = If(separator >= 0, logicalPath.Substring(0, separator), logicalPath)
            Return MakeRelativeDisplay(_scanRootFolder, root) & If(separator >= 0, logicalPath.Substring(separator), "")
        End Function

        Private Sub PublishSourceHit(hit As SearchHit)
            If hit.Details.Count = 0 Then Return
            If hit.PartialReason.Contains("limit", StringComparison.OrdinalIgnoreCase) Then Interlocked.Increment(_structuredLimitFiles)
            AddHit(hit)
        End Sub

        Private Async Function SearchSourceAsync(root As String, ct As CancellationToken) As Task
            If Directory.Exists(root) Then
                For Each filePath As String In FileSystemTraversal.EnumerateFiles(root, ct, _searchOptions.FollowReparsePoints,
                                                                    AddressOf HandleAccessDeniedPath, AddressOf RecordFileSystemIssue, _searchOptions)
                    ct.ThrowIfCancellationRequested()
                    If Volatile.Read(_resultLimitReached) Then Exit For
                    Await SearchDiskSourceAsync(filePath, Path.GetFullPath(filePath), ct)
                Next
            Else
                Await SearchDiskSourceAsync(root, Path.GetFullPath(root), ct)
            End If
        End Function

        Private Async Function SearchDiskSourceAsync(physicalPath As String, logicalPath As String, ct As CancellationToken,
                                                       Optional depth As Integer = 0) As Task
            ct.ThrowIfCancellationRequested()
            Dim extension = Path.GetExtension(physicalPath).ToLowerInvariant()
            Dim hit = NewSourceHit(physicalPath, logicalPath, HitKind.DiskTextFile)
            ThrottledSetCurrentFileDisplay(hit.DisplayName)
            Try
                If _supportedArchiveExt.Contains(extension) Then
                    PublishSourceHit(hit)
                    If Not Volatile.Read(_resultLimitReached) Then Await SearchArchiveSourceAsync(physicalPath, logicalPath, depth, ct)
                    Return
                End If
                _filesScanned += 1
                UpdateScanProgress()
                If _searchOptions.SearchFileContents Then
                    If _supportedTextExt.Contains(extension) Then
                        Using stream = OpenBoundedDiskFile(physicalPath, ct)
                            Await CollectTextDetailsAsync(stream, extension, hit, ct)
                        End Using
                    ElseIf _supportedEvtxExt.Contains(extension) Then
                        hit.Kind = HitKind.EvtxFileOnDisk
                        CollectEventDetails(physicalPath, hit, ct)
                    ElseIf _supportedHarExt.Contains(extension) Then
                        hit.Kind = HitKind.HarFileOnDisk
                        Using stream = OpenBoundedDiskFile(physicalPath, ct)
                            Await CollectHarDetailsAsync(stream, hit, ct)
                        End Using
                    End If
                End If
            Catch ex As Exception When TypeOf ex Is IOException OrElse TypeOf ex Is UnauthorizedAccessException OrElse
                                            TypeOf ex Is InvalidDataException OrElse TypeOf ex Is EventLogException OrElse TypeOf ex Is JsonException
                RecordFileSystemIssue(logicalPath, ex)
                hit.PartialReason = "content unavailable"
            End Try
            PublishSourceHit(hit)
        End Function

        Private Async Function CollectTextDetailsAsync(stream As Stream, extension As String, hit As SearchHit, ct As CancellationToken) As Task
            Dim document = extension = ".html" OrElse extension = ".xml" OrElse extension = ".json"
            Dim collected = Await SearchTextCollector.CollectAsync(stream, _activeQuery, _searchOptions.MaximumStructuredMatches,
                                                                   _searchOptions.StopAfterFirstMatchPerFile, document, ct)
            hit.Details.AddRange(collected.Details)
            hit.PartialReason = collected.PartialReason
        End Function

        Private Shared Function GetArchiveAttributes(entry As IArchiveEntry) As FileAttributes?
            Try
                Dim attributes = entry.Attrib
                If attributes.HasValue Then Return CType(attributes.Value, FileAttributes)
                Return Nothing
            Catch ex As NotImplementedException
                ' SharpCompress 0.38 does not expose attributes for every archive format.
                Return Nothing
            End Try
        End Function

        Private Async Function SearchArchiveSourceAsync(physicalPath As String, logicalPath As String, depth As Integer, ct As CancellationToken) As Task
            RememberSourcePath(physicalPath, logicalPath)
            If depth > _searchOptions.ArchiveNestingDepth Then
                RecordFileSystemIssue(logicalPath, New InvalidDataException("Archive nesting limit reached."))
                Return
            End If
            If Path.GetExtension(physicalPath).Equals(".cab", StringComparison.OrdinalIgnoreCase) Then
                Await SearchCabSourceAsync(physicalPath, logicalPath, depth, ct)
                Return
            End If
            Try
                Using archive = OpenArchive(physicalPath)
                    Dim budget As New ArchiveReadBudget(_searchOptions, New FileInfo(physicalPath).Length)
                    For Each entry In archive.Entries
                        ct.ThrowIfCancellationRequested()
                        If Volatile.Read(_resultLimitReached) Then Exit For
                        budget.Register(entry.Key, entry.Size, entry.CompressedSize, entry.IsEncrypted, entry.LinkTarget)
                        If entry.IsDirectory OrElse ArchiveSafety.IsExcluded(entry.Key, _searchOptions) Then Continue For
                        Dim entryAttributes = GetArchiveAttributes(entry)
                        If entryAttributes.HasValue Then
                            Dim attributes = entryAttributes.Value
                            If Not _searchOptions.IncludeHiddenFiles AndAlso (attributes And FileAttributes.Hidden) <> 0 Then Continue For
                            If Not _searchOptions.IncludeSystemFiles AndAlso (attributes And FileAttributes.System) <> 0 Then Continue For
                        End If
                        Dim logicalEntry = logicalPath & " | " & entry.Key
                        Dim extension = Path.GetExtension(entry.Key).ToLowerInvariant()
                        Dim hit = NewSourceHit(Nothing, logicalEntry, HitKind.ZipTextEntry, physicalPath, entry.Key)
                        ThrottledSetCurrentFileDisplay(hit.DisplayName)
                        If _supportedArchiveExt.Contains(extension) Then
                            PublishSourceHit(hit)
                            If depth < _searchOptions.ArchiveNestingDepth AndAlso Not Volatile.Read(_resultLimitReached) Then
                                Dim nested = ExtractArchiveEntryToTemp(entry, physicalPath, ct, budget)
                                Await SearchArchiveSourceAsync(nested, logicalEntry, depth + 1, ct)
                            ElseIf depth >= _searchOptions.ArchiveNestingDepth Then
                                RecordFileSystemIssue(logicalEntry, New InvalidDataException("Archive nesting limit reached."))
                            End If
                            Continue For
                        End If
                        _filesScanned += 1
                        UpdateScanProgress()
                        If _searchOptions.SearchFileContents Then
                            If _supportedTextExt.Contains(extension) Then
                                Using stream = OpenBoundedArchiveEntry(entry, physicalPath, ct, budget)
                                    Await CollectTextDetailsAsync(stream, extension, hit, ct)
                                End Using
                            ElseIf _supportedEvtxExt.Contains(extension) Then
                                hit.Kind = HitKind.EvtxFromZipTemp
                                hit.TempEvtxPath = ExtractArchiveEntryToTemp(entry, physicalPath, ct, budget)
                                CollectEventDetails(hit.TempEvtxPath, hit, ct)
                            ElseIf _supportedHarExt.Contains(extension) Then
                                hit.Kind = HitKind.HarFromZip
                                Using stream = OpenBoundedArchiveEntry(entry, physicalPath, ct, budget)
                                    Await CollectHarDetailsAsync(stream, hit, ct)
                                End Using
                            End If
                        End If
                        PublishSourceHit(hit)
                    Next
                End Using
            Catch ex As OperationCanceledException
                Throw
            Catch ex As System.Text.RegularExpressions.RegexMatchTimeoutException
                Throw
            Catch ex As Exception
                RecordFileSystemIssue(logicalPath, ex)
            End Try
        End Function

        Private Sub CollectEventDetails(filePath As String, hit As SearchHit, ct As CancellationToken)
            Using reader As New EventLogReader(filePath, PathType.FilePath)
                While True
                    ct.ThrowIfCancellationRequested()
                    Using record = reader.ReadEvent()
                        If record Is Nothing Then Exit While
                        Dim provider = ReadEventField(Function() record.ProviderName, "Unknown")
                        Dim level = ReadEventField(Function() record.LevelDisplayName, "Unknown")
                        Dim message = ReadEventField(Function() record.FormatDescription(), "")
                        Dim xml = record.ToXml()
                        Dim timestamp = record.TimeCreated
                        Dim location = $"Event {record.Id} · {If(timestamp.HasValue, timestamp.Value.ToString("g"), "no timestamp")}"
                        Dim content = String.Join(vbLf, {"Event ID " & record.Id.ToString(), provider, level,
                                                       If(timestamp.HasValue, timestamp.Value.ToString("O"), ""), message})
                        Dim detail = SearchDetail.FromRecord(_activeQuery, content, location,
                                                              recordIndex:=hit.MatchingEvents.Count, token:=ct)
                        If detail Is Nothing Then
                            detail = SearchDetail.FromRecord(_activeQuery, xml, location & " (XML)",
                                                             recordIndex:=hit.MatchingEvents.Count, token:=ct)
                            If detail IsNot Nothing Then
                                Dim readableXml = PreviewFormatting.FormatXml(xml, True, CLng(_searchOptions.MaximumPreviewSizeMb) * 1024 * 1024)
                                Dim xmlSpans = _activeQuery.FindHighlights(readableXml, token:=ct)
                                If xmlSpans.Count > 0 Then detail.Context = MatchContext.FromText(readableXml, xmlSpans, ct)
                                detail.ContextKind = "Event XML lines"
                                message &= vbCrLf & vbCrLf & "Event XML:" & vbCrLf & readableXml
                            End If
                        End If
                        If detail Is Nothing Then Continue While
                        If detail.ContextKind = "Record text" Then detail.ContextKind = "Event text lines"
                        If String.IsNullOrWhiteSpace(message) Then
                            message = "Message resources are unavailable. Raw event XML:" & vbCrLf &
                                      PreviewFormatting.FormatXml(xml, True, CLng(_searchOptions.MaximumPreviewSizeMb) * 1024 * 1024)
                        End If
                        hit.MatchingEvents.Add(New EventSummary With {.Provider = provider, .Level = level,
                            .EventId = record.Id, .TimeCreated = timestamp, .Message = message})
                        hit.Details.Add(detail)
                        If _searchOptions.StopAfterFirstMatchPerFile OrElse
                           hit.MatchingEvents.Count >= Math.Min(_searchOptions.MaximumStructuredMatches, _searchOptions.EvtxMaximumMatches) Then
                            hit.PartialReason = If(_searchOptions.StopAfterFirstMatchPerFile, "first matching event only", "event limit reached")
                            Exit While
                        End If
                    End Using
                End While
            End Using
        End Sub

        Private Shared Function ReadEventField(read As Func(Of String), fallback As String) As String
            Try
                Return If(read(), fallback)
            Catch ex As EventLogException
                Return fallback
            End Try
        End Function

        Private Async Function CollectHarDetailsAsync(stream As Stream, hit As SearchHit, ct As CancellationToken) As Task
            Using document = Await JsonDocument.ParseAsync(stream, cancellationToken:=ct)
                Dim log As JsonElement
                Dim entries As JsonElement
                If document.RootElement.ValueKind <> JsonValueKind.Object OrElse
                   Not document.RootElement.TryGetProperty("log", log) OrElse log.ValueKind <> JsonValueKind.Object OrElse
                   Not log.TryGetProperty("entries", entries) OrElse entries.ValueKind <> JsonValueKind.Array Then
                    Throw New InvalidDataException("HAR content must contain a log.entries array.")
                End If
                Dim ordinal As Integer = 0
                For Each entry In entries.EnumerateArray()
                    ct.ThrowIfCancellationRequested()
                    ordinal += 1
                    Dim request = ParseHarEntry(entry)
                    If request Is Nothing Then Continue For
                    Dim content = String.Join(vbLf, {request.Method, request.Url, request.StatusCode.ToString(), request.StatusText,
                        request.ServerIpAddress, request.RequestHeaders, request.ResponseHeaders, request.RequestBody, request.ResponseBody})
                    Dim detail = SearchDetail.FromRecord(_activeQuery, content, $"Request {ordinal:N0} · {request.Method} {request.StatusCode}",
                                                          recordIndex:=hit.MatchingRequests.Count, token:=ct)
                    If detail Is Nothing Then Continue For
                    detail.ContextKind = "Request text lines"
                    hit.Details.Add(detail)
                    hit.MatchingRequests.Add(request)
                    If _searchOptions.StopAfterFirstMatchPerFile OrElse
                       hit.MatchingRequests.Count >= Math.Min(_searchOptions.MaximumStructuredMatches, _searchOptions.HarMaximumMatches) Then
                        hit.PartialReason = If(_searchOptions.StopAfterFirstMatchPerFile, "first matching request only", "request limit reached")
                        Exit For
                    End If
                Next
            End Using
        End Function

        Private Async Function SearchCabSourceAsync(physicalPath As String, logicalPath As String, depth As Integer, ct As CancellationToken) As Task
            Dim directory = Path.Combine(Path.GetTempPath(), "BeaconCabExtract_" & Guid.NewGuid().ToString("N"))
            System.IO.Directory.CreateDirectory(directory)
            _tempDirectories.Add(directory)
            If Not SevenZipHelper.ExtractCab(physicalPath, directory, _searchOptions, ct, AddressOf RecordFileSystemIssue) Then Return
            For Each filePath As String In FileSystemTraversal.EnumerateFiles(directory, ct, False, Nothing, AddressOf RecordFileSystemIssue, _searchOptions)
                ct.ThrowIfCancellationRequested()
                If Volatile.Read(_resultLimitReached) Then Exit For
                Dim relative = Path.GetRelativePath(directory, filePath)
                Await SearchDiskSourceAsync(filePath, logicalPath & " | " & relative, ct, depth + 1)
            Next
        End Function
    End Class
End Namespace
