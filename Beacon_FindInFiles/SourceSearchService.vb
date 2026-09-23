Imports System.Diagnostics.Eventing.Reader
Imports System.IO
Imports System.Text.Json
Imports System.Text.RegularExpressions
Imports System.Threading
Imports System.Threading.Tasks
Imports SharpCompress.Archives

Namespace Beacon
    Public NotInheritable Class SourceSearchService
        Implements IDisposable
        Private ReadOnly _options As BeaconSettings
        Private ReadOnly _query As SearchQuery
        Private ReadOnly _archives As HashSet(Of String)
        Private ReadOnly _text As HashSet(Of String)
        Private ReadOnly _temporary As New SearchTemporarySources()
        Private ReadOnly _execution As SourceSearchExecution
        Private ReadOnly _publish As Action(Of SourceSearchResult)
        Private ReadOnly _progress As Action(Of String, Boolean)
        Private ReadOnly _report As Action(Of String, Exception, String, String)
        Private ReadOnly _access As Func(Of String, Boolean)
        Private ReadOnly _remember As Action(Of String, String)
        Private _counting As Boolean
        Private _started As Integer
        Private _disposed As Boolean
        Public ReadOnly Property FilesProcessed As Integer
        Public ReadOnly Property MatchingFiles As Integer
        Public ReadOnly Property ResultLimitReached As Boolean

        Public Sub New(query As SearchQuery, options As BeaconSettings, archiveExtensions As IEnumerable(Of String),
                       Optional publish As Action(Of SourceSearchResult) = Nothing,
                       Optional progress As Action(Of String, Boolean) = Nothing,
                       Optional report As Action(Of String, Exception, String, String) = Nothing,
                       Optional accessDenied As Func(Of String, Boolean) = Nothing,
                       Optional rememberSource As Action(Of String, String) = Nothing)
            _query = query
            _options = BeaconSettingsService.Clone(options)
            _archives = New HashSet(Of String)(archiveExtensions, StringComparer.OrdinalIgnoreCase)
            _text = New HashSet(Of String)(_options.IncludedExtensions.Split(";"c, StringSplitOptions.RemoveEmptyEntries).
                Select(Function(value) value.Trim()).Where(Function(value) Not _archives.Contains(value) AndAlso value <> ".har" AndAlso value <> ".evtx"), StringComparer.OrdinalIgnoreCase)
            _publish = publish
            _progress = progress
            _report = report
            _access = accessDenied
            _remember = rememberSource
        End Sub

        Friend Sub New(query As SearchQuery, options As BeaconSettings, archiveExtensions As IEnumerable(Of String), execution As SourceSearchExecution)
            Me.New(query, options, archiveExtensions, progress:=AddressOf execution.Progress,
                   report:=AddressOf execution.Report, rememberSource:=AddressOf execution.Remember)
            _temporary.Dispose()
            _temporary = execution.Temporary
            _execution = execution
        End Sub

        Public Async Function RunAsync(root As String, countingOnly As Boolean, token As CancellationToken) As Task
            ObjectDisposedException.ThrowIf(_disposed, Me)
            If Interlocked.Exchange(_started, 1) <> 0 Then Throw New InvalidOperationException("A source-search service can run only once.")
            _counting = countingOnly
            token.ThrowIfCancellationRequested()
            If Not countingOnly Then
                EvtxFilter.FromSettings(_options)
                HarFilter.FromSettings(_options)
            End If
            If Directory.Exists(root) Then
                For Each path In FileSystemTraversal.EnumerateFiles(root, token, _options.FollowReparsePoints, _access, AddressOf Issue, _options)
                    token.ThrowIfCancellationRequested()
                    If ResultLimitReached Then Exit For
                    Await SearchDiskAsync(path, IO.Path.GetFullPath(path), 0, token).ConfigureAwait(False)
                Next
            ElseIf FileSystemTraversal.IsFileAllowed(root, _options, AddressOf Issue) Then
                Await SearchDiskAsync(root, Path.GetFullPath(root), 0, token).ConfigureAwait(False)
            End If
        End Function

        Private Sub Issue(path As String, ex As Exception)
            _report?.Invoke(path, ex, "Search", "Warning")
        End Sub
        Private Function NewResult(physical As String, logical As String, kind As SourceResultKind,
                                   Optional archive As String = Nothing, Optional entry As String = Nothing) As SourceSearchResult
            _remember?.Invoke(physical, logical)
            _progress?.Invoke(logical, False)
            Dim result As New SourceSearchResult With {.PhysicalPath = physical, .LogicalPath = logical, .Kind = kind, .ArchivePath = archive, .EntryName = entry}
            If _counting Then Return result
            If _options.SearchFileNames Then
                Dim detail = SearchDetail.FromRecord(_query, Path.GetFileName(If(entry, physical)), "File name", metadata:=True)
                If detail IsNot Nothing Then result.Details.Add(detail)
            End If
            If _options.SearchFullPaths Then
                Dim detail = SearchDetail.FromRecord(_query, logical, "Full path", metadata:=True)
                If detail IsNot Nothing Then result.Details.Add(detail)
            End If
            Return result
        End Function
        Private Async Function PublishAsync(result As SourceSearchResult, token As CancellationToken) As Task
            If _counting OrElse result.Details.Count = 0 OrElse ResultLimitReached Then Return
            _MatchingFiles += 1
            If _execution IsNot Nothing Then
                Await _execution.PublishAsync(result, token).ConfigureAwait(False)
            Else
                _publish?.Invoke(result)
            End If
            _ResultLimitReached = MatchingFiles >= _options.MaximumTotalResults
        End Function
        Private Sub Processed(logical As String)
            _FilesProcessed += 1
            _progress?.Invoke(logical, True)
        End Sub

        Private Async Function SearchDiskAsync(physical As String, logical As String, depth As Integer, token As CancellationToken) As Task
            token.ThrowIfCancellationRequested()
            Dim extension = Path.GetExtension(physical).ToLowerInvariant()
            Dim result = NewResult(physical, logical, SourceResultKind.DiskText)
            If _execution IsNot Nothing Then Await _execution.FlushAsync(token).ConfigureAwait(False)
            Try
                If _archives.Contains(extension) Then
                    Await PublishAsync(result, token).ConfigureAwait(False)
                    If Not ResultLimitReached Then Await SearchArchiveAsync(physical, logical, depth, token).ConfigureAwait(False)
                    Return
                End If
                Processed(logical)
                If _counting Then Return
                If _options.SearchFileContents Then
                    If extension = ".evtx" Then
                        result.Kind = SourceResultKind.DiskEvent
                        CollectEvents(physical, result, token)
                    ElseIf extension = ".har" OrElse _text.Contains(extension) Then
                        If extension = ".har" Then result.Kind = SourceResultKind.DiskHar
                        Using stream As New BoundedReadStream(New FileStream(physical, FileMode.Open, FileAccess.Read, FileShare.ReadWrite),
                                                              CLng(_options.MaximumFileSizeMb) * 1024 * 1024, token,
                                                              onError:=Sub(ex) Issue(logical, ex))
                            Await CollectStreamAsync(stream, extension, result, token,
                                If(IsMaterialized(extension), New FileInfo(physical).Length, 0L)).ConfigureAwait(False)
                        End Using
                    End If
                End If
            Catch ex As Exception When TypeOf ex Is IOException OrElse TypeOf ex Is UnauthorizedAccessException OrElse
                                        TypeOf ex Is InvalidDataException OrElse TypeOf ex Is EventLogException OrElse TypeOf ex Is JsonException
                Issue(logical, ex)
                result.PartialReason = "content unavailable"
            End Try
            If _execution IsNot Nothing Then Await _execution.CompletedAsync(logical, token).ConfigureAwait(False)
            Await PublishAsync(result, token).ConfigureAwait(False)
        End Function

        Private Async Function SearchArchiveAsync(physical As String, logical As String, depth As Integer, token As CancellationToken) As Task
            Dim lease As IDisposable = Nothing
            Dim extension = Path.GetExtension(physical).ToLowerInvariant()
            If _execution IsNot Nothing AndAlso extension <> ".zip" AndAlso extension <> ".tar" Then
                lease = Await _execution.EnterHeavyAsync(CLng(_options.MaximumArchiveEntrySizeMb) * SearchConcurrencyPolicy.MiB * 8, token).ConfigureAwait(False)
            End If
            Using lease
                Await SearchArchiveCoreAsync(physical, logical, depth, token).ConfigureAwait(False)
            End Using
        End Function

        Private Async Function SearchArchiveCoreAsync(physical As String, logical As String, depth As Integer, token As CancellationToken) As Task
            _remember?.Invoke(physical, logical)
            If depth > _options.ArchiveNestingDepth Then
                Issue(logical, New InvalidDataException("Archive nesting limit reached."))
                Return
            End If
            If Path.GetExtension(physical).Equals(".cab", StringComparison.OrdinalIgnoreCase) Then
                Dim directory = _temporary.CreateDirectory()
                If Not SevenZipHelper.ExtractCab(physical, directory, _options, token, AddressOf Issue) Then Return
                For Each path In FileSystemTraversal.EnumerateFiles(directory, token, False, Nothing, AddressOf Issue, _options)
                    token.ThrowIfCancellationRequested()
                    If ResultLimitReached Then Exit For
                    Await SearchDiskAsync(path, logical & " | " & IO.Path.GetRelativePath(directory, path), depth + 1, token).ConfigureAwait(False)
                Next
                Return
            End If
            Try
                Using archive = ArchiveCompatibility.OpenArchive(physical, _options, token)
                    Dim budget As New ArchiveReadBudget(_options, New FileInfo(physical).Length)
                    For Each entry In archive.Entries
                        token.ThrowIfCancellationRequested()
                        If ResultLimitReached Then Exit For
                        budget.Register(entry.Key, entry.Size, entry.CompressedSize, entry.IsEncrypted, entry.LinkTarget)
                        If entry.IsDirectory OrElse ArchiveSafety.IsExcluded(entry.Key, _options) Then Continue For
                        Dim attributes = AttributesOf(entry)
                        If attributes.HasValue Then
                            If Not _options.IncludeHiddenFiles AndAlso (attributes.Value And FileAttributes.Hidden) <> 0 Then Continue For
                            If Not _options.IncludeSystemFiles AndAlso (attributes.Value And FileAttributes.System) <> 0 Then Continue For
                        End If
                        Dim logicalEntry = logical & " | " & entry.Key
                        Dim extension = Path.GetExtension(entry.Key).ToLowerInvariant()
                        Dim result = NewResult(Nothing, logicalEntry, SourceResultKind.ArchiveText, physical, entry.Key)
                        If _execution IsNot Nothing Then Await _execution.FlushAsync(token).ConfigureAwait(False)
                        If _archives.Contains(extension) Then
                            Await PublishAsync(result, token).ConfigureAwait(False)
                            If depth < _options.ArchiveNestingDepth AndAlso Not ResultLimitReached Then
                                Dim nested = Extract(entry, logicalEntry, budget, token)
                                Await SearchArchiveAsync(nested, logicalEntry, depth + 1, token).ConfigureAwait(False)
                            ElseIf depth >= _options.ArchiveNestingDepth Then
                                Issue(logicalEntry, New InvalidDataException("Archive nesting limit reached."))
                            End If
                            Continue For
                        End If
                        Processed(logicalEntry)
                        If _counting Then Continue For
                        If _options.SearchFileContents Then
                            If extension = ".evtx" Then
                                result.Kind = SourceResultKind.ArchiveEvent
                                result.TemporaryEventPath = Extract(entry, logicalEntry, budget, token)
                                CollectEvents(result.TemporaryEventPath, result, token)
                            ElseIf extension = ".har" OrElse _text.Contains(extension) Then
                                If extension = ".har" Then result.Kind = SourceResultKind.ArchiveHar
                                Using stream = OpenEntry(entry, logicalEntry, budget, token)
                                    Await CollectStreamAsync(stream, extension, result, token, entry.Size).ConfigureAwait(False)
                                End Using
                            End If
                        End If
                        If _execution IsNot Nothing Then Await _execution.CompletedAsync(logicalEntry, token).ConfigureAwait(False)
                        Await PublishAsync(result, token).ConfigureAwait(False)
                    Next
                End Using
            Catch ex As OperationCanceledException
                Throw
            Catch ex As RegexMatchTimeoutException
                Throw
            Catch ex As Exception
                Issue(logical, ex)
            End Try
        End Function

        Friend Shared Function AttributesOf(entry As IArchiveEntry) As FileAttributes?
            Try
                Dim value = entry.Attrib
                Return If(value.HasValue, CType(value.Value, FileAttributes), CType(Nothing, FileAttributes?))
            Catch ex As NotImplementedException
                Return Nothing
            End Try
        End Function
        Private Function OpenEntry(entry As IArchiveEntry, logical As String, budget As ArchiveReadBudget, token As CancellationToken) As Stream
            Return New BoundedReadStream(ArchiveCompatibility.OpenEntry(entry), Math.Min(entry.Size, budget.EntryLimit), token, budget, Sub(ex) Issue(logical, ex))
        End Function
        Private Function Extract(entry As IArchiveEntry, logical As String, budget As ArchiveReadBudget, token As CancellationToken) As String
            Dim path = IO.Path.Combine(_temporary.CreateDirectory(), IO.Path.GetFileName(entry.Key))
            Using input = OpenEntry(entry, logical, budget, token), output As New FileStream(path, FileMode.CreateNew, FileAccess.Write, FileShare.None)
                input.CopyTo(output)
            End Using
            Return path
        End Function
        Private Sub CollectEvents(path As String, result As SourceSearchResult, token As CancellationToken)
            Dim collected As New StructuredSearchResult(Of EventRecordSummary)()
            Try
                Dim service As New EvtxSearchService(_query, _options)
                service.Collect(path, collected, token, Sub(ex, stage, severity) _report?.Invoke(result.LogicalPath, ex, stage, severity))
            Finally
                result.Events.AddRange(collected.Records)
                result.Details.AddRange(collected.Details)
                result.PartialReason = collected.PartialReason
            End Try
        End Sub
        Private Shared Function IsMaterialized(extension As String) As Boolean
            Return extension = ".har" OrElse extension = ".html" OrElse extension = ".xml" OrElse extension = ".json"
        End Function

        Private Async Function CollectStreamAsync(stream As Stream, extension As String, result As SourceSearchResult,
                                                 token As CancellationToken, estimatedLength As Long) As Task
            Dim lease As IDisposable = Nothing
            If _execution IsNot Nothing AndAlso IsMaterialized(extension) Then
                Dim length = If(estimatedLength > 0, estimatedLength,
                    CLng(Math.Max(_options.MaximumFileSizeMb, _options.MaximumArchiveEntrySizeMb)) * SearchConcurrencyPolicy.MiB)
                lease = Await _execution.EnterHeavyAsync(SearchConcurrencyPolicy.WorkerBytes + length * 8, token).ConfigureAwait(False)
            End If
            Using lease
                Await CollectStreamCoreAsync(stream, extension, result, token).ConfigureAwait(False)
            End Using
        End Function

        Private Async Function CollectStreamCoreAsync(stream As Stream, extension As String, result As SourceSearchResult, token As CancellationToken) As Task
            If extension = ".har" Then
                Dim collected As New StructuredSearchResult(Of HarRecord)()
                Try
                    Dim service As New HarSearchService(_query, _options)
                    Await service.CollectAsync(stream, collected, token,
                        Sub(ex, stage, severity) _report?.Invoke(result.LogicalPath, ex, stage, severity)).ConfigureAwait(False)
                Finally
                    result.Requests.AddRange(collected.Records)
                    result.Details.AddRange(collected.Details)
                    result.PartialReason = collected.PartialReason
                End Try
            Else
                Dim resultText = Await SearchTextCollector.CollectAsync(stream, _query, _options.MaximumStructuredMatches,
                    _options.StopAfterFirstMatchPerFile, extension = ".html" OrElse extension = ".xml" OrElse extension = ".json", token).ConfigureAwait(False)
                result.Details.AddRange(resultText.Details)
                result.PartialReason = resultText.PartialReason
            End If
        End Function
        Public Sub Dispose() Implements IDisposable.Dispose
            If _disposed Then Return
            _disposed = True
            If _execution Is Nothing Then _temporary.Dispose()
        End Sub
    End Class
End Namespace
