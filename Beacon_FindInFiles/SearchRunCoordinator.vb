Imports System.IO
Imports System.Runtime.ExceptionServices
Imports System.Threading
Imports System.Threading.Channels
Imports System.Threading.Tasks

Namespace Beacon
    Public NotInheritable Class SearchRunCoordinator
        Implements IDisposable
        Private ReadOnly _query As SearchQuery
        Private ReadOnly _options As BeaconSettings
        Private ReadOnly _extensions As String()
        Private ReadOnly _policy As SearchConcurrencyPolicy
        Private ReadOnly _resources As SearchResourceGate
        Private ReadOnly _outputBudget As SearchOutputBudget
        Private ReadOnly _publish As Action(Of SourceSearchResult)
        Private ReadOnly _publishAsync As Func(Of SourceSearchResult, CancellationToken, Task)
        Private ReadOnly _progress As Action(Of String, Boolean)
        Private ReadOnly _report As Action(Of String, Exception, String, String)
        Private ReadOnly _access As Func(Of String, Boolean)
        Private ReadOnly _remember As Action(Of String, String)
        Private ReadOnly _retained As New List(Of SearchTemporarySources)()
        Private _state As Integer
        Private _disposed As Boolean
        Private _activeWorkers As Integer
        Private _peakWorkers As Integer
        Private _headOrdinal As Long
        Public ReadOnly Property FilesProcessed As Integer
        Public ReadOnly Property MatchingFiles As Integer
        Public ReadOnly Property ResultLimitReached As Boolean
        Friend ReadOnly Property WorkerCeiling As Integer
            Get
                Return _policy.MaximumWorkers
            End Get
        End Property
        Public ReadOnly Property PeakWorkers As Integer
            Get
                Return Volatile.Read(_peakWorkers)
            End Get
        End Property
        Friend Property BeforeSourceAsync As Func(Of String, Long, CancellationToken, Task)
        Friend Property EventQueuedForChecks As Action(Of Long, SourceSearchEventKind)

        Public Sub New(query As SearchQuery, options As BeaconSettings, archiveExtensions As IEnumerable(Of String),
                       Optional publish As Action(Of SourceSearchResult) = Nothing,
                       Optional progress As Action(Of String, Boolean) = Nothing,
                       Optional report As Action(Of String, Exception, String, String) = Nothing,
                       Optional accessDenied As Func(Of String, Boolean) = Nothing,
                       Optional rememberSource As Action(Of String, String) = Nothing,
                       Optional publishAsync As Func(Of SourceSearchResult, CancellationToken, Task) = Nothing)
            Me.New(query, options, archiveExtensions, New SearchConcurrencyPolicy(), publish, progress, report, accessDenied, rememberSource, publishAsync)
        End Sub

        Friend Sub New(query As SearchQuery, options As BeaconSettings, archiveExtensions As IEnumerable(Of String),
                       policy As SearchConcurrencyPolicy,
                       Optional publish As Action(Of SourceSearchResult) = Nothing,
                       Optional progress As Action(Of String, Boolean) = Nothing,
                       Optional report As Action(Of String, Exception, String, String) = Nothing,
                       Optional accessDenied As Func(Of String, Boolean) = Nothing,
                       Optional rememberSource As Action(Of String, String) = Nothing,
                       Optional publishAsync As Func(Of SourceSearchResult, CancellationToken, Task) = Nothing)
            If publish IsNot Nothing AndAlso publishAsync IsNot Nothing Then Throw New ArgumentException("Choose one result publisher.")
            _query = query
            _options = BeaconSettingsService.Clone(options)
            _extensions = archiveExtensions.ToArray()
            _policy = policy
            _resources = New SearchResourceGate(policy, Function() Volatile.Read(_headOrdinal))
            _outputBudget = New SearchOutputBudget(policy, Function() Volatile.Read(_headOrdinal))
            _publish = publish
            _publishAsync = publishAsync
            _progress = progress
            _report = report
            _access = accessDenied
            _remember = rememberSource
        End Sub

        Public Async Function RunAsync(root As String, token As CancellationToken) As Task
            ObjectDisposedException.ThrowIf(_disposed, Me)
            If Interlocked.CompareExchange(_state, 1, 0) <> 0 Then Throw New InvalidOperationException("A search run can execute only once.")
            Dim pending As New Queue(Of WorkItem)()
            Dim workers As New List(Of Task)()
            Dim discovery = Channel.CreateBounded(Of DiscoveryItem)(New BoundedChannelOptions(_policy.MaximumWorkers) With {
                .SingleReader = True, .SingleWriter = True, .AllowSynchronousContinuations = False})
            Dim work = Channel.CreateBounded(Of WorkItem)(New BoundedChannelOptions(_policy.MaximumWorkers) With {
                .SingleReader = False, .SingleWriter = True, .AllowSynchronousContinuations = False})
            Using cancellation = CancellationTokenSource.CreateLinkedTokenSource(token)
                Dim ct = cancellation.Token
                Dim producer = Task.Factory.StartNew(Sub() Discover(root, discovery.Writer, ct), CancellationToken.None,
                                                     TaskCreationOptions.LongRunning, TaskScheduler.Default)
                Dim failure As ExceptionDispatchInfo = Nothing
                Try
                    EvtxFilter.FromSettings(_options)
                    HarFilter.FromSettings(_options)
                    Dim barrier As DiscoveryItem = Nothing
                    Dim ended As Boolean
                    Dim ordinal As Long
                    While Not ResultLimitReached
                        ct.ThrowIfCancellationRequested()
                        While barrier Is Nothing AndAlso Not ended AndAlso pending.Count < _policy.AdmissionLimit()
                            If Not Await discovery.Reader.WaitToReadAsync(ct).ConfigureAwait(False) Then
                                ended = True
                                Exit While
                            End If
                            Dim found = Await discovery.Reader.ReadAsync(ct).ConfigureAwait(False)
                            If found.Decision IsNot Nothing OrElse found.Error IsNot Nothing Then
                                barrier = found
                                Exit While
                            End If
                            Dim job As New WorkItem(found.Path, ordinal)
                            ordinal += 1
                            job.Execution = New SourceSearchExecution(
                                Function(item, sourceToken) SendEventAsync(job, item, sourceToken),
                                Function(bytes, sourceToken) _resources.EnterAsync(job.Ordinal, bytes, sourceToken))
                            pending.Enqueue(job)
                            If workers.Count < pending.Count Then workers.Add(Task.Run(Function() WorkerAsync(work.Reader, ct)))
                            Await work.Writer.WriteAsync(job, ct).ConfigureAwait(False)
                        End While
                        If pending.Count > 0 Then
                            Dim job = pending.Peek()
                            Volatile.Write(_headOrdinal, job.Ordinal)
                            Await CommitJobAsync(job, ct).ConfigureAwait(False)
                            If ResultLimitReached Then Exit While
                            Release(pending.Dequeue())
                        ElseIf barrier IsNot Nothing Then
                            If barrier.Fatal Then
                                ExceptionDispatchInfo.Capture(barrier.Error).Throw()
                            ElseIf barrier.Decision IsNot Nothing Then
                                barrier.Decision.TrySetResult(_access IsNot Nothing AndAlso _access(barrier.Path))
                            Else
                                _report?.Invoke(barrier.Path, barrier.Error, "Search", "Warning")
                            End If
                            barrier = Nothing
                        ElseIf ended Then
                            Exit While
                        End If
                    End While
                    token.ThrowIfCancellationRequested()
                Catch ex As Exception
                    failure = ExceptionDispatchInfo.Capture(ex)
                End Try

                cancellation.Cancel()
                work.Writer.TryComplete()
                Try
                    Await Task.WhenAll(workers.Append(producer)).ConfigureAwait(False)
                Catch ex As Exception
                    If failure Is Nothing Then failure = ExceptionDispatchInfo.Capture(ex)
                End Try
                Try
                    While pending.Count > 0
                        Release(pending.Dequeue())
                    End While
                Finally
                    Volatile.Write(_state, 2)
                End Try
                If failure IsNot Nothing Then
                    If TypeOf failure.SourceException Is OperationCanceledException AndAlso Not token.IsCancellationRequested Then
                        Throw New InvalidOperationException("A source stopped without a search cancellation request.", failure.SourceException)
                    End If
                    failure.Throw()
                End If
            End Using
        End Function

        Private Sub Discover(root As String, writer As ChannelWriter(Of DiscoveryItem), token As CancellationToken)
            ' This single dedicated thread adapts the synchronous iterator and permission callback.
            ' Ordered barriers keep access prompts behind previously discovered searches.
            Dim send As Action(Of DiscoveryItem) = Sub(item) writer.WriteAsync(item, token).AsTask().GetAwaiter().GetResult()
            Dim issue As Action(Of String, Exception) = Sub(path, ex) send(New DiscoveryItem With {.Path = path, .Error = ex})
            Dim access As Func(Of String, Boolean) = Function(path)
                                                        Dim decision As New TaskCompletionSource(Of Boolean)(TaskCreationOptions.RunContinuationsAsynchronously)
                                                        send(New DiscoveryItem With {.Path = path, .Decision = decision})
                                                        Return decision.Task.WaitAsync(token).GetAwaiter().GetResult()
                                                    End Function
            Try
                token.ThrowIfCancellationRequested()
                If Directory.Exists(root) Then
                    For Each path In FileSystemTraversal.EnumerateFiles(root, token, _options.FollowReparsePoints,
                                                                       If(_access Is Nothing, Nothing, access), issue, _options)
                        send(New DiscoveryItem With {.Path = IO.Path.GetFullPath(path)})
                    Next
                ElseIf FileSystemTraversal.IsFileAllowed(root, _options, issue) Then
                    send(New DiscoveryItem With {.Path = IO.Path.GetFullPath(root)})
                End If
                writer.TryComplete()
            Catch ex As OperationCanceledException When token.IsCancellationRequested
                writer.TryComplete()
            Catch ex As Exception
                Try
                    send(New DiscoveryItem With {.Path = root, .Error = ex, .Fatal = True})
                Catch cancelled As OperationCanceledException When token.IsCancellationRequested
                Finally
                    writer.TryComplete()
                End Try
            End Try
        End Sub

        Private Async Function WorkerAsync(reader As ChannelReader(Of WorkItem), token As CancellationToken) As Task
            ' Regex has a single cached runner; each sequential worker reuses its own query.
            Dim query As New Lazy(Of SearchQuery)(Function() New SearchQuery(_query.Text, _query.Mode, _query.CaseSensitive), LazyThreadSafetyMode.None)
            Try
                While Await reader.WaitToReadAsync(token).ConfigureAwait(False)
                    Dim job As WorkItem = Nothing
                    While reader.TryRead(job)
                        token.ThrowIfCancellationRequested()
                        Dim active = Interlocked.Increment(_activeWorkers)
                        Dim previous = Volatile.Read(_peakWorkers)
                        While active > previous
                            Dim observed = Interlocked.CompareExchange(_peakWorkers, active, previous)
                            If observed = previous Then Exit While
                            previous = observed
                        End While
                        Try
                            Await ProcessJobAsync(job, query, token).ConfigureAwait(False)
                        Finally
                            Interlocked.Decrement(_activeWorkers)
                        End Try
                    End While
                End While
            Catch ex As OperationCanceledException When token.IsCancellationRequested
            End Try
        End Function

        Private Async Function SendEventAsync(job As WorkItem, item As SourceSearchEvent, token As CancellationToken) As Task
            If item.Kind = SourceSearchEventKind.Result Then
                Await job.ResultSlot.WaitAsync(token).ConfigureAwait(False)
                Dim lease As IDisposable = Nothing
                Try
                    lease = Await _outputBudget.EnterAsync(job.Ordinal, item.Result, token).ConfigureAwait(False)
                Catch
                    job.ResultSlot.Release()
                    Throw
                End Try
                item.OutputLease = New SearchResourceLease(Sub()
                                                              lease?.Dispose()
                                                              job.ResultSlot.Release()
                                                          End Sub)
            End If
            Try
                Await job.Events.Writer.WriteAsync(item, token).ConfigureAwait(False)
                EventQueuedForChecks?.Invoke(job.Ordinal, item.Kind)
            Catch
                item.OutputLease?.Dispose()
                Throw
            End Try
        End Function

        Private Async Function ProcessJobAsync(job As WorkItem, query As Lazy(Of SearchQuery), token As CancellationToken) As Task
            Try
                If BeforeSourceAsync IsNot Nothing Then Await BeforeSourceAsync(job.Path, job.Ordinal, token).ConfigureAwait(False)
                Using service As New SourceSearchService(query.Value, _options, _extensions, job.Execution)
                    Await service.RunAsync(job.Path, False, token).ConfigureAwait(False)
                End Using
            Catch ex As Exception
                job.Failure = ExceptionDispatchInfo.Capture(ex)
            End Try
            Try
                Await job.Execution.FlushAsync(token).ConfigureAwait(False)
            Catch ex As Exception
                If job.Failure Is Nothing Then job.Failure = ExceptionDispatchInfo.Capture(ex)
            End Try
            job.Events.Writer.TryComplete()
            job.Completion.TrySetResult(True)
        End Function

        Private Async Function CommitJobAsync(job As WorkItem, token As CancellationToken) As Task
            While Not ResultLimitReached AndAlso Await job.Events.Reader.WaitToReadAsync(token).ConfigureAwait(False)
                Dim item As SourceSearchEvent = Nothing
                While Not ResultLimitReached AndAlso job.Events.Reader.TryRead(item)
                    If token.IsCancellationRequested Then item.OutputLease?.Dispose()
                    token.ThrowIfCancellationRequested()
                    Select Case item.Kind
                        Case SourceSearchEventKind.Result
                            Try
                                job.Retain = _publish IsNot Nothing OrElse _publishAsync IsNot Nothing
                                _MatchingFiles += 1
                                If _publishAsync IsNot Nothing Then
                                    Await _publishAsync(item.Result, token).ConfigureAwait(False)
                                Else
                                    _publish?.Invoke(item.Result)
                                End If
                                _ResultLimitReached = MatchingFiles >= Math.Max(1, _options.MaximumTotalResults)
                            Finally
                                item.OutputLease?.Dispose()
                            End Try
                        Case SourceSearchEventKind.Completed
                            _FilesProcessed += 1
                            _progress?.Invoke(item.Path, True)
                        Case SourceSearchEventKind.Started
                            _progress?.Invoke(item.Path, False)
                        Case SourceSearchEventKind.Issue
                            _report?.Invoke(item.Path, item.Error, item.Stage, item.Severity)
                        Case SourceSearchEventKind.Remember
                            _remember?.Invoke(item.PhysicalPath, item.Path)
                    End Select
                End While
            End While
            If Not ResultLimitReached Then
                Await job.Completion.Task.ConfigureAwait(False)
                If job.Failure IsNot Nothing Then job.Failure.Throw()
            End If
        End Function

        Private Sub Release(job As WorkItem)
            Dim item As SourceSearchEvent = Nothing
            While job.Events.Reader.TryRead(item)
                item.OutputLease?.Dispose()
            End While
            job.ResultSlot.Dispose()
            If job.Retain AndAlso job.Execution.Temporary.HasSources Then
                _retained.Add(job.Execution.Temporary)
            Else
                job.Execution.Temporary.Dispose()
            End If
        End Sub

        Public Sub Dispose() Implements IDisposable.Dispose
            If _disposed Then Return
            If Volatile.Read(_state) = 1 Then Throw New InvalidOperationException("Await the search run before releasing its sources.")
            _disposed = True
            For Each source In _retained
                source.Dispose()
            Next
            _retained.Clear()
            _resources.Dispose()
        End Sub

        Private NotInheritable Class DiscoveryItem
            Public Property Path As String
            Public Property [Error] As Exception
            Public Property Fatal As Boolean
            Public Property Decision As TaskCompletionSource(Of Boolean)
        End Class

        Private NotInheritable Class WorkItem
            Public ReadOnly Path As String
            Public ReadOnly Ordinal As Long
            Public ReadOnly Events As Channel(Of SourceSearchEvent) = Channel.CreateBounded(Of SourceSearchEvent)(
                New BoundedChannelOptions(16) With {.SingleReader = True, .SingleWriter = True, .AllowSynchronousContinuations = False})
            Public ReadOnly Completion As New TaskCompletionSource(Of Boolean)(TaskCreationOptions.RunContinuationsAsynchronously)
            Public ReadOnly ResultSlot As New SemaphoreSlim(1, 1)
            Public Execution As SourceSearchExecution
            Public Failure As ExceptionDispatchInfo
            Public Retain As Boolean
            Public Sub New(path As String, ordinal As Long)
                Me.Path = path
                Me.Ordinal = ordinal
            End Sub
        End Class
    End Class
End Namespace
