Imports System.Runtime.ExceptionServices
Imports System.Threading
Imports System.Threading.Channels
Imports System.Threading.Tasks
Imports System.Windows.Threading

Namespace Beacon
    Partial Public Class MainWindow
        Private _scanGeneration As Long
        Private _publishingSearchBatch As Boolean
        Private _scanPublicationComplete As Boolean
        Private _latestScanProgress As ScanProgressSnapshot
        Private _progressDispatchPending As Integer
        Friend Property BeforeSourceForChecks As Func(Of String, Long, CancellationToken, Task)
        Friend Property ScanPolicyForChecks As SearchConcurrencyPolicy

        Private Async Function RunSourceWithPublicationAsync(root As String, token As CancellationToken) As Task
            Dim generation = If(_isScanning, Volatile.Read(_scanGeneration), Interlocked.Increment(_scanGeneration))
            Volatile.Write(_scanPublicationComplete, False)
            Dim output = Channel.CreateBounded(Of SourceSearchResult)(New BoundedChannelOptions(64) With {
                .SingleReader = True, .SingleWriter = True, .AllowSynchronousContinuations = False})
            Using cancellation = CancellationTokenSource.CreateLinkedTokenSource(token)
                Dim consumer = Dispatcher.InvokeAsync(Function() DrainSearchResultsAsync(output.Reader, generation, cancellation)).Task.Unwrap()
                Dim run As New SearchRunCoordinator(_activeQuery, _searchOptions, _supportedArchiveExt,
                    If(ScanPolicyForChecks, New SearchConcurrencyPolicy()),
                    progress:=Sub(logical, completed)
                                  If generation <> Volatile.Read(_scanGeneration) Then Return
                                  ThrottledSetCurrentFileDisplay(DisplaySourcePath(logical))
                                  If completed Then
                                      CountSourceFile()
                                  Else
                                      UpdateScanProgress()
                                  End If
                              End Sub,
                    report:=Sub(logical, ex, stage, severity)
                                If generation <> Volatile.Read(_scanGeneration) Then Return
                                If stage = "Search" Then
                                    RecordFileSystemIssue(logical, ex)
                                Else
                                    RecordDiagnostic(logical, ex, stage, severity)
                                End If
                            End Sub,
                    accessDenied:=Function(path) generation = Volatile.Read(_scanGeneration) AndAlso HandleAccessDeniedPath(path),
                    rememberSource:=Sub(physical, logical)
                                        If generation = Volatile.Read(_scanGeneration) Then RememberSourcePath(physical, logical)
                                    End Sub,
                    publishAsync:=Function(result, ct) output.Writer.WriteAsync(result, ct).AsTask())
                run.BeforeSourceAsync = BeforeSourceForChecks
                _sourceSearches.Add(run)
                Dim failure As ExceptionDispatchInfo = Nothing
                Try
                    Await run.RunAsync(root, cancellation.Token).ConfigureAwait(False)
                Catch ex As Exception
                    failure = ExceptionDispatchInfo.Capture(ex)
                Finally
                    If run.ResultLimitReached Then Volatile.Write(_resultLimitReached, True)
                    output.Writer.TryComplete()
                End Try
                Try
                    Await consumer.ConfigureAwait(False)
                Catch ex As Exception
                    If failure Is Nothing OrElse TypeOf failure.SourceException Is OperationCanceledException Then
                        failure = ExceptionDispatchInfo.Capture(ex)
                    End If
                End Try
                Volatile.Write(_scanPublicationComplete, failure Is Nothing AndAlso Not token.IsCancellationRequested AndAlso Not run.ResultLimitReached)
                UpdateScanProgress()
                Await Dispatcher.InvokeAsync(Sub() ApplyPendingScanProgress(), DispatcherPriority.Background).Task.ConfigureAwait(False)
                If failure IsNot Nothing Then failure.Throw()
            End Using
        End Function

        Private Async Function DrainSearchResultsAsync(reader As ChannelReader(Of SourceSearchResult), generation As Long,
                                                       cancellation As CancellationTokenSource) As Task
            Try
                While Await reader.WaitToReadAsync()
                    Dim batch As Integer = 0
                    Dim result As SourceSearchResult = Nothing
                    _publishingSearchBatch = True
                    Try
                        While batch < 32 AndAlso reader.TryRead(result)
                            If generation = Volatile.Read(_scanGeneration) AndAlso Not (_isResetting OrElse _isClosing) Then PublishServiceResult(result)
                            batch += 1
                        End While
                    Finally
                        _publishingSearchBatch = False
                    End Try
                    UpdateReportingButtons()
                    Await System.Windows.Threading.Dispatcher.Yield(DispatcherPriority.Background)
                End While
            Catch
                cancellation.Cancel()
                Throw
            End Try
        End Function

        Private Sub QueueScanProgress()
            Dim snapshot As New ScanProgressSnapshot(Volatile.Read(_scanGeneration), _isCountingFiles,
                Volatile.Read(_totalFilesToScan), Volatile.Read(_filesScanned), _pendingFileLabel, Volatile.Read(_scanPublicationComplete))
            Volatile.Write(_latestScanProgress, snapshot)
            If Interlocked.Exchange(_progressDispatchPending, 1) <> 0 Then Return
            Dispatcher.BeginInvoke(New Action(AddressOf ApplyPendingScanProgress), DispatcherPriority.Background)
        End Sub

        Private Sub ApplyPendingScanProgress()
            Interlocked.Exchange(_progressDispatchPending, 0)
            Dim snapshot = Volatile.Read(_latestScanProgress)
            If snapshot Is Nothing OrElse snapshot.Generation <> Volatile.Read(_scanGeneration) OrElse
               snapshot.Counting <> _isCountingFiles OrElse Not _isScanning OrElse _isResetting OrElse _isClosing Then Return
            If snapshot.Counting Then
                CurrentFile_lbl.Text = $"Counting files: {snapshot.Total} found — {snapshot.Label}"
                Return
            End If
            If snapshot.Total > 0 Then
                Dim percentage = (CDbl(snapshot.Completed) / snapshot.Total) * 100
                ScanProgress_pb.Value = Math.Max(ScanProgress_pb.Value, Math.Min(percentage, If(snapshot.Published, 100.0, 99.9)))
            End If
            If Not String.IsNullOrEmpty(snapshot.Label) Then
                CurrentFile_lbl.Text = If(snapshot.Total > 0,
                    $"Scanning {snapshot.Completed} of {snapshot.Total} counted files: {snapshot.Label}",
                    $"Scanning file {snapshot.Completed}: {snapshot.Label}")
            End If
        End Sub

        Private NotInheritable Class ScanProgressSnapshot
            Public ReadOnly Generation As Long
            Public ReadOnly Counting As Boolean
            Public ReadOnly Total As Integer
            Public ReadOnly Completed As Integer
            Public ReadOnly Label As String
            Public ReadOnly Published As Boolean
            Public Sub New(generation As Long, counting As Boolean, total As Integer, completed As Integer, label As String, published As Boolean)
                Me.Generation = generation
                Me.Counting = counting
                Me.Total = total
                Me.Completed = completed
                Me.Label = label
                Me.Published = published
            End Sub
        End Class
    End Class
End Namespace
