Imports System.Threading
Imports System.Threading.Tasks

Namespace Beacon
    Partial Public Class MainWindow
        Private _previewGeneration As Long
        Private _textNavigationGeneration As Long
        Private _currentPreviewOperation As PreviewOperation
        Private ReadOnly _previewOperations As New List(Of PreviewOperation)()

        Friend ReadOnly Property PendingPreviewTask As Task
            Get
                Return If(_currentPreviewOperation Is Nothing, Task.CompletedTask, _currentPreviewOperation.Completion.Task)
            End Get
        End Property

        Private Sub InvalidatePreviewRequests()
            Interlocked.Increment(_previewGeneration)
            Interlocked.Increment(_textNavigationGeneration)
            Interlocked.Increment(_webHighlightVersion)
            _currentPreviewOperation?.Cancellation.Cancel()
            _currentPreviewOperation = Nothing
            For index = _previewOperations.Count - 1 To 0 Step -1
                Dim operation = _previewOperations(index)
                If operation.Completion.Task.IsCompleted Then
                    operation.Cancellation.Dispose()
                    _previewOperations.RemoveAt(index)
                End If
            Next
        End Sub

        Private Function BeginPreviewOperation() As PreviewOperation
            InvalidatePreviewRequests()
            Dim operation As New PreviewOperation(Volatile.Read(_previewGeneration), TryCast(Results_lst.SelectedItem, SearchHit), _activeQuery)
            _previewOperations.Add(operation)
            _currentPreviewOperation = operation
            Return operation
        End Function

        Private Function IsPreviewSelectionCurrent(generation As Long, selected As Object) As Boolean
            Return Not (_isClosing OrElse _isResetting) AndAlso generation = Volatile.Read(_previewGeneration) AndAlso Results_lst.SelectedItem Is selected
        End Function

        Private Function IsPreviewCurrent(operation As PreviewOperation) As Boolean
            Return Not (_isClosing OrElse _isResetting OrElse operation.Cancellation.IsCancellationRequested) AndAlso
                operation.Generation = Volatile.Read(_previewGeneration) AndAlso Results_lst.SelectedItem Is operation.Hit AndAlso _activeQuery Is operation.Query
        End Function

        Private Async Function DrainPreviewOperationsAsync() As Task
            InvalidatePreviewRequests()
            Dim operations = _previewOperations.ToArray()
            _previewOperations.Clear()
            For Each operation In operations
                operation.Cancellation.Cancel()
            Next
            Try
                Await Task.WhenAll(operations.Select(Function(operation) operation.Completion.Task)).ConfigureAwait(False)
            Finally
                For Each operation In operations
                    operation.Cancellation.Dispose()
                Next
            End Try
        End Function

        Private NotInheritable Class PreviewOperation
            Public ReadOnly Generation As Long
            Public ReadOnly Hit As SearchHit
            Public ReadOnly Query As SearchQuery
            Public ReadOnly Cancellation As New CancellationTokenSource()
            Public ReadOnly Completion As New TaskCompletionSource(Of Boolean)(TaskCreationOptions.RunContinuationsAsynchronously)
            Public Sub New(generation As Long, hit As SearchHit, query As SearchQuery)
                Me.Generation = generation
                Me.Hit = hit
                Me.Query = query
            End Sub
        End Class
    End Class
End Namespace
