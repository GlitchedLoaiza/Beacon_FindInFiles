Imports System.Threading
Imports System.Threading.Tasks

Namespace Beacon
    Friend NotInheritable Class SearchResourceGate
        Implements IDisposable
        Private ReadOnly _policy As SearchConcurrencyPolicy
        Private ReadOnly _head As Func(Of Long)
        Private ReadOnly _background As New SemaphoreSlim(1, 1)

        Public Sub New(policy As SearchConcurrencyPolicy, head As Func(Of Long))
            _policy = policy
            _head = head
        End Sub

        Public Async Function EnterAsync(ordinal As Long, estimatedBytes As Long, token As CancellationToken) As Task(Of IDisposable)
            While ordinal <> _head()
                token.ThrowIfCancellationRequested()
                If _policy.CanStartHeavy(estimatedBytes) Then
                    If Await _background.WaitAsync(50, token).ConfigureAwait(False) Then
                        If ordinal = _head() Then
                            _background.Release()
                            Return Nothing
                        End If
                        If _policy.CanStartHeavy(estimatedBytes) Then Return New SearchResourceLease(Sub() _background.Release())
                        _background.Release()
                    End If
                Else
                    Await Task.Delay(50, token).ConfigureAwait(False)
                End If
            End While
            ' The ordered head has a reserved lane; speculative work cannot deadlock it.
            token.ThrowIfCancellationRequested()
            Return Nothing
        End Function

        Public Sub Dispose() Implements IDisposable.Dispose
            _background.Dispose()
        End Sub
    End Class

    Friend NotInheritable Class SearchResourceLease
        Implements IDisposable
        Private _release As Action
        Public Sub New(release As Action)
            _release = release
        End Sub
        Public Sub Dispose() Implements IDisposable.Dispose
            Dim release = Interlocked.Exchange(_release, Nothing)
            release?.Invoke()
        End Sub
    End Class
End Namespace
