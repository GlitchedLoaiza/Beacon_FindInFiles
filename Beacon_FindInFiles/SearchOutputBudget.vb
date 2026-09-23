Imports System.Threading
Imports System.Threading.Tasks

Namespace Beacon
    Friend NotInheritable Class SearchOutputBudget
        Private Const LimitBytes As Long = 32L * 1024 * 1024
        Private ReadOnly _head As Func(Of Long)
        Private ReadOnly _policy As SearchConcurrencyPolicy
        Private ReadOnly _sync As New Object()
        Private _reserved As Long

        Public Sub New(policy As SearchConcurrencyPolicy, head As Func(Of Long))
            _policy = policy
            _head = head
        End Sub

        Public Async Function EnterAsync(ordinal As Long, result As SourceSearchResult, token As CancellationToken) As Task(Of IDisposable)
            Dim bytes = Estimate(result)
            While ordinal <> _head()
                token.ThrowIfCancellationRequested()
                If _policy.CanStartHeavy(bytes) Then
                    SyncLock _sync
                        If bytes <= LimitBytes - _reserved Then
                            _reserved += bytes
                            Return New SearchResourceLease(Sub()
                                                               SyncLock _sync
                                                                   _reserved -= bytes
                                                               End SyncLock
                                                           End Sub)
                        End If
                    End SyncLock
                End If
                Await Task.Delay(50, token).ConfigureAwait(False)
            End While
            token.ThrowIfCancellationRequested()
            Return Nothing
        End Function

        Private Shared Function Estimate(result As SourceSearchResult) As Long
            Dim bytes As Long = 512
            For Each detail In result.Details
                bytes += 256 + TextBytes(detail.Excerpt) + TextBytes(detail.Location)
                For Each line In detail.Context
                    bytes += 128 + TextBytes(line.Text) + CLng(line.Highlights.Count) * 32
                Next
            Next
            For Each record In result.Events
                bytes += 256 + TextBytes(record.Message) + TextBytes(record.RawXml)
            Next
            For Each record In result.Requests
                bytes += 512 + TextBytes(record.Url) + TextBytes(record.RequestHeaders) + TextBytes(record.ResponseHeaders) +
                    TextBytes(record.RequestBody) + TextBytes(record.ResponseBody)
            Next
            Return bytes
        End Function

        Private Shared Function TextBytes(value As String) As Long
            Return If(value Is Nothing, 0L, 24L + CLng(value.Length) * 2)
        End Function
    End Class
End Namespace
