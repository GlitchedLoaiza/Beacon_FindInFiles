Imports System.IO
Imports System.Threading
Imports System.Threading.Tasks

Namespace Beacon
    Friend Enum SourceSearchEventKind
        Started
        Completed
        Result
        Issue
        Remember
    End Enum

    Friend NotInheritable Class SourceSearchEvent
        Public Property Kind As SourceSearchEventKind
        Public Property Path As String
        Public Property PhysicalPath As String
        Public Property Result As SourceSearchResult
        Public Property [Error] As Exception
        Public Property Stage As String
        Public Property Severity As String
        Public Property OutputLease As IDisposable
    End Class

    Friend NotInheritable Class SourceSearchExecution
        Private Const PendingIssueLimit As Integer = 256
        Private ReadOnly _send As Func(Of SourceSearchEvent, CancellationToken, Task)
        Private ReadOnly _heavy As Func(Of Long, CancellationToken, Task(Of IDisposable))
        Private ReadOnly _pending As New List(Of SourceSearchEvent)()
        Private _issues As Integer
        Private _omitted As Long
        Private _heavyDepth As Integer
        Public ReadOnly Property Temporary As New SearchTemporarySources()

        Public Sub New(send As Func(Of SourceSearchEvent, CancellationToken, Task),
                       Optional heavy As Func(Of Long, CancellationToken, Task(Of IDisposable)) = Nothing)
            _send = send
            _heavy = heavy
        End Sub

        Public Sub Progress(path As String, processed As Boolean)
            If Not processed Then _pending.Add(New SourceSearchEvent With {.Kind = SourceSearchEventKind.Started, .Path = path})
        End Sub

        Public Sub Remember(physical As String, logical As String)
            If Not String.IsNullOrEmpty(physical) Then
                _pending.Add(New SourceSearchEvent With {.Kind = SourceSearchEventKind.Remember, .PhysicalPath = physical, .Path = logical})
            End If
        End Sub

        Public Sub Report(path As String, ex As Exception, stage As String, severity As String)
            If _issues >= PendingIssueLimit Then
                _omitted += 1
                Return
            End If
            _issues += 1
            _pending.Add(New SourceSearchEvent With {.Kind = SourceSearchEventKind.Issue, .Path = path,
                         .Error = ex, .Stage = stage, .Severity = severity})
        End Sub

        Public Async Function FlushAsync(token As CancellationToken) As Task
            For Each item In _pending
                Await _send(item, token).ConfigureAwait(False)
            Next
            _pending.Clear()
            _issues = 0
            If _omitted > 0 Then
                Dim omitted = _omitted
                _omitted = 0
                Await _send(New SourceSearchEvent With {.Kind = SourceSearchEventKind.Issue, .Path = "",
                    .Error = New InvalidDataException($"{omitted:N0} additional source diagnostics omitted from this bounded batch."),
                    .Stage = "Search diagnostics", .Severity = "Warning"}, token).ConfigureAwait(False)
            End If
        End Function

        Public Async Function PublishAsync(result As SourceSearchResult, token As CancellationToken) As Task
            Await FlushAsync(token).ConfigureAwait(False)
            Await _send(New SourceSearchEvent With {.Kind = SourceSearchEventKind.Result, .Result = result}, token).ConfigureAwait(False)
        End Function

        Public Async Function CompletedAsync(path As String, token As CancellationToken) As Task
            Await FlushAsync(token).ConfigureAwait(False)
            Await _send(New SourceSearchEvent With {.Kind = SourceSearchEventKind.Completed, .Path = path}, token).ConfigureAwait(False)
        End Function

        Public Async Function EnterHeavyAsync(bytes As Long, token As CancellationToken) As Task(Of IDisposable)
            If _heavy Is Nothing Then Return Nothing
            Dim lease As IDisposable = Nothing
            If _heavyDepth = 0 Then lease = Await _heavy(bytes, token).ConfigureAwait(False)
            _heavyDepth += 1
            Return New SearchResourceLease(Sub()
                                               _heavyDepth -= 1
                                               lease?.Dispose()
                                           End Sub)
        End Function
    End Class
End Namespace
