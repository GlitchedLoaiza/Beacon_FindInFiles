Imports System.IO

Namespace Beacon
    Public NotInheritable Class ScanDiagnostic
        Public Property TimestampUtc As DateTimeOffset
        Public Property Severity As String
        Public Property Category As String
        Public Property Stage As String
        Public Property SourcePath As String
        Public Property Message As String
        Public Property ExceptionType As String
        Public Property ErrorCode As String
        Public Property TechnicalDetails As String
        Public Property Occurrences As Long = 1
        Public Property TextTruncated As Boolean
        Public ReadOnly Property Heading As String
            Get
                Return $"{Severity} · {Category} · {Stage}"
            End Get
        End Property
        <System.Text.Json.Serialization.JsonIgnore>
        Public ReadOnly Property TechnicalText As String
            Get
                Return $"{ExceptionType} ({ErrorCode})" & Environment.NewLine & TechnicalDetails
            End Get
        End Property
        Public Function Copy(Optional includeTechnical As Boolean = True) As ScanDiagnostic
            Dim clone = DirectCast(MemberwiseClone(), ScanDiagnostic)
            If Not includeTechnical Then
                clone.ExceptionType = Nothing
                clone.ErrorCode = Nothing
                clone.TechnicalDetails = Nothing
            End If
            Return clone
        End Function
    End Class

    Public NotInheritable Class DiagnosticSnapshot
        Public Property Items As New List(Of ScanDiagnostic)()
        Public Property TotalNotifications As Long
        Public Property OmittedNotifications As Long
    End Class

    Public NotInheritable Class ScanDiagnosticStore
        Public Const DefaultCapacity As Integer = 10000
        Private ReadOnly _gate As New Object()
        Private ReadOnly _capacity As Integer
        Private ReadOnly _items As New Dictionary(Of Tuple(Of String, String, String, String), ScanDiagnostic)()
        Private _total As Long
        Private _omitted As Long

        Public Sub New(Optional capacity As Integer = DefaultCapacity)
            If capacity < 1 OrElse capacity > DefaultCapacity Then Throw New ArgumentOutOfRangeException(NameOf(capacity))
            _capacity = capacity
        End Sub

        Public Sub Record(sourcePath As String, ex As Exception, stage As String, detailed As Boolean,
                          Optional severity As String = Nothing)
            If ex Is Nothing Then Throw New ArgumentNullException(NameOf(ex))
            Dim entry As New ScanDiagnostic With {
                .TimestampUtc = DateTimeOffset.UtcNow,
                .Severity = If(severity, If(TypeOf ex Is UnauthorizedAccessException OrElse TypeOf ex Is IOException OrElse TypeOf ex Is InvalidDataException, "Warning", "Error")),
                .Category = CategoryFor(ex), .Stage = Limit(stage, 64), .SourcePath = Limit(sourcePath, 4096),
                .Message = Limit(ex.Message, 2000), .ExceptionType = Limit(ex.GetType().FullName, 256),
                .ErrorCode = $"0x{ex.HResult:X8}", .TechnicalDetails = If(detailed, Limit(ex.StackTrace, 8000), Nothing)
            }
            entry.TextTruncated = entry.SourcePath <> If(sourcePath, "") OrElse entry.Message <> ex.Message OrElse
                                  (detailed AndAlso entry.TechnicalDetails <> If(ex.StackTrace, ""))
            Dim key = Tuple.Create(entry.Stage, entry.SourcePath, entry.ExceptionType, entry.Message)
            SyncLock _gate
                _total += 1
                Dim previous As ScanDiagnostic = Nothing
                If _items.TryGetValue(key, previous) Then
                    previous.Occurrences += 1
                    Return
                End If
                If _items.Count >= _capacity Then
                    _omitted += 1
                    Return
                End If
                _items.Add(key, entry)
            End SyncLock
        End Sub

        Public ReadOnly Property Count As Integer
            Get
                SyncLock _gate
                    Return _items.Count
                End SyncLock
            End Get
        End Property

        Public ReadOnly Property OmittedCount As Long
            Get
                SyncLock _gate
                    Return _omitted
                End SyncLock
            End Get
        End Property

        Public Function Snapshot() As DiagnosticSnapshot
            SyncLock _gate
                Return New DiagnosticSnapshot With {
                    .Items = _items.Values.Select(Function(item) item.Copy()).OrderBy(Function(item) item.TimestampUtc).ToList(),
                    .TotalNotifications = _total, .OmittedNotifications = _omitted
                }
            End SyncLock
        End Function

        Public Sub Clear()
            SyncLock _gate
                _items.Clear()
                _total = 0
                _omitted = 0
            End SyncLock
        End Sub

        Private Shared Function Limit(value As String, length As Integer) As String
            value = If(value, "")
            Return If(value.Length <= length, value, value.Substring(0, length - 12) & " [truncated]")
        End Function

        Private Shared Function CategoryFor(ex As Exception) As String
            If TypeOf ex Is UnauthorizedAccessException Then Return "Access denied"
            If TypeOf ex Is FileNotFoundException OrElse TypeOf ex Is DirectoryNotFoundException Then Return "Missing source"
            If TypeOf ex Is System.Text.RegularExpressions.RegexMatchTimeoutException Then Return "Query timeout"
            If TypeOf ex Is TimeoutException Then Return "Timeout"
            If TypeOf ex Is System.Text.Json.JsonException Then Return "Invalid JSON"
            If TypeOf ex Is InvalidDataException Then Return "Invalid or limited data"
            If TypeOf ex Is NotSupportedException Then Return "Unsupported source"
            If TypeOf ex Is IOException Then Return "File access"
            Return "Processing error"
        End Function
    End Class
End Namespace
