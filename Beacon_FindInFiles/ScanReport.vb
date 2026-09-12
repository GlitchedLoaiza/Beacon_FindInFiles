Namespace Beacon
    Public Enum ScanReportState
        Ready
        Running
        Completed
        Cancelled
        Failed
        ResultLimitReached
    End Enum

    Public NotInheritable Class ScanRunInfo
        Public Property SourceRoot As String
        Public Property QueryText As String
        Public Property Mode As SearchMode
        Public Property CaseSensitive As Boolean
        Public Property Options As BeaconSettings
        Public Property StartedUtc As DateTimeOffset?
        Public Property CompletedUtc As DateTimeOffset?
        Public Property State As ScanReportState
        Public Property FailureMessage As String

        Public Function Copy() As ScanRunInfo
            Dim result = DirectCast(MemberwiseClone(), ScanRunInfo)
            If Options IsNot Nothing Then result.Options = BeaconSettingsService.Clone(Options)
            Return result
        End Function
    End Class

    Public NotInheritable Class ReportMatch
        Public Property Location As String
        Public Property LineNumber As Integer
        Public Property RecordIndex As Integer
        Public Property IsMetadata As Boolean
        Public Property VisibleMatches As Integer
        Public Property Excerpt As String
        Public Property ContextKind As String
        Public Property Context As New List(Of ContextLine)()

        Public Function Copy(includeExcerpt As Boolean) As ReportMatch
            Dim result = DirectCast(MemberwiseClone(), ReportMatch)
            result.Context = If(includeExcerpt, Context.Select(Function(line) line.Copy()).ToList(), New List(Of ContextLine)())
            If Not includeExcerpt Then result.Excerpt = Nothing
            Return result
        End Function
    End Class

    Public NotInheritable Class ReportFile
        Public Property DisplayName As String
        Public Property SourcePath As String
        Public Property FileType As String
        Public Property PartialReason As String
        Public Property HarRedactionApplied As Boolean
        Public Property HarRedactionEnabled As Boolean
        Public Property Matches As New List(Of ReportMatch)()
        Public ReadOnly Property StoredDetailCount As Integer
            Get
                Return Matches.Count
            End Get
        End Property
    End Class

    Public NotInheritable Class ScanReportSnapshot
        Public Property SchemaVersion As Integer = 1
        Public Property ApplicationName As String = "Beacon"
        Public Property ApplicationVersion As String
        Public Property CapturedUtc As DateTimeOffset = DateTimeOffset.UtcNow
        Public Property Run As New ScanRunInfo With {.State = ScanReportState.Ready}
        Public Property ElapsedMilliseconds As Long
        Public Property FilesScanned As Integer
        Public Property EstimatedTotalFiles As Integer
        Public Property Results As New List(Of ReportFile)()
        Public Property Diagnostics As New DiagnosticSnapshot()
        Public Property IncludesExcerpts As Boolean
        Public Property IncludesTechnicalDetails As Boolean
        Public ReadOnly Property RedactionApplied As Boolean
            Get
                Return Results.Any(Function(file) file.HarRedactionApplied)
            End Get
        End Property
        Public ReadOnly Property IsComplete As Boolean
            Get
                Return Run.State = ScanReportState.Completed AndAlso
                       Results.All(Function(item) String.IsNullOrEmpty(item.PartialReason)) AndAlso
                       Diagnostics.OmittedNotifications = 0 AndAlso
                       Not Diagnostics.Items.Any(Function(item) item.Severity <> "Information")
            End Get
        End Property
        Public ReadOnly Property Coverage As String
            Get
                Dim reasons As New List(Of String)()
                If Run.State <> ScanReportState.Completed Then reasons.Add(Run.State.ToString())
                If Results.Any(Function(item) Not String.IsNullOrEmpty(item.PartialReason)) Then reasons.Add("per-file results are partial")
                If Diagnostics.Items.Any(Function(item) item.Severity <> "Information") Then reasons.Add("scan/preview issues were reported")
                If Diagnostics.OmittedNotifications > 0 Then reasons.Add("some diagnostics were omitted")
                Return If(reasons.Count = 0, "Complete within the configured search scope", "Incomplete snapshot: " & String.Join("; ", reasons))
            End Get
        End Property

        Public Function ForExport(includeExcerpts As Boolean, includeTechnical As Boolean) As ScanReportSnapshot
            Return New ScanReportSnapshot With {
                .SchemaVersion = SchemaVersion, .ApplicationName = ApplicationName, .ApplicationVersion = ApplicationVersion,
                .CapturedUtc = CapturedUtc, .Run = Run.Copy(), .ElapsedMilliseconds = ElapsedMilliseconds,
                .FilesScanned = FilesScanned, .EstimatedTotalFiles = EstimatedTotalFiles,
                .IncludesExcerpts = includeExcerpts, .IncludesTechnicalDetails = includeTechnical,
                .Results = Results.Select(Function(file) New ReportFile With {
                    .DisplayName = file.DisplayName, .SourcePath = file.SourcePath, .FileType = file.FileType,
                    .HarRedactionApplied = file.HarRedactionApplied, .HarRedactionEnabled = file.HarRedactionEnabled,
                    .PartialReason = file.PartialReason, .Matches = file.Matches.Select(Function(item) item.Copy(includeExcerpts)).ToList()
                }).ToList(),
                .Diagnostics = New DiagnosticSnapshot With {
                    .TotalNotifications = Diagnostics.TotalNotifications, .OmittedNotifications = Diagnostics.OmittedNotifications,
                    .Items = Diagnostics.Items.Select(Function(item) item.Copy(includeTechnical)).ToList()
                }
            }
        End Function
    End Class
End Namespace
