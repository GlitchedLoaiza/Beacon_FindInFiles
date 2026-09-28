Imports System.Threading

Namespace Beacon
    Public NotInheritable Class EvtxSearchService
        Private ReadOnly _options As BeaconSettings
        Private ReadOnly _query As SearchQuery

        Public Sub New(query As SearchQuery, options As BeaconSettings)
            _query = query
            _options = BeaconSettingsService.Clone(options)
            EvtxFilter.FromSettings(_options)
        End Sub

        Public Sub Collect(filePath As String, result As StructuredSearchResult(Of EventRecordSummary), token As CancellationToken,
                           Optional report As Action(Of Exception, String, String) = Nothing,
                           Optional fastCopyLimitBytes As Long = -1)
            EvtxWorkerHost.CollectIsolated(filePath, _query, _options, result, token, report, fastCopyLimitBytes)
        End Sub

        Public Function CollectRecord(record As EventRecordSummary, result As StructuredSearchResult(Of EventRecordSummary), token As CancellationToken) As Boolean
            Return EvtxWorkerHost.CollectRecordCore(_query, _options, record, result, token)
        End Function
    End Class
End Namespace
