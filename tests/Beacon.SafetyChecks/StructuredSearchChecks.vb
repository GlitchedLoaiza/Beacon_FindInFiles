Imports System.IO
Imports System.Text
Imports System.Threading
Imports Beacon

Module StructuredSearchChecks
    Private Const HarEntry As String = "{""request"":{""method"":""GET"",""url"":""https://example.invalid/""},""response"":{""status"":503,""content"":{""mimeType"":""application/json"",""text"":""{\""token\"":\""PRIVATE_SECRET\"",\""safe\"":\""error\""}""}}}"

    Public Sub Har()
        Dim options As New BeaconSettings With {.StopAfterFirstMatchPerFile = False, .MaximumStructuredMatches = 2, .HarMaximumMatches = 2, .RedactSensitiveHarData = True}
        Dim service As New HarSearchService(New SearchQuery("PRIVATE_SECRET", SearchMode.PlainText, False), options)
        options.RedactSensitiveHarData = False
        options.HarStatusCodes = "200"
        Dim result As New StructuredSearchResult(Of HarRecord)()
        Dim issues As New List(Of String)()
        Using stream As New MemoryStream(Encoding.UTF8.GetBytes("{""log"":{""entries"":[{}," & HarEntry & "," & HarEntry & "," & HarEntry & "]}}"))
            service.CollectAsync(stream, result, CancellationToken.None, Sub(ex, stage, severity) issues.Add(stage)).GetAwaiter().GetResult()
            Require(stream.CanRead, "HAR service disposed the caller-owned stream.")
        End Using
        Require(result.Records.Count = 2 AndAlso result.Details.Count = 2, "HAR record limits or settings snapshot failed.")
        Require(result.Details.Select(Function(detail) detail.RecordIndex).SequenceEqual({0, 1}), "HAR detail indexes are not aligned.")
        Require(result.Details(0).Location.StartsWith("Request 2"), "Malformed entries changed source ordinals.")
        Require(result.PartialReason.Contains("request limit") AndAlso result.PartialReason.Contains("malformed"), "HAR partial coverage was lost.")
        Require(issues.SequenceEqual({"HAR parsing"}), "HAR diagnostics changed.")
        Require(result.Records.All(Function(record) record.RedactionEnabled AndAlso Not record.SearchText().Contains("PRIVATE_SECRET")), "HAR service did not snapshot redaction settings.")
        Require(result.Details.All(Function(detail) Not detail.Excerpt.Contains("PRIVATE_SECRET") AndAlso detail.VisibleMatches = 0), "Sensitive context was exposed.")
        Dim second As New StructuredSearchResult(Of HarRecord)()
        Using stream As New MemoryStream(Encoding.UTF8.GetBytes("{""log"":{""entries"":[" & HarEntry & ",{}]}}"))
            Reject(Of IOException)(Sub() service.CollectAsync(stream, second, CancellationToken.None,
                Sub(ex, stage, severity)
                    Throw New IOException("Simulated later reporting failure")
                End Sub).GetAwaiter().GetResult())
        End Using
        Require(second.Records.Count = 1 AndAlso second.Details.Count = 1, "Later failure discarded prior HAR results.")
        Using cancelled As New CancellationTokenSource(), stream As New MemoryStream()
            cancelled.Cancel()
            Reject(Of OperationCanceledException)(Sub() service.CollectAsync(stream, New StructuredSearchResult(Of HarRecord)(), cancelled.Token).GetAwaiter().GetResult())
        End Using
        Using stream As New MemoryStream(Encoding.UTF8.GetBytes("{}"))
            Reject(Of InvalidDataException)(Sub() service.CollectAsync(stream, New StructuredSearchResult(Of HarRecord)(), CancellationToken.None).GetAwaiter().GetResult())
        End Using
    End Sub

    Public Sub Events()
        Dim options As New BeaconSettings With {.StopAfterFirstMatchPerFile = False, .EvtxEventIds = "100", .EvtxMaximumMatches = 2, .ShowRawXmlWhenMessageUnavailable = False}
        Dim service As New EvtxSearchService(New SearchQuery("error", SearchMode.PlainText, False), options)
        options.EvtxEventIds = "999"
        options.EvtxMaximumMatches = 1
        Dim result As New StructuredSearchResult(Of EventRecordSummary)()
        Dim source As New EventRecordSummary With {.EventId = 100, .Provider = "provider", .Message = "error", .RawXml = "<Event>safe</Event>"}
        Require(Not service.CollectRecord(source, result, CancellationToken.None), "EVTX settings snapshot failed.")
        source.Message = "changed"
        Require(result.Records(0).Message = "error", "Captured EVTX record aliases its source.")
        Require(service.CollectRecord(New EventRecordSummary With {.EventId = 100, .Message = "", .RawXml = "<Event>error</Event>"}, result, CancellationToken.None), "EVTX limit was not reached.")
        Require(result.Details(1).ContextKind = "Event XML lines" AndAlso result.Records(1).MessageUnavailable, "XML matching/fallback changed.")
        Require(result.Details.Select(Function(detail) detail.RecordIndex).SequenceEqual({0, 1}) AndAlso result.PartialReason = "event limit reached", "EVTX indexes/partial state changed.")
        Dim unmatched As New StructuredSearchResult(Of EventRecordSummary)()
        service.CollectRecord(New EventRecordSummary With {.EventId = 999, .Message = "error"}, unmatched, CancellationToken.None)
        Require(unmatched.Records.Count = 0, "EVTX filter was not preserved.")
        Dim fallback As New EvtxSearchService(New SearchQuery("provider", SearchMode.PlainText, False), New BeaconSettings With {.ShowRawXmlWhenMessageUnavailable = False})
        Dim missing As New StructuredSearchResult(Of EventRecordSummary)()
        fallback.CollectRecord(New EventRecordSummary With {.Provider = "provider", .RawXml = "<Event>xml payload</Event>"}, missing, CancellationToken.None)
        Require(Not missing.Records(0).Message.Contains("xml payload") AndAlso missing.Records(0).RawXml.Contains("xml payload"), "Explicit XML/fallback preference changed.")
        Using cancelled As New CancellationTokenSource()
            cancelled.Cancel()
            Reject(Of OperationCanceledException)(Sub() service.CollectRecord(source, result, cancelled.Token))
            Reject(Of OperationCanceledException)(Sub() service.Collect("does-not-exist.evtx", result, cancelled.Token))
        End Using
        Require(result.Records.Count = 2, "Cancellation changed accumulated EVTX records.")
    End Sub

    Private Sub Reject(Of T As Exception)(action As Action)
        Try
            action()
        Catch ex As T
            Return
        End Try
        Throw New InvalidOperationException("Expected " & GetType(T).Name)
    End Sub
    Private Sub Require(condition As Boolean, message As String)
        If Not condition Then Throw New InvalidOperationException(message)
    End Sub
End Module
