Imports System.Diagnostics.Eventing.Reader
Imports System.IO
Imports System.Threading

Namespace Beacon
    Public NotInheritable Class EvtxSearchService
        Private ReadOnly _options As BeaconSettings
        Private ReadOnly _query As SearchQuery
        Private ReadOnly _filter As EvtxFilter

        Public Sub New(query As SearchQuery, options As BeaconSettings)
            _query = query
            _options = BeaconSettingsService.Clone(options)
            _filter = EvtxFilter.FromSettings(_options)
        End Sub

        Public Sub Collect(filePath As String, result As StructuredSearchResult(Of EventRecordSummary), token As CancellationToken,
                           Optional report As Action(Of Exception, String, String) = Nothing)
            token.ThrowIfCancellationRequested()
            Dim missingProviders As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
            Using reader As New EventLogReader(filePath, PathType.FilePath)
                While True
                    token.ThrowIfCancellationRequested()
                    Using record = reader.ReadEvent()
                        If record Is Nothing Then Exit While
                        Dim provider = ReadField(Function() record.ProviderName, "Unknown")
                        Dim timestamp = record.TimeCreated
                        Dim levelNumber = record.Level
                        If Not _filter.Matches(record.Id, provider, levelNumber, timestamp) Then Continue While
                        Dim level = ReadField(Function() record.LevelDisplayName, "Unknown")
                        Dim message = ReadField(Function() record.FormatDescription(), "")
                        If String.IsNullOrWhiteSpace(message) AndAlso missingProviders.Add(provider) Then
                            report?.Invoke(New InvalidDataException($"Message text unavailable for provider '{provider}'. XML remains searchable. {EvtxFilter.ResourceGuidance}"), "EVTX message resources", "Warning")
                        End If
                        If CollectRecord(New EventRecordSummary With {.Provider = provider, .Level = level, .EventId = record.Id,
                            .TimeCreated = timestamp, .LevelNumber = levelNumber, .Message = message, .RawXml = record.ToXml()}, result, token) Then Exit While
                    End Using
                End While
            End Using
        End Sub

        Public Function CollectRecord(record As EventRecordSummary, result As StructuredSearchResult(Of EventRecordSummary), token As CancellationToken) As Boolean
            token.ThrowIfCancellationRequested()
            If Not _filter.Matches(record.EventId, record.Provider, record.LevelNumber, record.TimeCreated) Then Return False
            Dim message = If(record.Message, "")
            Dim xml = If(record.RawXml, "")
            Dim unavailable = String.IsNullOrWhiteSpace(message)
            Dim timestamp = record.TimeCreated
            Dim location = $"Event {record.EventId} · {If(timestamp.HasValue, timestamp.Value.ToString("g"), "no timestamp")}"
            Dim content = String.Join(vbLf, {"Event ID " & record.EventId.ToString(), record.Provider, record.Level,
                                          If(timestamp.HasValue, timestamp.Value.ToString("O"), ""), message})
            Dim detail = SearchDetail.FromRecord(_query, content, location, recordIndex:=result.Records.Count, token:=token)
            If detail Is Nothing Then
                detail = SearchDetail.FromRecord(_query, xml, location & " (XML)", recordIndex:=result.Records.Count, token:=token)
                If detail IsNot Nothing Then
                    Dim readableXml = PreviewFormatting.FormatXml(xml, True, CLng(_options.MaximumPreviewSizeMb) * 1024 * 1024)
                    Dim xmlSpans = _query.FindHighlights(readableXml, token:=token)
                    If xmlSpans.Count > 0 Then detail.Context = MatchContext.FromText(readableXml, xmlSpans, token)
                    detail.ContextKind = "Event XML lines"
                    message &= vbCrLf & vbCrLf & "Event XML:" & vbCrLf & readableXml
                End If
            End If
            If detail Is Nothing Then Return False
            If detail.ContextKind = "Record text" Then detail.ContextKind = "Event text lines"
            If String.IsNullOrWhiteSpace(message) Then
                message = "Message resources are unavailable. Use Raw XML or the offline resource guidance below."
                If _options.ShowRawXmlWhenMessageUnavailable Then
                    message &= vbCrLf & PreviewFormatting.FormatXml(xml, True, CLng(_options.MaximumPreviewSizeMb) * 1024 * 1024)
                End If
            End If
            Dim xmlLimit = CInt(Math.Min(65536L, CLng(_options.MaximumPreviewSizeMb) * 1024 * 1024))
            Dim capturedXml = If(xml.Length > xmlLimit, xml.Substring(0, xmlLimit), xml)
            result.Records.Add(New EventRecordSummary With {.Provider = record.Provider, .Level = record.Level,
                .EventId = record.EventId, .TimeCreated = timestamp, .Message = message, .LevelNumber = record.LevelNumber,
                .RawXml = capturedXml, .XmlShortened = xml.Length > xmlLimit, .MessageUnavailable = unavailable})
            result.Details.Add(detail)
            If _options.StopAfterFirstMatchPerFile OrElse result.Records.Count >= Math.Min(_options.MaximumStructuredMatches, _options.EvtxMaximumMatches) Then
                result.PartialReason = If(_options.StopAfterFirstMatchPerFile, "first matching event only", "event limit reached")
                Return True
            End If
            Return False
        End Function

        Private Shared Function ReadField(read As Func(Of String), fallback As String) As String
            Try
                Return If(read(), fallback)
            Catch ex As EventLogException
                Return fallback
            End Try
        End Function
    End Class
End Namespace
