Imports System.IO
Imports System.Text.Json
Imports System.Threading
Imports System.Threading.Tasks

Namespace Beacon
    Public NotInheritable Class HarSearchService
        Private ReadOnly _options As BeaconSettings
        Private ReadOnly _query As SearchQuery

        Public Sub New(query As SearchQuery, options As BeaconSettings)
            _query = query
            _options = BeaconSettingsService.Clone(options)
        End Sub

        Public Async Function CollectAsync(stream As Stream, result As StructuredSearchResult(Of HarRecord), token As CancellationToken,
                                            Optional report As Action(Of Exception, String, String) = Nothing) As Task
            token.ThrowIfCancellationRequested()
            Dim filter = HarFilter.FromSettings(_options)
            Dim incomplete As Boolean
            Using document = Await JsonDocument.ParseAsync(stream, cancellationToken:=token).ConfigureAwait(False)
                Dim log As JsonElement
                Dim entries As JsonElement
                If document.RootElement.ValueKind <> JsonValueKind.Object OrElse
                   Not document.RootElement.TryGetProperty("log", log) OrElse log.ValueKind <> JsonValueKind.Object OrElse
                   Not log.TryGetProperty("entries", entries) OrElse entries.ValueKind <> JsonValueKind.Array Then
                    Throw New InvalidDataException("HAR content must contain a log.entries array.")
                End If
                Dim ordinal As Integer = 0
                For Each entry In entries.EnumerateArray()
                    token.ThrowIfCancellationRequested()
                    ordinal += 1
                    Dim request As HarRecord
                    Try
                        request = HarReader.Read(entry, _options, token)
                    Catch ex As InvalidDataException
                        report?.Invoke(New InvalidDataException($"HAR request {ordinal}: invalid request structure."), "HAR parsing", "Warning")
                        incomplete = True
                        Continue For
                    End Try
                    If Not filter.Matches(request) Then Continue For
                    If request.BodyIncomplete Then
                        incomplete = True
                        report?.Invoke(New InvalidDataException(request.BodyNotice), "HAR body limits", "Warning")
                    End If
                    If Not _query.IsMatch(request.SearchText(), token) Then Continue For
                    request = HarRedaction.Present(request, _options.RedactSensitiveHarData, token)
                    Dim detail = HarRedaction.CaptureDetail(_query, request, $"Request {ordinal:N0} · {request.Method} {request.StatusCode}", result.Records.Count, token)
                    result.Details.Add(detail)
                    result.Records.Add(request)
                    If _options.StopAfterFirstMatchPerFile OrElse
                       result.Records.Count >= Math.Min(_options.MaximumStructuredMatches, _options.HarMaximumMatches) Then
                        result.PartialReason = If(_options.StopAfterFirstMatchPerFile, "first matching request only", "request limit reached")
                        Exit For
                    End If
                Next
            End Using
            If incomplete Then result.PartialReason = (result.PartialReason & "; HAR bodies or malformed requests omitted").Trim(";"c, " "c)
        End Function
    End Class
End Namespace
