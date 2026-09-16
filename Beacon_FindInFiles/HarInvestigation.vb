Imports System.Globalization
Imports System.IO
Imports System.Text
Imports System.Text.Json
Imports System.Text.Json.Nodes
Imports System.Threading

Namespace Beacon
    Public Class HarRecord
        Public Property Method As String = ""
        Public Property Url As String = ""
        Public Property StatusCode As Integer
        Public Property StatusText As String = ""
        Public Property StartedDateTime As DateTime?
        Public Property Time As Double
        Public Property RequestHeaders As String = ""
        Public Property ResponseHeaders As String = ""
        Public Property RequestBody As String = ""
        Public Property ResponseBody As String = ""
        Public Property ServerIpAddress As String = ""
        Public Property RequestMime As String = ""
        Public Property MimeType As String = ""
        Public Property BodyNotice As String = ""
        Public Property BodyIncomplete As Boolean
        Public Property RedactionApplied As Boolean
        Public Property RedactionEnabled As Boolean

        Public Function SearchText() As String
            Return String.Join(vbLf, {Method, Url, StatusCode.ToString(CultureInfo.InvariantCulture), StatusText, MimeType,
                ServerIpAddress, RequestHeaders, ResponseHeaders, RequestBody, ResponseBody})
        End Function
        Public Function Copy() As HarRecord
            Return DirectCast(MemberwiseClone(), HarRecord)
        End Function
    End Class

    Public NotInheritable Class HarFilter
        Private ReadOnly _method, _host, _mime As String
        Private ReadOnly _statuses As New HashSet(Of Integer)()
        Private ReadOnly _minimum As Double?
        Private ReadOnly _timeFilter As EvtxFilter

        Public Sub New(method As String, host As String, statuses As String, mime As String, minimum As String, fromUtc As String, toUtc As String)
            _method = If(method, "").Trim()
            _host = If(host, "").Trim()
            _mime = If(mime, "").Trim()
            If _method.Length > 32 OrElse _host.Length > 256 OrElse _mime.Length > 256 OrElse If(statuses, "").Length > 2048 Then
                Throw New ArgumentException("HAR filter text exceeds the supported length.")
            End If
            If Not String.IsNullOrWhiteSpace(statuses) Then
                For Each part In statuses.Split(","c)
                    Dim value As Integer
                    If Not Integer.TryParse(part.Trim(), NumberStyles.None, CultureInfo.InvariantCulture, value) OrElse value < 0 OrElse value > 599 Then
                        Throw New ArgumentException("HAR status codes must be comma-separated numbers from 0 to 599.")
                    End If
                    _statuses.Add(value)
                Next
            End If
            If Not String.IsNullOrWhiteSpace(minimum) Then
                Dim value As Double
                If Not Double.TryParse(minimum, NumberStyles.AllowDecimalPoint, CultureInfo.InvariantCulture, value) OrElse Not Double.IsFinite(value) OrElse value < 0 Then
                    Throw New ArgumentException("Minimum HAR duration must be a nonnegative number of milliseconds (use a decimal point).")
                End If
                _minimum = value
            End If
            Try
                _timeFilter = New EvtxFilter("", "", "", fromUtc, toUtc)
            Catch ex As ArgumentException
                Throw New ArgumentException(ex.Message.Replace("EVTX", "HAR"))
            End Try
        End Sub
        Public Shared Function FromSettings(settings As BeaconSettings) As HarFilter
            Return New HarFilter(settings.HarMethod, settings.HarHost, settings.HarStatusCodes, settings.HarMimeType,
                                 settings.HarMinimumDuration, settings.HarFromUtc, settings.HarToUtc)
        End Function
        Public Function Matches(record As HarRecord) As Boolean
            If _method.Length > 0 AndAlso Not String.Equals(_method, record.Method, StringComparison.OrdinalIgnoreCase) Then Return False
            If _host.Length > 0 Then
                Dim uri As Uri = Nothing
                If Not Uri.TryCreate(record.Url, UriKind.Absolute, uri) OrElse Not String.Equals(uri.Host, _host, StringComparison.OrdinalIgnoreCase) Then Return False
            End If
            If _statuses.Count > 0 AndAlso Not _statuses.Contains(record.StatusCode) Then Return False
            If _mime.Length > 0 AndAlso Not record.MimeType.Contains(_mime, StringComparison.OrdinalIgnoreCase) Then Return False
            If _minimum.HasValue AndAlso (Not Double.IsFinite(record.Time) OrElse record.Time < _minimum.Value) Then Return False
            Return _timeFilter.Matches(0, "", Nothing, record.StartedDateTime)
        End Function
        Public Shared Function Describe(settings As BeaconSettings) As String
            If settings Is Nothing Then Return "No scan options captured"
            Return $"Method: {settings.HarMethod}; Host: {settings.HarHost}; Status: {settings.HarStatusCodes}; MIME: {settings.HarMimeType}; Minimum ms: {settings.HarMinimumDuration}; UTC from: {settings.HarFromUtc}; UTC to: {settings.HarToUtc} (blank = any)"
        End Function
    End Class

    Public NotInheritable Class HarReader
        Private Sub New()
        End Sub
        Private Shared Function Field(element As JsonElement, name As String) As JsonElement
            Dim value As JsonElement
            If element.ValueKind = JsonValueKind.Object Then element.TryGetProperty(name, value)
            Return value
        End Function
        Private Shared Function Text(element As JsonElement, name As String) As String
            Dim value = Field(element, name)
            Return If(value.ValueKind = JsonValueKind.String, value.GetString(), "")
        End Function
        Public Shared Function Read(entry As JsonElement, settings As BeaconSettings, token As CancellationToken) As HarRecord
            token.ThrowIfCancellationRequested()
            Dim request = Field(entry, "request"), response = Field(entry, "response")
            If request.ValueKind <> JsonValueKind.Object OrElse response.ValueKind <> JsonValueKind.Object Then
                Throw New InvalidDataException("HAR entry must contain request and response objects.")
            End If
            Dim result As New HarRecord With {.Method = Text(request, "method"), .Url = Text(request, "url"),
                .StatusText = Text(response, "statusText"), .ServerIpAddress = Text(entry, "serverIPAddress")}
            Integer.TryParse(Field(response, "status").ToString(), NumberStyles.Integer, CultureInfo.InvariantCulture, result.StatusCode)
            Dim duration As Double
            result.Time = If(Double.TryParse(Field(entry, "time").ToString(), NumberStyles.Float, CultureInfo.InvariantCulture, duration) AndAlso Double.IsFinite(duration), duration, -1)
            Dim started As DateTimeOffset
            If DateTimeOffset.TryParse(Text(entry, "startedDateTime"), CultureInfo.InvariantCulture, DateTimeStyles.AssumeUniversal, started) Then result.StartedDateTime = started.UtcDateTime
            result.RequestHeaders = Headers(Field(request, "headers"), token)
            result.ResponseHeaders = Headers(Field(response, "headers"), token)
            Dim post = Field(request, "postData"), content = Field(response, "content")
            result.RequestMime = Text(post, "mimeType")
            result.MimeType = Text(content, "mimeType")
            Dim limit = CLng(settings.MaximumHarBodySizeMb) * 1024 * 1024
            result.RequestBody = Body(post, limit, settings.DecodeBase64HarBodies, result, "Request", token)
            result.ResponseBody = Body(content, limit, settings.DecodeBase64HarBodies, result, "Response", token)
            If result.RequestBody.Length = 0 AndAlso Field(post, "params").ValueKind = JsonValueKind.Array Then
                Dim form As New StringBuilder()
                For Each param In Field(post, "params").EnumerateArray()
                    token.ThrowIfCancellationRequested()
                    Dim pair = Uri.EscapeDataString(Text(param, "name")) & "=" & Uri.EscapeDataString(Text(param, "value"))
                    If form.Length + pair.Length + 1 > limit Then
                        Notice(result, "Request form exceeds body limit.")
                        Exit For
                    End If
                    If form.Length > 0 Then form.Append("&")
                    form.Append(pair)
                Next
                result.RequestBody = form.ToString()
                result.RequestMime = "application/x-www-form-urlencoded"
            End If
            Return result
        End Function
        Private Shared Function Headers(items As JsonElement, token As CancellationToken) As String
            If items.ValueKind <> JsonValueKind.Array Then Return ""
            Dim builder As New StringBuilder()
            For Each item In items.EnumerateArray()
                token.ThrowIfCancellationRequested()
                builder.Append(Text(item, "name")).Append(": ").AppendLine(Text(item, "value"))
            Next
            Return builder.ToString()
        End Function
        Private Shared Sub Notice(record As HarRecord, message As String)
            record.BodyIncomplete = True
            record.BodyNotice = (record.BodyNotice & " " & message).Trim()
        End Sub
        Private Shared Function Body(content As JsonElement, limit As Long, decode As Boolean, record As HarRecord, label As String, token As CancellationToken) As String
            Dim value = Text(content, "text")
            If value.Length = 0 Then Return ""
            token.ThrowIfCancellationRequested()
            Dim encoding = Text(content, "encoding")
            If encoding.Length > 0 AndAlso Not encoding.Equals("base64", StringComparison.OrdinalIgnoreCase) Then
                Notice(record, label & " body has an unsupported encoding; omitted.")
                Return ""
            End If
            If encoding.Equals("base64", StringComparison.OrdinalIgnoreCase) Then
                If Not decode Then
                    Notice(record, label & " Base64 body skipped (decoding disabled).")
                    Return ""
                End If
                If value.Length > ((limit + 2) \ 3) * 4 Then
                    Notice(record, label & " Base64 body exceeds decoded-size limit.")
                    Return ""
                End If
                Try
                    Dim bytes = Convert.FromBase64String(value)
                    token.ThrowIfCancellationRequested()
                    If bytes.LongLength > limit Then
                        Notice(record, label & " decoded body exceeds limit.")
                        Return ""
                    End If
                    Dim decoded = New UTF8Encoding(False, True).GetString(bytes)
                    If decoded.Contains(ChrW(0)) Then
                        Notice(record, label & " binary body omitted.")
                        Return ""
                    End If
                    record.BodyNotice = (record.BodyNotice & " " & label & " body decoded from Base64 (UTF-8).").Trim()
                    Return decoded
                Catch ex As FormatException
                    Notice(record, label & " body has invalid Base64.")
                    Return ""
                Catch ex As DecoderFallbackException
                    Notice(record, label & " body is not UTF-8 text; omitted.")
                    Return ""
                End Try
            End If
            If value.Length > limit OrElse System.Text.Encoding.UTF8.GetByteCount(value) > limit Then
                Notice(record, label & " body exceeds limit; omitted from search and preview.")
                Return ""
            End If
            Return value
        End Function
    End Class
End Namespace
