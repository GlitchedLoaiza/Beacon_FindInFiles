Imports System.IO
Imports System.Text
Imports System.Text.Json
Imports System.Text.Json.Nodes
Imports System.Threading

Namespace Beacon
    Public NotInheritable Class HarRedaction
        Public Const Notice As String = "HAR presentation redaction is enabled. Header values, URL query values/user information and recognized JSON/form secrets are hidden; unstructured bodies are withheld. Search uses original bounded data, so sensitive-only matches may have no visible highlight. This is not anonymization: paths, search terms and other data may remain sensitive."
        Private Sub New()
        End Sub
        Public Shared Function Present(original As HarRecord, enabled As Boolean, token As CancellationToken) As HarRecord
            Dim result = original.Copy()
            result.RedactionEnabled = enabled
            If Not enabled Then Return result
            result.Url = SafeUrl(original.Url)
            result.RequestHeaders = HideHeaders(original.RequestHeaders)
            result.ResponseHeaders = HideHeaders(original.ResponseHeaders)
            result.RequestBody = SafeBody(original.RequestBody, original.RequestMime, token)
            result.ResponseBody = SafeBody(original.ResponseBody, original.MimeType, token)
            result.StatusText = ""
            result.RedactionApplied = result.SearchText() <> original.SearchText()
            Return result
        End Function
        Private Shared Function HideHeaders(value As String) As String
            If String.IsNullOrEmpty(value) Then Return ""
            ' Header values can contain credentials in nonstandard fields or redirect URLs; withhold all values.
            Return String.Join(vbLf, value.Replace(vbCrLf, vbLf).Split(vbLf).Where(Function(line) line.Length > 0).
                Select(Function(line)
                           Dim separator = line.IndexOf(":"c)
                           Return If(separator > 0, line.Substring(0, separator) & ": [redacted]", "[redacted header]")
                       End Function))
        End Function
        Private Shared Function Sensitive(name As String) As Boolean
            Dim normalized = name.Replace("-", "").Replace("_", "").ToLowerInvariant()
            Return {"authorization", "cookie", "password", "passwd", "secret", "token", "apikey", "session", "credential", "signature"}.
                Any(Function(word) normalized.Contains(word, StringComparison.Ordinal))
        End Function
        Private Shared Function SafeUrl(value As String) As String
            If String.IsNullOrEmpty(value) Then Return ""
            Dim uri As Uri = Nothing
            If Not Uri.TryCreate(value, UriKind.Absolute, uri) OrElse (uri.Scheme <> "http" AndAlso uri.Scheme <> "https") Then Return "[redacted invalid URL]"
            Dim builder As New UriBuilder(uri) With {.UserName = "", .Password = "", .Fragment = "", .Query = ""}
            Dim result = builder.Uri.GetLeftPart(UriPartial.Path)
            If uri.Query.Length > 0 Then result &= "?" & String.Join("&", uri.Query.TrimStart("?"c).Split("&"c).Select(
                Function(pair)
                    Dim index = pair.IndexOf("="c)
                    Return If(index >= 0, pair.Substring(0, index), "parameter") & "=[redacted]"
                End Function))
            Return result
        End Function
        Private Shared Function SafeBody(value As String, mime As String, token As CancellationToken) As String
            If String.IsNullOrEmpty(value) Then Return ""
            token.ThrowIfCancellationRequested()
            If If(mime, "").Contains("json", StringComparison.OrdinalIgnoreCase) Then
                Try
                    Dim root = JsonNode.Parse(value)
                    Scrub(root, token)
                    Return If(root Is Nothing, "null", root.ToJsonString(New JsonSerializerOptions With {.WriteIndented = True}))
                Catch ex As JsonException
                    Return "[body withheld: invalid JSON]"
                End Try
            End If
            If If(mime, "").StartsWith("application/x-www-form-urlencoded", StringComparison.OrdinalIgnoreCase) Then
                Return String.Join("&", value.Split("&"c).Select(Function(pair)
                    token.ThrowIfCancellationRequested()
                    Dim separator = pair.IndexOf("="c)
                    If separator < 0 Then Return "[redacted field]"
                    Dim key = pair.Substring(0, separator)
                    Return If(Sensitive(Uri.UnescapeDataString(key.Replace("+", " "))), key & "=[redacted]", pair)
                End Function))
            End If
            Return "[body withheld: unstructured or unsupported content]"
        End Function
        Private Shared Sub Scrub(node As JsonNode, token As CancellationToken)
            token.ThrowIfCancellationRequested()
            If TypeOf node Is JsonObject Then
                Dim obj = DirectCast(node, JsonObject)
                For Each key In obj.Select(Function(pair) pair.Key).ToArray()
                    If Sensitive(key) Then
                        obj(key) = JsonValue.Create("[redacted]")
                    Else
                        Scrub(obj(key), token)
                    End If
                Next
            ElseIf TypeOf node Is JsonArray Then
                For Each child In DirectCast(node, JsonArray)
                    Scrub(child, token)
                Next
            End If
        End Sub

        Public Shared Function CaptureDetail(query As SearchQuery, presented As HarRecord, location As String, recordIndex As Integer, token As CancellationToken) As SearchDetail
            Dim text = presented.SearchText()
            Dim detail = SearchDetail.FromRecord(query, text, location, recordIndex:=recordIndex, token:=token)
            If detail IsNot Nothing Then
                detail.ContextKind = If(presented.RedactionEnabled, "HAR presentation lines (redaction enabled)", "Request text lines")
                Return detail
            End If
            Dim hidden = "Matched original request data; matching text is hidden by HAR presentation redaction."
            Return New SearchDetail With {.Location = location, .RecordIndex = recordIndex, .Excerpt = hidden,
                .ContextKind = "HAR redacted match", .Context = MatchContext.FromText(hidden, New List(Of SearchSpan)(), token)}
        End Function
    End Class
End Namespace
