Imports System.Globalization
Imports System.IO
Imports System.Net
Imports System.Text
Imports System.Text.Json
Imports System.Text.Json.Serialization
Imports System.Threading
Imports System.Threading.Tasks

Namespace Beacon
    Public Enum ScanReportFormat
        Json
        Csv
        Html
    End Enum

    Public NotInheritable Class ScanReportWriter
        Private Shared ReadOnly JsonOptions As JsonSerializerOptions = CreateJsonOptions()
        Private Sub New()
        End Sub
        Private Shared Function CreateJsonOptions() As JsonSerializerOptions
            Dim options As New JsonSerializerOptions With {.WriteIndented = True, .DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull}
            options.Converters.Add(New JsonStringEnumConverter())
            Return options
        End Function

        Public Shared Function SaveAsync(report As ScanReportSnapshot, destination As String, format As ScanReportFormat,
                                          includeExcerpts As Boolean, includeTechnical As Boolean, token As CancellationToken) As Task
            Return Task.Run(Async Function()
                                If Not [Enum].IsDefined(format) Then Throw New ArgumentException("Unsupported report format.")
                                Dim fullPath = Path.GetFullPath(destination)
                                ProtectSources(report, fullPath)
                                Dim snapshot = report.ForExport(includeExcerpts, includeTechnical)
                                Dim pending = Path.Combine(Path.GetDirectoryName(fullPath), ".beacon-report-" & Guid.NewGuid().ToString("N") & ".tmp")
                                Try
                                    token.ThrowIfCancellationRequested()
                                    Using output As New FileStream(pending, FileMode.CreateNew, FileAccess.Write, FileShare.None, 65536, FileOptions.Asynchronous)
                                        If format = ScanReportFormat.Json Then
                                            Await JsonSerializer.SerializeAsync(output, snapshot, JsonOptions, token).ConfigureAwait(False)
                                        Else
                                            Using writer As New StreamWriter(output, New UTF8Encoding(format = ScanReportFormat.Csv), 65536, leaveOpen:=True)
                                                If format = ScanReportFormat.Csv Then
                                                    WriteCsv(snapshot, writer, token)
                                                Else
                                                    WriteHtml(snapshot, writer, token)
                                                End If
                                                Await writer.FlushAsync(token).ConfigureAwait(False)
                                            End Using
                                        End If
                                        Await output.FlushAsync(token).ConfigureAwait(False)
                                    End Using
                                    token.ThrowIfCancellationRequested()
                                    File.Move(pending, fullPath, True)
                                Finally
                                    Try
                                        If File.Exists(pending) Then File.Delete(pending)
                                    Catch ex As IOException
                                        Diagnostics.Debug.WriteLine("Could not remove export temporary file: " & ex.Message)
                                    Catch ex As UnauthorizedAccessException
                                        Diagnostics.Debug.WriteLine("Could not remove export temporary file: " & ex.Message)
                                    End Try
                                End Try
                            End Function, token)
        End Function

        Private Shared Sub ProtectSources(report As ScanReportSnapshot, destination As String)
            Dim sources = report.Results.Select(Function(item) item.SourcePath).Append(report.Run.SourceRoot)
            For Each source In sources
                If String.IsNullOrWhiteSpace(source) Then Continue For
                Dim boundary = source.IndexOf(" | ", StringComparison.Ordinal)
                Dim physicalPath = If(boundary < 0, source, source.Substring(0, boundary))
                If String.Equals(Path.GetFullPath(physicalPath), destination, StringComparison.OrdinalIgnoreCase) Then
                    Throw New IOException("Choose a report destination different from the scanned source files.")
                End If
            Next
        End Sub

        Public Shared Function CopyPaths(report As ScanReportSnapshot) As String
            Return String.Join(Environment.NewLine, report.Results.Select(Function(item) item.SourcePath).Distinct(StringComparer.OrdinalIgnoreCase))
        End Function

        Public Shared Function CopyResult(file As ReportFile) As String
            Dim result As New StringBuilder()
            result.AppendLine(file.SourcePath)
            result.AppendLine($"{file.StoredDetailCount:N0} stored detail(s)" & If(String.IsNullOrEmpty(file.PartialReason), "", " · " & file.PartialReason))
            For Each match In file.Matches
                result.AppendLine($"{match.Location}: {match.VisibleMatches:N0} visible match(es)")
            Next
            Return result.ToString().TrimEnd()
        End Function

        Public Shared Function CopyDiagnostics(report As ScanReportSnapshot, technical As Boolean) As String
            Dim text As New StringBuilder()
            text.AppendLine(report.Coverage)
            For Each issue In report.Diagnostics.Items
                text.AppendLine($"{issue.TimestampUtc:O} [{issue.Heading}] {issue.SourcePath}")
                text.AppendLine(issue.Message & $" (reported {issue.Occurrences} time(s))")
                If technical Then text.AppendLine(issue.TechnicalText)
            Next
            If report.Diagnostics.OmittedNotifications > 0 Then text.AppendLine($"Omitted notifications: {report.Diagnostics.OmittedNotifications}")
            Return text.ToString()
        End Function

        Public Shared Function CsvField(value As String) As String
            value = If(value, "")
            Dim start = 0
            While start < value.Length AndAlso (Char.IsWhiteSpace(value(start)) OrElse value(start) = ChrW(&HFEFF))
                start += 1
            End While
            If (start < value.Length AndAlso "=+-@＝＋－＠".IndexOf(value(start)) >= 0) OrElse
               (value.Length > 0 AndAlso (value(0) = ControlChars.Tab OrElse value(0) = ControlChars.Cr OrElse value(0) = ControlChars.Lf)) Then
                value = "'" & value
            End If
            Return """" & value.Replace("""", """""") & """"
        End Function

        Private Shared Iterator Function Metadata(report As ScanReportSnapshot) As IEnumerable(Of KeyValuePair(Of String, String))
            Yield New KeyValuePair(Of String, String)("Application", report.ApplicationName & " " & report.ApplicationVersion)
            Yield New KeyValuePair(Of String, String)("CapturedUtc", report.CapturedUtc.ToString("O"))
            Yield New KeyValuePair(Of String, String)("StartedUtc", If(report.Run.StartedUtc.HasValue, report.Run.StartedUtc.Value.ToString("O"), ""))
            Yield New KeyValuePair(Of String, String)("CompletedUtc", If(report.Run.CompletedUtc.HasValue, report.Run.CompletedUtc.Value.ToString("O"), ""))
            Yield New KeyValuePair(Of String, String)("State", report.Run.State.ToString())
            Yield New KeyValuePair(Of String, String)("Coverage", report.Coverage)
            Yield New KeyValuePair(Of String, String)("Source", report.Run.SourceRoot)
            Yield New KeyValuePair(Of String, String)("Query", report.Run.QueryText)
            Yield New KeyValuePair(Of String, String)("Mode", report.Run.Mode.ToString())
            Yield New KeyValuePair(Of String, String)("CaseSensitive", report.Run.CaseSensitive.ToString())
            Yield New KeyValuePair(Of String, String)("ElapsedMilliseconds", report.ElapsedMilliseconds.ToString(CultureInfo.InvariantCulture))
            Yield New KeyValuePair(Of String, String)("FilesScanned", report.FilesScanned.ToString(CultureInfo.InvariantCulture))
            Yield New KeyValuePair(Of String, String)("EstimatedTotalFiles", report.EstimatedTotalFiles.ToString(CultureInfo.InvariantCulture))
            Yield New KeyValuePair(Of String, String)("MatchedFiles", report.Results.Count.ToString(CultureInfo.InvariantCulture))
            Yield New KeyValuePair(Of String, String)("DiagnosticNotifications", report.Diagnostics.TotalNotifications.ToString(CultureInfo.InvariantCulture))
            Yield New KeyValuePair(Of String, String)("OmittedDiagnosticNotifications", report.Diagnostics.OmittedNotifications.ToString(CultureInfo.InvariantCulture))
            Yield New KeyValuePair(Of String, String)("IncludesExcerpts", report.IncludesExcerpts.ToString())
            Yield New KeyValuePair(Of String, String)("IncludesTechnicalDetails", report.IncludesTechnicalDetails.ToString())
            Yield New KeyValuePair(Of String, String)("RedactionApplied", report.RedactionApplied.ToString())
            Yield New KeyValuePair(Of String, String)("Failure", report.Run.FailureMessage)
            Yield New KeyValuePair(Of String, String)("Options", JsonSerializer.Serialize(report.Run.Options, JsonOptions))
        End Function

        Private Shared Sub CsvRow(writer As TextWriter, ParamArray fields As String())
            Dim row(15) As String
            Array.Copy(fields, row, fields.Length)
            writer.WriteLine(String.Join(",", row.Select(AddressOf CsvField)))
        End Sub

        Private Shared Sub WriteCsv(report As ScanReportSnapshot, writer As TextWriter, token As CancellationToken)
            CsvRow(writer, "RecordType", "Source", "Location", "Count", "Excerpt", "PartialReason", "TimestampUtc", "Severity", "Category", "Message", "Stage", "ExceptionType", "ErrorCode", "TechnicalDetails", "MetadataKey", "MetadataValue")
            For Each pair In Metadata(report)
                token.ThrowIfCancellationRequested()
                CsvRow(writer, "metadata", "", "", "", "", "", "", "", "", "", "", "", "", "", pair.Key, pair.Value)
            Next
            For Each file In report.Results
                token.ThrowIfCancellationRequested()
                CsvRow(writer, "file", file.SourcePath, file.FileType, file.StoredDetailCount.ToString(CultureInfo.InvariantCulture), "", file.PartialReason)
                For Each match In file.Matches
                    token.ThrowIfCancellationRequested()
                    CsvRow(writer, "match", file.SourcePath, match.Location, match.VisibleMatches.ToString(CultureInfo.InvariantCulture), match.Excerpt, file.PartialReason)
                Next
            Next
            For Each issue In report.Diagnostics.Items
                token.ThrowIfCancellationRequested()
                CsvRow(writer, "diagnostic", issue.SourcePath, "", issue.Occurrences.ToString(CultureInfo.InvariantCulture), "", If(issue.TextTruncated, "diagnostic text truncated", ""),
                       issue.TimestampUtc.ToString("O"), issue.Severity, issue.Category, issue.Message, issue.Stage, issue.ExceptionType, issue.ErrorCode, issue.TechnicalDetails)
            Next
        End Sub

        Private Shared Function H(value As String) As String
            Return WebUtility.HtmlEncode(If(value, ""))
        End Function

        Private Shared Sub WriteHtml(report As ScanReportSnapshot, writer As TextWriter, token As CancellationToken)
            writer.WriteLine("<!doctype html><html lang='en'><head><meta charset='utf-8'><meta name='viewport' content='width=device-width,initial-scale=1'><meta http-equiv='Content-Security-Policy' content=""default-src 'none'; style-src 'unsafe-inline'""><title>Beacon — Search results</title><style>body{font:14px 'Segoe UI',sans-serif;max-width:1100px;margin:24px auto;padding:0 20px;background:#f3f3f3;color:#202020}header,section,details{background:white;border:1px solid #ddd;border-radius:8px;padding:18px;margin:16px 0}h1,h2,h3{font-weight:600}h1{margin:4px 0 16px}.brand{color:#0066cc;font-weight:600;letter-spacing:.06em}.muted,.path{color:#555}.path{overflow-wrap:anywhere;font-family:Consolas,monospace;font-size:12px}.notice{border-left:3px solid #0066cc;padding-left:10px}nav a{display:block;padding:4px 0;color:#0066cc;overflow-wrap:anywhere}article{margin:20px 0}table{border-collapse:collapse;width:100%;table-layout:fixed}th{width:3em;font-weight:400;color:#666;text-align:right;vertical-align:top;padding:6px 10px;border-right:1px solid #ddd}td{padding:6px 10px;vertical-align:top}.focus{background:#e5f1ff}pre{font:13px Consolas,monospace;white-space:pre-wrap;overflow-wrap:anywhere;tab-size:4;margin:0}mark{background:#ffe082;color:#202020;font-weight:600;border-radius:2px}summary{cursor:pointer;font-weight:600}dt{font-weight:600;margin-top:10px}dd{margin:4px 0;overflow-wrap:anywhere;white-space:pre-wrap}@media(prefers-color-scheme:dark){body{background:#202020;color:#e0e0e0}header,section,details{background:#2b2b2b;border-color:#505050}.brand,nav a{color:#60cfff}.muted,.path,th{color:#b0b0b0}.focus{background:#183c50}th{border-color:#505050}.notice{border-color:#60cfff}}</style></head><body><header><div class='brand'>BEACON</div><h1>Search results</h1>")
            Dim mode = SearchModeChoice.All().FirstOrDefault(Function(item) item.Value = report.Run.Mode)
            writer.WriteLine("<p>Search: <strong>" & H(report.Run.QueryText) & "</strong> · " & H(If(mode Is Nothing, report.Run.Mode.ToString(), mode.Label)) & If(report.Run.CaseSensitive, " · Case sensitive", " · Case insensitive") & "</p>")
            writer.WriteLine("<p class='path'>Source: " & H(report.Run.SourceRoot) & "</p><p>" & report.Results.Count.ToString(CultureInfo.InvariantCulture) & " matching file(s) · Saved " & H(report.CapturedUtc.ToString("yyyy-MM-dd HH:mm:ss 'UTC'")) & "</p>")
            writer.WriteLine("<p class='notice'>" & H(If(report.IsComplete, "Search completed within the selected scope.", report.Coverage.Replace("Incomplete snapshot:", "Partial results:"))) & "</p>")
            writer.WriteLine("<p class='muted'>" & If(report.IncludesExcerpts, "Includes unredacted matching text and context. Review before sharing.", "Text snippets are not included.") & "</p></header>")
            If report.Results.Count > 0 Then
                writer.WriteLine("<nav aria-label='Matching files'><h2>Matching files</h2>")
                For index = 0 To report.Results.Count - 1
                    token.ThrowIfCancellationRequested()
                    writer.WriteLine($"<a href='#file-{index}'>" & H(report.Results(index).DisplayName) & "</a>")
                Next
                writer.WriteLine("</nav>")
            Else
                writer.WriteLine("<section><p>No matching files were captured.</p></section>")
            End If
            Dim fileIndex As Integer = 0
            For Each file In report.Results
                token.ThrowIfCancellationRequested()
                writer.WriteLine($"<section id='file-{fileIndex}'><h2>" & H(file.DisplayName) & "</h2><p class='path'>" & H(file.SourcePath) & "</p><p>" & file.StoredDetailCount.ToString(CultureInfo.InvariantCulture) & " saved match location(s)</p>")
                If Not String.IsNullOrEmpty(file.PartialReason) Then writer.WriteLine("<p class='notice'>" & H(file.PartialReason) & " — additional matches may exist.</p>")
                For Each match In file.Matches
                    token.ThrowIfCancellationRequested()
                    writer.WriteLine("<article><h3>" & H(match.Location) & "</h3>")
                    If match.IsMetadata Then writer.WriteLine("<p class='muted'>This is a file-name/path match, not a content match.</p>")
                    If report.IncludesExcerpts Then
                        WriteContext(writer, match)
                    Else
                        writer.WriteLine("<p>" & match.VisibleMatches.ToString(CultureInfo.InvariantCulture) & " visible match(es)</p>")
                    End If
                    writer.WriteLine("</article>")
                Next
                writer.WriteLine("</section>")
                fileIndex += 1
            Next
            writer.WriteLine("<details><summary>Scan information</summary><dl>")
            For Each pair In Metadata(report).Where(Function(item) item.Key <> "Options")
                writer.WriteLine("<dt>" & H(pair.Key) & "</dt><dd>" & H(pair.Value) & "</dd>")
            Next
            writer.WriteLine("</dl></details>")
            If report.Diagnostics.Items.Count > 0 Then
                writer.WriteLine("<details><summary>Files skipped and scan problems</summary>")
                For Each issue In report.Diagnostics.Items
                    token.ThrowIfCancellationRequested()
                    writer.WriteLine("<h3>" & H(issue.Category) & "</h3><p class='path'>" & H(issue.SourcePath) & "</p><pre>" & H(issue.Message) & "</pre>")
                    If report.IncludesTechnicalDetails Then writer.WriteLine("<pre>" & H(issue.TechnicalText) & "</pre>")
                Next
                writer.WriteLine("</details>")
            End If
            writer.WriteLine("<p class='muted'>Generated by Beacon. Context was captured during the search; original files are not embedded.</p></body></html>")
        End Sub

        Private Shared Sub WriteContext(writer As TextWriter, match As ReportMatch)
            If match.Context.Count = 0 Then
                writer.WriteLine("<p class='muted'>Full line context was not captured for this result. Run the search again to include it.</p><pre>" & H(match.Excerpt) & "</pre>")
                Return
            End If
            writer.WriteLine("<p class='muted'>" & H(If(match.ContextKind, "Context")) & " · up to five lines" & If(match.Context.Any(Function(line) line.Shortened), " · long lines shortened", "") & "</p><table class='snippet' aria-label='Match and surrounding context'><tbody>")
            For Each line In match.Context.Take(MatchContext.LineLimit)
                writer.Write("<tr" & If(line.IsMatchLine, " class='focus'", "") & "><th scope='row'>" & line.LineNumber.ToString(CultureInfo.InvariantCulture) & "</th><td><pre>")
                Dim position As Integer = 0
                Dim text = If(line.Text, "")
                For Each span In line.Highlights.OrderBy(Function(item) item.Start)
                    If span.Start < position OrElse span.Start >= text.Length OrElse span.Length <= 0 OrElse span.Length > text.Length - span.Start Then Continue For
                    writer.Write(H(text.Substring(position, span.Start - position)))
                    writer.Write("<mark>" & H(text.Substring(span.Start, span.Length)) & "</mark>")
                    position = span.Start + span.Length
                Next
                writer.Write(H(text.Substring(position)))
                writer.WriteLine("</pre></td></tr>")
            Next
            writer.WriteLine("</tbody></table>")
        End Sub
    End Class
End Namespace
