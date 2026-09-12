Imports System.Collections.Concurrent
Imports System.IO
Imports System.Threading
Imports System.Threading.Tasks
Imports Microsoft.Win32
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Threading

Namespace Beacon
    Partial Public Class MainWindow
        Private ReadOnly _diagnostics As New ScanDiagnosticStore()
        Private ReadOnly _logicalSourcePaths As New ConcurrentDictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
        Private _scanRun As ScanRunInfo
        Private _diagnosticRefreshPending As Integer
        Private _htmlExportCancellation As CancellationTokenSource
        Private _htmlExportTask As Task

        Private Sub RememberSourcePath(physicalPath As String, logicalPath As String)
            If String.IsNullOrEmpty(physicalPath) Then Return
            _logicalSourcePaths(Path.GetFullPath(physicalPath)) = logicalPath
        End Sub

        Private Function DiagnosticSource(pathValue As String) As String
            If String.IsNullOrEmpty(pathValue) Then Return ""
            Dim boundary = pathValue.IndexOf(" | ", StringComparison.Ordinal)
            Dim physical = If(boundary < 0, pathValue, pathValue.Substring(0, boundary))
            Try
                Dim logical As String = Nothing
                If _logicalSourcePaths.TryGetValue(Path.GetFullPath(physical), logical) Then
                    Return logical & If(boundary < 0, "", pathValue.Substring(boundary))
                End If
            Catch ex As ArgumentException
            Catch ex As NotSupportedException
            End Try
            Return pathValue
        End Function

        Private Sub RecordDiagnostic(source As String, ex As Exception, stage As String, Optional severity As String = Nothing)
            Dim options = If(_isScanning, _searchOptions, _settings)
            _diagnostics.Record(DiagnosticSource(source), ex, stage,
                                options IsNot Nothing AndAlso options.Diagnostics = DiagnosticDetail.Detailed, severity)
            If _isClosing OrElse Interlocked.Exchange(_diagnosticRefreshPending, 1) <> 0 Then Return
            Dispatcher.BeginInvoke(New Action(Sub()
                                                  Interlocked.Exchange(_diagnosticRefreshPending, 0)
                                                  If Not _isClosing Then UpdateReportingButtons()
                                              End Sub), DispatcherPriority.Background)
        End Sub

        Private Sub UpdateReportingButtons()
            Dim available = Not (_isResetting OrElse _isClosing)
            Dim hasReport = _scanRun IsNot Nothing OrElse _diagnostics.Count > 0 OrElse _hits.Count > 0
            Dim diagnosticButton = TryCast(FindName("Diagnostics_btn"), Button)
            If diagnosticButton IsNot Nothing Then
                diagnosticButton.Content = If(_diagnostics.Count = 0, "Diagnostics", $"Diagnostics ({_diagnostics.Count:N0}{If(_diagnostics.OmittedCount > 0, "+", "")})")
                diagnosticButton.IsEnabled = available AndAlso hasReport
            End If
            Dim exportButton = TryCast(FindName("ExportResults_btn"), Button)
            If exportButton IsNot Nothing Then
                exportButton.Content = If(_htmlExportCancellation Is Nothing, "Export…", "Cancel export")
                exportButton.IsEnabled = available AndAlso (hasReport OrElse _htmlExportCancellation IsNot Nothing)
            End If
            Dim copyPaths = TryCast(FindName("CopyPaths_btn"), Button)
            If copyPaths IsNot Nothing Then copyPaths.IsEnabled = available AndAlso _hits.Count > 0
        End Sub

        Private Sub OpenDiagnostics(sender As Object, e As RoutedEventArgs)
            If _isResetting OrElse _isClosing Then Return
            Dim reportWindow As New ScanReportWindow(Me, BuildReportSnapshot(), AddressOf BuildReportSnapshot,
                                                     _settings.Diagnostics = DiagnosticDetail.Detailed, _isDarkMode)
            reportWindow.ShowDialog()
        End Sub

        Private Async Sub ExportHtmlReport(sender As Object, e As RoutedEventArgs)
            If _htmlExportCancellation IsNot Nothing Then
                _htmlExportCancellation.Cancel()
                Return
            End If
            If _isResetting OrElse _isClosing Then Return
            Dim dialog As New SaveFileDialog With {
                .Title = "Export HTML results — review sensitive data before sharing",
                .FileName = "Beacon-results-" & DateTime.Now.ToString("yyyyMMdd-HHmmss") & ".html",
                .Filter = "HTML search report (*.html)|*.html", .DefaultExt = ".html",
                .AddExtension = True, .OverwritePrompt = True
            }
            If dialog.ShowDialog(Me) <> True Then Return
            Dim report = BuildReportSnapshot()
            Dim cancellation As New CancellationTokenSource()
            _htmlExportCancellation = cancellation
            UpdateReportingButtons()
            Status("Exporting HTML report…")
            Try
                _htmlExportTask = ScanReportWriter.SaveAsync(report, dialog.FileName, ScanReportFormat.Html, True, False, cancellation.Token)
                Await _htmlExportTask
                If Not _isClosing Then Status("HTML report saved: " & Path.GetFileName(dialog.FileName))
            Catch ex As OperationCanceledException
                If Not _isClosing Then Status("Export cancelled")
            Catch ex As Exception
                If Not _isClosing Then MessageBox.Show(Me, ex.Message, "Could not export HTML report", MessageBoxButton.OK, MessageBoxImage.Warning)
            Finally
                _htmlExportTask = Nothing
                _htmlExportCancellation = Nothing
                cancellation.Dispose()
                If Not _isClosing Then UpdateReportingButtons()
            End Try
        End Sub

        Private Sub CopyResultPaths(sender As Object, e As RoutedEventArgs)
            If _hits.Count = 0 Then Return
            Try
                Dim report As New ScanReportSnapshot With {
                    .Results = _hits.Select(Function(hit) New ReportFile With {.SourcePath = If(hit.LogicalPath, hit.FilePath)}).ToList()
                }
                Clipboard.SetText(ScanReportWriter.CopyPaths(report))
                Status("Result paths copied")
            Catch ex As Exception
                Status("Could not copy paths: " & ex.Message)
            End Try
        End Sub

        Private Function BuildReportSnapshot() As ScanReportSnapshot
            Dispatcher.VerifyAccess()
            Return New ScanReportSnapshot With {
                .ApplicationVersion = GetType(MainWindow).Assembly.GetName().Version.ToString(),
                .CapturedUtc = DateTimeOffset.UtcNow, .Run = If(_scanRun?.Copy(), New ScanRunInfo With {.State = ScanReportState.Ready}),
                .ElapsedMilliseconds = _scanElapsedStopwatch.ElapsedMilliseconds,
                .FilesScanned = Volatile.Read(_filesScanned), .EstimatedTotalFiles = Volatile.Read(_totalFilesToScan),
                .Results = _hits.Select(AddressOf CreateReportFile).ToList(), .Diagnostics = _diagnostics.Snapshot()
            }
        End Function

        Private Shared Function CreateReportFile(hit As SearchHit) As ReportFile
            Dim logical = If(hit.LogicalPath, If(hit.FilePath, hit.DisplayName))
            Dim leaf = logical.Split(New String() {" | "}, StringSplitOptions.None).Last()
            Return New ReportFile With {
                .DisplayName = hit.DisplayName, .SourcePath = logical,
                .HarRedactionApplied = hit.MatchingRequests.Any(Function(request) request.RedactionApplied),
                .HarRedactionEnabled = hit.MatchingRequests.Any(Function(request) request.RedactionEnabled),
                .FileType = Path.GetExtension(leaf).TrimStart("."c).ToUpperInvariant(), .PartialReason = hit.PartialReason,
                .Matches = hit.Details.Select(Function(detail) New ReportMatch With {
                    .Location = detail.Location, .LineNumber = detail.LineNumber, .RecordIndex = detail.RecordIndex,
                    .IsMetadata = detail.IsMetadata, .VisibleMatches = detail.VisibleMatches, .Excerpt = detail.Excerpt,
                    .ContextKind = detail.ContextKind, .Context = detail.Context.Select(Function(line) line.Copy()).ToList()
                }).ToList()
            }
        End Function
    End Class
End Namespace
