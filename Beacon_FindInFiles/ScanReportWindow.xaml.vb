Imports System.Windows
Imports System.Windows.Media

Namespace Beacon
    Partial Public Class ScanReportWindow
        Inherits Window

        Private _report As ScanReportSnapshot
        Private ReadOnly _refresh As Func(Of ScanReportSnapshot)
        Private ReadOnly _showTechnical As Boolean

        Public Sub New(owner As Window, report As ScanReportSnapshot, refresh As Func(Of ScanReportSnapshot),
                       detailed As Boolean, dark As Boolean)
            InitializeComponent()
            _report = report
            _refresh = refresh
            _showTechnical = detailed
            Tag = detailed
            If owner IsNot Nothing Then
                If owner.IsVisible Then Me.Owner = owner
                For Each key In owner.Resources.Keys
                    Dim brush = TryCast(owner.Resources(key), SolidColorBrush)
                    If brush IsNot Nothing Then Resources(key) = brush.CloneCurrentValue()
                Next
            End If
            Resources("CheckmarkBrush") = New SolidColorBrush(If(dark, Color.FromRgb(&H20, &H20, &H20), Colors.White))
            NativeCaptionTheme.Apply(Me, dark)
            Resources("SelectionBackgroundBrush") = New SolidColorBrush(If(dark, Color.FromRgb(&H18, &H3C, &H50), Color.FromRgb(&HE5, &HF1, &HFF)))
            Severity_cmb.ItemsSource = {"All", "Warning", "Error", "Information"}
            Severity_cmb.SelectedIndex = 0
            AddHandler Filter_txt.TextChanged, Sub() ApplyFilter()
            AddHandler Severity_cmb.SelectionChanged, Sub() ApplyFilter()
            AddHandler Refresh_btn.Click, AddressOf RefreshReport
            AddHandler CopyIssues_btn.Click, AddressOf CopyIssues
            AddHandler Close_btn.Click, Sub() Close()
            PresentReport()
        End Sub

        Private Sub PresentReport()
            Source_txt.Text = "Source: " & If(_report.Run.SourceRoot, "No scan captured")
            Source_txt.ToolTip = Source_txt.Text
            Coverage_txt.Text = If(_report.Diagnostics.Items.Count = 0, "No scan problems recorded.", $"{_report.Diagnostics.Items.Count:N0} issue(s) recorded — some files may not have been fully searched.")
            Dim state = If(_report.Run.State = ScanReportState.ResultLimitReached, "Result limit reached", _report.Run.State.ToString())
            Summary_txt.Text = $"{state} · {_report.FilesScanned:N0} files processed · Updated {_report.CapturedUtc:HH:mm:ss} UTC"
            Query_txt.Text = $"Search: {_report.Run.QueryText}"
            Query_txt.ToolTip = Query_txt.Text
            ReportStatus_txt.Text = If(_report.Run.State = ScanReportState.Running, "Search is still running. Select Refresh to update this list.", "Use the source and message above to understand skipped files and errors.")
            ApplyFilter()
            Refresh_btn.IsEnabled = _refresh IsNot Nothing
        End Sub

        Private Sub ApplyFilter()
            If _report Is Nothing Then Return
            Dim term = Filter_txt.Text.Trim()
            Dim severity = TryCast(Severity_cmb.SelectedItem, String)
            Dim filtered = _report.Diagnostics.Items.Where(Function(item)
                                                               If severity IsNot Nothing AndAlso severity <> "All" AndAlso item.Severity <> severity Then Return False
                                                               Return term.Length = 0 OrElse String.Join(" ", {item.SourcePath, item.Message, item.Category, item.Stage}).
                                                                   Contains(term, StringComparison.OrdinalIgnoreCase)
                                                           End Function).ToList()
            Diagnostics_lst.ItemsSource = filtered
            EmptyDiagnostics_txt.Visibility = If(filtered.Count = 0, Visibility.Visible, Visibility.Collapsed)
            EmptyDiagnostics_txt.Text = If(_report.Diagnostics.Items.Count = 0, "No diagnostics recorded.", "No diagnostics match this filter.")
            DiagnosticCount_txt.Text = $"Showing {filtered.Count:N0} of {_report.Diagnostics.Items.Count:N0} recorded issue(s) · {_report.Diagnostics.TotalNotifications:N0} notification(s)"
            If _report.Diagnostics.OmittedNotifications > 0 Then DiagnosticCount_txt.Text &= $" · {_report.Diagnostics.OmittedNotifications:N0} omitted after the diagnostic limit"
            CopyIssues_btn.IsEnabled = _report.Diagnostics.Items.Count > 0
        End Sub

        Private Sub RefreshReport(sender As Object, e As RoutedEventArgs)
            If _refresh Is Nothing Then Return
            _report = _refresh()
            PresentReport()
        End Sub

        Private Sub CopyIssues(sender As Object, e As RoutedEventArgs)
            Try
                Clipboard.SetText(ScanReportWriter.CopyDiagnostics(_report, _showTechnical))
                ReportStatus_txt.Text = "Diagnostics copied. Review sensitive data before sharing."
            Catch ex As Exception
                ReportStatus_txt.Text = "Could not copy diagnostics: " & ex.Message
            End Try
        End Sub

    End Class
End Namespace
