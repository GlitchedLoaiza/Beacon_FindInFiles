Imports System.Windows
Imports System.Windows.Controls

Namespace Beacon
    Partial Public Class MainWindow
        Private ReadOnly _harPreviewEditor As New HarFilterEditor()
        Private _harPreviewFilter As New HarFilter("", "", "", "", "", "", "")

        Private Function VisibleHarIndexes(hit As SearchHit) As List(Of Integer)
            Return Enumerable.Range(0, hit.MatchingRequests.Count).Where(Function(index) _harPreviewFilter.Matches(hit.MatchingRequests(index))).ToList()
        End Function
        Private Sub ApplyHarFilter(sender As Object, e As RoutedEventArgs) Handles ApplyHarFilter_btn.Click
            Try
                _harPreviewFilter = _harPreviewEditor.ReadFilter()
                HarFilterStatus_txt.Text = "Preview filters applied; captured results and exports are unchanged."
                Dim hit = TryCast(Results_lst.SelectedItem, SearchHit)
                If hit IsNot Nothing Then RenderFirstHarRequest(hit)
            Catch ex As ArgumentException
                HarFilterStatus_txt.Text = ex.Message
            End Try
        End Sub
        Private Sub ClearHarFilter(sender As Object, e As RoutedEventArgs) Handles ClearHarFilter_btn.Click
            _harPreviewEditor.Clear()
            ApplyHarFilter(sender, e)
        End Sub
        Private Sub NavigateHarRequest(forward As Boolean)
            Dim hit = TryCast(Results_lst.SelectedItem, SearchHit)
            If hit Is Nothing Then Return
            Dim visible = VisibleHarIndexes(hit)
            If visible.Count = 0 Then Return
            Dim position = visible.IndexOf(hit.CurrentRequestIndex)
            hit.CurrentRequestIndex = visible(If(position < 0, 0, (position + If(forward, 1, visible.Count - 1)) Mod visible.Count))
            RenderHarRequest(hit)
        End Sub
        Private Sub ClearHarPreview()
            For Each block In {HarMethod_txt, HarUrl_txt, HarStatus_txt, HarTime_txt, HarDuration_txt, HarServerIp_txt,
                               HarRequestHeaders_txt, HarResponseHeaders_txt, HarRequestBody_txt, HarResponseBody_txt}
                block.Inlines.Clear()
                block.Text = ""
            Next
            HarBodyNotice_txt.Text = "No captured requests match the preview filters."
            HarMatchCounter_lbl.Text = "0 highlighted spans"
        End Sub
    End Class
End Namespace
