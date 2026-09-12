Namespace Beacon
    Partial Public Class HarFilterEditor
        Private ReadOnly _from As New BeaconFilterControls.UtcDateTimePicker()
        Private ReadOnly _to As New BeaconFilterControls.UtcDateTimePicker()
        Public Sub New()
            InitializeComponent()
            Method_cmb.ItemsSource = {"", "GET", "POST", "PUT", "PATCH", "DELETE", "HEAD", "OPTIONS", "CONNECT", "TRACE"}
            From_host.Content = _from
            To_host.Content = _to
        End Sub
        Public Sub LoadSettings(settings As BeaconSettings)
            Method_cmb.Text = settings.HarMethod
            Host_txt.Text = settings.HarHost
            Status_txt.Text = settings.HarStatusCodes
            Mime_txt.Text = settings.HarMimeType
            Duration_txt.Text = settings.HarMinimumDuration
            _from.Text = settings.HarFromUtc
            _to.Text = settings.HarToUtc
        End Sub
        Public Function ReadFilter() As HarFilter
            Return New HarFilter(Method_cmb.Text, Host_txt.Text, Status_txt.Text, Mime_txt.Text, Duration_txt.Text, _from.Text, _to.Text)
        End Function
        Public Sub SaveSettings(settings As BeaconSettings)
            ReadFilter()
            settings.HarMethod = Method_cmb.Text.Trim()
            settings.HarHost = Host_txt.Text.Trim()
            settings.HarStatusCodes = Status_txt.Text.Trim()
            settings.HarMimeType = Mime_txt.Text.Trim()
            settings.HarMinimumDuration = Duration_txt.Text.Trim()
            settings.HarFromUtc = _from.Text.Trim()
            settings.HarToUtc = _to.Text.Trim()
        End Sub
        Public Sub Clear()
            LoadSettings(New BeaconSettings())
        End Sub
    End Class
End Namespace
