Imports System.Globalization

Namespace Beacon
    Public NotInheritable Class EvtxFilter
        Public Const ResourceGuidance As String = "Beacon uses Windows-installed event provider resources and never downloads or loads arbitrary DLLs. Missing messages do not make the raw event XML unavailable. On the source computer, export the relevant log with Event Viewer and include display information when offered, or archive a COPY of the exported EVTX with wevtutil al in a trusted working directory without untrusted symlinks or junctions. Keep the EVTX and its LocaleMetaData folder together when transferring them. Open that EVTX directly from disk: Beacon's archive-entry extraction does not carry sibling locale metadata. Run archive operations on copies, not original evidence. See Microsoft documentation for wevtutil archive-log and EventLogSession.ExportLogAndMessages."
        Private ReadOnly _ids As HashSet(Of Integer)
        Private ReadOnly _levels As HashSet(Of Integer)
        Private ReadOnly _provider As String
        Private ReadOnly _fromUtc As DateTimeOffset?
        Private ReadOnly _toUtc As DateTimeOffset?

        Public Sub New(ids As String, provider As String, levels As String, fromUtc As String, toUtc As String)
            _ids = ParseNumbers(ids, 65535, "Event IDs")
            _levels = ParseNumbers(levels, 5, "Severity levels")
            _provider = If(provider, "").Trim()
            If _provider.Length > 256 Then Throw New ArgumentException("Provider name must not exceed 256 characters.")
            _fromUtc = ParseUtc(fromUtc)
            _toUtc = ParseUtc(toUtc)
            If _fromUtc.HasValue AndAlso _toUtc.HasValue AndAlso _fromUtc > _toUtc Then
                Throw New ArgumentException("EVTX start time must not be later than end time.")
            End If
        End Sub

        Public Shared Function FromSettings(settings As BeaconSettings) As EvtxFilter
            Return New EvtxFilter(settings.EvtxEventIds, settings.EvtxProvider, settings.EvtxLevels, settings.EvtxFromUtc, settings.EvtxToUtc)
        End Function

        Public Function Matches(id As Integer, provider As String, level As Byte?, time As DateTime?) As Boolean
            If _ids.Count > 0 AndAlso Not _ids.Contains(id) Then Return False
            If _provider.Length > 0 AndAlso Not String.Equals(_provider, provider, StringComparison.OrdinalIgnoreCase) Then Return False
            If _levels.Count > 0 AndAlso (Not level.HasValue OrElse Not _levels.Contains(CInt(level.Value))) Then Return False
            If _fromUtc.HasValue OrElse _toUtc.HasValue Then
                If Not time.HasValue Then Return False
                Dim utc = New DateTimeOffset(time.Value.ToUniversalTime())
                If _fromUtc.HasValue AndAlso utc < _fromUtc.Value Then Return False
                If _toUtc.HasValue AndAlso utc > _toUtc.Value Then Return False
            End If
            Return True
        End Function

        Private Shared Function ParseNumbers(value As String, maximum As Integer, label As String) As HashSet(Of Integer)
            Dim result As New HashSet(Of Integer)()
            If String.IsNullOrWhiteSpace(value) Then Return result
            If value.Length > 2048 Then Throw New ArgumentException(label & " list is too long.")
            For Each part In value.Split(","c)
                Dim number As Integer
                If Not Integer.TryParse(part.Trim(), NumberStyles.None, CultureInfo.InvariantCulture, number) OrElse number < 0 OrElse number > maximum Then
                    Throw New ArgumentException($"{label} must be comma-separated numbers between 0 and {maximum}.")
                End If
                result.Add(number)
            Next
            Return result
        End Function

        Private Shared Function ParseUtc(value As String) As DateTimeOffset?
            If String.IsNullOrWhiteSpace(value) Then Return Nothing
            Dim result As DateTimeOffset
            If Not DateTimeOffset.TryParseExact(value.Trim(), {"yyyy-MM-dd HH:mm:ss", "yyyy-MM-ddTHH:mm:ss'Z'"},
                                               CultureInfo.InvariantCulture, DateTimeStyles.AssumeUniversal Or DateTimeStyles.AdjustToUniversal, result) Then
                Throw New ArgumentException("EVTX times must be UTC in yyyy-MM-dd HH:mm:ss format, or yyyy-MM-ddTHH:mm:ssZ.")
            End If
            Return result
        End Function

        Public Shared Function Describe(settings As BeaconSettings) As String
            If settings Is Nothing Then Return "No scan options captured"
            Return $"IDs: {OrAll(settings.EvtxEventIds)}; Provider: {OrAll(settings.EvtxProvider)}; Levels: {OrAll(settings.EvtxLevels)}; UTC from: {OrAll(settings.EvtxFromUtc)}; UTC to: {OrAll(settings.EvtxToUtc)}"
        End Function

        Private Shared Function OrAll(value As String) As String
            Return If(String.IsNullOrWhiteSpace(value), "Any", value.Trim())
        End Function
    End Class
End Namespace
