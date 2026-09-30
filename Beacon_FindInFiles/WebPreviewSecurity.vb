Imports System.IO
Imports System.Text
Imports System.Text.Json
Imports System.Threading.Tasks
Imports Microsoft.Web.WebView2.Core

Namespace Beacon
    Friend NotInheritable Class WebPreviewSecurity
        Implements IDisposable

        Friend Const ContentSecurityPolicy As String = "default-src 'none'; script-src 'none'; style-src 'unsafe-inline'; img-src data:; " &
            "connect-src 'none'; object-src 'none'; base-uri 'none'; form-action 'none'; frame-src 'none'; frame-ancestors 'none'; sandbox allow-same-origin"
        Private Const ResponseHeaders As String = "Content-Type: text/html; charset=utf-8" & vbCrLf &
            "Content-Security-Policy: " & ContentSecurityPolicy & vbCrLf &
            "Cache-Control: no-store" & vbCrLf & "X-Content-Type-Options: nosniff" & vbCrLf &
            "Referrer-Policy: no-referrer" & vbCrLf & "X-DNS-Prefetch-Control: off"
        Private Shared ReadOnly DocumentStart As Byte() = Encoding.UTF8.GetBytes("<!doctype html><html><head><meta charset='utf-8'></head><body><script id='beacon-source-data' type='application/json'>")
        Private Shared ReadOnly DocumentEnd As Byte() = Encoding.UTF8.GetBytes("</script></body></html>")
        Private Const InstallStaticDocument As String = "(function(expectedUri) {
    if (location.href !== expectedUri) return false;
    var payload = document.getElementById('beacon-source-data');
    if (!payload) return false;
    var template = document.createElement('template');
    template.innerHTML = JSON.parse(payload.textContent);
    var blocked = new Set(['script','link','base','meta','iframe','frame','frameset','object','embed','applet','portal','fencedframe',
                           'audio','video','source','track','template','animate','animatemotion','animatetransform','set']);
    var elements = template.content.querySelectorAll('*');
    for (var index = 0; index < elements.length; index++) {
        var element = elements[index], name = element.localName.toLowerCase();
        if (blocked.has(name)) { element.remove(); continue; }
        var attributes = Array.from(element.attributes);
        for (var a = 0; a < attributes.length; a++) {
            var attribute = attributes[a], key = attribute.name.toLowerCase(), value = attribute.value.trim();
            var remove = key.startsWith('on') || ['autofocus','contenteditable','srcset','ping','action','formaction','target','poster','data'].includes(key);
            var raster = /^data:image\/(png|jpeg|gif|webp|bmp|avif);base64,/i.test(value);
            if (key === 'src') remove = !(name === 'img' && raster);
            if (key === 'href' || key === 'xlink:href') {
                remove = !(value.startsWith('#') || (name === 'image' && raster));
            }
            if (remove) element.removeAttributeNode(attribute);
        }
        if (['input','button','select','textarea'].includes(name)) element.disabled = true;
    }
    document.body.replaceChildren(template.content);
    return true;
})(__BEACON_URI__);"
        Private ReadOnly _core As CoreWebView2
        Private ReadOnly _networkReady As Task
        Private ReadOnly _origin As String = "https://beacon-preview-" & Guid.NewGuid().ToString("N") & ".invalid/"
        Private _document As Stream
        Private _completion As TaskCompletionSource(Of Boolean)
        Private _allowNavigation As Boolean
        Private _allowResponse As Boolean
        Private _navigationId As ULong
        Private _disposed As Boolean

        Public Sub New(core As CoreWebView2)
            ArgumentNullException.ThrowIfNull(core)
            _core = core
            With core.Settings
                .IsScriptEnabled = False
                .IsWebMessageEnabled = False
                .AreHostObjectsAllowed = False
                .AreDevToolsEnabled = False
                .AreDefaultContextMenusEnabled = False
                .AreBrowserAcceleratorKeysEnabled = False
                .IsGeneralAutofillEnabled = False
                .IsPasswordAutosaveEnabled = False
                .IsBuiltInErrorPageEnabled = False
                .IsStatusBarEnabled = False
            End With
            core.IsMuted = True
            core.AddWebResourceRequestedFilter("*", CoreWebView2WebResourceContext.All, CoreWebView2WebResourceRequestSourceKinds.All)
            AddHandler core.WebResourceRequested, AddressOf ResourceRequested
            AddHandler core.NavigationStarting, AddressOf NavigationStarting
            AddHandler core.FrameNavigationStarting, AddressOf FrameNavigationStarting
            AddHandler core.NavigationCompleted, AddressOf NavigationCompleted
            AddHandler core.NewWindowRequested, AddressOf NewWindowRequested
            AddHandler core.DownloadStarting, AddressOf DownloadStarting
            AddHandler core.PermissionRequested, AddressOf PermissionRequested
            AddHandler core.LaunchingExternalUriScheme, AddressOf ExternalSchemeRequested
            AddHandler core.BasicAuthenticationRequested, AddressOf AuthenticationRequested
            AddHandler core.ServerCertificateErrorDetected, AddressOf CertificateError
            AddHandler core.ProcessFailed, AddressOf ProcessFailed
            _networkReady = DisableNetworkAsync()
        End Sub

        Private Async Function DisableNetworkAsync() As Task
            Await _core.CallDevToolsProtocolMethodAsync("Network.enable", "{}")
            ObjectDisposedException.ThrowIf(_disposed, Me)
            Await _core.CallDevToolsProtocolMethodAsync("Network.emulateNetworkConditions",
                "{""offline"":true,""latency"":0,""downloadThroughput"":-1,""uploadThroughput"":-1}")
        End Function

        Friend ReadOnly Property CurrentUri As String
        Friend ReadOnly Property BlockedRequests As Integer
        Friend ReadOnly Property BlockedNavigations As Integer

        Public Async Function ShowHtmlAsync(html As String) As Task(Of Boolean)
            ObjectDisposedException.ThrowIf(_disposed, Me)
            ArgumentNullException.ThrowIfNull(html)
            Await _networkReady
            ObjectDisposedException.ThrowIf(_disposed, Me)
            _navigationId = 0UL
            _core.Stop()
            _completion?.TrySetResult(False)
            ReleaseDocument()
            _document = CreateDocumentStream(html)
            _CurrentUri = _origin & Guid.NewGuid().ToString("N") & "/index.html"
            _allowNavigation = True
            _allowResponse = True
            Dim completion As New TaskCompletionSource(Of Boolean)(TaskCreationOptions.RunContinuationsAsynchronously)
            _completion = completion
            Dim documentUri = CurrentUri
            Try
                If Not TryNavigateDocument(documentUri) Then Throw New InvalidOperationException("The preview navigation was not authorized.")
                If Not Await completion.Task.WaitAsync(TimeSpan.FromSeconds(20)) OrElse _completion IsNot completion OrElse _disposed Then Return False
                Dim installed = Await _core.ExecuteScriptAsync(InstallStaticDocument.Replace("__BEACON_URI__", JsonSerializer.Serialize(documentUri)))
                Return installed = "true" AndAlso _completion Is completion AndAlso Not _disposed
            Catch
                If _completion Is completion Then
                    _allowNavigation = False
                    _allowResponse = False
                    _core.Stop()
                    ReleaseDocument()
                End If
                Throw
            End Try
        End Function

        Friend Function TryNavigateDocument(destination As String) As Boolean
            If _disposed OrElse Not _allowNavigation OrElse CurrentUri Is Nothing OrElse
               Not String.Equals(destination, CurrentUri, StringComparison.Ordinal) Then Return False
            _core.Navigate(destination)
            Return True
        End Function

        Private Shared Function CreateDocumentStream(html As String) As Stream
            Dim stream As Stream = If(html.Length <= 200000, DirectCast(New MemoryStream(), Stream),
                                      PrivateTemporaryDirectory.CreateTemporaryStream("BeaconPreview_", ".html"))
            Try
                stream.Write(DocumentStart)
                stream.WriteByte(34)
                Dim offset As Integer
                While offset < html.Length
                    Dim count = Math.Min(8192, html.Length - offset)
                    If offset + count < html.Length AndAlso Char.IsHighSurrogate(html(offset + count - 1)) Then count -= 1
                    Dim chunk = JsonEncodedText.Encode(html.AsSpan(offset, count))
                    stream.Write(chunk.EncodedUtf8Bytes)
                    offset += count
                End While
                stream.WriteByte(34)
                stream.Write(DocumentEnd)
                stream.Position = 0
                Return stream
            Catch
                stream.Dispose()
                Throw
            End Try
        End Function

        Private Sub ResourceRequested(sender As Object, e As CoreWebView2WebResourceRequestedEventArgs)
            If Not _disposed AndAlso _allowResponse AndAlso _document IsNot Nothing AndAlso
               e.ResourceContext = CoreWebView2WebResourceContext.Document AndAlso
               String.Equals(e.Request.Uri, CurrentUri, StringComparison.Ordinal) AndAlso e.Request.Method = "GET" Then
                _allowResponse = False
                e.Response = _core.Environment.CreateWebResourceResponse(_document, 200, "OK", ResponseHeaders)
            Else
                _BlockedRequests += 1
                e.Response = _core.Environment.CreateWebResourceResponse(Nothing, 403, "Blocked by Beacon", "Cache-Control: no-store")
            End If
        End Sub

        Private Sub NavigationStarting(sender As Object, e As CoreWebView2NavigationStartingEventArgs)
            If Not _disposed AndAlso _allowNavigation AndAlso String.Equals(e.Uri, CurrentUri, StringComparison.Ordinal) Then
                _allowNavigation = False
                _navigationId = e.NavigationId
                Return
            End If
            If Not _disposed AndAlso e.IsUserInitiated AndAlso IsCurrentFragment(e.Uri) Then Return
            e.Cancel = True
            _BlockedNavigations += 1
        End Sub

        Private Function IsCurrentFragment(value As String) As Boolean
            If CurrentUri Is Nothing OrElse String.IsNullOrEmpty(value) Then Return False
            Dim fragment = value.IndexOf("#"c)
            Return fragment >= 0 AndAlso String.Equals(value.Substring(0, fragment), CurrentUri, StringComparison.Ordinal)
        End Function

        Private Sub FrameNavigationStarting(sender As Object, e As CoreWebView2NavigationStartingEventArgs)
            e.Cancel = True
        End Sub

        Private Sub NavigationCompleted(sender As Object, e As CoreWebView2NavigationCompletedEventArgs)
            If e.NavigationId <> _navigationId Then Return
            _allowResponse = False
            _completion?.TrySetResult(e.IsSuccess)
            ReleaseDocument()
        End Sub

        Private Sub NewWindowRequested(sender As Object, e As CoreWebView2NewWindowRequestedEventArgs)
            e.Handled = True
        End Sub

        Private Sub DownloadStarting(sender As Object, e As CoreWebView2DownloadStartingEventArgs)
            e.Cancel = True
            e.Handled = True
        End Sub

        Private Sub PermissionRequested(sender As Object, e As CoreWebView2PermissionRequestedEventArgs)
            e.State = CoreWebView2PermissionState.Deny
            e.SavesInProfile = False
        End Sub

        Private Sub ExternalSchemeRequested(sender As Object, e As CoreWebView2LaunchingExternalUriSchemeEventArgs)
            e.Cancel = True
        End Sub

        Private Sub AuthenticationRequested(sender As Object, e As CoreWebView2BasicAuthenticationRequestedEventArgs)
            e.Cancel = True
        End Sub

        Private Sub CertificateError(sender As Object, e As CoreWebView2ServerCertificateErrorDetectedEventArgs)
            e.Action = CoreWebView2ServerCertificateErrorAction.Cancel
        End Sub

        Private Sub ProcessFailed(sender As Object, e As CoreWebView2ProcessFailedEventArgs)
            _allowNavigation = False
            _allowResponse = False
            _completion?.TrySetResult(False)
            ReleaseDocument()
        End Sub

        Private Sub ReleaseDocument()
            _document?.Dispose()
            _document = Nothing
        End Sub

        Public Sub Dispose() Implements IDisposable.Dispose
            If _disposed Then Return
            _disposed = True
            _completion?.TrySetResult(False)
            Try
                _core.Stop()
            Finally
                ReleaseDocument()
                RemoveHandler _core.WebResourceRequested, AddressOf ResourceRequested
                RemoveHandler _core.NavigationStarting, AddressOf NavigationStarting
                RemoveHandler _core.FrameNavigationStarting, AddressOf FrameNavigationStarting
                RemoveHandler _core.NavigationCompleted, AddressOf NavigationCompleted
                RemoveHandler _core.NewWindowRequested, AddressOf NewWindowRequested
                RemoveHandler _core.DownloadStarting, AddressOf DownloadStarting
                RemoveHandler _core.PermissionRequested, AddressOf PermissionRequested
                RemoveHandler _core.LaunchingExternalUriScheme, AddressOf ExternalSchemeRequested
                RemoveHandler _core.BasicAuthenticationRequested, AddressOf AuthenticationRequested
                RemoveHandler _core.ServerCertificateErrorDetected, AddressOf CertificateError
                RemoveHandler _core.ProcessFailed, AddressOf ProcessFailed
            End Try
        End Sub
    End Class
End Namespace
