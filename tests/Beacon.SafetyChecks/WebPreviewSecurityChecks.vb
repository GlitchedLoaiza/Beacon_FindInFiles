Imports System.IO
Imports System.Linq
Imports System.Net
Imports System.Net.Sockets
Imports System.Text
Imports System.Text.Json
Imports System.Threading
Imports System.Threading.Tasks
Imports System.Windows
Imports System.Windows.Threading
Imports Microsoft.Web.WebView2.Core
Imports Microsoft.Web.WebView2.Wpf
Imports Beacon

Module WebPreviewSecurityChecks
    Public Sub Run()
        If Application.Current Is Nothing Then
            Dim application As New Application With {.ShutdownMode = ShutdownMode.OnExplicitShutdown}
        End If
        Dim uiDispatcher = Dispatcher.CurrentDispatcher
        Dim previous = SynchronizationContext.Current
        SynchronizationContext.SetSynchronizationContext(New DispatcherSynchronizationContext(uiDispatcher))
        Try
            Dim pending = ExerciseAsync()
            While Not pending.IsCompleted
                uiDispatcher.Invoke(Sub()
                                    End Sub, DispatcherPriority.ContextIdle)
                Thread.Sleep(5)
            End While
            pending.GetAwaiter().GetResult()
        Finally
            SynchronizationContext.SetSynchronizationContext(previous)
        End Try
    End Sub

    Private Async Function ExerciseAsync() As Task
        Using profile = PrivateTemporaryDirectory.Create("BeaconSecureBrowserChecks_"), trap As New NetworkTrap()
            Dim browser As New WebView2 With {.AllowExternalDrop = False}
            Dim window As New Window With {.Content = browser, .Width = 720, .Height = 440, .Left = -10000, .Top = -10000,
                                           .ShowActivated = False, .ShowInTaskbar = False, .WindowStartupLocation = WindowStartupLocation.Manual}
            Dim security As WebPreviewSecurity = Nothing
            Try
                window.Show()
                Dim environment = Await CoreWebView2Environment.CreateAsync(Nothing, profile.DirectoryPath).WaitAsync(TimeSpan.FromSeconds(20))
                Await browser.EnsureCoreWebView2Async(environment).WaitAsync(TimeSpan.FromSeconds(20))
                Dim core = browser.CoreWebView2
                security = New WebPreviewSecurity(core)
                Require(Not core.Settings.IsScriptEnabled AndAlso Not core.Settings.IsWebMessageEnabled AndAlso Not core.Settings.AreHostObjectsAllowed,
                        "The preview exposes source scripts or host communication.")
                Require(Not core.Settings.IsGeneralAutofillEnabled AndAlso Not core.Settings.IsPasswordAutosaveEnabled,
                        "The evidence preview must not inject or retain browser credentials.")
                Dim policyHeader As String = Nothing
                Dim headerHandler As EventHandler(Of CoreWebView2WebResourceResponseReceivedEventArgs) =
                    Sub(sender, e)
                        If e.Request.Uri = security.CurrentUri Then policyHeader = e.Response.Headers.GetHeader("Content-Security-Policy")
                    End Sub
                AddHandler core.WebResourceResponseReceived, headerHandler
                Dim url = trap.Url
                Dim localFile = Path.Combine(profile.DirectoryPath, "local-secret.txt")
                File.WriteAllText(localFile, "synthetic local data that must not be loaded")
                Dim fileUri = New Uri(localFile).AbsoluteUri
                Dim html = "<!doctype html><html><head>" &
                    "<meta http-equiv='Content-Security-Policy' content=""default-src * 'unsafe-inline' data:; script-src 'unsafe-inline' *"">" &
                    "<base href='" & url & "'><link rel='stylesheet' href='" & url & "style.css'>" &
                    "<link rel='preconnect' href='" & url & "'><link rel='dns-prefetch' href='" & url & "'>" &
                    "<style>@import url('" & url & "import.css'); #proof{color:rgb(1,2,3);background-image:url('" & url & "css.png')}</style>" &
                    "<script>window.pageScriptRan=true;fetch('" & url & "script');</script><script src='" & url & "source.js'></script>" &
                    "</head><body><p id='proof'>confidential synthetic evidence</p>" &
                    "<img src='" & url & "image.png' onerror='window.pageScriptRan=true'>" &
                    "<img src='" & fileUri & "'><iframe src='" & url & "frame'></iframe>" &
                    "<iframe srcdoc='&lt;script&gt;parent.pageScriptRan=true&lt;/script&gt;'></iframe>" &
                    "<object data='" & url & "object'></object><audio autoplay src='" & url & "audio'></audio>" &
                    "<img id='inline' src='data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9Wl6iZsAAAAASUVORK5CYII='>" &
                    "<a id='link' href='" & url & "navigate'>external</a><a id='popup' target='_blank' href='" & url & "popup'>popup</a>" &
                    "<a id='download' download='fixture.txt' href='data:text/plain,synthetic'>download</a>" &
                    "<a id='scheme' href='beacon-security-fixture://not-installed'>external scheme</a>" &
                    "<form id='form' action='" & url & "submit' method='post'><input name='data' value='synthetic'></form></body></html>"
                Require(Await security.ShowHtmlAsync(html), "The hostile-document fixture could not be displayed as static evidence.")
                Await Task.Delay(150)
                Console.WriteLine($"Preview network probe after document installation: {trap.Connections}")
                Require(policyHeader = WebPreviewSecurity.ContentSecurityPolicy, "The preview response did not carry its restrictive CSP header.")
                Require(Await browser.ExecuteScriptAsync("typeof window.pageScriptRan") = """undefined""", "A script supplied by the document executed.")
                Require(Await browser.ExecuteScriptAsync("getComputedStyle(document.getElementById('proof')).color") = """rgb(1, 2, 3)""", "Safe inline document styling was lost.")
                Require(Await browser.ExecuteScriptAsync("document.getElementById('inline').naturalWidth") = "1", "Safe embedded images no longer load.")
                Require(Await browser.ExecuteScriptAsync("document.querySelectorAll('script,link,base,iframe,object,meta[http-equiv]').length") = "0",
                        "Active elements or resource hints reached the live document.")
                Require(Await browser.ExecuteScriptAsync("document.querySelector('input').disabled && !document.getElementById('link').hasAttribute('href')") = "true",
                        "Imported controls or external links retained active capabilities.")
                Dim captureScript = WebSearchScripts.Capture(10000, "secure").TrimEnd(";"c)
                Dim captureJson = Await browser.ExecuteScriptAsync("(()=>{try{return " & captureScript & ";}catch(e){return {captureError:String(e)};}})()")
                Using result = JsonDocument.Parse(captureJson)
                    Require(result.RootElement.ValueKind = JsonValueKind.String, "Isolated text capture failed: " & captureJson)
                End Using
                Dim captured = JsonSerializer.Deserialize(Of String)(captureJson)
                Require(captured.Contains("confidential synthetic evidence"), "The isolated document is no longer searchable.")
                Dim ranges = New SearchQuery("synthetic evidence", SearchMode.PlainText, False).FindHighlights(captured)
                Require(Await browser.ExecuteScriptAsync(WebSearchScripts.Apply(ranges, "secure")) = "1", "Host-injected highlighting stopped working with source scripting disabled.")

                Dim current = security.CurrentUri
                Await browser.ExecuteScriptAsync("document.getElementById('link').click(); document.getElementById('popup').click(); document.getElementById('download').click(); document.getElementById('scheme').click(); document.getElementById('form').submit();")
                Await Task.Delay(150)
                Console.WriteLine($"Preview network probe after imported-element actions: {trap.Connections}")
                Await browser.ExecuteScriptAsync("fetch('" & url & "fetch').catch(()=>{});")
                Await Task.Delay(150)
                Console.WriteLine($"Preview network probe after host-injected fetch: {trap.Connections}")
                Await browser.ExecuteScriptAsync("try { new WebSocket('" & url.Replace("http:", "ws:") & "socket'); } catch(e) {}")
                Await Task.Delay(200)
                Console.WriteLine($"Preview network probe after host-injected WebSocket: {trap.Connections}")
                Require(core.Source = current, "An unsolicited page navigation was allowed.")
                Require(Not security.TryNavigateDocument(fileUri), "The preview controller accepted a local-file host destination.")
                Await Task.Delay(100)
                Console.WriteLine($"Preview network probe after host file navigation: {trap.Connections}")
                Require(core.Source = current, "The preview navigated to a local file.")
                Require(Not security.TryNavigateDocument(url & "direct"), "The preview controller accepted an external host destination.")
                Await Task.Delay(100)
                Console.WriteLine($"Preview network probe after host HTTP navigation: {trap.Connections}")
                Require(core.Source = current, "The preview navigated to a remote address.")
                Require(trap.Connections = 0, "The isolated preview contacted a network endpoint.")

                Dim unicodePrefix = "<p id='unicode'>"
                Dim unicodeHtml = unicodePrefix & New String("x"c, 8191 - unicodePrefix.Length) & Char.ConvertFromUtf32(&H1F600) & " Ω</p>"
                Require(Await security.ShowHtmlAsync(unicodeHtml), "The Unicode chunk-boundary preview failed.")
                Dim unicodeText = JsonSerializer.Deserialize(Of String)(Await browser.ExecuteScriptAsync("document.getElementById('unicode').textContent"))
                Require(unicodeText.EndsWith(Char.ConvertFromUtf32(&H1F600) & " Ω", StringComparison.Ordinal), "Chunked source encoding split a surrogate pair.")

                Dim before = Directory.GetDirectories(Path.GetTempPath(), "BeaconPreview_*").ToHashSet(StringComparer.OrdinalIgnoreCase)
                Dim large = "<html><body><p id='large'>" & New String("x"c, 1700000) & "</p></body></html>"
                Require(Await security.ShowHtmlAsync(large), "A large preview failed through the protected response stream.")
                Require(Await browser.ExecuteScriptAsync("document.getElementById('large').textContent.length") = "1700000", "The large document was truncated by preview navigation.")
                Require(core.Source.StartsWith("https://beacon-preview-", StringComparison.Ordinal) AndAlso Not core.Source.StartsWith("file:", StringComparison.OrdinalIgnoreCase),
                        "A large preview bypassed isolation with local-file navigation.")
                Require(Directory.GetDirectories(Path.GetTempPath(), "BeaconPreview_*").All(Function(path) before.Contains(path)), "The completed large preview retained its response file.")
                Dim obsolete = security.ShowHtmlAsync("<p>old preview</p>")
                Dim latest = security.ShowHtmlAsync("<p id='latest'>new preview</p>")
                Await obsolete
                Require(Await latest, "Superseding a preview canceled the latest document.")
                Require(Await browser.ExecuteScriptAsync("document.getElementById('latest').textContent") = """new preview""", "An old navigation replaced the latest document.")
                Require(trap.Connections = 0, "A later preview escaped the request policy.")
                Using bypassProbe As New NetworkTrap()
                    Dim expected = security.CurrentUri
                    core.Navigate(bypassProbe.Url & "trusted-sdk-bypass")
                    Await Task.Delay(200)
                    Require(core.Source = expected AndAlso security.BlockedNavigations > 0, "The secondary navigation guard did not cancel a raw SDK navigation.")
                    Console.WriteLine($"Raw SDK bypass characterization: navigation canceled; {bypassProbe.Connections} speculative connections. Production must use the validated controller entry point.")
                End Using
                RemoveHandler core.WebResourceResponseReceived, headerHandler
            Finally
                security?.Dispose()
                browser.Dispose()
                window.Close()
            End Try
        End Using
    End Function

    Private NotInheritable Class NetworkTrap
        Implements IDisposable
        Private ReadOnly _listener As New TcpListener(IPAddress.Loopback, 0)
        Private ReadOnly _cancel As New CancellationTokenSource()
        Private ReadOnly _work As Task
        Private _connections As Integer

        Public Sub New()
            _listener.Start()
            Url = "http://127.0.0.1:" & DirectCast(_listener.LocalEndpoint, IPEndPoint).Port.ToString(Globalization.CultureInfo.InvariantCulture) & "/"
            _work = Task.Run(Async Function()
                                 Try
                                     While True
                                         Using client = Await _listener.AcceptTcpClientAsync(_cancel.Token).ConfigureAwait(False)
                                             Interlocked.Increment(_connections)
                                             Dim response = Encoding.ASCII.GetBytes("HTTP/1.1 204 No Content" & vbCrLf & "Content-Length: 0" & vbCrLf & "Connection: close" & vbCrLf & vbCrLf)
                                             Await client.GetStream().WriteAsync(response, _cancel.Token).ConfigureAwait(False)
                                         End Using
                                     End While
                                 Catch ex As Exception When _cancel.IsCancellationRequested OrElse TypeOf ex Is IOException OrElse TypeOf ex Is SocketException
                                 End Try
                             End Function)
        End Sub

        Public ReadOnly Property Url As String
        Public ReadOnly Property Connections As Integer
            Get
                Return Volatile.Read(_connections)
            End Get
        End Property

        Public Sub Dispose() Implements IDisposable.Dispose
            _cancel.Cancel()
            _listener.Stop()
            _work.GetAwaiter().GetResult()
            _cancel.Dispose()
        End Sub
    End Class

    Private Sub Require(condition As Boolean, message As String)
        If Not condition Then Throw New InvalidOperationException(message)
    End Sub
End Module
