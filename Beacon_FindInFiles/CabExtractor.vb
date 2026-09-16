Imports System.Diagnostics
Imports System.Globalization
Imports System.IO
Imports System.Text
Imports System.Threading
Imports System.Threading.Tasks

Namespace Beacon
    Public NotInheritable Class CabExtractor
        Private Sub New()
        End Sub

        Public Shared Async Function ExtractAsync(executable As String, archivePath As String, outputDirectory As String,
                                                  settings As BeaconSettings, token As CancellationToken) As Task
            token.ThrowIfCancellationRequested()
            Using deadline = CancellationTokenSource.CreateLinkedTokenSource(token)
                deadline.CancelAfter(TimeSpan.FromSeconds(settings.ArchiveProcessTimeoutSeconds))
                Dim ct = deadline.Token
                Dim budget As New ArchiveReadBudget(settings, New FileInfo(archivePath).Length)
                Dim entries As New List(Of KeyValuePair(Of String, Long))()
                Using listing As New MemoryStream()
                    Dim listingLimit = Math.Min(64L * 1024 * 1024, CLng(settings.MaximumArchiveEntries) * 8192 + 1048576)
                    Await RunAsync(executable, {"l", "-slt", "-sccUTF-8", "-pBeacon-No-Password", "--", Path.GetFullPath(archivePath)}, listing, listingLimit, Nothing, ct)
                    Dim text = Encoding.UTF8.GetString(listing.ToArray()).Replace(vbCrLf, vbLf)
                    Dim marker = text.IndexOf(vbLf & "----------" & vbLf, StringComparison.Ordinal)
                    If marker < 0 OrElse Not text.Substring(0, marker).Contains("Type = Cab", StringComparison.OrdinalIgnoreCase) Then
                        Throw New InvalidDataException("The file is not a supported CAB archive.")
                    End If
                    Dim seen As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
                    For Each block In text.Substring(marker + 12).Split({vbLf & vbLf}, StringSplitOptions.RemoveEmptyEntries)
                        ct.ThrowIfCancellationRequested()
                        Dim fields As New Dictionary(Of String, String)(StringComparer.Ordinal)
                        For Each line In block.Split(vbLf)
                            If String.IsNullOrWhiteSpace(line) Then Continue For
                            Dim separator = line.IndexOf(" = ", StringComparison.Ordinal)
                            If separator <= 0 Then Throw New InvalidDataException("Invalid CAB metadata.")
                            If Not fields.TryAdd(line.Substring(0, separator), line.Substring(separator + 3)) Then
                                Throw New InvalidDataException("Ambiguous CAB metadata.")
                            End If
                        Next
                        Dim name As String = Nothing
                        Dim sizeText As String = Nothing
                        Dim size As Long
                        If Not fields.TryGetValue("Path", name) OrElse Not fields.TryGetValue("Size", sizeText) OrElse
                           Not Long.TryParse(sizeText, NumberStyles.None, CultureInfo.InvariantCulture, size) Then
                            Throw New InvalidDataException("Missing CAB entry metadata.")
                        End If
                        Dim encrypted = fields.ContainsKey("Encrypted") AndAlso fields("Encrypted") <> "-"
                        budget.Register(name, size, 0, encrypted)
                        If Not seen.Add(name.Replace("/", "\")) Then Throw New InvalidDataException("Duplicate CAB entry path.")
                        If ArchiveSafety.IsExcluded(name, settings) Then Continue For
                        Dim attributes = If(fields.ContainsKey("Attributes"), fields("Attributes"), "")
                        If Not settings.IncludeHiddenFiles AndAlso attributes.Contains("H"c) Then Continue For
                        If Not settings.IncludeSystemFiles AndAlso attributes.Contains("S"c) Then Continue For
                        If fields.ContainsKey("Folder") AndAlso fields("Folder") = "+" Then Continue For
                        entries.Add(New KeyValuePair(Of String, Long)(name, size))
                    Next
                End Using

                Directory.CreateDirectory(outputDirectory)
                Dim root = Path.GetFullPath(outputDirectory).TrimEnd(Path.DirectorySeparatorChar) & Path.DirectorySeparatorChar
                For Each entry In entries
                    ct.ThrowIfCancellationRequested()
                    Dim destination = Path.GetFullPath(Path.Combine(root, entry.Key))
                    If Not destination.StartsWith(root, StringComparison.OrdinalIgnoreCase) Then Throw New InvalidDataException("CAB path escapes output directory.")
                    Directory.CreateDirectory(Path.GetDirectoryName(destination))
                    Using output As New FileStream(destination, FileMode.CreateNew, FileAccess.Write, FileShare.None)
                        Await RunAsync(executable, {"e", "-so", "-spd", "-y", "-pBeacon-No-Password", "--", Path.GetFullPath(archivePath), entry.Key}, output, entry.Value, budget, ct)
                        If output.Length <> entry.Value Then Throw New InvalidDataException("CAB entry size does not match metadata.")
                    End Using
                Next
            End Using
        End Function

        Private Shared Async Function RunAsync(executable As String, arguments As IEnumerable(Of String), output As Stream,
                                                limit As Long, budget As ArchiveReadBudget, token As CancellationToken) As Task
            token.ThrowIfCancellationRequested()
            Dim startInfo As New ProcessStartInfo(executable) With {
                .UseShellExecute = False, .CreateNoWindow = True,
                .RedirectStandardInput = True, .RedirectStandardOutput = True, .RedirectStandardError = True
            }
            For Each argument In arguments
                startInfo.ArgumentList.Add(argument)
            Next
            Using child = Process.Start(startInfo)
                If child Is Nothing Then Throw New IOException("Could not start the archive reader.")
                child.StandardInput.Close()
                Using abort = CancellationTokenSource.CreateLinkedTokenSource(token)
                    Using registration = abort.Token.Register(Sub()
                                                                  Try
                                                                      If Not child.HasExited Then child.Kill(True)
                                                                  Catch ex As InvalidOperationException
                                                                  Catch ex As ComponentModel.Win32Exception
                                                                  End Try
                                                              End Sub)
                        Dim copyTask = CopyAsync(child.StandardOutput.BaseStream, output, limit, budget, abort)
                        Dim errorTask = CopyAsync(child.StandardError.BaseStream, Stream.Null, 65536, Nothing, abort)
                        Try
                            Await Task.WhenAll(copyTask, errorTask, child.WaitForExitAsync(abort.Token))
                            token.ThrowIfCancellationRequested()
                            If child.ExitCode <> 0 Then Throw New InvalidDataException($"Archive reader failed (exit code {child.ExitCode}).")
                        Finally
                            If Not child.HasExited Then
                                child.Kill(True)
                                child.WaitForExit()
                            End If
                        End Try
                    End Using
                End Using
            End Using
        End Function

        Private Shared Async Function CopyAsync(input As Stream, output As Stream, limit As Long,
                                                  budget As ArchiveReadBudget, abort As CancellationTokenSource) As Task
            Dim buffer(81919) As Byte
            Dim total As Long = 0
            Try
                While True
                    Dim count = Await input.ReadAsync(buffer.AsMemory(0, CInt(Math.Min(buffer.Length, limit - total + 1))), abort.Token)
                    If count = 0 Then Exit While
                    If count > limit - total Then Throw New InvalidDataException("Archive output limit reached.")
                    budget?.Consume(count)
                    Await output.WriteAsync(buffer.AsMemory(0, count), abort.Token)
                    total += count
                End While
            Catch
                abort.Cancel()
                Throw
            End Try
        End Function
    End Class
End Namespace
