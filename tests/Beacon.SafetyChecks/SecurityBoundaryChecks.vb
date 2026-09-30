Imports System.IO
Imports System.Linq
Imports System.Reflection
Imports System.Runtime.InteropServices
Imports System.Security.AccessControl
Imports System.Security.Principal
Imports System.Text
Imports System.Text.Json
Imports Beacon

Module SecurityBoundaryChecks
    Public Sub PrivateStorage()
        Dim prefix = "BeaconSecurityChecks_"
        Dim ownedPath As String
        Using directory = PrivateTemporaryDirectory.Create(prefix)
            ownedPath = directory.DirectoryPath
            Dim security = New DirectoryInfo(ownedPath).GetAccessControl(AccessControlSections.Owner Or AccessControlSections.Access)
            Using identity = WindowsIdentity.GetCurrent()
                Require(security.AreAccessRulesProtected AndAlso identity.User.Equals(security.GetOwner(GetType(SecurityIdentifier))), "Temporary storage did not receive a private owner/DACL at creation.")
                For Each rule As FileSystemAccessRule In security.GetAccessRules(True, True, GetType(SecurityIdentifier))
                    Require(rule.AccessControlType <> AccessControlType.Allow OrElse rule.IdentityReference.Equals(identity.User) OrElse
                            rule.IdentityReference.Equals(New SecurityIdentifier(WellKnownSidType.LocalSystemSid, Nothing)), "Temporary storage grants access to an unexpected identity.")
                Next
            End Using
            Using file = directory.CreateFile("evidence.bin")
                file.Write(Encoding.UTF8.GetBytes("private fixture"))
                Reject(Of IOException)(Sub()
                                           Using duplicate = directory.CreateFile("evidence.bin")
                                           End Using
                                       End Sub)
            End Using
            Reject(Of ArgumentException)(Sub()
                                             Using invalid = directory.CreateFile("../outside.bin")
                                             End Using
                                         End Sub)
            Reject(Of IOException)(Sub() IO.Directory.Move(ownedPath, ownedPath & "-moved"))
            PrivateTemporaryDirectory.CleanupStale(prefix)
            Require(IO.File.Exists(Path.Combine(ownedPath, "evidence.bin")), "Stale cleanup removed an active private session.")
            Using file = PrivateTemporaryDirectory.OpenRead(Path.Combine(ownedPath, "evidence.bin"))
                Reject(Of IOException)(Sub() IO.File.WriteAllText(Path.Combine(ownedPath, "evidence.bin"), "replaced"))
                Reject(Of IOException)(Sub() IO.File.Delete(Path.Combine(ownedPath, "evidence.bin")))
            End Using
        End Using
        Require(Not Directory.Exists(ownedPath), "Private session cleanup leaked its directory.")

        Dim streamPath As String
        Using temporary = PrivateTemporaryDirectory.CreateTemporaryStream("BeaconSecurityStream_", ".tar")
            streamPath = temporary.Name
            temporary.Write(Encoding.UTF8.GetBytes("temporary stream"))
            Reject(Of IOException)(Sub()
                                       Using other = File.OpenRead(streamPath)
                                       End Using
                                   End Sub)
        End Using
        Require(Not File.Exists(streamPath) AndAlso Not Directory.Exists(Path.GetDirectoryName(streamPath)), "Temporary stream did not release its file and private directory.")

        Dim abandoned = PrivateTemporaryDirectory.Create(prefix)
        Dim abandonedPath = abandoned.DirectoryPath
        Using busy = abandoned.CreateFile("busy.bin")
            abandoned.Dispose()
            Require(Directory.Exists(abandonedPath), "The cleanup fixture was not retained by its locked file.")
        End Using
        PrivateTemporaryDirectory.CleanupStale(prefix)
        Require(Not Directory.Exists(abandonedPath), "Private stale cleanup did not reclaim an inactive session.")
        CheckRedirectedPaths()
    End Sub

    Private Sub CheckRedirectedPaths()
        Using outside = PrivateTemporaryDirectory.Create("BeaconSecurityOutside_"),
              directory = PrivateTemporaryDirectory.Create("BeaconSecurityLinks_")
            Dim sentinel = Path.Combine(outside.DirectoryPath, "keep.txt")
            IO.File.WriteAllText(sentinel, "outside evidence")
            Dim link = Path.Combine(directory.DirectoryPath, "redirect")
            Try
                IO.Directory.CreateSymbolicLink(link, outside.DirectoryPath)
            Catch ex As Exception When TypeOf ex Is UnauthorizedAccessException OrElse TypeOf ex Is IOException
                Dim start As New System.Diagnostics.ProcessStartInfo(Path.Combine(Environment.SystemDirectory, "cmd.exe")) With {
                    .Arguments = $"/d /c mklink /J ""{link}"" ""{outside.DirectoryPath}""", .UseShellExecute = False, .CreateNoWindow = True,
                    .RedirectStandardOutput = True, .RedirectStandardError = True
                }
                Using process = System.Diagnostics.Process.Start(start)
                    If Not process.WaitForExit(10000) Then
                        process.Kill(True)
                        Throw New TimeoutException("Creating the isolated junction fixture timed out.")
                    End If
                    Require(process.ExitCode = 0, "The reparse-point test could not create a junction: " & process.StandardError.ReadToEnd())
                End Using
            End Try
            Reject(Of IOException)(Sub()
                                       Using lease = PrivateTemporaryDirectory.LockDirectory(link)
                                       End Using
                                   End Sub)
            Reject(Of IOException)(Sub()
                                       Using lease = PrivateTemporaryDirectory.LockDirectory(Path.Combine(link, "child"), createMissing:=True)
                                       End Using
                                   End Sub)
            Require(Not IO.Directory.Exists(Path.Combine(outside.DirectoryPath, "child")), "A redirected output path created data outside the private directory.")
            Dim fileLink = Path.Combine(directory.DirectoryPath, "redirect.txt")
            Try
                File.CreateSymbolicLink(fileLink, sentinel)
                Reject(Of IOException)(Sub()
                                           Using input = PrivateTemporaryDirectory.OpenRead(fileLink)
                                           End Using
                                       End Sub)
            Catch ex As Exception When TypeOf ex Is UnauthorizedAccessException OrElse TypeOf ex Is IOException
                Console.WriteLine("SKIP file-symlink fixture: Windows requires additional privilege; directory-junction rejection and cleanup were exercised.")
            End Try
            directory.Dispose()
            Require(File.ReadAllText(sentinel) = "outside evidence", "Temporary cleanup followed a link into external evidence.")
        End Using
    End Sub

    Public Sub HelperIntegrity()
        Dim before = Directory.GetDirectories(Path.GetTempPath(), "BeaconTool_*").ToHashSet(StringComparer.OrdinalIgnoreCase)
        Dim bytes = Encoding.UTF8.GetBytes("synthetic executable image for locking tests")
        Dim executable As String
        Using source As New MemoryStream(bytes), helper = SevenZipHelper.VerifiedHelper.Create(source)
            executable = helper.ExecutablePath
            Require(File.ReadAllBytes(executable).SequenceEqual(bytes), "Verified helper bytes differ from the embedded fixture.")
            Reject(Of IOException)(Sub() File.WriteAllBytes(executable, {CByte(0)}))
            Reject(Of IOException)(Sub() File.Delete(executable))
            Reject(Of IOException)(Sub() File.Move(executable, executable & ".old"))
            Reject(Of IOException)(Sub() Directory.Move(Path.GetDirectoryName(executable), Path.GetDirectoryName(executable) & "-moved"))
            Require(File.ReadAllBytes(executable).SequenceEqual(bytes), "A cached helper image was replaced while verified.")
        End Using
        Require(Not File.Exists(executable), "Disposing the helper did not release and remove its executable.")
        Using changed As New ChangingImageStream(bytes)
            Reject(Of InvalidDataException)(Sub()
                                               Using helper = SevenZipHelper.VerifiedHelper.Create(changed)
                                               End Using
                                           End Sub)
        End Using
        Require(Directory.GetDirectories(Path.GetTempPath(), "BeaconTool_*").All(Function(path) before.Contains(path)), "Failed helper verification leaked a usable executable directory.")
    End Sub

    Public Sub WorkerRequests()
        Using directory = PrivateTemporaryDirectory.Create("BeaconRequestChecks_")
            Dim request As New EvtxWorkerHost.EvtxWorkerRequest With {
                .FilePath = Path.Combine(directory.DirectoryPath, "sample.evtx"), .QueryText = "error", .QueryMode = SearchMode.PlainText,
                .Settings = New BeaconSettings(), .RemainingMatches = 300
            }
            Dim good = WriteRequest(directory, JsonSerializer.SerializeToUtf8Bytes(request))
            Dim restored = EvtxWorkerHost.ReadRequest(good)
            Require(restored.FilePath = request.FilePath AndAlso restored.QueryText = request.QueryText AndAlso restored.RemainingMatches = 300,
                    "Bounded request parsing changed valid worker input.")
            Reject(Of InvalidDataException)(Sub() EvtxWorkerHost.ReadRequest("relative.json"))
            Dim empty = WriteRequest(directory, Array.Empty(Of Byte)())
            Reject(Of InvalidDataException)(Sub() EvtxWorkerHost.ReadRequest(empty))
            Dim oversized = WriteRequest(directory, New Byte(EvtxWorkerHost.MaximumRequestBytes) {})
            Reject(Of InvalidDataException)(Sub() EvtxWorkerHost.ReadRequest(oversized))
            Dim nested = WriteRequest(directory, Encoding.UTF8.GetBytes("{""padding"":" & New String("["c, 24) & "0" & New String("]"c, 24) & "}"))
            Reject(Of JsonException)(Sub() EvtxWorkerHost.ReadRequest(nested))
            For Each change In New Action(Of EvtxWorkerHost.EvtxWorkerRequest)() {
                Sub(value) value.FilePath = "relative.evtx",
                Sub(value) value.QueryText = New String("x"c, 4097),
                Sub(value) value.QueryMode = CType(999, SearchMode),
                Sub(value) value.StartSequence = -1,
                Sub(value) value.RemainingMatches = 100001,
                Sub(value) value.RemainingMatches = -1
            }
                Dim invalid = JsonSerializer.Deserialize(Of EvtxWorkerHost.EvtxWorkerRequest)(JsonSerializer.SerializeToUtf8Bytes(request))
                change(invalid)
                Dim filePath = WriteRequest(directory, JsonSerializer.SerializeToUtf8Bytes(invalid))
                Reject(Of InvalidDataException)(Sub() EvtxWorkerHost.ReadRequest(filePath))
            Next
            Using locked = PrivateTemporaryDirectory.OpenRead(good)
                Reject(Of IOException)(Sub() File.WriteAllText(good, "{}"))
                Require(EvtxWorkerHost.ReadRequest(good).QueryText = "error", "A locked request could not be read by a second reader.")
            End Using
        End Using
    End Sub

    Private Function WriteRequest(directory As PrivateTemporaryDirectory, content As Byte()) As String
        Dim name = Guid.NewGuid().ToString("N") & ".json"
        Using output = directory.CreateFile(name)
            output.Write(content)
        End Using
        Return Path.Combine(directory.DirectoryPath, name)
    End Function

    Public Sub NativeImports()
        Dim count As Integer
        For Each root In {GetType(MainWindow), GetType(NativeCaptionTheme), GetType(SearchConcurrencyPolicy), GetType(PrivateTemporaryDirectory), GetType(SearchCompletionNotifier)}
            For Each type In TypesIncludingNested(root)
                For Each method In type.GetMethods(BindingFlags.Public Or BindingFlags.NonPublic Or BindingFlags.Static Or BindingFlags.Instance Or BindingFlags.DeclaredOnly)
                    If method.GetCustomAttribute(Of DllImportAttribute)() Is Nothing Then Continue For
                    count += 1
                    Dim paths = method.GetCustomAttribute(Of DefaultDllImportSearchPathsAttribute)()
                    Require(paths IsNot Nothing AndAlso paths.Paths = DllImportSearchPath.System32, "Windows import allows uncontrolled DLL lookup: " & type.Name & "." & method.Name)
                Next
            Next
        Next
        Require(count >= 12, "The native import audit missed expected production APIs.")
    End Sub

    Private Iterator Function TypesIncludingNested(root As Type) As IEnumerable(Of Type)
        Yield root
        For Each nested In root.GetNestedTypes(BindingFlags.Public Or BindingFlags.NonPublic)
            For Each type In TypesIncludingNested(nested)
                Yield type
            Next
        Next
    End Function

    Private NotInheritable Class ChangingImageStream
        Inherits MemoryStream
        Public Sub New(bytes As Byte())
            MyBase.New(bytes)
        End Sub
        Public Overrides Sub CopyTo(destination As Stream, bufferSize As Integer)
            destination.Write(Encoding.UTF8.GetBytes("substituted after verification"))
        End Sub
    End Class

    Private Sub Reject(Of T As Exception)(action As Action)
        Try
            action()
        Catch ex As T
            Return
        End Try
        Throw New InvalidOperationException("Expected " & GetType(T).Name & ".")
    End Sub

    Private Sub Require(condition As Boolean, message As String)
        If Not condition Then Throw New InvalidOperationException(message)
    End Sub
End Module
