Imports System.Formats.Tar
Imports System.IO
Imports System.IO.Compression
Imports System.Reflection
Imports System.Text
Imports System.Threading
Imports Beacon

Module ArchiveChecks
    Private ReadOnly Payload As String = "first line" & vbLf & "archive upgrade marker café 日本語" & vbLf & "last line"

    Public Sub Formats(sevenZip As String)
        WithFixture(Sub(root)
                        Dim names = {"logs with spaces/日本語.log", "second.log"}
                        Dim tarPath = Path.Combine(root, "logs.tar")
                        WriteTar(tarPath, names)
                        Dim zipPath = Path.Combine(root, "logs.zip")
                        Using zip = ZipFile.Open(zipPath, ZipArchiveMode.Create)
                            For Each name In names
                                Using writer As New StreamWriter(zip.CreateEntry(name).Open())
                                    writer.Write(Payload)
                                End Using
                            Next
                            zip.CreateEntry("empty/")
                        End Using
                        Dim input = Path.Combine(root, "input")
                        Directory.CreateDirectory(input)
                        For Each name In names
                            Dim file = Path.Combine(input, name)
                            Directory.CreateDirectory(Path.GetDirectoryName(file))
                            System.IO.File.WriteAllText(file, Payload, New UTF8Encoding(False))
                        Next
                        Dim sevenPath = Path.Combine(root, "logs.7z")
                        Run7Zip(sevenZip, input, {"a", "-t7z", "-m0=LZMA2", "-ms=on", sevenPath, ".\*"})
                        Dim archives As New List(Of String) From {zipPath, sevenPath, tarPath}
                        For Each formatName As String In {"gzip", "bzip2", "xz"}
                            Dim extension = If(formatName = "gzip", "gz", If(formatName = "bzip2", "bz2", "xz"))
                            Dim compressed = tarPath & "." & extension
                            Run7Zip(sevenZip, root, {"a", "-t" & formatName, compressed, tarPath})
                            archives.Add(compressed)
                            Dim aliasPath = Path.Combine(root, "alias." & If(formatName = "gzip", "tgz", If(formatName = "bzip2", "tbz2", "txz")))
                            File.Copy(compressed, aliasPath)
                            archives.Add(aliasPath)
                        Next
                        ' No GZIP original-filename header: compressed TAR detection must use bytes.
                        Dim nameless = Path.Combine(root, "no-header-name.gz")
                        Using output = File.Create(nameless)
                            Using gzip As New GZipStream(output, CompressionMode.Compress)
                                Using inputTar = File.OpenRead(tarPath)
                                    inputTar.CopyTo(gzip)
                                End Using
                            End Using
                        End Using
                        archives.Add(nameless)
                        For Each archivePath In archives
                            VerifyEntries(archivePath, names)
                            SearchChecks.CheckPipeline(archivePath, "archive upgrade marker", SearchMode.PlainText, Options(),
                                Sub(hits)
                                    Ensure(hits.Count = 2, Path.GetFileName(archivePath) & ": count/search missed an entry.")
                                    For Each hit In hits
                                        Dim logical = CStr(hit.GetType().GetProperty("LogicalPath").GetValue(hit))
                                        Ensure(logical.StartsWith(archivePath & " | ", StringComparison.Ordinal) AndAlso Not logical.Contains("BeaconTar_"), "Temporary TAR name leaked into results.")
                                    Next
                                End Sub, expectedCount:=2)
                        Next
                        ' GZIP containing text must remain a normal single-entry archive.
                        Dim plain = Path.Combine(root, "standalone.log")
                        File.WriteAllText(plain, Payload, New UTF8Encoding(False))
                        Dim gzipPath = plain & ".gz"
                        Run7Zip(sevenZip, root, {"a", "-tgzip", gzipPath, plain})
                        VerifyEntries(gzipPath, {"standalone.log"})
                        SearchChecks.CheckPipeline(gzipPath, "archive upgrade marker", SearchMode.PlainText, Options(),
                            Sub(hits) Ensure(hits.Count = 1, "Standalone GZIP text was not searched."), expectedCount:=1)
                    End Sub)
    End Sub

    Public Sub RarPayloads()
        Dim expected As New Dictionary(Of String, String) From {
            {"exe/test.exe", "8557928804F57ECC340B3BB38B095A3607474EC8DEB0076F316FCFE02B562106"},
            {"jpg/test.jpg", "B251C7501FB0F55DD4A92FEABE0A6F5733BC40A02679498155FAE9B30138FC53"},
            {"тест.txt", "4D581D93D369F6E1C9B295FF38D82DABD577F927DFAF0C35818C015C85E322D9"}}
        For Each name In {"Rar.rar", "Rar.solid.rar", "Rar5.rar", "Rar5.solid.rar"}
            Dim path = IO.Path.Combine(AppContext.BaseDirectory, "Fixtures", "Archives", name)
            Using archive = ArchiveCompatibility.OpenArchive(path, Options(), CancellationToken.None)
                Dim count As Integer = 0
                For Each entry In archive.Entries.Where(Function(item) Not item.IsDirectory).ToArray()
                    Using input = ArchiveCompatibility.OpenEntry(entry)
                        Dim hash = Convert.ToHexString(System.Security.Cryptography.SHA256.HashData(input))
                        Ensure(expected(entry.Key.Replace("\", "/")) = hash, name & ": decompressed content hash mismatch.")
                        count += 1
                    End Using
                Next
                Ensure(count = 3, name & ": missing fixture entries.")
            End Using
            SearchChecks.CheckPipeline(path, "тест.txt", SearchMode.PlainText,
                New BeaconSettings With {.SearchFileContents = False, .SearchFileNames = True},
                Sub(hits) Ensure(hits.Count = 1, "RAR filename search failed."), expectedCount:=3)
        Next
    End Sub

    Public Sub DirectoryTraversalAndPreviewLimits()
        WithFixture(Sub(root)
            Dim path = IO.Path.Combine(root, "directory-traversal.zip")
            Using zip = ZipFile.Open(path, ZipArchiveMode.Create)
                zip.CreateEntry("../outside/")
                Using writer As New StreamWriter(zip.CreateEntry("safe.log").Open())
                    writer.Write(Payload)
                End Using
            End Using
            SearchChecks.CheckPipeline(path, "archive upgrade marker", SearchMode.PlainText, ArchiveChecks.Options(),
                Sub(hits) Ensure(hits.Count = 0, "Unsafe directory entry was not rejected before scanning."), expectedCount:=0)
            Ensure(Not Directory.Exists(IO.Path.Combine(root, "..", "outside")), "Traversal created a directory outside the fixture.")
            Dim large = IO.Path.Combine(root, "preview.zip")
            Using zip = ZipFile.Open(large, ZipArchiveMode.Create)
                Using writer As New StreamWriter(zip.CreateEntry("large.log").Open())
                    writer.Write(New String("x"c, 2 * 1024 * 1024))
                End Using
            End Using
            Dim options = ArchiveChecks.Options()
            options.MaximumPreviewSizeMb = 1
            options.SearchFileContents = False
            options.SearchFileNames = True
            SearchChecks.CheckPipeline(large, "large.log", SearchMode.PlainText, options, Sub(hits) Ensure(hits.Count = 1, "Preview limit fixture missing."),
                Sub(window, hits)
                    GetType(MainWindow).GetMethod("LoadTextFromArchive", BindingFlags.Instance Or BindingFlags.NonPublic).Invoke(window, {large, "large.log"})
                    PreviewTestHelpers.WaitForPreview(window)
                    Dim preview = DirectCast(window.FindName("TextPreview_rtb"), System.Windows.Controls.RichTextBox)
                    Dim text = New System.Windows.Documents.TextRange(preview.Document.ContentStart, preview.Document.ContentEnd).Text
                    Ensure(text.StartsWith(New String("x"c, 1000), StringComparison.Ordinal) AndAlso
                           text.Contains("Preview truncated", StringComparison.Ordinal) AndAlso
                           text.Length >= 1024 * 1024 AndAlso text.Length < 1024 * 1024 + 500,
                           "Oversized archive preview lost readable content or bypassed its display cap.")
                    Ensure(Not text.Contains("Error loading", StringComparison.Ordinal), "A normal preview cap was reported as an archive error.")
                    Dim disk = IO.Path.Combine(root, "large.log")
                    File.WriteAllText(disk, New String("x"c, 2 * 1024 * 1024))
                    GetType(MainWindow).GetMethod("LoadTextFromDisk", BindingFlags.Instance Or BindingFlags.NonPublic).Invoke(window, {disk})
                    PreviewTestHelpers.WaitForPreview(window)
                    text = New System.Windows.Documents.TextRange(preview.Document.ContentStart, preview.Document.ContentEnd).Text
                    Ensure(text.StartsWith(New String("x"c, 1000), StringComparison.Ordinal) AndAlso text.Contains("Preview truncated"), "Disk preview did not retain bounded content.")
                    Dim truncatedPreview As New PreviewText("<html><script>not executed</script><p>readable prefix", True, 100)
                    Dim work = DirectCast(GetType(MainWindow).GetMethod("ShowWebPreviewAsync", BindingFlags.Instance Or BindingFlags.NonPublic).Invoke(window, {truncatedPreview, ".html"}), System.Threading.Tasks.Task)
                    work.GetAwaiter().GetResult()
                    PreviewTestHelpers.WaitForPreview(window)
                    text = New System.Windows.Documents.TextRange(preview.Document.ContentStart, preview.Document.ContentEnd).Text
                    Ensure(text.Contains(truncatedPreview.Text) AndAlso text.Contains("Preview truncated") AndAlso preview.IsVisible,
                           "Truncated HTML was not shown as readable, labeled plain text.")
                End Sub, expectedCount:=1)
        End Sub)
    End Sub

    Public Sub PreviewAndNested(sevenZip As String)
        WithFixture(Sub(root)
                        Dim tarPath = Path.Combine(root, "logs.tar")
                        WriteTar(tarPath, {"sample.log"})
                        Dim compressed = tarPath & ".gz"
                        Run7Zip(sevenZip, root, {"a", "-tgzip", compressed, tarPath})
                        SearchChecks.CheckPipeline(compressed, "archive upgrade marker", SearchMode.PlainText, Options(),
                            Sub(hits) Ensure(hits.Count = 1, "Compressed TAR search failed."),
                            Sub(window, hits)
                                Dim flags = BindingFlags.Instance Or BindingFlags.NonPublic
                                GetType(MainWindow).GetMethod("LoadTextFromArchive", flags).Invoke(window, {compressed, "sample.log"})
                                PreviewTestHelpers.WaitForPreview(window)
                                Dim preview = DirectCast(window.FindName("TextPreview_rtb"), System.Windows.Controls.RichTextBox)
                                Dim text = New System.Windows.Documents.TextRange(preview.Document.ContentStart, preview.Document.ContentEnd).Text
                                Ensure(text.Contains("archive upgrade marker café 日本語"), "Compressed TAR preview could not reopen the entry.")
                            End Sub, expectedCount:=1)
                        Dim nested = Path.Combine(root, "outer.zip")
                        Using zip = ZipFile.Open(nested, ZipArchiveMode.Create)
                            zip.CreateEntryFromFile(compressed, "inner.tar.gz")
                        End Using
                        SearchChecks.CheckPipeline(nested, "archive upgrade marker", SearchMode.PlainText, Options(),
                            Sub(hits) Ensure(hits.Count = 1, "Nested compressed TAR search failed."), expectedCount:=1)
                        Dim noNesting = Options()
                        noNesting.ArchiveNestingDepth = 0
                        SearchChecks.CheckPipeline(nested, "archive upgrade marker", SearchMode.PlainText, noNesting,
                            Sub(hits) Ensure(hits.Count = 0, "Nested TAR bypassed the nesting limit."), expectedCount:=0)
                        ' A compression wrapper itself must not consume archive nesting depth.
                        SearchChecks.CheckPipeline(compressed, "archive upgrade marker", SearchMode.PlainText, noNesting,
                            Sub(hits) Ensure(hits.Count = 1, "Compression wrapper consumed nesting depth."), expectedCount:=1)
                    End Sub)
    End Sub

    Public Sub ZipAttributes()
        WithFixture(Sub(root)
                        Dim path = IO.Path.Combine(root, "attributes.zip")
                        Using zip = ZipFile.Open(path, ZipArchiveMode.Create)
                            For Each pair In New Dictionary(Of String, FileAttributes) From {{"visible.log", FileAttributes.Normal}, {"hidden.log", FileAttributes.Hidden}, {"system.log", FileAttributes.System}}
                                Dim entry = zip.CreateEntry(pair.Key)
                                entry.ExternalAttributes = CInt(pair.Value)
                                Using writer As New StreamWriter(entry.Open())
                                    writer.Write(Payload)
                                End Using
                            Next
                        End Using
                        For Each include In {False, True}
                            Dim settings = Options()
                            settings.IncludeHiddenFiles = include
                            settings.IncludeSystemFiles = include
                            Dim expected = If(include, 3, 1)
                            SearchChecks.CheckPipeline(path, "archive upgrade marker", SearchMode.PlainText, settings,
                                Sub(hits) Ensure(hits.Count = expected, "ZIP hidden/system settings were not honored."), expectedCount:=expected)
                        Next
                    End Sub)
    End Sub

    Public Sub RejectionAndChecksums(sevenZip As String)
        WithFixture(Sub(root)
                        Dim plain = Path.Combine(root, "sample.log")
                        File.WriteAllText(plain, Payload)
                        For Each formatName As String In {"zip", "7z"}
                            Dim encrypted = Path.Combine(root, "encrypted." & formatName)
                            Run7Zip(sevenZip, root, {"a", "-t" & formatName, "-pfixture-only", encrypted, plain})
                            Using archive = ArchiveCompatibility.OpenArchive(encrypted, Options(), CancellationToken.None)
                                Dim entry = archive.Entries.First(Function(item) Not item.IsDirectory)
                                Ensure(entry.IsEncrypted, "Encrypted " & formatName & " metadata was lost.")
                                Dim budget As New ArchiveReadBudget(Options(), New FileInfo(encrypted).Length)
                                Expect(Of InvalidDataException)(Sub() budget.Register(entry.Key, entry.Size, entry.CompressedSize, entry.IsEncrypted, entry.LinkTarget))
                            End Using
                        Next
                        Dim unsafePath = Path.Combine(root, "unsafe.zip")
                        Using zip = ZipFile.Open(unsafePath, ZipArchiveMode.Create)
                            Using writer As New StreamWriter(zip.CreateEntry("../escape.log").Open())
                                writer.Write(Payload)
                            End Using
                        End Using
                        Using archive = ArchiveCompatibility.OpenArchive(unsafePath, Options(), CancellationToken.None)
                            Dim entry = archive.Entries.Single()
                            Dim budget As New ArchiveReadBudget(Options(), 10000)
                            Expect(Of InvalidDataException)(Sub() budget.Register(entry.Key, entry.Size, entry.CompressedSize, entry.IsEncrypted, entry.LinkTarget))
                        End Using
                        Dim linked = Path.Combine(root, "links.tar")
                        Using output = File.Create(linked)
                            Using tar As New TarWriter(output, TarEntryFormat.Pax, True)
                                tar.WriteEntry(New PaxTarEntry(TarEntryType.SymbolicLink, "linked.log") With {.LinkName = "../outside.log"})
                            End Using
                        End Using
                        Using archive = ArchiveCompatibility.OpenArchive(linked, Options(), CancellationToken.None)
                            Dim entry = archive.Entries.Single()
                            Ensure(Not String.IsNullOrEmpty(entry.LinkTarget), "TAR link target was lost.")
                            Dim budget As New ArchiveReadBudget(Options(), 10000)
                            Expect(Of InvalidDataException)(Sub() budget.Register(entry.Key, entry.Size, entry.CompressedSize, entry.IsEncrypted, entry.LinkTarget))
                        End Using
                        Dim corrupt = Path.Combine(root, "checksum.tar.gz")
                        Using output = File.Create(corrupt)
                            Using gzip As New GZipStream(output, CompressionMode.Compress)
                                Using data As New MemoryStream(Encoding.UTF8.GetBytes(Payload))
                                    gzip.Write(data.ToArray())
                                End Using
                            End Using
                        End Using
                        Dim bytes = File.ReadAllBytes(corrupt)
                        ' GZIP's trailer contains CRC32 followed by ISIZE.
                        bytes(bytes.Length - 8) = bytes(bytes.Length - 8) Xor CByte(1)
                        File.WriteAllBytes(corrupt, bytes)
                        Expect(Of InvalidDataException)(Sub()
                            Using archive = ArchiveCompatibility.OpenArchive(corrupt, Options(), CancellationToken.None)
                                Throw New InvalidOperationException("Corrupt GZIP checksum was accepted.")
                            End Using
                        End Sub)
                    End Sub)
    End Sub

    Public Sub LimitsAndCleanup(sevenZip As String)
        WithFixture(Sub(root)
                        Dim tarPath = Path.Combine(root, "large.tar")
                        Using output = File.Create(tarPath)
                            Using tar As New TarWriter(output, TarEntryFormat.Pax, True)
                                Using data As New MemoryStream(Encoding.UTF8.GetBytes(New String("x"c, 2 * 1024 * 1024)))
                                    tar.WriteEntry(New PaxTarEntry(TarEntryType.RegularFile, "large.log") With {.DataStream = data})
                                End Using
                            End Using
                        End Using
                        Dim compressed = tarPath & ".gz"
                        Run7Zip(sevenZip, root, {"a", "-tgzip", compressed, tarPath})
                        Dim before = Directory.GetFiles(Path.GetTempPath(), "BeaconTar_*.tar").ToHashSet(StringComparer.OrdinalIgnoreCase)
                        Dim settings = Options()
                        settings.MaximumArchiveExpandedSizeMb = 1
                        Expect(Of InvalidDataException)(Sub()
                            Using archive = ArchiveCompatibility.OpenArchive(compressed, settings, CancellationToken.None)
                                Throw New InvalidOperationException("Over-limit TAR opened.")
                            End Using
                        End Sub)
                        settings = Options()
                        settings.MaximumCompressionRatio = 2
                        Expect(Of InvalidDataException)(Sub()
                            Using archive = ArchiveCompatibility.OpenArchive(compressed, settings, CancellationToken.None)
                                Throw New InvalidOperationException("Compression ratio limit bypassed.")
                            End Using
                        End Sub)
                        Using cancelled As New CancellationTokenSource()
                            cancelled.Cancel()
                            Expect(Of OperationCanceledException)(Sub()
                                Using archive = ArchiveCompatibility.OpenArchive(compressed, Options(), cancelled.Token)
                                    Throw New InvalidOperationException("Cancellation ignored.")
                                End Using
                            End Sub)
                        End Using
                        Using archive = ArchiveCompatibility.OpenArchive(compressed, Options(), CancellationToken.None)
                            Dim entry = archive.Entries.Single()
                            Dim entrySettings = Options()
                            entrySettings.MaximumArchiveEntrySizeMb = 1
                            Dim budget As New ArchiveReadBudget(entrySettings, New FileInfo(compressed).Length)
                            Expect(Of InvalidDataException)(Sub() budget.Register(entry.Key, entry.Size, entry.CompressedSize, entry.IsEncrypted, entry.LinkTarget))
                        End Using
                        Ensure(Directory.GetFiles(Path.GetTempPath(), "BeaconTar_*.tar").All(Function(path) before.Contains(path)), "Temporary TAR leaked after success/failure/cancellation.")
                        Using exclusive As New FileStream(compressed, FileMode.Open, FileAccess.ReadWrite, FileShare.None)
                            Ensure(exclusive.Length > 0, "Archive file handle was not released.")
                        End Using
                    End Sub)
    End Sub

    Private Sub VerifyEntries(path As String, names As String())
        ' Search and preview use separate archive instances; mirror that ownership model.
        For Each asyncRead In {False, True}
            Using archive = ArchiveCompatibility.OpenArchive(path, Options(), CancellationToken.None)
                Dim entries = archive.Entries.Where(Function(entry) Not entry.IsDirectory).ToArray()
                Ensure(entries.Length = names.Length, IO.Path.GetFileName(path) & ": unexpected entry count.")
                For Each entry In entries.Reverse()
                    Ensure(names.Contains(entry.Key.Replace("\", "/")), "Entry name changed: " & entry.Key)
                    Using input = ArchiveCompatibility.OpenEntry(entry), output As New MemoryStream()
                        If asyncRead Then
                            input.CopyToAsync(output).GetAwaiter().GetResult()
                            Ensure(Encoding.UTF8.GetString(output.ToArray()) = Payload, IO.Path.GetFileName(path) & ": async/reopened content changed.")
                        Else
                            input.CopyTo(output)
                            Ensure(Encoding.UTF8.GetString(output.ToArray()) = Payload, IO.Path.GetFileName(path) & ": sync content changed.")
                        End If
                    End Using
                Next
            End Using
        Next
        Using exclusive As New FileStream(path, FileMode.Open, FileAccess.ReadWrite, FileShare.None)
            Ensure(exclusive.Length > 0, "Original archive remained locked.")
        End Using
    End Sub

    Private Sub WriteTar(path As String, names As String())
        Using output = File.Create(path)
            Using tar As New TarWriter(output, TarEntryFormat.Pax, True)
                For Each name In names
                    Using data As New MemoryStream(Encoding.UTF8.GetBytes(Payload))
                        tar.WriteEntry(New PaxTarEntry(TarEntryType.RegularFile, name) With {.DataStream = data})
                    End Using
                Next
            End Using
        End Using
    End Sub

    Private Function Options() As BeaconSettings
        Return New BeaconSettings With {.StopAfterFirstMatchPerFile = False, .MaximumCompressionRatio = 100000}
    End Function

    Private Sub Run7Zip(exe As String, directory As String, arguments As String())
        Dim start As New Diagnostics.ProcessStartInfo(exe) With {.WorkingDirectory = directory, .UseShellExecute = False, .CreateNoWindow = True}
        For Each argument In arguments
            start.ArgumentList.Add(argument)
        Next
        Using child = Diagnostics.Process.Start(start)
            If Not child.WaitForExit(15000) Then
                child.Kill(True)
                Throw New TimeoutException("Archive fixture creation timed out.")
            End If
            Ensure(child.ExitCode = 0, "7-Zip fixture creation failed.")
        End Using
    End Sub

    Private Sub WithFixture(check As Action(Of String))
        Dim root = Path.Combine(Path.GetTempPath(), "BeaconArchiveChecks-" & Guid.NewGuid().ToString("N"))
        Directory.CreateDirectory(root)
        Try
            check(root)
        Finally
            Directory.Delete(root, True)
        End Try
    End Sub

    Private Sub Ensure(condition As Boolean, message As String)
        If Not condition Then Throw New InvalidOperationException(message)
    End Sub

    Private Sub Expect(Of T As Exception)(action As Action)
        Try
            action()
        Catch ex As T
            Return
        End Try
        Throw New InvalidOperationException("Expected " & GetType(T).Name)
    End Sub
End Module
