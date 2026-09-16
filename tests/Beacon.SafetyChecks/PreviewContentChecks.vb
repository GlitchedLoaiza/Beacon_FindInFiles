Imports System.Diagnostics
Imports System.IO
Imports System.IO.Compression
Imports System.Text
Imports System.Threading
Imports Beacon

Module PreviewContentChecks
    Public Sub Run()
        Dim root = Path.Combine(Path.GetTempPath(), "BeaconPreviewService-" & Guid.NewGuid().ToString("N"))
        Directory.CreateDirectory(root)
        Try
            Dim options As New BeaconSettings With {.MaximumPreviewSizeMb = 1, .MaximumFileSizeMb = 4,
                .MaximumArchiveEntrySizeMb = 4, .MaximumCompressionRatio = 10000}
            Dim service As New PreviewContentService(options)
            options.MaximumPreviewSizeMb = 3
            Dim large = Path.Combine(root, "large.log")
            File.WriteAllText(large, New String("a"c, 1048575) & "😀tail", New UTF8Encoding(False))
            Dim preview = service.ReadFile(large)
            Require(preview.IsTruncated AndAlso preview.Text.Length = 1048575 AndAlso preview.ByteLimit = 1048576, "Service lost settings snapshot or Unicode-safe prefix.")
            Dim small = Path.Combine(root, "sample.log")
            Dim text = "error café 日本語" & vbLf & "last line"
            File.WriteAllText(small, text, Encoding.Unicode)
            Require(service.ReadFile(small).Text = text, "Disk preview decoding changed.")
            Dim zipPath = Path.Combine(root, "sample.zip")
            Using zip = ZipFile.Open(zipPath, ZipArchiveMode.Create)
                zip.CreateEntryFromFile(small, "logs/sample.log")
                zip.CreateEntryFromFile(large, "large.log")
            End Using
            Require(service.ReadArchive(zipPath, "logs/sample.log").Text = text, "Archive preview changed selected content.")
            Require(service.ReadArchive(zipPath, "absent.log") Is Nothing, "Missing archive entry did not return Nothing.")
            Require(service.ReadArchive(zipPath, "large.log").IsTruncated, "Archive preview lost soft truncation.")
            Using exclusive As New FileStream(zipPath, FileMode.Open, FileAccess.ReadWrite, FileShare.None)
                Require(exclusive.Length > 0, "Archive preview left a handle open.")
            End Using
            Dim errors As New List(Of String)()
            Dim hard As New PreviewContentService(New BeaconSettings With {.MaximumFileSizeMb = 1, .MaximumPreviewSizeMb = 3}, Sub(source, ex) errors.Add(source))
            Reject(Of InvalidDataException)(Sub() hard.ReadFile(large))
            Require(errors.SequenceEqual({large}), "Hard read failure lost source diagnostics.")
            Using cancelled As New CancellationTokenSource()
                cancelled.Cancel()
                Reject(Of OperationCanceledException)(Sub() service.ReadFile(small, cancelled.Token))
                Reject(Of OperationCanceledException)(Sub() service.ReadArchive(zipPath, "logs/sample.log", cancelled.Token))
                Reject(Of OperationCanceledException)(Sub() service.ReadCab("missing.cab", "sample.log", cancelled.Token))
            End Using
            Dim cab = Path.Combine(root, "sample.cab")
            Dim start As New ProcessStartInfo(Path.Combine(Environment.SystemDirectory, "makecab.exe")) With {.UseShellExecute = False, .CreateNoWindow = True}
            start.ArgumentList.Add(small)
            start.ArgumentList.Add(cab)
            Using child = Process.Start(start)
                If Not child.WaitForExit(15000) Then
                    child.Kill(True)
                    Throw New TimeoutException("CAB preview fixture creation timed out.")
                End If
                Require(child.ExitCode = 0, "CAB fixture creation failed.")
            End Using
            Dim before = Directory.GetDirectories(Path.GetTempPath(), "BeaconCabPreview_*").ToHashSet(StringComparer.OrdinalIgnoreCase)
            Require(service.ReadCab(cab, "sample.log").Text = text, "CAB preview decoding changed.")
            Require(service.ReadCab(cab, "missing.log") Is Nothing, "Missing CAB entry was not handled.")
            Reject(Of InvalidDataException)(Sub() service.ReadCab(cab, "../outside.log"))
            Dim corrupt = Path.Combine(root, "bad.cab")
            File.WriteAllText(corrupt, "invalid cabinet")
            Reject(Of InvalidDataException)(Sub() service.ReadCab(corrupt, "sample.log"))
            Require(Directory.GetDirectories(Path.GetTempPath(), "BeaconCabPreview_*").All(Function(folder) before.Contains(folder)), "CAB preview temporary directory leaked.")
        Finally
            Directory.Delete(root, True)
        End Try
    End Sub

    Private Sub Require(condition As Boolean, message As String)
        If Not condition Then Throw New InvalidOperationException(message)
    End Sub
    Private Sub Reject(Of T As Exception)(action As Action)
        Try
            action()
        Catch ex As T
            Return
        End Try
        Throw New InvalidOperationException("Expected " & GetType(T).Name)
    End Sub
End Module
