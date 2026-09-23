Imports System.IO
Imports System.IO.Compression
Imports System.Reflection
Imports System.Text
Imports System.Windows.Controls
Imports Beacon

Module ArchiveDisplayChecks
    Private Const Marker As String = "archive display marker"

    Public Sub Run()
        Dim root = Path.Combine(Path.GetTempPath(), "BeaconArchiveDisplay-" & Guid.NewGuid().ToString("N"))
        Directory.CreateDirectory(root)
        Try
            Dim deepest = ZipBytes(New Dictionary(Of String, Byte()) From {{"deep.log", Encoding.UTF8.GetBytes(Marker)}})
            Dim inner = ZipBytes(New Dictionary(Of String, Byte()) From {
                {"logs/nested.log", Encoding.UTF8.GetBytes(Marker)}, {"outer.zip", deepest}})
            Dim archivePath = Path.Combine(root, "outer.zip")
            File.WriteAllBytes(archivePath, ZipBytes(New Dictionary(Of String, Byte()) From {
                {"logs with spaces/café.log", Encoding.UTF8.GetBytes(Marker)}, {"bundles/inner.zip", inner}}))
            Dim expected = New HashSet(Of String)(StringComparer.Ordinal) From {
                "logs with spaces/café.log", "bundles/inner.zip | logs/nested.log", "bundles/inner.zip | outer.zip | deep.log"}
            For Each absolute In {False, True}
                Dim settings As New BeaconSettings With {.ArchiveNestingDepth = 3, .DisplayAbsolutePaths = absolute}
                SearchChecks.CheckPipeline(archivePath, Marker, SearchMode.PlainText, settings,
                    Sub(hits)
                        Require(hits.Count = 3, "Archive display fixture lost matches.")
                        Require(expected.SetEquals(hits.Select(Function(hit) Value(hit, "DisplayName"))), "Selected outer archive leaked into labels or nested names were lost.")
                        For Each hit In hits
                            Require(Value(hit, "LogicalPath") = archivePath & " | " & Value(hit, "DisplayName"), "Full logical source path was modified.")
                            Require(Path.IsPathFullyQualified(Value(hit, "ZipPath")), "Archive navigation path was shortened.")
                        Next
                    End Sub,
                    Sub(window, hits)
                        Dim preview = DirectCast(window.FindName("TextPreview_rtb"), RichTextBox)
                        For Each hit In hits
                            GetType(MainWindow).GetMethod("LoadTextFromArchive", BindingFlags.Instance Or BindingFlags.NonPublic).
                                Invoke(window, {Value(hit, "ZipPath"), Value(hit, "ZipEntryName")})
                            PreviewTestHelpers.WaitForPreview(window)
                            Dim text = New System.Windows.Documents.TextRange(preview.Document.ContentStart, preview.Document.ContentEnd).Text
                            Require(text.Contains(Marker), "Shortened labels broke direct or nested preview reopening.")
                        Next
                    End Sub, expectedCount:=3)
                SearchChecks.CheckPipeline(root, Marker, SearchMode.PlainText, settings,
                    Sub(hits)
                        Dim prefix = If(absolute, archivePath, "outer.zip") & " | "
                        Require(expected.SetEquals(hits.Select(Function(hit) Value(hit, "DisplayName").Substring(prefix.Length))), "Folder-started labels changed nested paths.")
                        Require(hits.All(Function(hit) Value(hit, "DisplayName").StartsWith(prefix, StringComparison.Ordinal)), "Folder-started labels lost the outer archive.")
                    End Sub, expectedCount:=3)
            Next
            Dim pathOnly As New BeaconSettings With {.ArchiveNestingDepth = 3, .SearchFileContents = False, .SearchFullPaths = True}
            SearchChecks.CheckPipeline(archivePath, "outer.zip", SearchMode.PlainText, pathOnly,
                Sub(hits)
                    Require(hits.Any(Function(hit) Value(hit, "LogicalPath") = archivePath & " | logs with spaces/café.log" AndAlso
                                                  Value(hit, "DisplayName") = "logs with spaces/café.log"), "Full-path matching no longer sees the selected archive name.")
                End Sub, expectedCount:=3)
        Finally
            Directory.Delete(root, True)
        End Try
    End Sub

    Private Function ZipBytes(entries As Dictionary(Of String, Byte())) As Byte()
        Using output As New MemoryStream()
            Using archive As New ZipArchive(output, ZipArchiveMode.Create, True)
                For Each entry In entries
                    Using stream = archive.CreateEntry(entry.Key).Open()
                        stream.Write(entry.Value, 0, entry.Value.Length)
                    End Using
                Next
            End Using
            Return output.ToArray()
        End Using
    End Function

    Private Function Value(hit As Object, name As String) As String
        Return CStr(hit.GetType().GetProperty(name).GetValue(hit))
    End Function

    Private Sub Require(condition As Boolean, message As String)
        If Not condition Then Throw New InvalidOperationException(message)
    End Sub
End Module
