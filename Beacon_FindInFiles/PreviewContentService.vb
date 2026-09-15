Imports System.Diagnostics
Imports System.IO
Imports System.Threading
Imports SharpCompress.Archives

Namespace Beacon
    Public NotInheritable Class PreviewContentService
        Private ReadOnly _options As BeaconSettings
        Private ReadOnly _report As Action(Of String, Exception)

        Public Sub New(options As BeaconSettings, Optional report As Action(Of String, Exception) = Nothing)
            _options = BeaconSettingsService.Clone(options)
            _report = report
        End Sub

        Public Function ReadFile(filePath As String, Optional token As CancellationToken = Nothing) As PreviewText
            token.ThrowIfCancellationRequested()
            Using stream As New BoundedReadStream(New FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite),
                                                 CLng(_options.MaximumFileSizeMb) * 1024 * 1024, token,
                                                 onError:=Sub(ex) _report?.Invoke(filePath, ex))
                Return PreviewFormatting.ReadTextPreview(stream, CLng(_options.MaximumPreviewSizeMb) * 1024 * 1024)
            End Using
        End Function

        Public Function ReadArchive(archivePath As String, entryName As String, Optional token As CancellationToken = Nothing) As PreviewText
            token.ThrowIfCancellationRequested()
            Using archive = ArchiveCompatibility.OpenArchive(archivePath, _options, token)
                For Each entry In archive.Entries
                    token.ThrowIfCancellationRequested()
                    If entry.Key = entryName Then Return ReadEntry(entry, archivePath, token)
                Next
            End Using
            Return Nothing
        End Function

        Private Function ReadEntry(entry As IArchiveEntry, archivePath As String, token As CancellationToken) As PreviewText
            Dim budget As New ArchiveReadBudget(_options, New FileInfo(archivePath).Length)
            budget.Register(entry.Key, entry.Size, entry.CompressedSize, entry.IsEncrypted, entry.LinkTarget)
            Using stream As New BoundedReadStream(ArchiveCompatibility.OpenEntry(entry), Math.Min(entry.Size, budget.EntryLimit), token, budget,
                                                 Sub(ex) _report?.Invoke(archivePath & " | " & entry.Key, ex))
                Return PreviewFormatting.ReadTextPreview(stream, CLng(_options.MaximumPreviewSizeMb) * 1024 * 1024)
            End Using
        End Function

        Public Function ReadCab(cabPath As String, entryName As String, Optional token As CancellationToken = Nothing) As PreviewText
            token.ThrowIfCancellationRequested()
            ArchiveSafety.ValidateEntryName(entryName)
            Dim directory = Path.Combine(Path.GetTempPath(), "BeaconCabPreview_" & Guid.NewGuid().ToString("N"))
            Try
                If Not SevenZipHelper.ExtractCab(cabPath, directory, _options, token, _report) Then
                    Throw New InvalidDataException("Failed to extract CAB file.")
                End If
                token.ThrowIfCancellationRequested()
                Dim extracted = Path.GetFullPath(Path.Combine(directory, entryName))
                Dim root = Path.GetFullPath(directory).TrimEnd(Path.DirectorySeparatorChar) & Path.DirectorySeparatorChar
                If Not extracted.StartsWith(root, StringComparison.OrdinalIgnoreCase) Then Throw New InvalidDataException("Unsafe CAB preview entry path.")
                If Not File.Exists(extracted) Then Return Nothing
                Dim reader As New PreviewContentService(_options, Sub(source, ex) _report?.Invoke(cabPath & " | " & entryName, ex))
                Return reader.ReadFile(extracted, token)
            Finally
                If System.IO.Directory.Exists(directory) Then
                    Try
                        System.IO.Directory.Delete(directory, True)
                    Catch ex As IOException
                        Debug.WriteLine($"Could not remove CAB preview directory: {ex.Message}")
                    Catch ex As UnauthorizedAccessException
                        Debug.WriteLine($"Could not remove CAB preview directory: {ex.Message}")
                    End Try
                End If
            End Try
        End Function
    End Class
End Namespace
