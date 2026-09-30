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
            Using directory = PrivateTemporaryDirectory.Create("BeaconCabPreview_")
                If Not SevenZipHelper.ExtractCab(cabPath, directory.DirectoryPath, _options, token, _report) Then
                    Throw New InvalidDataException("Failed to extract CAB file.")
                End If
                token.ThrowIfCancellationRequested()
                Dim extracted = Path.GetFullPath(Path.Combine(directory.DirectoryPath, entryName))
                Dim root = directory.DirectoryPath.TrimEnd(Path.DirectorySeparatorChar) & Path.DirectorySeparatorChar
                If Not extracted.StartsWith(root, StringComparison.OrdinalIgnoreCase) Then Throw New InvalidDataException("Unsafe CAB preview entry path.")
                If Not File.Exists(extracted) Then Return Nothing
                Dim reader As New PreviewContentService(_options, Sub(source, ex) _report?.Invoke(cabPath & " | " & entryName, ex))
                Return reader.ReadFile(extracted, token)
            End Using
        End Function
    End Class
End Namespace
