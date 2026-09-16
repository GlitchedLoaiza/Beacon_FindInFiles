Imports System.IO
Imports System.Threading
Imports System.Threading.Tasks

Namespace Beacon
    Partial Public Class MainWindow
        Private _isCountingFiles As Boolean
        Private ReadOnly _sourceSearches As New List(Of SourceSearchService)()

        Private Async Function CountSourceFilesAsync(root As String, ct As CancellationToken) As Task
            _isCountingFiles = True
            _totalFilesToScan = 0
            Try
                Using service = CreateSourceService()
                    Await service.RunAsync(root, True, ct)
                End Using
            Finally
                CleanupTemp()
                _isCountingFiles = False
            End Try
        End Function

        Private Sub CountSourceFile()
            If _isCountingFiles Then
                _totalFilesToScan += 1
            Else
                _filesScanned += 1
            End If
            UpdateScanProgress()
        End Sub

        Private Function DisplaySourcePath(logicalPath As String) As String
            Dim separator = logicalPath.IndexOf(" | ", StringComparison.Ordinal)
            ' A file-started scan has no folder root. Its outer archive is already
            ' selected, so show entry paths and retain only nested archive prefixes.
            If String.IsNullOrEmpty(_scanRootFolder) AndAlso separator >= 0 Then
                Return logicalPath.Substring(separator + 3)
            End If
            If _searchOptions.DisplayAbsolutePaths Then Return logicalPath
            Dim root = If(separator >= 0, logicalPath.Substring(0, separator), logicalPath)
            Return MakeRelativeDisplay(_scanRootFolder, root) & If(separator >= 0, logicalPath.Substring(separator), "")
        End Function

        Private Sub PublishSourceHit(hit As SearchHit)
            If _isCountingFiles Then Return
            If hit.Details.Count = 0 Then Return
            If hit.PartialReason.Contains("limit", StringComparison.OrdinalIgnoreCase) Then Interlocked.Increment(_structuredLimitFiles)
            AddHit(hit)
        End Sub

        Private Async Function SearchSourceAsync(root As String, ct As CancellationToken) As Task
            Dim service = CreateSourceService()
            _sourceSearches.Add(service)
            Try
                Await service.RunAsync(root, False, ct)
            Finally
                If service.ResultLimitReached Then Volatile.Write(_resultLimitReached, True)
            End Try
        End Function

        Private Function CreateSourceService() As SourceSearchService
            Return New SourceSearchService(_activeQuery, _searchOptions, _supportedArchiveExt,
                AddressOf PublishServiceResult,
                Sub(logical, processed)
                    ThrottledSetCurrentFileDisplay(DisplaySourcePath(logical))
                    If processed Then CountSourceFile()
                End Sub,
                Sub(logical, ex, stage, severity)
                    If stage = "Search" Then
                        RecordFileSystemIssue(logical, ex)
                    Else
                        RecordDiagnostic(logical, ex, stage, severity)
                    End If
                End Sub, AddressOf HandleAccessDeniedPath, AddressOf RememberSourcePath)
        End Function

        Private Sub PublishServiceResult(source As SourceSearchResult)
            Dim kind As HitKind
            Select Case source.Kind
                Case SourceResultKind.ArchiveText : kind = HitKind.ZipTextEntry
                Case SourceResultKind.DiskEvent : kind = HitKind.EvtxFileOnDisk
                Case SourceResultKind.ArchiveEvent : kind = HitKind.EvtxFromZipTemp
                Case SourceResultKind.DiskHar : kind = HitKind.HarFileOnDisk
                Case SourceResultKind.ArchiveHar : kind = HitKind.HarFromZip
                Case Else : kind = HitKind.DiskTextFile
            End Select
            Dim hit As New SearchHit With {.Kind = kind, .FilePath = source.PhysicalPath, .LogicalPath = source.LogicalPath,
                .DisplayName = DisplaySourcePath(source.LogicalPath), .ZipPath = source.ArchivePath, .ZipEntryName = source.EntryName,
                .TempEvtxPath = source.TemporaryEventPath, .PartialReason = source.PartialReason}
            hit.Details.AddRange(source.Details)
            hit.MatchingEvents.AddRange(source.Events)
            hit.MatchingRequests.AddRange(source.Requests)
            PublishSourceHit(hit)
        End Sub

    End Class
End Namespace
