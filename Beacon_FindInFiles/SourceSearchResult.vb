Imports System.Diagnostics
Imports System.IO

Namespace Beacon
    Public Enum SourceResultKind
        DiskText
        ArchiveText
        DiskEvent
        ArchiveEvent
        DiskHar
        ArchiveHar
    End Enum

    Public NotInheritable Class SourceSearchResult
        Public Property PhysicalPath As String
        Public Property LogicalPath As String
        Public Property ArchivePath As String
        Public Property EntryName As String
        Public Property TemporaryEventPath As String
        Public Property Kind As SourceResultKind
        Public Property PartialReason As String = ""
        Public ReadOnly Property Details As New List(Of SearchDetail)()
        Public ReadOnly Property Events As New List(Of EventRecordSummary)()
        Public ReadOnly Property Requests As New List(Of HarRecord)()
    End Class

    Friend NotInheritable Class SearchTemporarySources
        Implements IDisposable
        Private ReadOnly _directories As New List(Of String)()
        Private _disposed As Boolean

        Public ReadOnly Property HasSources As Boolean
            Get
                Return _directories.Count > 0
            End Get
        End Property

        Public Function CreateDirectory() As String
            ObjectDisposedException.ThrowIf(_disposed, Me)
            Dim path = IO.Path.Combine(IO.Path.GetTempPath(), "BeaconSearch_" & Guid.NewGuid().ToString("N"))
            Directory.CreateDirectory(path)
            _directories.Add(path)
            Return path
        End Function

        Public Sub Dispose() Implements IDisposable.Dispose
            If _disposed Then Return
            _disposed = True
            For Each path In _directories
                Try
                    Directory.Delete(path, True)
                Catch ex As IOException
                    Debug.WriteLine($"Could not remove search temporary directory: {ex.Message}")
                Catch ex As UnauthorizedAccessException
                    Debug.WriteLine($"Could not remove search temporary directory: {ex.Message}")
                End Try
            Next
            _directories.Clear()
        End Sub
    End Class
End Namespace
