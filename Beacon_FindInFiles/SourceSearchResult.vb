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
        Private ReadOnly _directories As New List(Of PrivateTemporaryDirectory)()
        Private _disposed As Boolean

        Public ReadOnly Property HasSources As Boolean
            Get
                Return _directories.Count > 0
            End Get
        End Property

        Public Function CreateDirectory() As String
            ObjectDisposedException.ThrowIf(_disposed, Me)
            Dim directory = PrivateTemporaryDirectory.Create("BeaconSearch_")
            _directories.Add(directory)
            Return directory.DirectoryPath
        End Function

        Public Sub Dispose() Implements IDisposable.Dispose
            If _disposed Then Return
            _disposed = True
            For Each directory In _directories
                directory.Dispose()
            Next
            _directories.Clear()
        End Sub
    End Class
End Namespace
