Imports System.IO
Imports System.Threading

Namespace Beacon
    Public NotInheritable Class ArchiveReadBudget
        Private ReadOnly _settings As BeaconSettings
        Private ReadOnly _archiveSize As Long
        Private _entries As Integer
        Private _declaredBytes As Long
        Private _readBytes As Long

        Public Sub New(settings As BeaconSettings, archiveSize As Long)
            _settings = settings
            _archiveSize = Math.Max(1, archiveSize)
        End Sub

        Public Sub Register(name As String, size As Long, compressedSize As Long, encrypted As Boolean,
                            Optional linkTarget As String = Nothing)
            ArchiveSafety.ValidateEntryName(name)
            If encrypted Then Throw New InvalidDataException("Encrypted archive entries are not scanned.")
            If Not String.IsNullOrEmpty(linkTarget) Then Throw New InvalidDataException("Archive links are not extracted.")
            _entries += 1
            If _entries > _settings.MaximumArchiveEntries Then Throw New InvalidDataException("Archive entry count limit reached.")
            If size < 0 OrElse size > EntryLimit Then Throw New InvalidDataException("Archive entry size limit reached.")
            If compressedSize > 0 AndAlso CDec(size) > CDec(compressedSize) * _settings.MaximumCompressionRatio Then
                Throw New InvalidDataException("Archive compression ratio limit reached.")
            End If
            If size > TotalLimit - _declaredBytes Then Throw New InvalidDataException("Archive expanded size or compression ratio limit reached.")
            _declaredBytes += size
        End Sub

        Public ReadOnly Property EntryLimit As Long
            Get
                Return CLng(Math.Min(_settings.MaximumFileSizeMb, _settings.MaximumArchiveEntrySizeMb)) * 1024 * 1024
            End Get
        End Property

        Private ReadOnly Property TotalLimit As Long
            Get
                Return CLng(Math.Min(CDec(_settings.MaximumArchiveExpandedSizeMb) * 1024 * 1024,
                                     CDec(_archiveSize) * _settings.MaximumCompressionRatio))
            End Get
        End Property

        Public Sub Consume(count As Integer)
            If count > TotalLimit - _readBytes Then Throw New InvalidDataException("Actual archive expansion limit reached.")
            _readBytes += count
        End Sub
    End Class

    Public NotInheritable Class ArchiveSafety
        Private Sub New()
        End Sub

        Public Shared Sub ValidateEntryName(name As String)
            If String.IsNullOrWhiteSpace(name) OrElse Path.IsPathRooted(name) OrElse name.Contains(":"c) Then
                Throw New InvalidDataException("Unsafe archive entry path.")
            End If
            For Each part In name.Replace("\", "/").TrimEnd("/"c).Split("/"c)
                If part = "." Then Continue For
                If part.Length = 0 OrElse part = ".." OrElse
                   part.EndsWith(" ", StringComparison.Ordinal) OrElse part.EndsWith(".", StringComparison.Ordinal) OrElse
                   part.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0 Then
                    Throw New InvalidDataException("Unsafe archive entry path.")
                End If
                Dim stem = part.Split("."c)(0).ToUpperInvariant()
                If {"CON", "PRN", "AUX", "NUL", "CONIN$", "CONOUT$"}.Contains(stem) OrElse
                   System.Text.RegularExpressions.Regex.IsMatch(stem, "^(COM|LPT)[1-9¹²³]$") Then
                    Throw New InvalidDataException("Reserved device name in archive.")
                End If
            Next
        End Sub

        Public Shared Function IsExcluded(name As String, settings As BeaconSettings) As Boolean
            Dim parts = name.Replace("\", "/").Split("/"c)
            Dim excluded As New HashSet(Of String)(settings.ExcludedDirectories.Split(";"c, StringSplitOptions.RemoveEmptyEntries), StringComparer.OrdinalIgnoreCase)
            Return parts.Take(Math.Max(0, parts.Length - 1)).Any(Function(part) excluded.Contains(part))
        End Function
    End Class

    Public NotInheritable Class BoundedReadStream
        Inherits Stream

        Private ReadOnly _inner As Stream
        Private ReadOnly _limit As Long
        Private ReadOnly _budget As ArchiveReadBudget
        Private ReadOnly _token As CancellationToken
        Private ReadOnly _onError As Action(Of Exception)
        Private _bytesRead As Long

        Public Sub New(inner As Stream, limit As Long, token As CancellationToken,
                       Optional budget As ArchiveReadBudget = Nothing, Optional onError As Action(Of Exception) = Nothing)
            _inner = inner
            _limit = limit
            _token = token
            _budget = budget
            _onError = onError
        End Sub

        Public Overrides Function Read(buffer As Byte(), offset As Integer, count As Integer) As Integer
            _token.ThrowIfCancellationRequested()
            Try
                Dim readCount = _inner.Read(buffer, offset, CInt(Math.Min(CLng(count), Math.Max(0, _limit - _bytesRead) + 1)))
                If readCount > _limit - _bytesRead Then Throw New InvalidDataException("Actual entry/file read limit reached.")
                _budget?.Consume(readCount)
                _bytesRead += readCount
                Return readCount
            Catch ex As Exception When TypeOf ex Is IOException OrElse TypeOf ex Is InvalidDataException
                _onError?.Invoke(ex)
                Throw
            End Try
        End Function

        Public Overrides ReadOnly Property CanRead As Boolean = True
        Public Overrides ReadOnly Property CanSeek As Boolean = False
        Public Overrides ReadOnly Property CanWrite As Boolean = False
        Public Overrides ReadOnly Property Length As Long
            Get
                Throw New NotSupportedException()
            End Get
        End Property
        Public Overrides Property Position As Long
            Get
                Return _bytesRead
            End Get
            Set(value As Long)
                Throw New NotSupportedException()
            End Set
        End Property
        Public Overrides Sub Flush()
        End Sub
        Public Overrides Function Seek(offset As Long, origin As SeekOrigin) As Long
            Throw New NotSupportedException()
        End Function
        Public Overrides Sub SetLength(value As Long)
            Throw New NotSupportedException()
        End Sub
        Public Overrides Sub Write(buffer As Byte(), offset As Integer, count As Integer)
            Throw New NotSupportedException()
        End Sub
        Protected Overrides Sub Dispose(disposing As Boolean)
            If disposing Then _inner.Dispose()
            MyBase.Dispose(disposing)
        End Sub
    End Class
End Namespace
