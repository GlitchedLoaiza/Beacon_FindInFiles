Imports System.IO
Imports System.Threading
Imports SharpCompress.Archives
Imports SharpCompress.Archives.Tar
Imports SharpCompress.Common
Imports SharpCompress.Providers
Imports SharpCompress.Readers

Namespace Beacon
    ' SharpCompress 0.50 no longer unwraps compressed TARs through ArchiveFactory.
    ' Keep random-access entries for search and preview without unbounded buffering.
    Friend NotInheritable Class ArchiveCompatibility
        Private Shared ReadOnly Providers As CompressionProviderRegistry = CompressionProviderRegistry.Default.With(
            New SharpCompress.Providers.System.SystemGZipCompressionProvider())

        Private Sub New()
        End Sub

        Public Shared Function OpenArchive(path As String, settings As BeaconSettings, token As CancellationToken) As IArchive
            token.ThrowIfCancellationRequested()
            Using source = File.OpenRead(path)
                Dim signature(5) As Byte
                Dim signatureLength = ReadPrefix(source, signature)
                source.Position = 0
                Dim compression As CompressionType
                If signatureLength >= 2 AndAlso signature(0) = &H1F AndAlso signature(1) = &H8B Then
                    compression = CompressionType.GZip
                ElseIf signatureLength >= 3 AndAlso signature(0) = &H42 AndAlso signature(1) = &H5A AndAlso signature(2) = &H68 Then
                    compression = CompressionType.BZip2
                ElseIf signatureLength = 6 AndAlso signature.SequenceEqual(New Byte() {&HFD, &H37, &H7A, &H58, &H5A, &H0}) Then
                    compression = CompressionType.Xz
                Else
                    Return ArchiveFactory.OpenArchive(path, New ReaderOptions With {.Providers = Providers})
                End If

                Dim temporaryPath = IO.Path.Combine(IO.Path.GetTempPath(), "BeaconTar_" & Guid.NewGuid().ToString("N") & ".tar")
                Dim temporary As New FileStream(temporaryPath, FileMode.CreateNew, FileAccess.ReadWrite, FileShare.None,
                                                81920, FileOptions.DeleteOnClose)
                Dim archiveOwnsTemporary As Boolean
                Try
                    Dim budget As New ArchiveReadBudget(settings, source.Length)
                    ' Avoid SharpCompress's GZIP partial-read disposal behavior masking a
                    ' quota/cancellation exception. This only unwraps the compression layer;
                    ' SharpCompress still provides the TAR entries and random-access reader.
                    Using decompressed = Providers.CreateDecompressStream(compression, source)
                        Using bounded As New BoundedReadStream(decompressed, CLng(settings.MaximumArchiveExpandedSizeMb) * 1024 * 1024, token, budget)
                            bounded.CopyTo(temporary)
                        End Using
                    End Using
                    token.ThrowIfCancellationRequested()
                    temporary.Position = 0
                    Dim header(511) As Byte
                    Dim headerLength = ReadPrefix(temporary, header)
                    temporary.Position = 0
                    If headerLength < header.Length OrElse Not IsTarHeader(header) Then
                        token.ThrowIfCancellationRequested()
                        Return ArchiveFactory.OpenArchive(path, New ReaderOptions With {.Providers = Providers})
                    End If
                    token.ThrowIfCancellationRequested()
                    temporary.Position = 0
                    Dim archive = TarArchive.OpenArchive(temporary, New ReaderOptions With {.LeaveStreamOpen = False})
                    archiveOwnsTemporary = True
                    Return archive
                Finally
                    If Not archiveOwnsTemporary Then temporary.Dispose()
                End Try
            End Using
        End Function

        Private Shared Function ReadPrefix(source As Stream, buffer As Byte()) As Integer
            Dim count = 0
            While count < buffer.Length
                Dim read = source.Read(buffer, count, buffer.Length - count)
                If read = 0 Then Exit While
                count += read
            End While
            Return count
        End Function

        Private Shared Function IsTarHeader(header As Byte()) As Boolean
            If header.All(Function(value) value = 0) Then Return True

            Dim checksum As Integer = 0
            Dim foundDigit As Boolean
            For index = 148 To 155
                Dim value = header(index)
                If value = 0 OrElse value = 32 Then Continue For
                If value < 48 OrElse value > 55 Then Return False
                foundDigit = True
                checksum = checksum * 8 + value - 48
            Next

            Dim unsignedSum As Integer = 0
            Dim signedSum As Integer = 0
            For index = 0 To 511
                Dim value = If(index >= 148 AndAlso index <= 155, 32, CInt(header(index)))
                unsignedSum += value
                signedSum += If(value > 127, value - 256, value)
            Next

            Return foundDigit AndAlso (checksum = unsignedSum OrElse checksum = signedSum)
        End Function

        Public Shared Function OpenEntry(entry As IArchiveEntry) As Stream
            Dim input = entry.OpenEntryStream()
            If entry.Archive IsNot Nothing AndAlso entry.Archive.Type = ArchiveType.Zip Then
                Return New VerifiedZipStream(input, entry.Size, entry.Crc)
            End If
            Return input
        End Function
    End Class

    Friend NotInheritable Class VerifiedZipStream
        Inherits Stream

        Private ReadOnly _inner As Stream
        Private ReadOnly _size As Long
        Private ReadOnly _expectedCrc As Long
        Private ReadOnly _crc As New SharpCompress.Compressors.Deflate.CRC32()
        Private _read As Long
        Private _completed As Boolean

        Public Sub New(inner As Stream, size As Long, crc As Long)
            _inner = inner
            _size = size
            _expectedCrc = crc And &HFFFFFFFFL
        End Sub

        Public Overrides Function Read(buffer As Byte(), offset As Integer, count As Integer) As Integer
            Dim actual = _inner.Read(buffer, offset, count)
            If count = 0 Then Return actual
            If actual > 0 Then
                _crc.SlurpBlock(buffer, offset, actual)
                _read += actual
                If _read > _size Then
                    Throw New InvalidDataException("ZIP entry size or CRC32 mismatch.")
                End If
            Else
                ValidateCompleted()
            End If
            Return actual
        End Function

        Private Sub ValidateCompleted()
            If _completed Then Return
            Dim actualCrc = CLng(_crc.Crc32Result) And &HFFFFFFFFL
            If _read <> _size OrElse actualCrc <> _expectedCrc Then
                Throw New InvalidDataException("ZIP entry size or CRC32 mismatch.")
            End If
            _completed = True
        End Sub

        Public Overrides ReadOnly Property CanRead As Boolean
            Get
                Return _inner.CanRead
            End Get
        End Property

        Public Overrides ReadOnly Property CanSeek As Boolean
            Get
                Return False
            End Get
        End Property

        Public Overrides ReadOnly Property CanWrite As Boolean
            Get
                Return False
            End Get
        End Property

        Public Overrides ReadOnly Property Length As Long
            Get
                Throw New NotSupportedException()
            End Get
        End Property

        Public Overrides Property Position As Long
            Get
                Return _read
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
