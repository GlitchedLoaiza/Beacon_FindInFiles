Imports System.IO
Imports System.IO.Compression
Imports System.Text
Imports System.Threading
Imports Beacon

Module ZipChecksumChecks
    Public Sub Run()
        ' Standard CRC32 check vector; expected value does not use the implementation under test.
        Dim payload = Encoding.ASCII.GetBytes("123456789")
        For Each asyncRead In {False, True}
            Using input As New VerifiedZipStream(New MemoryStream(payload), payload.Length, &HCBF43926L)
                Dim buffer(15) As Byte
                Require(input.Read(buffer, 0, 0) = 0, "Zero-length read changed state.")
                Using output As New MemoryStream()
                    If asyncRead Then
                        input.CopyToAsync(output).GetAwaiter().GetResult()
                    Else
                        input.CopyTo(output)
                    End If
                    Require(output.ToArray().SequenceEqual(payload), "Verified ZIP bytes changed.")
                End Using
                Require(input.Read(buffer, 0, 1) = 0, "Repeated valid EOF failed.")
            End Using
        Next
        Using input As New VerifiedZipStream(New MemoryStream(), 0, 0)
            Require(input.ReadByte() = -1, "Empty ZIP entry failed CRC validation.")
        End Using
        For Each size In {CLng(payload.Length), CLng(payload.Length + 1), CLng(payload.Length - 1)}
            Using input As New VerifiedZipStream(New MemoryStream(payload), size, 0)
                Reject(Sub() input.CopyTo(Stream.Null))
                Reject(Sub() input.ReadByte())
            End Using
        Next
        ' A bounded preview must not drain the rest of an entry just to validate CRC.
        Using input As New VerifiedZipStream(New MemoryStream(payload), payload.Length, 0)
            Require(input.ReadByte() = payload(0), "Partial ZIP read failed prematurely.")
        End Using

        Dim root = Path.Combine(Path.GetTempPath(), "BeaconZipChecksum-" & Guid.NewGuid().ToString("N"))
        Directory.CreateDirectory(root)
        Try
            Dim path = IO.Path.Combine(root, "corrupt.zip")
            Using zip = ZipFile.Open(path, ZipArchiveMode.Create)
                Using output = zip.CreateEntry("sample.log", CompressionLevel.NoCompression).Open()
                    output.Write(payload, 0, payload.Length)
                End Using
            End Using
            Dim bytes = File.ReadAllBytes(path)
            Dim start = 30 + CInt(BitConverter.ToUInt16(bytes, 26)) + CInt(BitConverter.ToUInt16(bytes, 28))
            bytes(start) = bytes(start) Xor CByte(1)
            File.WriteAllBytes(path, bytes)
            Using archive = ArchiveCompatibility.OpenArchive(path, New BeaconSettings(), CancellationToken.None)
                Using input = ArchiveCompatibility.OpenEntry(archive.Entries.Single())
                    Reject(Sub() input.CopyToAsync(Stream.Null).GetAwaiter().GetResult())
                End Using
            End Using
        Finally
            Directory.Delete(root, True)
        End Try
    End Sub

    Private Sub Require(condition As Boolean, message As String)
        If Not condition Then Throw New InvalidOperationException(message)
    End Sub

    Private Sub Reject(action As Action)
        Try
            action()
        Catch ex As InvalidDataException
            Return
        End Try
        Throw New InvalidOperationException("Expected ZIP size/CRC rejection.")
    End Sub
End Module
