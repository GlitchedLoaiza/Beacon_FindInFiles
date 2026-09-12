Imports System.IO
Imports System.Text
Imports System.Threading
Imports Beacon

Module PreviewLimitChecks
    Public Sub Run()
        For Each content In {"", "abc", "abcd"}
            Using source As New MemoryStream(Encoding.UTF8.GetBytes(content))
                Dim result = PreviewFormatting.ReadTextPreview(source, 4)
                Require(result.Text = content AndAlso Not result.IsTruncated AndAlso result.DisplayText = content, "Complete/exact-limit preview changed.")
            End Using
        Next
        Using source As New MemoryStream(Encoding.UTF8.GetBytes("abcdef"))
            Dim result = PreviewFormatting.ReadTextPreview(source, 4)
            Require(result.Text = "abcd" AndAlso result.IsTruncated AndAlso result.DisplayText.Contains("Preview truncated"), "Oversized preview lost its prefix or notice.")
            Require(source.Position = 5, "Preview read beyond its one-byte lookahead.")
        End Using
        For Each encoding As Encoding In {New UTF8Encoding(True), Encoding.Unicode, Encoding.BigEndianUnicode, Encoding.UTF32, New UTF32Encoding(True, True)}
            Dim content = "first line" & vbLf & "café 日本語 😀"
            Dim bytes = encoding.GetPreamble().Concat(encoding.GetBytes(content)).ToArray()
            Using source As New MemoryStream(bytes)
                Dim result = PreviewFormatting.ReadTextPreview(source, bytes.Length)
                Require(result.Text = content AndAlso Not result.IsTruncated, "BOM-aware preview decoding changed.")
            End Using
            Dim prefix = encoding.GetPreamble().Concat(encoding.GetBytes("A😀B")).ToArray()
            Dim cap = encoding.GetPreamble().Length + encoding.GetByteCount("A") + 1
            Using source As New MemoryStream(prefix)
                Dim result = PreviewFormatting.ReadTextPreview(source, cap)
                Require(result.Text = "A" AndAlso result.IsTruncated, "Truncated preview fabricated a partial Unicode character.")
            End Using
        Next
        ' Chunk boundaries and short underlying reads must not change decoding.
        Dim longText = New String("a"c, 8191) & "😀日本語"
        Using source As New MemoryStream(Encoding.UTF8.GetBytes(longText))
            Require(PreviewFormatting.ReadTextPreview(source, 9000).Text = longText, "Decoder lost a character across read chunks.")
        End Using
        Using source As New ShortReadStream(Encoding.UTF8.GetBytes("café😀"))
            Require(PreviewFormatting.ReadTextPreview(source, 100).Text = "café😀", "Short reads were mistaken for EOF.")
        End Using
        ' Only the display cap is soft: hard bounds and CRC failures must still propagate.
        Using source As New BoundedReadStream(New MemoryStream(Encoding.UTF8.GetBytes("abcde")), 3, CancellationToken.None)
            Reject(Of InvalidDataException)(Sub() PreviewFormatting.ReadTextPreview(source, 10))
        End Using
        Using source As New VerifiedZipStream(New MemoryStream(Encoding.ASCII.GetBytes("123456789")), 9, 0)
            Reject(Of InvalidDataException)(Sub() PreviewFormatting.ReadTextPreview(source, 9))
        End Using
        Using cancelled As New CancellationTokenSource()
            cancelled.Cancel()
            Using source As New BoundedReadStream(New MemoryStream({CByte(1)}), 1, cancelled.Token)
                Reject(Of OperationCanceledException)(Sub() PreviewFormatting.ReadTextPreview(source, 10))
            End Using
        End Using
    End Sub

    Private NotInheritable Class ShortReadStream
        Inherits MemoryStream
        Public Sub New(bytes As Byte())
            MyBase.New(bytes)
        End Sub
        Public Overrides Function Read(buffer As Byte(), offset As Integer, count As Integer) As Integer
            Return MyBase.Read(buffer, offset, Math.Min(count, 1))
        End Function
    End Class

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
