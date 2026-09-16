Imports System.IO
Imports System.Text
Imports System.Xml
Imports System.Xml.Linq

Namespace Beacon
    Public NotInheritable Class PreviewText
        Public ReadOnly Property Text As String
        Public ReadOnly Property IsTruncated As Boolean
        Public ReadOnly Property ByteLimit As Long

        Public Sub New(text As String, truncated As Boolean, byteLimit As Long)
            Me.Text = text
            IsTruncated = truncated
            Me.ByteLimit = byteLimit
        End Sub

        Public ReadOnly Property DisplayText As String
            Get
                If Not IsTruncated Then Return Text
                Return Text & vbCrLf & vbCrLf & $"[Preview truncated at {ByteLimit:N0} bytes. More content is available in the source file. Increase Maximum preview size in Settings to view more.]"
            End Get
        End Property
    End Class

    Public NotInheritable Class PreviewFormatting
        Private Sub New()
        End Sub

        Public Const HtmlCanvasStyle As String = "<style id='beacon-canvas'>@layer beacon-preview-defaults { html { background-color: white; color: black; color-scheme: light; } }</style>"

        Public Shared Function ReadTextPreview(source As Stream, byteLimit As Long) As PreviewText
            If byteLimit < 1 OrElse byteLimit > Integer.MaxValue Then Throw New ArgumentOutOfRangeException(NameOf(byteLimit))
            Using prefix As New MemoryStream()
                Dim buffer(8191) As Byte
                While prefix.Length < byteLimit
                    Dim count = source.Read(buffer, 0, CInt(Math.Min(buffer.Length, byteLimit - prefix.Length)))
                    If count = 0 Then Exit While
                    prefix.Write(buffer, 0, count)
                End While
                ' A single lookahead distinguishes a complete exact-limit file from a prefix.
                ' The supplied source still enforces hard budgets, cancellation and integrity.
                Dim truncated = prefix.Length = byteLimit AndAlso source.ReadByte() <> -1
                Dim bytes = prefix.GetBuffer()
                Dim length = CInt(prefix.Length)
                Dim encoding As Encoding = Encoding.UTF8
                Dim offset As Integer
                If length >= 4 AndAlso bytes(0) = &HFF AndAlso bytes(1) = &HFE AndAlso bytes(2) = 0 AndAlso bytes(3) = 0 Then
                    encoding = Encoding.UTF32
                    offset = 4
                ElseIf length >= 4 AndAlso bytes(0) = 0 AndAlso bytes(1) = 0 AndAlso bytes(2) = &HFE AndAlso bytes(3) = &HFF Then
                    encoding = New UTF32Encoding(True, True)
                    offset = 4
                ElseIf length >= 3 AndAlso bytes(0) = &HEF AndAlso bytes(1) = &HBB AndAlso bytes(2) = &HBF Then
                    offset = 3
                ElseIf length >= 2 AndAlso bytes(0) = &HFF AndAlso bytes(1) = &HFE Then
                    encoding = Encoding.Unicode
                    offset = 2
                ElseIf length >= 2 AndAlso bytes(0) = &HFE AndAlso bytes(1) = &HFF Then
                    encoding = Encoding.BigEndianUnicode
                    offset = 2
                End If
                Dim decoder = encoding.GetDecoder()
                Dim chars(encoding.GetMaxCharCount(buffer.Length) - 1) As Char
                Dim text As New StringBuilder()
                While offset < length
                    Dim count = Math.Min(buffer.Length, length - offset)
                    Dim decoded = decoder.GetChars(bytes, offset, count, chars, 0, Not truncated AndAlso offset + count = length)
                    text.Append(chars, 0, decoded)
                    offset += count
                End While
                ' Do not flush an incomplete code point at a truncated byte boundary.
                Return New PreviewText(text.ToString(), truncated, byteLimit)
            End Using
        End Function

        Public Shared Function FormatXml(content As String, prettyPrint As Boolean, maxCharacters As Long) As String
            If Not prettyPrint Then Return content
            Try
                Dim settings As New XmlReaderSettings With {
                    .DtdProcessing = DtdProcessing.Prohibit,
                    .XmlResolver = Nothing,
                    .MaxCharactersInDocument = maxCharacters
                }
                Using source As New StringReader(content)
                    Using reader = XmlReader.Create(source, settings)
                        Dim document = XDocument.Load(reader)
                        Return If(document.Declaration Is Nothing, "", document.Declaration.ToString() & Environment.NewLine) & document.ToString()
                    End Using
                End Using
            Catch ex As XmlException
                Return content
            End Try
        End Function
    End Class
End Namespace
