Imports System.IO
Imports System.Xml
Imports System.Xml.Linq

Namespace Beacon
    Public NotInheritable Class PreviewFormatting
        Private Sub New()
        End Sub

        Public Const HtmlCanvasStyle As String = "<style id='beacon-canvas'>@layer beacon-preview-defaults { html { background-color: white; color: black; color-scheme: light; } }</style>"

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
