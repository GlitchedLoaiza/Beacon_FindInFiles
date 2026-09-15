Namespace Beacon
    Public Class EventRecordSummary
        Public Property Level As String
        Public Property Provider As String
        Public Property EventId As Integer
        Public Property TimeCreated As DateTime?
        Public Property Message As String
        Public Property LevelNumber As Byte?
        Public Property RawXml As String = ""
        Public Property XmlShortened As Boolean
        Public Property MessageUnavailable As Boolean
    End Class

    Public NotInheritable Class StructuredSearchResult(Of T)
        Public ReadOnly Property Records As New List(Of T)()
        Public ReadOnly Property Details As New List(Of SearchDetail)()
        Public Property PartialReason As String = ""
    End Class
End Namespace
