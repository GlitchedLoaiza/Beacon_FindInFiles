Imports System.Diagnostics
Imports System.IO
Imports System.Text

Namespace Beacon
    Public NotInheritable Class WelcomeTourState
        Private Sub New()
        End Sub

        Public Shared ReadOnly Property MarkerPath As String
            Get
                Return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Beacon", "welcome-tour.offered")
            End Get
        End Property

        Public Shared Function TryClaimOffer(Optional marker As String = Nothing) As Boolean
            Dim destination = If(marker, MarkerPath)
            Try
                Directory.CreateDirectory(Path.GetDirectoryName(destination))
                ' CreateNew is atomic: skipping, closing or crashing after the offer must not prompt again.
                Using stream As New FileStream(destination, FileMode.CreateNew, FileAccess.Write, FileShare.None)
                    Dim bytes = Encoding.UTF8.GetBytes("Beacon welcome tour offered " & DateTimeOffset.UtcNow.ToString("O"))
                    stream.Write(bytes, 0, bytes.Length)
                End Using
                Return True
            Catch ex As Exception When TypeOf ex Is IOException OrElse TypeOf ex Is UnauthorizedAccessException
                Debug.WriteLine($"Welcome tour not offered: {ex.Message}")
                Return False
            End Try
        End Function
    End Class

    Public NotInheritable Class TourStep
        Public ReadOnly Property Title As String
        Public ReadOnly Property Target As String
        Public ReadOnly Property Instructions As String
        Public Sub New(title As String, target As String, instructions As String)
            Me.Title = title
            Me.Target = target
            Me.Instructions = instructions
        End Sub
    End Class

    Public NotInheritable Class WelcomeTour
        Private Sub New()
        End Sub

        Public Shared Function Steps() As TourStep()
            Return {
                New TourStep("Choose where to search", "SourceActionsPanel", "Click Folder to choose a folder, or Archive to choose a compressed archive. Path displays your choice." & vbCrLf & vbCrLf & "You can try the controls during this tour, or use Next just to learn. The tour never selects files for you."),
                New TourStep("Enter what you want to find", "Search_txt", "Enter a word such as ""error"" in Search for. Do not type the surrounding quotation marks in this example." & vbCrLf & vbCrLf & "Nothing is searched yet. You decide when to start."),
                New TourStep("Choose a simple search mode", "SearchMode_cmb", "Start with Literal text. Leave Case sensitive unchecked to find both ""error"" and ""ERROR""." & vbCrLf & vbCrLf & "You can explore other modes later. Help includes a beginner regex guide and examples."),
                New TourStep("Start the search", "Scan_btn", "Click Scan when your source and search text are ready. If Scan is disabled, choose a valid source and enter search text first." & vbCrLf & vbCrLf & "Beacon counts eligible files, then searches them. Large archives may take time. During scanning, the button becomes Cancel; you can also press Esc to stop. Click Next whenever you want to continue the tour."),
                New TourStep("Read the matching files", "Results_lst", "Matching files appear in the left pane. Select one to see its preview on the right. No results may mean your word was not found or a filter/limit excluded it." & vbCrLf & vbCrLf & "A + beside a count means the results are partial; hover for the reason. The tour does not generate sample results."),
                New TourStep("Jump to an individual match", "MatchDetails_exp", "Expand Match details below the file list, then select a line, event or request to jump to it. Use the preview navigation buttons to explore further matches." & vbCrLf & vbCrLf & "Event tools and HAR tools narrow the results you see, not the results included in an export. If no file is selected yet, these details may be empty."),
                New TourStep("Export a readable report", "ExportResults_btn", "After collecting results, click Export… below the preview. Choose an HTML file name and location in the save dialog." & vbCrLf & vbCrLf & "The report includes matching files and nearby context, and opens in a browser. Export may be disabled until results exist. The tour never opens the save dialog or writes a report for you. Review private data before sharing—even when HAR redaction is enabled."),
                New TourStep("Check problems and get more help", "Diagnostics_btn", "Diagnostics explains skipped files, errors and limits. Use it when you expected more results." & vbCrLf & vbCrLf & "The ? Help button beside Settings provides the complete offline guide. Settings controls filters, limits, appearance and manual update checks. You are ready to use Beacon; click Finish to close the tour.")
            }
        End Function
    End Class
End Namespace
