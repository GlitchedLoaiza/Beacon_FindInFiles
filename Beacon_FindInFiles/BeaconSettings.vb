Imports System.IO
Imports System.Text.Json
Imports System.Text.Json.Serialization

Namespace Beacon

    Public Enum AccessDeniedAction
        SkipAndReport
        AskToTakeOwnership
    End Enum

    Public Enum SearchMode
        PlainText
        ExactWord
        RegularExpression
        AnyTerm
        AllTerms
    End Enum

    Public Enum EvtxResourcePolicy
        OfflineOnly
        AskForOfficialSource
        AutomaticOfficialSources
    End Enum

    Public Enum DiagnosticDetail
        Summary
        Detailed
    End Enum

    Public Class BeaconSettings
        Public Property MaximumStructuredMatches As Integer = 300
        Public Property MaximumTotalResults As Integer = 10000
        Public Property MaximumFileSizeMb As Integer = 500
        Public Property IncludedExtensions As String = ".txt;.log;.json;.xml;.csv;.html;.reg;.ini;.cfg;.config;.nfo"
        Public Property ExcludedDirectories As String = ".git;bin;obj;.vs;node_modules"
        Public Property IncludeHiddenFiles As Boolean = True
        Public Property IncludeSystemFiles As Boolean = False
        Public Property FollowReparsePoints As Boolean = False
        Public Property AccessDeniedBehavior As AccessDeniedAction = AccessDeniedAction.AskToTakeOwnership

        Public Property ArchiveNestingDepth As Integer = 1
        Public Property MaximumArchiveEntries As Integer = 25000
        Public Property MaximumArchiveEntrySizeMb As Integer = 500
        Public Property MaximumArchiveExpandedSizeMb As Integer = 2048
        Public Property MaximumCompressionRatio As Integer = 100
        Public Property ArchiveProcessTimeoutSeconds As Integer = 60

        Public Property EvtxMaximumMatches As Integer = 300
        Public Property ShowRawXmlWhenMessageUnavailable As Boolean = True
        Public Property EvtxMessageResourceBehavior As EvtxResourcePolicy = EvtxResourcePolicy.OfflineOnly

        Public Property HarMaximumMatches As Integer = 300
        Public Property MaximumHarBodySizeMb As Integer = 25
        Public Property DecodeBase64HarBodies As Boolean = True
        Public Property RedactSensitiveHarData As Boolean = True

        Public Property PreviewFontFamily As String = "Consolas"
        Public Property PreviewFontSize As Integer = 13
        Public Property PreviewWordWrap As Boolean = False
        Public Property MaximumPreviewSizeMb As Integer = 25
        Public Property PrettyPrintJsonAndXml As Boolean = True
        Public Property SelectFirstResult As Boolean = True
        Public Property DisplayAbsolutePaths As Boolean = False

        Public Property DefaultSearchMode As SearchMode = SearchMode.PlainText
        Public Property SearchFileContents As Boolean = True
        Public Property SearchFileNames As Boolean = False
        Public Property SearchFullPaths As Boolean = False
        Public Property StopAfterFirstMatchPerFile As Boolean = True
        Public Property ScanWorkerCount As Integer = 0
        Public Property Diagnostics As DiagnosticDetail = DiagnosticDetail.Summary

        Public Shared Function CreateDefaults() As BeaconSettings
            Return New BeaconSettings()
        End Function
    End Class

    Public NotInheritable Class BeaconSettingsService
        Private Shared ReadOnly SerializerOptions As New JsonSerializerOptions With {
            .WriteIndented = True,
            .PropertyNameCaseInsensitive = True
        }

        Shared Sub New()
            SerializerOptions.Converters.Add(New JsonStringEnumConverter())
        End Sub

        Private Sub New()
        End Sub

        Public Shared ReadOnly Property SettingsPath As String
            Get
                Return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                                    "Beacon", "settings.json")
            End Get
        End Property

        Public Shared Function Load() As BeaconSettings
            Try
                If Not File.Exists(SettingsPath) Then Return BeaconSettings.CreateDefaults()
                Dim settings = JsonSerializer.Deserialize(Of BeaconSettings)(File.ReadAllText(SettingsPath), SerializerOptions)
                Return Validate(If(settings, BeaconSettings.CreateDefaults()))
            Catch
                Return BeaconSettings.CreateDefaults()
            End Try
        End Function

        Public Shared Sub Save(settings As BeaconSettings)
            Dim validated = Validate(Clone(settings))
            Dim directoryPath = Path.GetDirectoryName(SettingsPath)
            Directory.CreateDirectory(directoryPath)
            Dim pendingPath = SettingsPath & ".new"
            File.WriteAllText(pendingPath, JsonSerializer.Serialize(validated, SerializerOptions))
            File.Move(pendingPath, SettingsPath, True)
        End Sub

        Public Shared Function Clone(settings As BeaconSettings) As BeaconSettings
            Dim json = JsonSerializer.Serialize(settings, SerializerOptions)
            Return JsonSerializer.Deserialize(Of BeaconSettings)(json, SerializerOptions)
        End Function

        Public Shared Function Validate(settings As BeaconSettings) As BeaconSettings
            If settings Is Nothing Then settings = BeaconSettings.CreateDefaults()

            settings.MaximumStructuredMatches = Math.Clamp(settings.MaximumStructuredMatches, 1, 100000)
            settings.MaximumTotalResults = Math.Clamp(settings.MaximumTotalResults, 1, 1000000)
            settings.MaximumFileSizeMb = Math.Clamp(settings.MaximumFileSizeMb, 1, 102400)
            settings.ArchiveNestingDepth = Math.Clamp(settings.ArchiveNestingDepth, 0, 5)
            settings.MaximumArchiveEntries = Math.Clamp(settings.MaximumArchiveEntries, 1, 1000000)
            settings.MaximumArchiveEntrySizeMb = Math.Clamp(settings.MaximumArchiveEntrySizeMb, 1, 102400)
            settings.MaximumArchiveExpandedSizeMb = Math.Clamp(settings.MaximumArchiveExpandedSizeMb, 1, 1024000)
            settings.MaximumCompressionRatio = Math.Clamp(settings.MaximumCompressionRatio, 1, 10000)
            settings.ArchiveProcessTimeoutSeconds = Math.Clamp(settings.ArchiveProcessTimeoutSeconds, 5, 3600)
            settings.EvtxMaximumMatches = Math.Clamp(settings.EvtxMaximumMatches, 1, settings.MaximumStructuredMatches)
            settings.HarMaximumMatches = Math.Clamp(settings.HarMaximumMatches, 1, settings.MaximumStructuredMatches)
            settings.MaximumHarBodySizeMb = Math.Clamp(settings.MaximumHarBodySizeMb, 1, 10240)
            settings.PreviewFontSize = Math.Clamp(settings.PreviewFontSize, 8, 48)
            settings.MaximumPreviewSizeMb = Math.Clamp(settings.MaximumPreviewSizeMb, 1, 1024)
            settings.ScanWorkerCount = Math.Clamp(settings.ScanWorkerCount, 0, 64)
            settings.PreviewFontFamily = If(String.IsNullOrWhiteSpace(settings.PreviewFontFamily), "Consolas", settings.PreviewFontFamily.Trim())
            settings.IncludedExtensions = NormalizeList(settings.IncludedExtensions, ensureExtensionPrefix:=True)
            settings.ExcludedDirectories = NormalizeList(settings.ExcludedDirectories, ensureExtensionPrefix:=False)
            Return settings
        End Function

        Private Shared Function NormalizeList(value As String, ensureExtensionPrefix As Boolean) As String
            If String.IsNullOrWhiteSpace(value) Then Return String.Empty

            Dim values = value.Split(New Char() {";"c, ","c, ControlChars.Cr, ControlChars.Lf}, StringSplitOptions.RemoveEmptyEntries).
                Select(Function(item) item.Trim()).
                Where(Function(item) item.Length > 0).
                Select(Function(item) If(ensureExtensionPrefix AndAlso Not item.StartsWith("."), "." & item, item)).
                Distinct(StringComparer.OrdinalIgnoreCase)

            Return String.Join(";", values)
        End Function
    End Class

End Namespace
