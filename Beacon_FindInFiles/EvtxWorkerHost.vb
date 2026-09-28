Imports System.Diagnostics
Imports System.Diagnostics.Eventing.Reader
Imports System.IO
Imports System.Reflection
Imports System.Text
Imports System.Text.Json
Imports System.Text.RegularExpressions
Imports System.Threading
Imports System.Threading.Tasks
Imports System.Xml.Linq

Namespace Beacon
    Friend NotInheritable Class EvtxWorkerHost
        Private Const WorkerSwitch As String = "--beacon-evtx-worker"
        Private Const ProtocolPrefix As String = "BEACON-EVTX-1 "
        Private Const ProtocolComplete As String = "complete"
        Private Const ProtocolCheckpoint As String = "checkpoint"
        Private Const ProtocolMatch As String = "match"
        Private Const ProtocolError As String = "error"
        Private Const ProtocolWarning As String = "warning"
        Private Const ProtocolRegexTimeout As String = "regex-timeout"
        Private Const MaxLineChars As Integer = 262144
        Private Const NormalSearchCoverageWarning As String = "EVTX normal search: Deep Search is disabled, so sibling archived message resources were not searched; local message resources and raw XML were used, so message-only matches can be missed"
        Private Shared ReadOnly DefaultRenderingBudget As TimeSpan = TimeSpan.FromSeconds(180)
        Friend Shared RenderingBudget As TimeSpan? = Nothing
        Friend Shared XmlFallbackBudget As TimeSpan = TimeSpan.FromSeconds(180)
        Friend Shared WorkerExecutableOverride As String

        Private Sub New()
        End Sub

        Friend Shared Function IsWorkerInvocation(args As String()) As Boolean
            Return args IsNot Nothing AndAlso args.Length >= 2 AndAlso args(0) = WorkerSwitch
        End Function

        Friend Shared Function RunWorker(args As String()) As Integer
            Try
                If Not IsWorkerInvocation(args) Then Return 2
                Dim requestPath = args(1)
                If String.IsNullOrWhiteSpace(requestPath) OrElse Not File.Exists(requestPath) Then Return 2
                Dim request = JsonSerializer.Deserialize(Of EvtxWorkerRequest)(File.ReadAllText(requestPath, Encoding.UTF8), JsonOptions())
                If request Is Nothing OrElse String.IsNullOrWhiteSpace(request.FilePath) OrElse Not File.Exists(request.FilePath) Then Return 2
                Dim query As New SearchQuery(request.QueryText, request.QueryMode, request.CaseSensitive)
                Dim options = If(request.Settings, New BeaconSettings())
                options.MaximumStructuredMatches = Math.Max(0, Math.Min(options.MaximumStructuredMatches, Math.Max(0, request.RemainingMatches)))
                options.EvtxMaximumMatches = Math.Max(0, Math.Min(options.EvtxMaximumMatches, Math.Max(0, request.RemainingMatches)))
                If options.MaximumStructuredMatches = 0 OrElse options.EvtxMaximumMatches = 0 Then
                    WriteMessage(New EvtxWorkerMessage With {.Type = ProtocolComplete})
                    Return 0
                End If
                If request.XmlOnly Then
                    CollectXmlOnly(request.FilePath, query, options, request.StartSequence, CancellationToken.None)
                Else
                    CollectRendered(request.FilePath, query, options, request.StartSequence, CancellationToken.None)
                End If
                WriteMessage(New EvtxWorkerMessage With {.Type = ProtocolComplete})
                Return 0
            Catch ex As RegexMatchTimeoutException
                WriteMessage(New EvtxWorkerMessage With {.Type = ProtocolRegexTimeout})
                Return 1
            Catch ex As Exception When TypeOf ex Is IOException OrElse TypeOf ex Is UnauthorizedAccessException OrElse TypeOf ex Is InvalidDataException OrElse TypeOf ex Is EventLogException OrElse TypeOf ex Is ArgumentException OrElse TypeOf ex Is JsonException
                Try
                    WriteMessage(New EvtxWorkerMessage With {.Type = ProtocolError, .ErrorMessage = ex.Message})
                Catch writeEx As Exception When TypeOf writeEx Is IOException OrElse TypeOf writeEx Is ObjectDisposedException OrElse TypeOf writeEx Is InvalidOperationException
                    Console.Error.WriteLine("EVTX worker could not write protocol error: " & writeEx.Message)
                End Try
                Return 1
            End Try
        End Function

        Friend Shared Sub CollectIsolated(filePath As String, query As SearchQuery, settings As BeaconSettings,
                                          result As StructuredSearchResult(Of EventRecordSummary), token As CancellationToken,
                                          report As Action(Of Exception, String, String),
                                          Optional fastCopyLimitBytes As Long = -1)
            token.ThrowIfCancellationRequested()
            Dim scanPath = filePath
            Dim fastCopy As EvtxFastModeCopy = Nothing
            If settings IsNot Nothing AndAlso settings.EvtxFastMode Then
                fastCopy = TryCreateFastModeCopy(filePath, fastCopyLimitBytes, token, report)
                If fastCopy IsNot Nothing Then
                    scanPath = fastCopy.FilePath
                Else
                    report?.Invoke(New IOException("EVTX normal-search copy was unavailable; searching the original file with Deep Search resource discovery."), "EVTX normal search unavailable", "Warning")
                End If
            End If
            Using fastCopy
                If fastCopy IsNot Nothing Then
                    AppendPartial(result, NormalSearchCoverageWarning)
                    report?.Invoke(New InvalidDataException(NormalSearchCoverageWarning & "."), "EVTX normal search coverage", "Warning")
                End If
                Dim sourceIdentity = EvtxSourceIdentity.Capture(scanPath)
                Dim matchedSequences As New HashSet(Of Long)()
                Dim renderBudget = ResolveRenderingBudget(settings)
                Dim render = RunPhase(scanPath, query, settings, result, token, report, xmlOnly:=False, startSequence:=0, budget:=renderBudget, matchedSequences:=matchedSequences)
                If render.LimitReached Then Return
                If render.Completed Then Return
                If render.TimedOut Then
                    Dim limitText = FormatBudget(renderBudget)
                    AppendPartial(result, "message rendering timed out after " & limitText & "; continued with XML-only search, so message-only coverage is incomplete")
                    report?.Invoke(New TimeoutException("EVTX message rendering exceeded the " & limitText & " per-file budget. Beacon preserved rendered matches and continued with raw XML search."), "EVTX rendering timeout", "Warning")
                Else
                    AppendPartial(result, "message rendering worker failed; continued with XML-only search, so message-only coverage is incomplete")
                End If
                If Not sourceIdentity.Matches(scanPath) Then
                    AppendPartial(result, "EVTX file changed while scanning; XML fallback was not started to avoid mixing event sequences")
                    Return
                End If
                Dim fallback = RunPhase(scanPath, query, settings, result, token, report, xmlOnly:=True, startSequence:=render.LastCheckpoint + 1, budget:=XmlFallbackBudget, matchedSequences:=matchedSequences)
                If fallback.LimitReached Then Return
                If Not fallback.Completed Then
                    If fallback.TimedOut Then
                        AppendPartial(result, "XML fallback also timed out before the file was fully searched")
                        report?.Invoke(New TimeoutException("EVTX raw XML fallback did not finish within its bounded 3-minute budget."), "EVTX XML fallback", "Warning")
                    Else
                        AppendPartial(result, "XML fallback also failed before the file was fully searched")
                    End If
                End If
            End Using
        End Sub

        Private Shared Function TryCreateFastModeCopy(filePath As String, copyLimitBytes As Long, token As CancellationToken,
                                                      report As Action(Of Exception, String, String)) As EvtxFastModeCopy
            Dim directory = Path.Combine(Path.GetTempPath(), "BeaconEvtxFast_" & Guid.NewGuid().ToString("N"))
            Try
                token.ThrowIfCancellationRequested()
                Dim sourceIdentity = EvtxSourceIdentity.Capture(filePath)
                Dim sourceInfo As New FileInfo(filePath)
                Dim limit = If(copyLimitBytes > 0, copyLimitBytes, sourceInfo.Length)
                If sourceInfo.Length > limit Then Throw New InvalidDataException("EVTX file exceeds the normal-search copy limit.")
                IO.Directory.CreateDirectory(directory)
                Dim target = Path.Combine(directory, Path.GetFileName(filePath))
                Dim copied As Long
                Dim buffer(81919) As Byte
                Using input As New FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.Read, buffer.Length, FileOptions.SequentialScan),
                      bounded As New BoundedReadStream(input, limit, token),
                      output As New FileStream(target, FileMode.CreateNew, FileAccess.Write, FileShare.None, buffer.Length, FileOptions.SequentialScan)
                    While True
                        Dim count = bounded.Read(buffer, 0, buffer.Length)
                        If count = 0 Then Exit While
                        token.ThrowIfCancellationRequested()
                        output.Write(buffer, 0, count)
                        copied += count
                    End While
                End Using
                If copied <> sourceInfo.Length Then Throw New IOException("EVTX normal-search copy did not preserve the source byte count.")
                If Not sourceIdentity.Matches(filePath) Then Throw New IOException("EVTX file changed while Beacon was preparing the normal-search copy.")
                Return New EvtxFastModeCopy(directory, target, report)
            Catch ex As OperationCanceledException
                CleanupFastModeDirectory(directory, report)
                Throw
            Catch ex As Exception When TypeOf ex Is IOException OrElse TypeOf ex Is UnauthorizedAccessException OrElse TypeOf ex Is InvalidDataException
                CleanupFastModeDirectory(directory, report)
                report?.Invoke(ex, "EVTX normal search copy", "Warning")
                Return Nothing
            End Try
        End Function

        Private Shared Sub CleanupFastModeDirectory(directory As String, report As Action(Of Exception, String, String))
            If String.IsNullOrWhiteSpace(directory) OrElse Not IO.Directory.Exists(directory) Then Return
            Try
                IO.Directory.Delete(directory, True)
            Catch ex As Exception When TypeOf ex Is IOException OrElse TypeOf ex Is UnauthorizedAccessException
                report?.Invoke(ex, "EVTX normal search cleanup", "Warning")
            End Try
        End Sub

        Private Shared Function ResolveRenderingBudget(settings As BeaconSettings) As TimeSpan?
            If RenderingBudget.HasValue Then Return RenderingBudget
            If settings IsNot Nothing AndAlso settings.EvtxDeepSearch Then
                If settings.EvtxLimitDeepSearchBeforeXml Then Return TimeSpan.FromSeconds(Math.Clamp(settings.EvtxDeepSearchTimeoutSeconds, 1, 86400))
                Return Nothing
            End If
            Return DefaultRenderingBudget
        End Function

        Private Shared Function FormatBudget(budget As TimeSpan?) As String
            If Not budget.HasValue Then Return "the configured limit"
            Dim totalSeconds = Math.Max(1, CInt(Math.Ceiling(budget.Value.TotalSeconds)))
            Return totalSeconds.ToString(Globalization.CultureInfo.InvariantCulture) & " second" & If(totalSeconds = 1, "", "s")
        End Function

        Friend Shared Function EffectiveRenderingBudgetForTest(settings As BeaconSettings) As TimeSpan?
            Return ResolveRenderingBudget(BeaconSettingsService.Validate(BeaconSettingsService.Clone(settings)))
        End Function

        Private Shared Function RunPhase(filePath As String, query As SearchQuery, settings As BeaconSettings,
                                         result As StructuredSearchResult(Of EventRecordSummary), token As CancellationToken,
                                         report As Action(Of Exception, String, String), xmlOnly As Boolean, startSequence As Long,
                                         budget As TimeSpan?, matchedSequences As HashSet(Of Long)) As EvtxPhaseResult
            Dim phase As New EvtxPhaseResult With {.LastCheckpoint = startSequence - 1}
            If settings.StopAfterFirstMatchPerFile AndAlso result.Records.Count > 0 Then
                phase.Completed = True
                phase.LimitReached = True
                Return phase
            End If
            Dim remaining = Math.Min(settings.MaximumStructuredMatches, settings.EvtxMaximumMatches) - result.Records.Count
            If remaining <= 0 Then
                phase.Completed = True
                phase.LimitReached = True
                Return phase
            End If
            Dim launch = ResolveWorkerLaunch()
            Dim replayedMatches = Enumerable.Count(matchedSequences, Function(sequence) sequence >= startSequence)
            Dim request As New EvtxWorkerRequest With {.FilePath = filePath, .QueryText = query.Text, .QueryMode = query.Mode,
                .CaseSensitive = query.CaseSensitive, .Settings = BeaconSettingsService.Clone(settings), .XmlOnly = xmlOnly,
                .StartSequence = startSequence, .RemainingMatches = remaining + replayedMatches}
            Dim requestPath = Path.Combine(Path.GetTempPath(), "BeaconEvtxWorker_" & Guid.NewGuid().ToString("N") & ".json")
            File.WriteAllText(requestPath, JsonSerializer.Serialize(request, JsonOptions()), Encoding.UTF8)
            Dim process As Process = Nothing
            Dim processStarted As Boolean
            Dim outputTask As Task = Nothing
            Dim stderrTask As Task = Nothing
            Dim stderr As New StringBuilder()
            Dim killed As Boolean
            Dim readerCancellation = CancellationTokenSource.CreateLinkedTokenSource(token)
            Dim deadline = Stopwatch.StartNew()
            Try
                Dim start As New ProcessStartInfo(launch.FileName) With {.UseShellExecute = False, .CreateNoWindow = True,
                    .RedirectStandardOutput = True, .RedirectStandardError = True,
                    .StandardOutputEncoding = New UTF8Encoding(False), .StandardErrorEncoding = New UTF8Encoding(False)}
                For Each argument In launch.Arguments
                    start.ArgumentList.Add(argument)
                Next
                start.ArgumentList.Add(WorkerSwitch)
                start.ArgumentList.Add(requestPath)
                process = New Process With {.StartInfo = start}
                processStarted = process.Start()
                If Not processStarted Then Throw New IOException("EVTX worker process did not start.")

                Dim state As New EvtxProtocolState With {.Phase = phase, .Query = query, .Settings = settings,
                    .Filter = EvtxFilter.FromSettings(settings), .Result = result, .Token = readerCancellation.Token, .Report = report,
                    .MatchedSequences = matchedSequences, .StartSequence = startSequence}
                outputTask = Task.Run(Sub() ReadBoundedLines(process.StandardOutput, Sub(line) ProcessProtocolLine(line, state), readerCancellation.Token))
                stderrTask = Task.Run(Sub() ReadBoundedLines(process.StandardError,
                    Sub(line)
                        If stderr.Length < 4096 Then
                            Dim remainingError = 4096 - stderr.Length
                            stderr.AppendLine(If(line.Length > remainingError, line.Substring(0, remainingError), line))
                        End If
                    End Sub, readerCancellation.Token))

                While True
                    token.ThrowIfCancellationRequested()
                    If outputTask.IsFaulted OrElse stderrTask.IsFaulted OrElse phase.LimitReached Then
                        If Not killed Then
                            readerCancellation.Cancel()
                            EnsureWorkerTerminated(process, report)
                            killed = True
                        End If
                    End If
                    If budget.HasValue AndAlso deadline.Elapsed >= budget.Value AndAlso Not phase.LimitReached Then
                        phase.TimedOut = True
                        readerCancellation.Cancel()
                        If Not killed Then
                            EnsureWorkerTerminated(process, report)
                            killed = True
                        End If
                    End If
                    If process.HasExited AndAlso outputTask.IsCompleted AndAlso stderrTask.IsCompleted Then Exit While
                    If killed AndAlso outputTask.IsCompleted AndAlso stderrTask.IsCompleted Then Exit While
                    Thread.Sleep(25)
                End While

                token.ThrowIfCancellationRequested()
                Dim outputError = If(outputTask.IsFaulted, outputTask.Exception.GetBaseException(), Nothing)
                Dim stderrError = If(stderrTask.IsFaulted, stderrTask.Exception.GetBaseException(), Nothing)
                If Not phase.LimitReached Then
                    If outputError IsNot Nothing AndAlso Not TypeOf outputError Is OperationCanceledException Then Throw outputError
                    If stderrError IsNot Nothing AndAlso Not TypeOf stderrError Is OperationCanceledException Then Throw stderrError
                End If
                If phase.TimedOut Then Return phase
                If Not phase.LimitReached Then
                    If Not state.Complete Then Throw New InvalidDataException("EVTX worker stdout ended before the required completion message.")
                    If process.ExitCode <> 0 Then
                        Dim err = stderr.ToString().Trim()
                        Throw New InvalidDataException("EVTX worker failed" & If(err.Length > 0, ": " & err, "."))
                    End If
                    phase.Completed = True
                End If
                Return phase
            Catch ex As OperationCanceledException When token.IsCancellationRequested
                Throw
            Catch ex As RegexMatchTimeoutException
                Throw
            Catch ex As Exception When TypeOf ex Is IOException OrElse TypeOf ex Is UnauthorizedAccessException OrElse TypeOf ex Is InvalidDataException OrElse TypeOf ex Is InvalidOperationException OrElse TypeOf ex Is ComponentModel.Win32Exception OrElse TypeOf ex Is JsonException
                phase.FailureMessage = ex.Message
                report?.Invoke(ex, If(xmlOnly, "EVTX XML fallback", "EVTX worker"), "Warning")
                Return phase
            Finally
                readerCancellation.Cancel()
                Try
                    If processStarted AndAlso Not process.HasExited Then EnsureWorkerTerminated(process, report)
                    If Not WaitReaderTasks(outputTask, stderrTask) Then Throw New InvalidOperationException("EVTX worker stream readers did not finish; Beacon will not start a replacement worker for this file.")
                Finally
                    process?.Dispose()
                    readerCancellation.Dispose()
                    Try
                        File.Delete(requestPath)
                    Catch ex As Exception When TypeOf ex Is IOException OrElse TypeOf ex Is UnauthorizedAccessException
                        report?.Invoke(ex, "EVTX worker cleanup", "Warning")
                    End Try
                End Try
            End Try
        End Function

        Private Shared Sub ReadBoundedLines(reader As TextReader, consume As Action(Of String), token As CancellationToken)
            Dim builder As New StringBuilder(Math.Min(4096, MaxLineChars))
            While True
                token.ThrowIfCancellationRequested()
                Dim value = reader.Read()
                If value < 0 Then
                    If builder.Length > 0 Then Throw New InvalidDataException("EVTX worker ended an unterminated protocol line.")
                    Exit While
                End If
                Dim ch = ChrW(value)
                If ch = ControlChars.Lf Then
                    If builder.Length > 0 AndAlso builder(builder.Length - 1) = ControlChars.Cr Then builder.Length -= 1
                    consume(builder.ToString())
                    builder.Clear()
                Else
                    builder.Append(ch)
                    If builder.Length > MaxLineChars Then Throw New InvalidDataException("EVTX worker output line exceeded Beacon's safety limit.")
                End If
            End While
        End Sub

        Private Shared Sub ProcessProtocolLine(line As String, state As EvtxProtocolState)
            state.Token.ThrowIfCancellationRequested()
            If state.Complete Then Throw New InvalidDataException("EVTX worker returned output after the completion message.")
            If Not line.StartsWith(ProtocolPrefix, StringComparison.Ordinal) Then Throw New InvalidDataException("EVTX worker returned an unrecognized protocol line.")
            Dim message = JsonSerializer.Deserialize(Of EvtxWorkerMessage)(line.Substring(ProtocolPrefix.Length), JsonOptions())
            If message Is Nothing OrElse String.IsNullOrEmpty(message.Type) Then Throw New InvalidDataException("EVTX worker returned an empty protocol message.")
            Select Case message.Type
                Case ProtocolMatch
                    ValidateSequence(message.Sequence, state.Phase.LastCheckpoint, "match")
                    If message.Record Is Nothing Then Throw New InvalidDataException("EVTX worker returned a match without event data.")
                    message.Record.EventSequence = message.Sequence
                    If state.MatchedSequences.Add(message.Sequence) Then
                        If (Not state.Settings.StopAfterFirstMatchPerFile OrElse state.Result.Records.Count = 0) AndAlso state.Result.Records.Count < Math.Min(state.Settings.MaximumStructuredMatches, state.Settings.EvtxMaximumMatches) Then
                            Dim before = state.Result.Records.Count
                            Dim stopNow = CollectRecordCore(state.Query, state.Settings, state.Filter, message.Record, state.Result, state.Token)
                            If state.Result.Records.Count = before Then state.MatchedSequences.Remove(message.Sequence)
                            If stopNow Then state.Phase.LimitReached = True
                        End If
                    End If
                Case ProtocolCheckpoint
                    ValidateSequence(message.Sequence, state.Phase.LastCheckpoint, "checkpoint")
                    state.Phase.LastCheckpoint = message.Sequence
                Case ProtocolComplete
                    state.Complete = True
                Case ProtocolRegexTimeout
                    Throw New RegexMatchTimeoutException()
                Case ProtocolError
                    Throw New InvalidDataException(If(message.ErrorMessage, "EVTX worker reported an error."))
                Case ProtocolWarning
                    state.Report?.Invoke(New InvalidDataException(If(message.ErrorMessage, "EVTX worker reported a warning.")), "EVTX message resources", "Warning")
                Case Else
                    Throw New InvalidDataException("EVTX worker returned an unknown protocol message.")
            End Select
        End Sub

        Private Shared Sub ValidateSequence(sequence As Long, lastCheckpoint As Long, messageType As String)
            If sequence < 0 Then Throw New InvalidDataException("EVTX worker returned a negative " & messageType & " sequence.")
            Dim expected = lastCheckpoint + 1
            If sequence < expected Then Throw New InvalidDataException("EVTX worker returned a backward " & messageType & " sequence.")
            If sequence > expected Then Throw New InvalidDataException("EVTX worker returned a future " & messageType & " sequence.")
        End Sub

        Private Shared Function WaitReaderTasks(outputTask As Task, stderrTask As Task) As Boolean
            Dim tasks = {outputTask, stderrTask}.Where(Function(item) item IsNot Nothing).ToArray()
            If tasks.Length = 0 Then Return True
            Try
                Task.WaitAll(tasks, TimeSpan.FromSeconds(2))
            Catch ex As AggregateException
                ' The caller inspects reader task faults; this helper only proves callbacks have stopped.
            End Try
            Return tasks.All(Function(item) item.IsCompleted)
        End Function

        Private Shared Sub EnsureWorkerTerminated(process As Process, report As Action(Of Exception, String, String))
            If Not KillWorker(process, report) Then Throw New InvalidOperationException("EVTX worker could not be terminated; Beacon will not start a replacement worker for this file.")
        End Sub

        Private Shared Function KillWorker(process As Process, report As Action(Of Exception, String, String)) As Boolean
            Try
                If Not process.HasExited Then process.Kill(entireProcessTree:=True)
            Catch ex As Exception When TypeOf ex Is InvalidOperationException OrElse TypeOf ex Is ComponentModel.Win32Exception
                report?.Invoke(ex, "EVTX worker cleanup", "Warning")
            End Try
            Try
                If process.WaitForExit(5000) Then Return True
                report?.Invoke(New TimeoutException("EVTX worker did not exit after termination."), "EVTX worker cleanup", "Warning")
                Return False
            Catch ex As Exception When TypeOf ex Is InvalidOperationException OrElse TypeOf ex Is ComponentModel.Win32Exception
                report?.Invoke(ex, "EVTX worker cleanup", "Warning")
                Return False
            End Try
        End Function

        Private Shared Function ResolveWorkerLaunch() As EvtxWorkerLaunch
            If Not String.IsNullOrWhiteSpace(WorkerExecutableOverride) AndAlso File.Exists(WorkerExecutableOverride) Then
                Return LaunchFromPath(WorkerExecutableOverride, allowCurrentDotnetHost:=True)
            End If

            Dim entry = Assembly.GetEntryAssembly()
            Dim entryPath = If(entry Is Nothing, Nothing, entry.Location)
            Dim entryName = If(entry Is Nothing, "", entry.GetName().Name)
            If IsKnownWorkerAssembly(entryName) AndAlso Not String.IsNullOrWhiteSpace(Environment.ProcessPath) AndAlso File.Exists(Environment.ProcessPath) Then
                If IsDotnetHost(Environment.ProcessPath) Then Return LaunchFromPath(entryPath, allowCurrentDotnetHost:=True)
                Return New EvtxWorkerLaunch With {.FileName = Environment.ProcessPath}
            End If

            Dim appName = Path.Combine(AppContext.BaseDirectory, "Beacon.exe")
            If File.Exists(appName) Then Return New EvtxWorkerLaunch With {.FileName = appName}

            Dim beaconDll = Path.Combine(AppContext.BaseDirectory, "Beacon.dll")
            If File.Exists(beaconDll) Then Return LaunchFromPath(beaconDll, allowCurrentDotnetHost:=False)

            Throw New FileNotFoundException("Beacon could not locate a worker-capable EVTX host. Run Beacon.exe or provide Beacon.dll with its runtimeconfig.")
        End Function

        Private Shared Function LaunchFromPath(path As String, allowCurrentDotnetHost As Boolean) As EvtxWorkerLaunch
            If String.IsNullOrWhiteSpace(path) OrElse Not File.Exists(path) Then Throw New FileNotFoundException("Beacon could not locate its EVTX worker host.")
            If path.EndsWith(".dll", StringComparison.OrdinalIgnoreCase) Then
                Dim runtimeConfig = IO.Path.ChangeExtension(path, ".runtimeconfig.json")
                If Not File.Exists(runtimeConfig) Then Throw New FileNotFoundException("Beacon EVTX worker DLL is missing its runtimeconfig: " & runtimeConfig)
                Return New EvtxWorkerLaunch With {.FileName = ResolveDotnetHost(allowCurrentDotnetHost), .Arguments = New List(Of String) From {path}}
            End If
            If IsDotnetHost(path) Then
                Dim entry = Assembly.GetEntryAssembly()
                Dim entryPath = If(entry Is Nothing, Nothing, entry.Location)
                If String.IsNullOrWhiteSpace(entryPath) OrElse Not File.Exists(entryPath) Then Throw New FileNotFoundException("Beacon could not locate the current worker assembly for dotnet hosting.")
                Return LaunchFromPath(entryPath, allowCurrentDotnetHost:=True)
            End If
            Return New EvtxWorkerLaunch With {.FileName = path}
        End Function

        Private Shared Function ResolveDotnetHost(allowCurrentDotnetHost As Boolean) As String
            If allowCurrentDotnetHost AndAlso Not String.IsNullOrWhiteSpace(Environment.ProcessPath) AndAlso File.Exists(Environment.ProcessPath) AndAlso IsDotnetHost(Environment.ProcessPath) Then Return Environment.ProcessPath
            If Not String.IsNullOrWhiteSpace(Environment.ProcessPath) Then
                Dim processDirectory = IO.Path.GetDirectoryName(Environment.ProcessPath)
                If Not String.IsNullOrWhiteSpace(processDirectory) Then
                    Dim dotnet = IO.Path.Combine(processDirectory, "dotnet.exe")
                    If File.Exists(dotnet) Then Return dotnet
                End If
            End If
            Return "dotnet"
        End Function

        Private Shared Function IsDotnetHost(path As String) As Boolean
            Dim name = IO.Path.GetFileNameWithoutExtension(path)
            Return String.Equals(name, "dotnet", StringComparison.OrdinalIgnoreCase)
        End Function

        Private Shared Function IsKnownWorkerAssembly(name As String) As Boolean
            Return String.Equals(name, "Beacon", StringComparison.OrdinalIgnoreCase) OrElse
                String.Equals(name, "Beacon.SafetyChecks", StringComparison.OrdinalIgnoreCase)
        End Function

        Private Shared Sub CollectRendered(filePath As String, query As SearchQuery, options As BeaconSettings, startSequence As Long, token As CancellationToken)
            Dim filter = EvtxFilter.FromSettings(options)
            Dim missingProviders As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
            Dim sequence As Long = -1
            Dim matchedCount As Integer
            Dim matchLimit = Math.Min(options.MaximumStructuredMatches, options.EvtxMaximumMatches)
            Using reader As New EventLogReader(filePath, PathType.FilePath)
                While True
                    token.ThrowIfCancellationRequested()
                    Using record = reader.ReadEvent()
                        If record Is Nothing Then Exit While
                        sequence += 1
                        If sequence < startSequence Then Continue While
                        Dim provider = ReadField(Function() record.ProviderName, "Unknown")
                        Dim timestamp = record.TimeCreated
                        Dim levelNumber = record.Level
                        If filter.Matches(record.Id, provider, levelNumber, timestamp) Then
                            Dim level = ReadField(Function() record.LevelDisplayName, "Unknown")
                            Dim message = ReadField(Function() record.FormatDescription(), "")
                            If String.IsNullOrWhiteSpace(message) AndAlso missingProviders.Add(provider) Then
                                WriteMessage(New EvtxWorkerMessage With {.Type = ProtocolWarning, .ErrorMessage = $"Message text unavailable for provider '{provider}'. XML remains searchable. {EvtxFilter.ResourceGuidance}"})
                            End If
                            Dim summary As New EventRecordSummary With {.Provider = provider, .Level = level, .EventId = record.Id,
                                .TimeCreated = timestamp, .LevelNumber = levelNumber, .Message = message, .RawXml = record.ToXml(), .EventSequence = sequence}
                            If EmitIfMatch(query, options, filter, summary, sequence, token) Then
                                matchedCount += 1
                                If options.StopAfterFirstMatchPerFile OrElse matchedCount >= matchLimit Then Return
                            End If
                        End If
                        WriteMessage(New EvtxWorkerMessage With {.Type = ProtocolCheckpoint, .Sequence = sequence})
                    End Using
                End While
            End Using
        End Sub

        Private Shared Sub CollectXmlOnly(filePath As String, query As SearchQuery, options As BeaconSettings, startSequence As Long, token As CancellationToken)
            Dim filter = EvtxFilter.FromSettings(options)
            Dim sequence As Long = -1
            Dim matchedCount As Integer
            Dim matchLimit = Math.Min(options.MaximumStructuredMatches, options.EvtxMaximumMatches)
            Using reader As New EventLogReader(filePath, PathType.FilePath)
                While True
                    token.ThrowIfCancellationRequested()
                    Using record = reader.ReadEvent()
                        If record Is Nothing Then Exit While
                        sequence += 1
                        If sequence < startSequence Then Continue While
                        Dim xml = record.ToXml()
                        Dim summary = SummaryFromXml(xml, sequence)
                        If filter.Matches(summary.EventId, summary.Provider, summary.LevelNumber, summary.TimeCreated) AndAlso EmitIfMatch(query, options, filter, summary, sequence, token) Then
                            matchedCount += 1
                            If options.StopAfterFirstMatchPerFile OrElse matchedCount >= matchLimit Then Return
                        End If
                        WriteMessage(New EvtxWorkerMessage With {.Type = ProtocolCheckpoint, .Sequence = sequence})
                    End Using
                End While
            End Using
        End Sub

        Private Shared Function EmitIfMatch(query As SearchQuery, options As BeaconSettings, filter As EvtxFilter, summary As EventRecordSummary, sequence As Long, token As CancellationToken) As Boolean
            Dim probe As New StructuredSearchResult(Of EventRecordSummary)()
            CollectRecordCore(query, options, filter, summary, probe, token)
            If probe.Records.Count = 0 Then Return False
            WriteMessage(New EvtxWorkerMessage With {.Type = ProtocolMatch, .Sequence = sequence, .Record = summary})
            Return True
        End Function

        Private Shared Function SummaryFromXml(xml As String, sequence As Long) As EventRecordSummary
            Try
                Dim document = XDocument.Parse(xml, LoadOptions.None)
                If document.Root Is Nothing Then Throw New InvalidDataException("EVTX XML record is missing its root element.")
                Dim ns = document.Root.Name.Namespace
                Dim system = document.Root.Element(ns + "System")
                If system Is Nothing Then Throw New InvalidDataException("EVTX XML record is missing the System element.")

                Dim provider = ""
                Dim providerElement = system.Element(ns + "Provider")
                If providerElement IsNot Nothing Then provider = If(CStr(providerElement.Attribute("Name")), "")

                Dim eventId As Integer
                Dim idText = CStr(system.Element(ns + "EventID"))
                If Not Integer.TryParse(idText, Globalization.NumberStyles.Integer, Globalization.CultureInfo.InvariantCulture, eventId) Then Throw New InvalidDataException("EVTX XML record has an invalid EventID.")

                Dim levelNumber As Byte? = Nothing
                Dim levelText = CStr(system.Element(ns + "Level"))
                If Not String.IsNullOrWhiteSpace(levelText) Then
                    Dim parsedLevel As Byte
                    If Not Byte.TryParse(levelText, Globalization.NumberStyles.Integer, Globalization.CultureInfo.InvariantCulture, parsedLevel) Then Throw New InvalidDataException("EVTX XML record has an invalid Level.")
                    levelNumber = parsedLevel
                End If

                Dim timestamp As DateTime? = Nothing
                Dim timeElement = system.Element(ns + "TimeCreated")
                If timeElement IsNot Nothing Then
                    Dim systemTime = CStr(timeElement.Attribute("SystemTime"))
                    If Not String.IsNullOrWhiteSpace(systemTime) Then
                        Dim parsed As DateTime
                        If Not DateTime.TryParse(systemTime, Globalization.CultureInfo.InvariantCulture, Globalization.DateTimeStyles.AdjustToUniversal Or Globalization.DateTimeStyles.AssumeUniversal, parsed) Then Throw New InvalidDataException("EVTX XML record has an invalid TimeCreated value.")
                        timestamp = parsed
                    End If
                End If

                Return New EventRecordSummary With {.Provider = provider, .Level = "", .EventId = eventId, .LevelNumber = levelNumber, .TimeCreated = timestamp,
                    .Message = "", .RawXml = xml, .MessageUnavailable = True, .EventSequence = sequence}
            Catch ex As Xml.XmlException
                Throw New InvalidDataException("EVTX XML record is malformed.", ex)
            End Try
        End Function

        Friend Shared Function CollectRecordCore(query As SearchQuery, options As BeaconSettings, record As EventRecordSummary,
                                                 result As StructuredSearchResult(Of EventRecordSummary), token As CancellationToken) As Boolean
            Return CollectRecordCore(query, options, EvtxFilter.FromSettings(options), record, result, token)
        End Function

        Private Shared Function CollectRecordCore(query As SearchQuery, options As BeaconSettings, filter As EvtxFilter, record As EventRecordSummary,
                                                  result As StructuredSearchResult(Of EventRecordSummary), token As CancellationToken) As Boolean
            token.ThrowIfCancellationRequested()
            If Not filter.Matches(record.EventId, record.Provider, record.LevelNumber, record.TimeCreated) Then Return False
            Dim message = If(record.Message, "")
            Dim xml = If(record.RawXml, "")
            Dim unavailable = String.IsNullOrWhiteSpace(message)
            Dim timestamp = record.TimeCreated
            Dim location = $"Event {record.EventId} · {If(timestamp.HasValue, timestamp.Value.ToString("g"), "no timestamp")}"
            Dim content = String.Join(vbLf, {"Event ID " & record.EventId.ToString(), record.Provider, record.Level,
                                          If(timestamp.HasValue, timestamp.Value.ToString("O"), ""), message})
            Dim detail = SearchDetail.FromRecord(query, content, location, recordIndex:=result.Records.Count, token:=token)
            If detail Is Nothing Then
                detail = SearchDetail.FromRecord(query, xml, location & " (XML)", recordIndex:=result.Records.Count, token:=token)
                If detail IsNot Nothing Then
                    Dim readableXml = PreviewFormatting.FormatXml(xml, True, CLng(options.MaximumPreviewSizeMb) * 1024 * 1024)
                    Dim xmlSpans = query.FindHighlights(readableXml, token:=token)
                    If xmlSpans.Count > 0 Then detail.Context = MatchContext.FromText(readableXml, xmlSpans, token)
                    detail.ContextKind = "Event XML lines"
                    message &= vbCrLf & vbCrLf & "Event XML:" & vbCrLf & readableXml
                End If
            End If
            If detail Is Nothing Then Return False
            If detail.ContextKind = "Record text" Then detail.ContextKind = "Event text lines"
            If String.IsNullOrWhiteSpace(message) Then
                message = "Message resources are unavailable. Use Raw XML or the offline resource guidance below."
                If options.ShowRawXmlWhenMessageUnavailable Then message &= vbCrLf & PreviewFormatting.FormatXml(xml, True, CLng(options.MaximumPreviewSizeMb) * 1024 * 1024)
            End If
            Dim xmlLimit = CInt(Math.Min(65536L, CLng(options.MaximumPreviewSizeMb) * 1024 * 1024))
            Dim capturedXml = If(xml.Length > xmlLimit, xml.Substring(0, xmlLimit), xml)
            result.Records.Add(New EventRecordSummary With {.Provider = record.Provider, .Level = record.Level,
                .EventId = record.EventId, .TimeCreated = timestamp, .Message = message, .LevelNumber = record.LevelNumber,
                .RawXml = capturedXml, .XmlShortened = xml.Length > xmlLimit, .MessageUnavailable = unavailable, .EventSequence = record.EventSequence})
            result.Details.Add(detail)
            If options.StopAfterFirstMatchPerFile OrElse result.Records.Count >= Math.Min(options.MaximumStructuredMatches, options.EvtxMaximumMatches) Then
                AppendPartial(result, If(options.StopAfterFirstMatchPerFile, "first matching event only", "event limit reached"))
                Return True
            End If
            Return False
        End Function

        Private Shared Function ReadField(read As Func(Of String), fallback As String) As String
            Try
                Return If(read(), fallback)
            Catch ex As EventLogException
                Return fallback
            End Try
        End Function

        Friend Shared Sub WriteMessage(message As EvtxWorkerMessage)
            Dim line = ProtocolPrefix & JsonSerializer.Serialize(message, JsonOptions())
            Console.Out.WriteLine(line)
            Console.Out.Flush()
        End Sub

        Friend Shared Sub AppendPartial(result As StructuredSearchResult(Of EventRecordSummary), reason As String)
            If String.IsNullOrWhiteSpace(reason) Then Return
            If String.IsNullOrWhiteSpace(result.PartialReason) Then
                result.PartialReason = reason
            ElseIf Not result.PartialReason.Contains(reason, StringComparison.OrdinalIgnoreCase) Then
                result.PartialReason &= "; " & reason
            End If
        End Sub

        Private Shared ReadOnly s_jsonOptions As New JsonSerializerOptions With {.PropertyNameCaseInsensitive = True}

        Private Shared Function JsonOptions() As JsonSerializerOptions
            Return s_jsonOptions
        End Function

        Private NotInheritable Class EvtxWorkerLaunch
            Public Property FileName As String
            Public Property Arguments As New List(Of String)()
        End Class

        Private NotInheritable Class EvtxProtocolState
            Public Property Phase As EvtxPhaseResult
            Public Property Query As SearchQuery
            Public Property Settings As BeaconSettings
            Public Property Filter As EvtxFilter
            Public Property Result As StructuredSearchResult(Of EventRecordSummary)
            Public Property Token As CancellationToken
            Public Property Report As Action(Of Exception, String, String)
            Public Property MatchedSequences As HashSet(Of Long)
            Public Property StartSequence As Long
            Public Property Complete As Boolean
        End Class

        Private NotInheritable Class EvtxSourceIdentity
            Private ReadOnly _length As Long
            Private ReadOnly _lastWriteUtc As DateTime

            Private Sub New(length As Long, lastWriteUtc As DateTime)
                _length = length
                _lastWriteUtc = lastWriteUtc
            End Sub

            Public Shared Function Capture(filePath As String) As EvtxSourceIdentity
                Dim info As New FileInfo(filePath)
                Return New EvtxSourceIdentity(info.Length, info.LastWriteTimeUtc)
            End Function

            Public Function Matches(filePath As String) As Boolean
                Dim info As New FileInfo(filePath)
                Return info.Length = _length AndAlso info.LastWriteTimeUtc = _lastWriteUtc
            End Function
        End Class

        Private NotInheritable Class EvtxFastModeCopy
            Implements IDisposable

            Private ReadOnly _directory As String
            Private ReadOnly _report As Action(Of Exception, String, String)
            Private _disposed As Boolean

            Public Sub New(directory As String, filePath As String, report As Action(Of Exception, String, String))
                _directory = directory
                _FilePath = filePath
                _report = report
            End Sub

            Public ReadOnly Property FilePath As String

            Public Sub Dispose() Implements IDisposable.Dispose
                If _disposed Then Return
                _disposed = True
                CleanupFastModeDirectory(_directory, _report)
            End Sub
        End Class

        Friend NotInheritable Class EvtxWorkerRequest
            Public Property FilePath As String
            Public Property QueryText As String
            Public Property QueryMode As SearchMode
            Public Property CaseSensitive As Boolean
            Public Property Settings As BeaconSettings
            Public Property XmlOnly As Boolean
            Public Property StartSequence As Long
            Public Property RemainingMatches As Integer
        End Class

        Friend NotInheritable Class EvtxWorkerMessage
            Public Property Type As String
            Public Property Sequence As Long
            Public Property Record As EventRecordSummary
            Public Property ErrorMessage As String
        End Class

        Private NotInheritable Class EvtxPhaseResult
            Public Property Completed As Boolean
            Public Property LimitReached As Boolean
            Public Property LastCheckpoint As Long
            Public Property TimedOut As Boolean
            Public Property FailureMessage As String
        End Class
    End Class
End Namespace
