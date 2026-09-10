Imports System.IO
Imports System.Diagnostics
Imports System.Reflection
Imports System.Security.Cryptography

''' <summary>
''' Helper class for extracting CAB files using 7-Zip command line tool
''' Extracts embedded 7za.exe from resources at runtime
''' </summary>
Public Class SevenZipHelper

    Private Shared _extractedSevenZipPath As String = Nothing
    Private Shared ReadOnly _lockObject As New Object()

    ''' <summary>
    ''' Extracts embedded 7za.exe to a temporary location (once per application run)
    ''' </summary>
    ''' <returns>Path to extracted 7za.exe, or Nothing if extraction failed</returns>
    Private Shared Function EnsureSevenZipExtracted() As String
        SyncLock _lockObject
            ' If already extracted, return cached path
            If Not String.IsNullOrEmpty(_extractedSevenZipPath) AndAlso File.Exists(_extractedSevenZipPath) Then
                Return _extractedSevenZipPath
            End If

            Try
                ' Try to extract 7za.exe from embedded resources
                Dim assembly As Assembly = Assembly.GetExecutingAssembly()
                Dim resourceName As String = assembly.GetManifestResourceNames().FirstOrDefault(Function(r) r.EndsWith("7za.exe"))

                If String.IsNullOrEmpty(resourceName) Then
                    Debug.WriteLine("ERROR: 7za.exe not found in embedded resources")
                    Debug.WriteLine("Available resources: " & String.Join(", ", assembly.GetManifestResourceNames()))
                    Return Nothing
                End If

                ' Extract to an application-owned, versioned directory.
                Dim version = assembly.GetName().Version?.ToString()
                If String.IsNullOrWhiteSpace(version) Then version = "unknown"
                Dim tempDir = Path.Combine(Path.GetTempPath(), "Beacon", "Tools", version)
                Directory.CreateDirectory(tempDir)

                _extractedSevenZipPath = Path.Combine(tempDir, "7za.exe")

                Using resourceStream = assembly.GetManifestResourceStream(resourceName)
                    If resourceStream Is Nothing Then
                        Debug.WriteLine("ERROR: Could not open resource stream for " & resourceName)
                        Return Nothing
                    End If

                    Dim embeddedHash = SHA256.HashData(resourceStream)
                    Dim existingIsValid = File.Exists(_extractedSevenZipPath) AndAlso
                        CryptographicOperations.FixedTimeEquals(embeddedHash, SHA256.HashData(File.ReadAllBytes(_extractedSevenZipPath)))

                    If Not existingIsValid Then
                        resourceStream.Position = 0
                        Dim pendingPath = _extractedSevenZipPath & ".new"
                        Using fileStream As New FileStream(pendingPath, FileMode.Create, FileAccess.Write, FileShare.None)
                            resourceStream.CopyTo(fileStream)
                            fileStream.Flush(True)
                        End Using
                        File.Move(pendingPath, _extractedSevenZipPath, True)
                        Debug.WriteLine("✓ Extracted verified 7za.exe to: " & _extractedSevenZipPath)
                    Else
                        Debug.WriteLine("✓ Verified existing 7za.exe at: " & _extractedSevenZipPath)
                    End If
                End Using

                Return _extractedSevenZipPath

            Catch ex As Exception
                Debug.WriteLine("ERROR extracting 7za.exe: " & ex.Message)
                Debug.WriteLine("Stack trace: " & ex.StackTrace)
                Return Nothing
            End Try
        End SyncLock
    End Function

    ''' <summary>
    ''' Extracts a CAB file to the specified output directory using 7z.exe
    ''' </summary>
    ''' <param name="cabPath">Full path to the CAB file</param>
    ''' <param name="outputDir">Directory where files should be extracted</param>
    ''' <returns>True if extraction succeeded, False otherwise</returns>
    Public Shared Function ExtractCab(cabPath As String, outputDir As String,
                                      Optional settings As Beacon.BeaconSettings = Nothing,
                                      Optional token As Threading.CancellationToken = Nothing,
                                      Optional report As Action(Of String, Exception) = Nothing) As Boolean
        Try
            Dim sevenZipPath As String = EnsureSevenZipExtracted()
            If String.IsNullOrEmpty(sevenZipPath) OrElse Not File.Exists(sevenZipPath) Then
                Throw New IOException("Embedded 7-Zip reader is unavailable.")
            End If
            ' Run the async pipe readers off the WPF dispatcher even for synchronous preview callers.
            Threading.Tasks.Task.Run(Function() Beacon.CabExtractor.ExtractAsync(sevenZipPath, cabPath, outputDir,
                                       If(settings, Beacon.BeaconSettings.CreateDefaults()), token)).GetAwaiter().GetResult()
            Return True
        Catch ex As OperationCanceledException
            If token.IsCancellationRequested Then Throw
            report?.Invoke(cabPath, New TimeoutException("CAB extraction exceeded its configured timeout.", ex))
            Return False
        Catch ex As Exception
            Debug.WriteLine("Exception during CAB extraction: " & ex.Message)
            report?.Invoke(cabPath, ex)
            Return False
        End Try
    End Function

    ''' <summary>
    ''' Lists files in a CAB archive without extracting them
    ''' Parses 7-Zip list output format: Date Time Attr Size Name (5 parts minimum)
    ''' Filters out directories (end with \) and summary lines (contain "files"/"folders" keywords)
    ''' </summary>
    ''' <param name="cabPath">Full path to the CAB file</param>
    ''' <returns>List of file names (relative paths) in the archive</returns>
    Public Shared Function ListCabFiles(cabPath As String) As List(Of String)
        Dim files As New List(Of String)

        Try
            ' Get path to 7za.exe (extracted from embedded resources)
            Dim sevenZipPath As String = EnsureSevenZipExtracted()

            If String.IsNullOrEmpty(sevenZipPath) OrElse Not File.Exists(sevenZipPath) Then
                Debug.WriteLine("ERROR: 7za.exe not available for listing")
                Return files
            End If

            ' l = list files
            ' -slt = show technical information (one file per section)
            Dim psi As New ProcessStartInfo() With {
                .FileName = sevenZipPath,
                .UseShellExecute = False,
                .CreateNoWindow = True,
                .RedirectStandardOutput = True,
                .RedirectStandardError = True
            }
            psi.ArgumentList.Add("l")
            psi.ArgumentList.Add(cabPath)

            Dim process As Process = Process.Start(psi)
            Using process
                Dim outputTask = process.StandardOutput.ReadToEndAsync()
                Dim errorTask = process.StandardError.ReadToEndAsync()
                If Not process.WaitForExit(60000) Then
                    process.Kill(True)
                    Debug.WriteLine("7z listing timed out")
                    Return files
                End If
                Dim output = outputTask.GetAwaiter().GetResult()
                errorTask.GetAwaiter().GetResult()

                If process.ExitCode = 0 Then
                    ' Parse output - 7z list format has file names after the header
                    Dim lines() As String = output.Split(New String() {Environment.NewLine}, StringSplitOptions.None)
                    Dim inFileList As Boolean = False

                    For Each line As String In lines
                        ' Detect start of file listing (after the "---" separator)
                        If line.StartsWith("-------------------") Then
                            inFileList = True
                            Continue For
                        End If

                        ' Stop at end of file listing (another "---" separator)
                        If inFileList AndAlso line.StartsWith("-------------------") Then
                            Exit For
                        End If

                        ' Parse file entries (skip directories)
                        If inFileList AndAlso Not String.IsNullOrWhiteSpace(line) Then
                            ' 7z list format: Date Time Attr Size Name
                            ' We want the Name (starting from 5th column, index 4)
                            Dim parts = line.Split(New Char() {" "c}, StringSplitOptions.RemoveEmptyEntries)
                            ' Files have at least 5 parts: Date, Time, Attr, Size, Name
                            If parts.Length >= 5 Then
                                ' Extract filename (everything from index 4 onwards)
                                Dim fileName = String.Join(" ", parts.Skip(4))
                                ' Skip if it's a directory (ends with \) or summary line (contains digits followed by "files")
                                If Not String.IsNullOrEmpty(fileName) AndAlso 
                                   Not fileName.EndsWith("\") AndAlso 
                                   Not fileName.EndsWith("files") AndAlso
                                   Not fileName.EndsWith("folders") Then
                                    files.Add(fileName)
                                End If
                            End If
                        End If
                    Next
                End If
            End Using

        Catch ex As Exception
            Debug.WriteLine("Exception listing CAB files: " & ex.Message)
        End Try

        Return files
    End Function

End Class
