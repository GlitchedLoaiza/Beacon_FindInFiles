Imports System.IO
Imports System.Diagnostics
Imports System.Reflection

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

                ' Extract to temp directory
                Dim tempDir = Path.Combine(Path.GetTempPath(), "Beacon_7zip")
                Directory.CreateDirectory(tempDir)

                _extractedSevenZipPath = Path.Combine(tempDir, "7za.exe")

                ' Only extract if not already present
                If Not File.Exists(_extractedSevenZipPath) Then
                    Using resourceStream = assembly.GetManifestResourceStream(resourceName)
                        If resourceStream Is Nothing Then
                            Debug.WriteLine("ERROR: Could not open resource stream for " & resourceName)
                            Return Nothing
                        End If

                        Using fileStream As New FileStream(_extractedSevenZipPath, FileMode.Create, FileAccess.Write)
                            resourceStream.CopyTo(fileStream)
                        End Using
                    End Using

                    Debug.WriteLine("✓ Extracted 7za.exe to: " & _extractedSevenZipPath)
                Else
                    Debug.WriteLine("✓ Using existing 7za.exe at: " & _extractedSevenZipPath)
                End If

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
    Public Shared Function ExtractCab(cabPath As String, outputDir As String) As Boolean
        Try
            ' Ensure output directory exists
            Directory.CreateDirectory(outputDir)

            ' Get path to 7za.exe (extracted from embedded resources)
            Dim sevenZipPath As String = EnsureSevenZipExtracted()

            If String.IsNullOrEmpty(sevenZipPath) OrElse Not File.Exists(sevenZipPath) Then
                Debug.WriteLine("ERROR: 7za.exe not available")
                Return False
            End If

            Debug.WriteLine("Using 7za.exe from: " & sevenZipPath)

            ' Build process start info
            ' x = extract with full paths
            ' -o = output directory
            ' -y = assume Yes on all queries (overwrite without prompting)
            Dim psi As New ProcessStartInfo() With {
                .FileName = sevenZipPath,
                .Arguments = $"x ""{cabPath}"" -o""{outputDir}"" -y",
                .UseShellExecute = False,
                .CreateNoWindow = True,
                .RedirectStandardOutput = True,
                .RedirectStandardError = True
            }

            Debug.WriteLine("Executing: " & psi.FileName & " " & psi.Arguments)

            Dim process As Process = Process.Start(psi)
            Using process
                Dim output = process.StandardOutput.ReadToEnd()
                Dim errorOutput = process.StandardError.ReadToEnd()
                process.WaitForExit()

                If process.ExitCode = 0 Then
                    Debug.WriteLine("7z extraction successful")
                    Return True
                Else
                    Debug.WriteLine("7z extraction failed with exit code " & process.ExitCode.ToString())
                    If Not String.IsNullOrEmpty(errorOutput) Then
                        Debug.WriteLine("Error output: " & errorOutput)
                    End If
                    Return False
                End If
            End Using

        Catch ex As Exception
            Debug.WriteLine("Exception during CAB extraction: " & ex.Message)
            Debug.WriteLine("Stack trace: " & ex.StackTrace)
            Return False
        End Try
    End Function

    ''' <summary>
    ''' Lists files in a CAB archive without extracting them
    ''' </summary>
    ''' <param name="cabPath">Full path to the CAB file</param>
    ''' <returns>List of file names in the archive</returns>
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
                .Arguments = $"l ""{cabPath}""",
                .UseShellExecute = False,
                .CreateNoWindow = True,
                .RedirectStandardOutput = True,
                .RedirectStandardError = True
            }

            Dim process As Process = Process.Start(psi)
            Using process
                Dim output = process.StandardOutput.ReadToEnd()
                process.WaitForExit()

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
                            ' 7z list format: Date Time Attr Size Compressed Name
                            ' We just want the Name (last column)
                            Dim parts = line.Split(New Char() {" "c}, StringSplitOptions.RemoveEmptyEntries)
                            If parts.Length >= 6 Then
                                Dim fileName = String.Join(" ", parts.Skip(5))
                                If Not String.IsNullOrEmpty(fileName) AndAlso Not fileName.EndsWith("\") Then
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
