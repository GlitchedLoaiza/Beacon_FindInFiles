Imports System.IO
Imports System.Diagnostics

''' <summary>
''' Helper class for extracting CAB files using 7-Zip command line tool
''' Uses the 7-Zip.CommandLine NuGet package (includes 7z.exe in output directory)
''' </summary>
Public Class SevenZipHelper

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

            ' 7za.exe is included in the NuGet package and copied to output directory via build target
            Dim sevenZipPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "7za.exe")

            ' Fallback: check in x64 subdirectory
            If Not File.Exists(sevenZipPath) Then
                sevenZipPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "x64", "7za.exe")
            End If

            ' Fallback: check in tools subdirectory
            If Not File.Exists(sevenZipPath) Then
                sevenZipPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "tools", "x64", "7za.exe")
            End If

            ' Fallback: try NuGet packages folder directly (for debugging)
            If Not File.Exists(sevenZipPath) Then
                Dim nugetPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), ".nuget", "packages", "7-zip.commandline", "25.1.0", "tools", "x64", "7za.exe")
                If File.Exists(nugetPath) Then
                    sevenZipPath = nugetPath
                End If
            End If

            If Not File.Exists(sevenZipPath) Then
                Debug.WriteLine($"ERROR: 7za.exe not found. Expected at: {sevenZipPath}")
                Return False
            End If

            Debug.WriteLine($"Using 7z.exe from: {sevenZipPath}")

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

            Debug.WriteLine($"Executing: {psi.FileName} {psi.Arguments}")

            Dim process As Process = Process.Start(psi)
            Using process
                Dim output = process.StandardOutput.ReadToEnd()
                Dim errorOutput = process.StandardError.ReadToEnd()
                process.WaitForExit()

                If process.ExitCode = 0 Then
                    Debug.WriteLine("7z extraction successful")
                    Return True
                Else
                    Debug.WriteLine($"7z extraction failed with exit code {process.ExitCode}")
                    If Not String.IsNullOrEmpty(errorOutput) Then
                        Debug.WriteLine("Error output: " & errorOutput)
                    End If
                    Return False
                End If
            End Using

        Catch ex As Exception
            Debug.WriteLine($"Exception during CAB extraction: {ex.Message}")
            Debug.WriteLine($"Stack trace: {ex.StackTrace}")
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
            Dim sevenZipPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "7za.exe")

            If Not File.Exists(sevenZipPath) Then
                sevenZipPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "x64", "7za.exe")
            End If

            If Not File.Exists(sevenZipPath) Then
                sevenZipPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "tools", "x64", "7za.exe")
            End If

            ' Fallback: try NuGet packages folder directly (for debugging)
            If Not File.Exists(sevenZipPath) Then
                Dim nugetPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), ".nuget", "packages", "7-zip.commandline", "25.1.0", "tools", "x64", "7za.exe")
                If File.Exists(nugetPath) Then
                    sevenZipPath = nugetPath
                End If
            End If

            If Not File.Exists(sevenZipPath) Then
                Debug.WriteLine("ERROR: 7za.exe not found for listing")
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
            Debug.WriteLine($"Exception listing CAB files: {ex.Message}")
        End Try

        Return files
    End Function

End Class
