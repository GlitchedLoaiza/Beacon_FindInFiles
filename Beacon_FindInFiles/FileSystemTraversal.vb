Imports System.IO
Imports System.Threading

Namespace Beacon

    Public NotInheritable Class FileSystemTraversal

        Private Sub New()
        End Sub

        Public Shared Iterator Function EnumerateFiles(rootPath As String,
                                                        cancellationToken As CancellationToken,
                                                        followReparsePoints As Boolean,
                                                        accessDeniedHandler As Func(Of String, Boolean),
                                                        errorHandler As Action(Of String, Exception),
                                                        Optional settings As BeaconSettings = Nothing) As IEnumerable(Of String)
            Dim pendingDirectories As New Stack(Of String)()
            Dim visited As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
            Dim excluded As New HashSet(Of String)(If(settings?.ExcludedDirectories, "").Split(";"c, StringSplitOptions.RemoveEmptyEntries), StringComparer.OrdinalIgnoreCase)
            pendingDirectories.Push(Path.GetFullPath(rootPath))

            While pendingDirectories.Count > 0
                cancellationToken.ThrowIfCancellationRequested()
                Dim currentDirectory = pendingDirectories.Pop()
                Try
                    Dim info As New DirectoryInfo(currentDirectory)
                    If (info.Attributes And FileAttributes.ReparsePoint) <> 0 Then
                        If Not followReparsePoints Then Continue While
                        Dim target = info.ResolveLinkTarget(True)
                        If target Is Nothing Then
                            errorHandler?.Invoke(currentDirectory, New IOException("Cannot resolve reparse point safely."))
                            Continue While
                        End If
                        currentDirectory = target.FullName
                    End If
                    If Not visited.Add(Path.GetFullPath(currentDirectory).TrimEnd(Path.DirectorySeparatorChar)) Then Continue While
                Catch ex As Exception When TypeOf ex Is IOException OrElse TypeOf ex Is UnauthorizedAccessException
                    errorHandler?.Invoke(currentDirectory, ex)
                    Continue While
                End Try
                Dim files As String() = Nothing
                Dim directories As String() = Nothing

                If Not TryReadDirectory(currentDirectory, files, directories, accessDeniedHandler, errorHandler) Then
                    Continue While
                End If

                For Each filePath In files
                    cancellationToken.ThrowIfCancellationRequested()
                    If settings IsNot Nothing AndAlso Not IsFileAllowed(filePath, settings, errorHandler) Then Continue For
                    Yield filePath
                Next

                For Each directoryPath In directories
                    cancellationToken.ThrowIfCancellationRequested()
                    If excluded.Contains(Path.GetFileName(directoryPath)) Then Continue For
                    Try
                        Dim attributes = File.GetAttributes(directoryPath)
                        If Not followReparsePoints AndAlso (attributes And FileAttributes.ReparsePoint) <> 0 Then Continue For
                        If settings IsNot Nothing AndAlso Not AttributesAllowed(attributes, settings) Then Continue For
                    Catch ex As Exception When TypeOf ex Is IOException OrElse TypeOf ex Is UnauthorizedAccessException
                        errorHandler?.Invoke(directoryPath, ex)
                        Continue For
                    End Try

                    pendingDirectories.Push(directoryPath)
                Next
            End While
        End Function

        Public Shared Function IsFileAllowed(filePath As String, settings As BeaconSettings,
                                             errorHandler As Action(Of String, Exception)) As Boolean
            Try
                Dim info As New FileInfo(filePath)
                If Not AttributesAllowed(info.Attributes, settings) Then Return False
                If (info.Attributes And FileAttributes.ReparsePoint) <> 0 Then
                    If Not settings.FollowReparsePoints Then Return False
                    Dim target = info.ResolveLinkTarget(True)
                    If target Is Nothing Then Throw New IOException("Cannot resolve file reparse point safely.")
                    info = New FileInfo(target.FullName)
                    If Not AttributesAllowed(info.Attributes, settings) Then Return False
                End If
                If info.Length > CLng(settings.MaximumFileSizeMb) * 1024 * 1024 Then
                    Throw New IOException($"File exceeds the {settings.MaximumFileSizeMb} MB scan limit.")
                End If
                Return True
            Catch ex As Exception When TypeOf ex Is IOException OrElse TypeOf ex Is UnauthorizedAccessException
                errorHandler?.Invoke(filePath, ex)
                Return False
            End Try
        End Function

        Private Shared Function AttributesAllowed(attributes As FileAttributes, settings As BeaconSettings) As Boolean
            Return (settings.IncludeHiddenFiles OrElse (attributes And FileAttributes.Hidden) = 0) AndAlso
                   (settings.IncludeSystemFiles OrElse (attributes And FileAttributes.System) = 0)
        End Function

        Private Shared Function TryReadDirectory(directoryPath As String,
                                                 ByRef files As String(),
                                                 ByRef directories As String(),
                                                 accessDeniedHandler As Func(Of String, Boolean),
                                                 errorHandler As Action(Of String, Exception)) As Boolean
            Try
                files = Directory.GetFiles(directoryPath, "*", SearchOption.TopDirectoryOnly)
                directories = Directory.GetDirectories(directoryPath, "*", SearchOption.TopDirectoryOnly)
                Return True
            Catch ex As UnauthorizedAccessException
                If accessDeniedHandler IsNot Nothing AndAlso accessDeniedHandler(directoryPath) Then
                    Try
                        files = Directory.GetFiles(directoryPath, "*", SearchOption.TopDirectoryOnly)
                        directories = Directory.GetDirectories(directoryPath, "*", SearchOption.TopDirectoryOnly)
                        Return True
                    Catch retryException As Exception When TypeOf retryException Is UnauthorizedAccessException OrElse
                                                              TypeOf retryException Is IOException
                        errorHandler?.Invoke(directoryPath, retryException)
                        Return False
                    End Try
                End If

                errorHandler?.Invoke(directoryPath, ex)
                Return False
            Catch ex As IOException
                errorHandler?.Invoke(directoryPath, ex)
                Return False
            End Try
        End Function

    End Class

End Namespace
