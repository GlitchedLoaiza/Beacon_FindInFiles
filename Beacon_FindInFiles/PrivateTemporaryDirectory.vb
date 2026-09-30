Imports System.ComponentModel
Imports System.Diagnostics
Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Security.AccessControl
Imports System.Security.Principal
Imports System.Threading
Imports Microsoft.Win32.SafeHandles

Namespace Beacon
    Friend NotInheritable Class PrivateTemporaryDirectory
        Implements IDisposable

        Private Const DeleteAccess As UInteger = &H10000UI
        Private Const GenericRead As UInteger = &H80000000UI
        Private Const OpenExisting As UInteger = 3UI
        Private Const OpenReparsePoint As UInteger = &H200000UI
        Private Const BackupSemantics As UInteger = &H2000000UI
        Private Const AttributeTagInfo As Integer = 9
        Private Shared ReadOnly CurrentUser As SecurityIdentifier = GetCurrentUser()
        Private ReadOnly _lease As DirectoryLease
        Private _disposed As Integer

        <StructLayout(LayoutKind.Sequential)>
        Private Structure SecurityAttributes
            Public Length As Integer
            Public Descriptor As IntPtr
            Public InheritHandle As Integer
        End Structure

        <StructLayout(LayoutKind.Sequential)>
        Private Structure AttributeTag
            Public Attributes As FileAttributes
            Public ReparseTag As UInteger
        End Structure

        <DllImport("kernel32.dll", EntryPoint:="CreateDirectoryW", CharSet:=CharSet.Unicode, ExactSpelling:=True, SetLastError:=True), DefaultDllImportSearchPaths(DllImportSearchPath.System32)>
        Private Shared Function CreateDirectoryNative(path As String, ByRef security As SecurityAttributes) As Boolean
        End Function

        <DllImport("kernel32.dll", EntryPoint:="CreateFileW", CharSet:=CharSet.Unicode, ExactSpelling:=True, SetLastError:=True), DefaultDllImportSearchPaths(DllImportSearchPath.System32)>
        Private Shared Function CreateFileNative(path As String, access As UInteger, share As FileShare, security As IntPtr,
                                                 disposition As UInteger, flags As UInteger, template As IntPtr) As SafeFileHandle
        End Function

        <DllImport("kernel32.dll", ExactSpelling:=True, SetLastError:=True), DefaultDllImportSearchPaths(DllImportSearchPath.System32)>
        Private Shared Function GetFileInformationByHandleEx(handle As SafeFileHandle, informationClass As Integer,
                                                             ByRef information As AttributeTag, size As UInteger) As Boolean
        End Function

        Private Sub New(path As String, lease As DirectoryLease)
            DirectoryPath = path
            _lease = lease
        End Sub

        Public ReadOnly Property DirectoryPath As String

        Public Shared Function Create(prefix As String) As PrivateTemporaryDirectory
            ValidateLeaf(prefix)
            Dim parent = IO.Path.GetFullPath(IO.Path.GetTempPath())
            Dim lease = LockDirectory(parent)
            Dim path = IO.Path.Combine(parent, prefix & Guid.NewGuid().ToString("N"))
            Try
                CreatePrivateDirectory(path)
                lease.Add(path)
                If Not HasPrivatePermissions(path) Then Throw New IOException("The temporary filesystem did not apply Beacon's private permissions.")
                Return New PrivateTemporaryDirectory(path, lease)
            Catch
                lease.Dispose()
                Throw
            End Try
        End Function

        Public Function CreateFile(name As String, Optional share As FileShare = FileShare.None,
                                   Optional options As FileOptions = FileOptions.None) As FileStream
            ObjectDisposedException.ThrowIf(Volatile.Read(_disposed) <> 0, Me)
            ValidateLeaf(name)
            Return New FileStream(Path.Combine(DirectoryPath, name), FileMode.CreateNew, FileAccess.ReadWrite, share, 81920, options)
        End Function

        Public Shared Function CreateTemporaryStream(prefix As String, extension As String) As FileStream
            ValidateLeaf("content" & extension)
            Dim directory = Create(prefix)
            Try
                Return New OwnedFileStream(Path.Combine(directory.DirectoryPath, "content" & extension), directory)
            Catch
                directory.Dispose()
                Throw
            End Try
        End Function

        Public Shared Function OpenRead(path As String, Optional share As FileShare = FileShare.Read) As FileStream
            Dim handle = CreateFileNative(NativePath(IO.Path.GetFullPath(path)), GenericRead, share, IntPtr.Zero, OpenExisting, OpenReparsePoint, IntPtr.Zero)
            If handle.IsInvalid Then
                Dim errorCode = Marshal.GetLastWin32Error()
                handle.Dispose()
                Throw NativeError("Could not open the protected file.", errorCode)
            End If
            Try
                VerifyHandle(handle, False)
                Return New FileStream(handle, FileAccess.Read, 81920, False)
            Catch
                handle.Dispose()
                Throw
            End Try
        End Function

        Public Shared Function LockDirectory(path As String, Optional createMissing As Boolean = False) As DirectoryLease
            Dim fullPath = IO.Path.TrimEndingDirectorySeparator(IO.Path.GetFullPath(path))
            Dim root = IO.Path.GetPathRoot(fullPath)
            If String.IsNullOrEmpty(root) OrElse root.StartsWith("\\", StringComparison.Ordinal) Then
                Throw New IOException("Beacon temporary storage requires a local, non-redirected directory.")
            End If
            Dim lease As New DirectoryLease()
            Try
                lease.Add(root)
                Dim current = root
                For Each part In fullPath.Substring(root.Length).Split(IO.Path.DirectorySeparatorChar, StringSplitOptions.RemoveEmptyEntries)
                    current = IO.Path.Combine(current, part)
                    Dim created = createMissing AndAlso Not Directory.Exists(current)
                    If created Then CreatePrivateDirectory(current)
                    lease.Add(current)
                    If created AndAlso Not HasPrivatePermissions(current) Then Throw New IOException("The temporary filesystem did not apply Beacon's private permissions.")
                Next
                Return lease
            Catch
                lease.Dispose()
                Throw
            End Try
        End Function

        Public Shared Sub CleanupStale(prefix As String)
            ValidateLeaf(prefix)
            Try
                Using parent = LockDirectory(IO.Path.GetTempPath())
                    For Each path In Directory.EnumerateDirectories(IO.Path.GetTempPath(), prefix & "*", SearchOption.TopDirectoryOnly)
                        Dim suffix = IO.Path.GetFileName(path).Substring(prefix.Length)
                        Dim id As Guid
                        If Not Guid.TryParseExact(suffix, "N", id) Then Continue For
                        Try
                            Using lease = LockDirectory(IO.Path.GetDirectoryName(path))
                                ' A live session holds a handle that denies this DELETE-access claim.
                                lease.Add(path, True)
                                If Not HasPrivatePermissions(path) Then Continue For
                                DeleteContents(path)
                                lease.ReleaseLast()
                                Directory.Delete(path, False)
                            End Using
                        Catch ex As Exception When TypeOf ex Is IOException OrElse TypeOf ex Is UnauthorizedAccessException
                            Debug.WriteLine($"Beacon temporary cleanup skipped an active or unavailable directory: {ex.Message}")
                        End Try
                    Next
                End Using
            Catch ex As Exception When TypeOf ex Is IOException OrElse TypeOf ex Is UnauthorizedAccessException
                Debug.WriteLine($"Beacon temporary cleanup unavailable: {ex.Message}")
            End Try
        End Sub

        Public Sub Dispose() Implements IDisposable.Dispose
            If Interlocked.Exchange(_disposed, 1) <> 0 Then Return
            Try
                DeleteContents(DirectoryPath)
                _lease.ReleaseLast()
                Directory.Delete(DirectoryPath, False)
            Catch ex As Exception When TypeOf ex Is IOException OrElse TypeOf ex Is UnauthorizedAccessException
                Debug.WriteLine($"Could not remove Beacon private temporary directory: {ex.Message}")
            Finally
                _lease.Dispose()
            End Try
        End Sub

        Private Shared Sub DeleteContents(path As String)
            For Each entry In Directory.EnumerateFileSystemEntries(path)
                Dim attributes = File.GetAttributes(entry)
                If (attributes And FileAttributes.Directory) = 0 Then
                    File.Delete(entry)
                ElseIf (attributes And FileAttributes.ReparsePoint) <> 0 Then
                    Directory.Delete(entry, False)
                Else
                    Using child As New DirectoryLease()
                        child.Add(entry)
                        DeleteContents(entry)
                    End Using
                    ' Nonrecursive deletion cannot follow a replacement link after the handle closes.
                    Directory.Delete(entry, False)
                End If
            Next
        End Sub

        Private Shared Sub CreatePrivateDirectory(path As String)
            Dim descriptor = PrivateSecurity().GetSecurityDescriptorBinaryForm()
            Dim memory = Marshal.AllocHGlobal(descriptor.Length)
            Try
                Marshal.Copy(descriptor, 0, memory, descriptor.Length)
                Dim security As New SecurityAttributes With {.Length = Marshal.SizeOf(Of SecurityAttributes)(), .Descriptor = memory}
                If Not CreateDirectoryNative(NativePath(path), security) Then
                    Throw NativeError("Could not exclusively create the private temporary directory.", Marshal.GetLastWin32Error())
                End If
            Finally
                Marshal.FreeHGlobal(memory)
            End Try
        End Sub

        Private Shared Function PrivateSecurity() As DirectorySecurity
            Dim security As New DirectorySecurity()
            security.SetOwner(CurrentUser)
            security.SetAccessRuleProtection(True, False)
            Dim inheritance = InheritanceFlags.ContainerInherit Or InheritanceFlags.ObjectInherit
            security.AddAccessRule(New FileSystemAccessRule(CurrentUser, FileSystemRights.FullControl, inheritance, PropagationFlags.None, AccessControlType.Allow))
            Dim system = New SecurityIdentifier(WellKnownSidType.LocalSystemSid, Nothing)
            If Not CurrentUser.Equals(system) Then security.AddAccessRule(New FileSystemAccessRule(system, FileSystemRights.FullControl, inheritance, PropagationFlags.None, AccessControlType.Allow))
            Return security
        End Function

        Private Shared Function HasPrivatePermissions(path As String) As Boolean
            Dim security = New DirectoryInfo(path).GetAccessControl(AccessControlSections.Owner Or AccessControlSections.Access)
            If Not security.AreAccessRulesProtected OrElse Not CurrentUser.Equals(security.GetOwner(GetType(SecurityIdentifier))) Then Return False
            For Each rule As FileSystemAccessRule In security.GetAccessRules(True, True, GetType(SecurityIdentifier))
                If rule.AccessControlType = AccessControlType.Allow AndAlso Not rule.IdentityReference.Equals(CurrentUser) AndAlso
                   Not rule.IdentityReference.Equals(New SecurityIdentifier(WellKnownSidType.LocalSystemSid, Nothing)) Then Return False
            Next
            Return True
        End Function

        Private Shared Function GetCurrentUser() As SecurityIdentifier
            Using identity = WindowsIdentity.GetCurrent()
                If identity.User Is Nothing Then Throw New UnauthorizedAccessException("The current Windows identity has no user SID.")
                Return identity.User
            End Using
        End Function

        Private Shared Sub ValidateLeaf(name As String)
            If String.IsNullOrWhiteSpace(name) OrElse name = "." OrElse name = ".." OrElse
               name.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0 OrElse name.EndsWith(".", StringComparison.Ordinal) OrElse name.EndsWith(" ", StringComparison.Ordinal) Then
                Throw New ArgumentException("A temporary resource must use a simple file name.", NameOf(name))
            End If
        End Sub

        Private Shared Function NativePath(path As String) As String
            Return If(path.StartsWith("\\?\", StringComparison.Ordinal), path, "\\?\" & path)
        End Function

        Private Shared Function NativeError(message As String, errorCode As Integer) As IOException
            Return New IOException(message, New Win32Exception(errorCode))
        End Function

        Private Shared Sub VerifyHandle(handle As SafeFileHandle, directory As Boolean)
            Dim information As AttributeTag
            If Not GetFileInformationByHandleEx(handle, AttributeTagInfo, information, CUInt(Marshal.SizeOf(Of AttributeTag)())) Then
                Throw NativeError("Could not inspect the temporary resource handle.", Marshal.GetLastWin32Error())
            End If
            If (information.Attributes And FileAttributes.ReparsePoint) <> 0 OrElse
               ((information.Attributes And FileAttributes.Directory) <> 0) <> directory Then
                Throw New IOException("Beacon refused a redirected or unexpected temporary resource.")
            End If
        End Sub

        Friend NotInheritable Class DirectoryLease
            Implements IDisposable
            Private ReadOnly _handles As New List(Of SafeFileHandle)()

            Friend Sub Add(path As String, Optional claimDeletion As Boolean = False)
                Dim access = GenericRead Or If(claimDeletion, DeleteAccess, 0UI)
                Dim handle = CreateFileNative(NativePath(path), access, FileShare.Read Or FileShare.Write, IntPtr.Zero,
                                              OpenExisting, BackupSemantics Or OpenReparsePoint, IntPtr.Zero)
                If handle.IsInvalid Then
                    Dim errorCode = Marshal.GetLastWin32Error()
                    handle.Dispose()
                    Throw NativeError("Could not lock the temporary directory path.", errorCode)
                End If
                Try
                    VerifyHandle(handle, True)
                    _handles.Add(handle)
                Catch
                    handle.Dispose()
                    Throw
                End Try
            End Sub

            Friend Sub ReleaseLast()
                If _handles.Count = 0 Then Return
                Dim index = _handles.Count - 1
                _handles(index).Dispose()
                _handles.RemoveAt(index)
            End Sub

            Public Sub Dispose() Implements IDisposable.Dispose
                While _handles.Count > 0
                    ReleaseLast()
                End While
            End Sub
        End Class

        Private NotInheritable Class OwnedFileStream
            Inherits FileStream
            Private _directory As PrivateTemporaryDirectory

            Public Sub New(path As String, directory As PrivateTemporaryDirectory)
                MyBase.New(path, FileMode.CreateNew, FileAccess.ReadWrite, FileShare.None, 81920, FileOptions.DeleteOnClose)
                _directory = directory
            End Sub

            Protected Overrides Sub Dispose(disposing As Boolean)
                Try
                    MyBase.Dispose(disposing)
                Finally
                    If disposing Then Interlocked.Exchange(_directory, Nothing)?.Dispose()
                End Try
            End Sub
        End Class
    End Class
End Namespace
