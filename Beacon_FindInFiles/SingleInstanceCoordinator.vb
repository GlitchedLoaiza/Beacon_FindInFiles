Imports System.Diagnostics
Imports System.Security.Principal
Imports System.Threading

Namespace Beacon
    Public NotInheritable Class SingleInstanceCoordinator
        Implements IDisposable

        Private ReadOnly _mutex As Mutex
        Private ReadOnly _activation As EventWaitHandle
        Private ReadOnly _ownsMutex As Boolean
        Private _wait As RegisteredWaitHandle
        Private _disposed As Integer

        Public Shared Function InstanceName() As String
            Using identity = WindowsIdentity.GetCurrent(), currentProcess = Process.GetCurrentProcess()
                Return "Local\Beacon.FindInFiles." & identity.User.Value & "." & currentProcess.SessionId.ToString()
            End Using
        End Function

        Public Sub New(name As String, activate As Action)
            _mutex = New Mutex(False, name & ".Mutex")
            Try
                ' Create the event before ownership is checked so a second launch cannot lose its signal during startup.
                _activation = New EventWaitHandle(False, EventResetMode.AutoReset, name & ".Activate")
                Try
                    _ownsMutex = _mutex.WaitOne(0)
                Catch ex As AbandonedMutexException
                    _ownsMutex = True
                End Try
                If _ownsMutex Then
                    _wait = ThreadPool.RegisterWaitForSingleObject(_activation,
                        Sub(state, timedOut)
                            If Volatile.Read(_disposed) <> 0 Then Return
                            Try
                                activate()
                            Catch ex As Exception
                                Debug.WriteLine($"Beacon activation request failed: {ex.Message}")
                            End Try
                        End Sub, Nothing, Timeout.Infinite, False)
                Else
                    _activation.Set()
                End If
            Catch
                If _ownsMutex Then _mutex.ReleaseMutex()
                _activation?.Dispose()
                _mutex.Dispose()
                Throw
            End Try
        End Sub

        Public ReadOnly Property IsPrimary As Boolean
            Get
                Return _ownsMutex
            End Get
        End Property

        Public Sub Dispose() Implements IDisposable.Dispose
            If Interlocked.Exchange(_disposed, 1) <> 0 Then Return
            _wait?.Unregister(Nothing)
            If _ownsMutex Then _mutex.ReleaseMutex()
            _activation.Dispose()
            _mutex.Dispose()
        End Sub
    End Class
End Namespace
