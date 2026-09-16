Imports System.Diagnostics
Imports System.Threading
Imports System.Windows
Imports Beacon

Module SingleInstanceChecks
    Public Function RunChildMode(args As String()) As Integer
        Using coordinator As New SingleInstanceCoordinator(args(1), Sub() Console.WriteLine("ACTIVATE"))
            Console.WriteLine(If(coordinator.IsPrimary, "PRIMARY", "SECONDARY"))
            If args(2) = "hold" AndAlso coordinator.IsPrimary Then Thread.Sleep(Timeout.Infinite)
            Return If(coordinator.IsPrimary, 10, 0)
        End Using
    End Function

    Private Function StartChild(name As String, mode As String) As Process
        Dim start As New ProcessStartInfo(Environment.ProcessPath) With {
            .UseShellExecute = False, .CreateNoWindow = True, .RedirectStandardOutput = True, .RedirectStandardError = True
        }
        start.ArgumentList.Add("--single-instance-check")
        start.ArgumentList.Add(name)
        start.ArgumentList.Add(mode)
        Return Process.Start(start)
    End Function

    Private Sub Finish(child As Process, expectedCode As Integer)
        If Not child.WaitForExit(10000) Then
            child.Kill(True)
            Throw New TimeoutException("Single-instance child did not exit.")
        End If
        Require(child.ExitCode = expectedCode, "Unexpected child ownership: " & child.StandardOutput.ReadToEnd() & child.StandardError.ReadToEnd())
    End Sub

    Public Sub Run()
        Dim name = "Local\Beacon.InstanceTests." & Guid.NewGuid().ToString("N")
        Dim count As Integer
        Using signaled As New AutoResetEvent(False)
            Using primary As New SingleInstanceCoordinator(name, Sub()
                                                                    Interlocked.Increment(count)
                                                                    signaled.Set()
                                                                End Sub)
                Require(primary.IsPrimary, "First launch did not own the instance.")
                Using child = StartChild(name, "probe")
                    Finish(child, 0)
                End Using
                Require(signaled.WaitOne(5000), "Second launch did not request activation.")
                Dim children As New List(Of Process)()
                Try
                    For index = 1 To 4
                        children.Add(StartChild(name, "probe"))
                    Next
                    For Each child In children
                        Finish(child, 0)
                    Next
                    Require(signaled.WaitOne(5000), "Concurrent secondary launches did not request activation.")
                Finally
                    For Each child In children
                        If Not child.HasExited Then child.Kill(True)
                        child.Dispose()
                    Next
                End Try
            End Using
            Using child = StartChild(name, "probe")
                Finish(child, 10)
            End Using
        End Using
        Using owner = StartChild(name, "hold")
            Try
                Dim ready = owner.StandardOutput.ReadLineAsync()
                Require(ready.Wait(10000) AndAlso ready.Result = "PRIMARY", "Child did not acquire primary ownership.")
                Using secondary As New SingleInstanceCoordinator(name, Sub() Throw New InvalidOperationException("Secondary cannot receive activation."))
                    Require(Not secondary.IsPrimary, "Two processes became primary.")
                    owner.Kill(True)
                    Require(owner.WaitForExit(5000), "Terminated owner did not exit.")
                    Using replacement As New SingleInstanceCoordinator(name, Sub()
                                                                          End Sub)
                        Require(replacement.IsPrimary, "Abandoned owner prevented relaunch.")
                    End Using
                End Using
            Finally
                If Not owner.HasExited Then owner.Kill(True)
            End Try
        End Using
        Require(SingleInstanceCoordinator.InstanceName().StartsWith("Local\Beacon.FindInFiles.S-1-"), "Production instance scope does not include the Windows user.")
        CheckActivation()
    End Sub

    Private Sub CheckActivation()
        Dim app = System.Windows.Application.Current
        Dim previousMain = app.MainWindow
        Dim window As New Window With {.Title = "Instance activation test", .Width = 300, .Height = 150,
            .WindowStartupLocation = WindowStartupLocation.Manual, .Left = -10000, .Top = -10000,
            .ShowActivated = False, .ShowInTaskbar = False}
        Dim dialog As Window = Nothing
        Try
            app.MainWindow = window
            window.Show()
            window.WindowState = WindowState.Minimized
            InstanceWindowActivation.Activate(app)
            app.Dispatcher.Invoke(Sub()
                                  End Sub, System.Windows.Threading.DispatcherPriority.ContextIdle)
            Require(window.WindowState <> WindowState.Minimized, "Existing window was not restored.")
            dialog = New Window With {.Owner = window, .Title = "Owned dialog", .Width = 200, .Height = 100, .ShowActivated = False, .ShowInTaskbar = False}
            dialog.Show()
            window.IsEnabled = False
            Dim before = app.Windows.Count
            InstanceWindowActivation.Activate(app)
            Require(app.Windows.Count = before AndAlso dialog.IsVisible AndAlso Not window.IsEnabled, "Activation bypassed the owned dialog or created another window.")
        Finally
            dialog?.Close()
            window.Close()
            app.MainWindow = previousMain
        End Try
    End Sub

    Private Sub Require(condition As Boolean, message As String)
        If Not condition Then Throw New InvalidOperationException(message)
    End Sub
End Module
