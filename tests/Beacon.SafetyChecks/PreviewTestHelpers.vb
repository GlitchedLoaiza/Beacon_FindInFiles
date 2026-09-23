Imports System.Diagnostics
Imports System.Threading
Imports System.Windows.Threading
Imports Beacon

Friend Module PreviewTestHelpers
    Public Sub WaitForPreview(window As MainWindow)
        Dim timeout = Stopwatch.StartNew()
        Do
            window.Dispatcher.Invoke(Sub()
                                     End Sub, DispatcherPriority.Background)
            Dim pending = window.PendingPreviewTask
            If pending.IsCompleted Then
                pending.GetAwaiter().GetResult()
                window.UpdateLayout()
                window.Dispatcher.Invoke(Sub()
                                         End Sub, DispatcherPriority.Background)
                If window.PendingPreviewTask.IsCompleted Then Return
            End If
            If timeout.Elapsed > TimeSpan.FromSeconds(20) Then Throw New TimeoutException("Preview did not finish or release its reader.")
            Thread.Sleep(1)
        Loop
    End Sub
End Module
