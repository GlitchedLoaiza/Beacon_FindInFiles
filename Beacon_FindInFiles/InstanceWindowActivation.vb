Imports System.Linq
Imports System.Windows

Namespace Beacon
    Public NotInheritable Class InstanceWindowActivation
        Private Sub New()
        End Sub

        Public Shared Sub Activate(application As System.Windows.Application)
            Dim windows = application.Windows.OfType(Of Window)().Where(Function(window) window.IsVisible).ToList()
            Dim main = application.MainWindow
            If main IsNot Nothing AndAlso main.IsVisible AndAlso main.WindowState = WindowState.Minimized Then
                SystemCommands.RestoreWindow(main)
            End If
            Dim target = windows.LastOrDefault(Function(window) window.IsActive AndAlso window.IsEnabled)
            If target Is Nothing Then target = windows.LastOrDefault(Function(window) window.Owner IsNot Nothing AndAlso window.IsEnabled)
            If target Is Nothing Then target = windows.LastOrDefault(Function(window) window.IsEnabled)
            If target Is Nothing Then Return
            If target.WindowState = WindowState.Minimized Then SystemCommands.RestoreWindow(target)
            target.Activate()
        End Sub
    End Class
End Namespace
