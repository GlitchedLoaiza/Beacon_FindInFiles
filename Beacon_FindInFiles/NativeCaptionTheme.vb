Imports System.ComponentModel
Imports System.Diagnostics
Imports System.Runtime.InteropServices
Imports System.Windows
Imports System.Windows.Interop

Namespace Beacon
    Public NotInheritable Class NativeCaptionTheme
        Private Const ImmersiveDarkMode As Integer = 20
        Private Const CaptionColor As Integer = 35
        Private Const TextColor As Integer = 36
        Private Shared ReadOnly StateProperty As DependencyProperty = DependencyProperty.RegisterAttached(
            "State", GetType(NativeCaptionTheme), GetType(NativeCaptionTheme), New PropertyMetadata(Nothing))
        Private ReadOnly _window As Window
        Private _dark As Boolean
        Private _closed As Boolean

        <DllImport("dwmapi.dll", PreserveSig:=True)>
        Private Shared Function DwmSetWindowAttribute(handle As IntPtr, attribute As Integer, ByRef value As Integer, size As Integer) As Integer
        End Function

        Private Sub New(window As Window)
            _window = window
            AddHandler window.SourceInitialized, AddressOf SourceReady
            AddHandler window.Closed, AddressOf WindowClosed
            AddHandler SystemParameters.StaticPropertyChanged, AddressOf SystemPreferenceChanged
        End Sub

        Public Shared Sub Apply(window As Window, dark As Boolean)
            Dim state = TryCast(window.GetValue(StateProperty), NativeCaptionTheme)
            If state Is Nothing Then
                state = New NativeCaptionTheme(window)
                window.SetValue(StateProperty, state)
            End If
            state._dark = dark
            state.ApplyColors()
        End Sub

        Private Sub SourceReady(sender As Object, e As EventArgs)
            ApplyColors()
        End Sub

        Private Sub ApplyColors()
            If _closed OrElse Not OperatingSystem.IsWindowsVersionAtLeast(10, 0, 22000) Then Return
            Dim handle = New WindowInteropHelper(_window).Handle
            If handle = IntPtr.Zero Then Return
            Dim highContrast = SystemParameters.HighContrast
            Try
                SetAttribute(handle, ImmersiveDarkMode, If(_dark AndAlso Not highContrast, 1, 0))
                SetAttribute(handle, CaptionColor, If(highContrast, -1, If(_dark, &H202020, &HF3F3F3)))
                SetAttribute(handle, TextColor, If(highContrast, -1, If(_dark, &HE0E0E0, &H202020)))
            Catch ex As Exception When TypeOf ex Is DllNotFoundException OrElse TypeOf ex Is EntryPointNotFoundException
                Debug.WriteLine($"Native caption theming unavailable: {ex.Message}")
            End Try
        End Sub

        Private Shared Sub SetAttribute(handle As IntPtr, attribute As Integer, value As Integer)
            Dim result = DwmSetWindowAttribute(handle, attribute, value, Marshal.SizeOf(Of Integer)())
            If result < 0 Then Debug.WriteLine($"Native caption attribute {attribute} unavailable: 0x{result:X8}")
        End Sub

        Private Sub SystemPreferenceChanged(sender As Object, e As PropertyChangedEventArgs)
            If e.PropertyName <> NameOf(SystemParameters.HighContrast) AndAlso Not String.IsNullOrEmpty(e.PropertyName) Then Return
            If _window.Dispatcher.HasShutdownStarted OrElse _window.Dispatcher.HasShutdownFinished Then Return
            _window.Dispatcher.BeginInvoke(Sub() ApplyColors())
        End Sub

        Private Sub WindowClosed(sender As Object, e As EventArgs)
            _closed = True
            RemoveHandler SystemParameters.StaticPropertyChanged, AddressOf SystemPreferenceChanged
            RemoveHandler _window.SourceInitialized, AddressOf SourceReady
            RemoveHandler _window.Closed, AddressOf WindowClosed
        End Sub
    End Class
End Namespace
