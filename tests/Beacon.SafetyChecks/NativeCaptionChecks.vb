Imports System.Runtime.InteropServices
Imports System.Windows
Imports System.Windows.Interop
Imports Beacon

Module NativeCaptionChecks
    <DllImport("dwmapi.dll", PreserveSig:=True)>
    Private Function DwmGetWindowAttribute(handle As IntPtr, attribute As Integer, ByRef value As Integer, size As Integer) As Integer
    End Function

    <DllImport("dwmapi.dll", PreserveSig:=True)>
    Private Function DwmSetWindowAttribute(handle As IntPtr, attribute As Integer, ByRef value As Integer, size As Integer) As Integer
    End Function

    <DllImport("user32.dll", EntryPoint:="GetWindowLongW")>
    Private Function GetWindowLong(handle As IntPtr, index As Integer) As Integer
    End Function

    Public Sub Verify(window As Window, dark As Boolean)
        Require(window.WindowStyle <> WindowStyle.None AndAlso Not window.AllowsTransparency, "Native frame was replaced by borderless chrome.")
        Dim handle = New WindowInteropHelper(window).Handle
        Require(handle <> IntPtr.Zero, "Caption test requires a native window.")
        Dim style = GetWindowLong(handle, -16)
        Require((style And &HC00000) = &HC00000 AndAlso (style And &H80000) <> 0, "Native caption or system menu is missing.")
        If Not OperatingSystem.IsWindowsVersionAtLeast(10, 0, 22000) Then Return
        Dim highContrast = SystemParameters.HighContrast
        For Each item In {(20, If(dark AndAlso Not highContrast, 1, 0)),
                          (35, If(highContrast, -1, If(dark, &H202020, &HF3F3F3))),
                          (36, If(highContrast, -1, If(dark, &HE0E0E0, &H202020)))}
            Dim actual As Integer
            If item.Item1 = 20 Then
                Dim result = DwmGetWindowAttribute(handle, item.Item1, actual, Marshal.SizeOf(Of Integer)())
                Require(result >= 0 AndAlso actual = item.Item2, $"Native dark caption does not match the theme: HRESULT={result:X8}, value={actual:X8}.")
            Else
                actual = item.Item2
                Dim result = DwmSetWindowAttribute(handle, item.Item1, actual, Marshal.SizeOf(Of Integer)())
                Require(result >= 0, $"Windows rejected caption color attribute {item.Item1}: HRESULT={result:X8}.")
            End If
        Next
    End Sub

    Public Sub Transitions()
        Dim window As New SettingsWindow(New BeaconSettings(), False) With {
            .WindowStartupLocation = WindowStartupLocation.Manual, .Left = -10000, .Top = -10000,
            .ShowActivated = False, .ShowInTaskbar = False
        }
        Try
            Dim title = window.Title
            Dim icon = window.Icon
            Dim resize = window.ResizeMode
            window.Show()
            Dim handle = New WindowInteropHelper(window).Handle
            Dim originalStyle = GetWindowLong(handle, -16)
            For Each dark In {False, True, False, True}
                NativeCaptionTheme.Apply(window, dark)
                Verify(window, dark)
                Require(GetWindowLong(handle, -16) = originalStyle, "Theme application changed native window controls.")
                Require(window.Title = title AndAlso window.Icon Is icon AndAlso window.ResizeMode = resize, "Theme application changed the title, logo or resize behavior.")
            Next
        Finally
            window.Close()
        End Try
    End Sub

    Private Sub Require(condition As Boolean, message As String)
        If Not condition Then Throw New InvalidOperationException(message)
    End Sub
End Module
