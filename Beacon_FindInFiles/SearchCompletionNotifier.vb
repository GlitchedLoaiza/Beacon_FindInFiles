Imports System.Diagnostics
Imports System.Runtime.InteropServices
Imports System.Windows
Imports System.Windows.Interop

Namespace Beacon
    Friend Interface ISearchCompletionNotificationPlatform
        ReadOnly Property CanNotify As Boolean
        ReadOnly Property IsApplicationForeground As Boolean
        Sub PlayCompletionSound()
        Sub FlashTaskbar()
        Sub StopFlashing()
    End Interface

    Public NotInheritable Class SearchCompletionNotifier
        Implements IDisposable

        Friend Const NotificationSoundAlias As String = "Notification.IM"
        Private Const SoundAsync As UInteger = &H1UI
        Private Const SoundNoDefault As UInteger = &H2UI
        Private Const SoundAlias As UInteger = &H10000UI
        Private Const SoundSystem As UInteger = &H200000UI
        Friend Const NotificationSoundFlags As UInteger = SoundAsync Or SoundNoDefault Or SoundAlias Or SoundSystem
        Private ReadOnly _window As Window
        Private ReadOnly _application As System.Windows.Application
        Private ReadOnly _platform As ISearchCompletionNotificationPlatform
        Private _generation As Long = Long.MinValue
        Private _completionHandled As Boolean = True
        Private _flashing As Boolean
        Private _disposed As Boolean

        Public Sub New(window As Window)
            Me.New(window, New WindowsNotificationPlatform(window))
        End Sub

        Friend Sub New(window As Window, platform As ISearchCompletionNotificationPlatform)
            ArgumentNullException.ThrowIfNull(platform)
            _window = window
            _platform = platform
            If window IsNot Nothing Then
                _application = System.Windows.Application.Current
                AddHandler window.Activated, AddressOf Acknowledge
                AddHandler window.Closed, AddressOf WindowClosed
                If _application IsNot Nothing Then AddHandler _application.Activated, AddressOf Acknowledge
            End If
        End Sub

        Public Sub BeginSearch(generation As Long)
            If _disposed OrElse generation <= _generation Then Return
            StopFlashing()
            _generation = generation
            _completionHandled = False
        End Sub

        Public Sub Complete(generation As Long, state As ScanReportState, settings As BeaconSettings)
            If _disposed OrElse generation <> _generation OrElse _completionHandled Then Return
            If state = ScanReportState.Ready OrElse state = ScanReportState.Running Then Return
            _completionHandled = True
            If state <> ScanReportState.Completed AndAlso state <> ScanReportState.ResultLimitReached Then Return
            If settings Is Nothing OrElse Not settings.NotifyOnSearchCompletion Then Return

            Dim foreground As Boolean
            Try
                If Not _platform.CanNotify Then Return
                foreground = _platform.IsApplicationForeground
            Catch ex As Exception
                Debug.WriteLine($"Search-completion notification unavailable: {ex.Message}")
                Return
            End Try
            If settings.NotifyOnlyInBackground AndAlso foreground Then Return

            TryEffect(AddressOf _platform.PlayCompletionSound)
            If Not foreground Then _flashing = TryEffect(AddressOf _platform.FlashTaskbar)
        End Sub

        Public Sub CancelPending()
            _completionHandled = True
            StopFlashing()
        End Sub

        Private Sub Acknowledge(sender As Object, e As EventArgs)
            StopFlashing()
        End Sub

        Private Sub StopFlashing()
            If Not _flashing Then Return
            _flashing = False
            TryEffect(AddressOf _platform.StopFlashing)
        End Sub

        Private Shared Function TryEffect(effect As Action) As Boolean
            Try
                effect()
                Return True
            Catch ex As Exception
                Debug.WriteLine($"Search-completion notification unavailable: {ex.Message}")
                Return False
            End Try
        End Function

        Private Sub WindowClosed(sender As Object, e As EventArgs)
            Dispose()
        End Sub

        Public Sub Dispose() Implements IDisposable.Dispose
            If _disposed Then Return
            _disposed = True
            CancelPending()
            If _window IsNot Nothing Then
                RemoveHandler _window.Activated, AddressOf Acknowledge
                RemoveHandler _window.Closed, AddressOf WindowClosed
            End If
            If _application IsNot Nothing Then RemoveHandler _application.Activated, AddressOf Acknowledge
        End Sub

        Private NotInheritable Class WindowsNotificationPlatform
            Implements ISearchCompletionNotificationPlatform

            Private Const FlashTray As UInteger = &H2UI
            Private Const FlashUntilForeground As UInteger = &HCUI
            Private ReadOnly _window As Window

            <StructLayout(LayoutKind.Sequential)>
            Private Structure FlashInfo
                Public Size As UInteger
                Public Handle As IntPtr
                Public Flags As UInteger
                Public Count As UInteger
                Public Timeout As UInteger
            End Structure

            <DllImport("winmm.dll", EntryPoint:="PlaySoundW", CharSet:=CharSet.Unicode, ExactSpelling:=True), DefaultDllImportSearchPaths(DllImportSearchPath.System32)>
            Private Shared Function PlaySound(name As String, moduleHandle As IntPtr, flags As UInteger) As Boolean
            End Function

            <DllImport("user32.dll", ExactSpelling:=True), DefaultDllImportSearchPaths(DllImportSearchPath.System32)>
            Private Shared Function FlashWindowEx(ByRef info As FlashInfo) As Boolean
            End Function

            <DllImport("user32.dll", ExactSpelling:=True), DefaultDllImportSearchPaths(DllImportSearchPath.System32)>
            Private Shared Function GetForegroundWindow() As IntPtr
            End Function

            <DllImport("user32.dll", ExactSpelling:=True), DefaultDllImportSearchPaths(DllImportSearchPath.System32)>
            Private Shared Function GetWindowThreadProcessId(handle As IntPtr, ByRef processId As UInteger) As UInteger
            End Function

            Public Sub New(window As Window)
                ArgumentNullException.ThrowIfNull(window)
                _window = window
            End Sub

            Public ReadOnly Property CanNotify As Boolean Implements ISearchCompletionNotificationPlatform.CanNotify
                Get
                    Return Not _window.Dispatcher.HasShutdownStarted AndAlso Not _window.Dispatcher.HasShutdownFinished AndAlso
                           _window.IsVisible AndAlso _window.ShowInTaskbar AndAlso New WindowInteropHelper(_window).Handle <> IntPtr.Zero
                End Get
            End Property

            Public ReadOnly Property IsApplicationForeground As Boolean Implements ISearchCompletionNotificationPlatform.IsApplicationForeground
                Get
                    Dim foreground = GetForegroundWindow()
                    If foreground = IntPtr.Zero Then Return False
                    Dim processId As UInteger
                    GetWindowThreadProcessId(foreground, processId)
                    Return processId = CUInt(Environment.ProcessId)
                End Get
            End Property

            Public Sub PlayCompletionSound() Implements ISearchCompletionNotificationPlatform.PlayCompletionSound
                PlaySound(NotificationSoundAlias, IntPtr.Zero, NotificationSoundFlags)
            End Sub

            Public Sub FlashTaskbar() Implements ISearchCompletionNotificationPlatform.FlashTaskbar
                SetFlash(FlashTray Or FlashUntilForeground)
            End Sub

            Public Sub StopFlashing() Implements ISearchCompletionNotificationPlatform.StopFlashing
                SetFlash(0UI)
            End Sub

            Private Sub SetFlash(flags As UInteger)
                Dim handle = New WindowInteropHelper(_window).Handle
                If handle = IntPtr.Zero Then Return
                Dim info As New FlashInfo With {
                    .Size = CUInt(Marshal.SizeOf(Of FlashInfo)()),
                    .Handle = handle,
                    .Flags = flags,
                    .Count = If(flags = 0UI, 0UI, UInteger.MaxValue),
                    .Timeout = 0UI
                }
                ' FlashWindowEx reports prior activation, not whether flashing succeeded.
                FlashWindowEx(info)
            End Sub
        End Class
    End Class
End Namespace
