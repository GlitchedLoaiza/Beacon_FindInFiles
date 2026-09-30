Imports System.Collections.Generic
Imports System.ComponentModel
Imports System.Linq
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Input
Imports System.Windows.Media
Imports System.Windows.Media.Effects
Imports System.Windows.Shell

Namespace Beacon
    Public NotInheritable Class AeroTheme
        Private Const Marker As String = "UsesAeroTheme"
        Private Shared ReadOnly StateProperty As DependencyProperty = DependencyProperty.RegisterAttached(
            "State", GetType(ThemeState), GetType(AeroTheme), New PropertyMetadata(Nothing))

        Private Sub New()
        End Sub

        Public Shared Function IsEnabled(resources As ResourceDictionary) As Boolean
            Return resources.Contains(Marker) AndAlso TypeOf resources(Marker) Is Boolean AndAlso CBool(resources(Marker))
        End Function

        Public Shared Sub Apply(window As Window, enabled As Boolean)
            ArgumentNullException.ThrowIfNull(window)
            Dim state = TryCast(window.GetValue(StateProperty), ThemeState)
            If Not enabled Then
                state?.Dispose()
                Return
            End If

            If state Is Nothing Then
                state = New ThemeState(window)
                window.SetValue(StateProperty, state)
            End If
            Try
                state.Refresh()
            Catch
                state.Dispose()
                Throw
            End Try
        End Sub

        Friend Shared Function CreateResources(highContrast As Boolean, Optional animationsEnabled As Boolean? = Nothing) As ResourceDictionary
            Dim assemblyName = GetType(AeroTheme).Assembly.GetName().Name
            Dim resources As New ResourceDictionary With {
                .Source = New Uri("/" & assemblyName & ";component/AeroStyles.xaml", UriKind.Relative)
            }
            Dim animate = Not highContrast AndAlso animationsEnabled.GetValueOrDefault(SystemParameters.ClientAreaAnimation)
            resources("AeroProgressTemplate") = resources(If(animate, "AeroProgressAnimatedTemplate", "AeroProgressStaticTemplate"))
            If Not highContrast Then Return resources

            For Each key In resources.Keys.OfType(Of String)().ToArray()
                If Not TypeOf resources(key) Is SolidColorBrush AndAlso Not TypeOf resources(key) Is GradientBrush Then Continue For
                Select Case key
                    Case "TextSelectionBrush", "AccentBrush", "AeroProgressFillBrush"
                        resources(key) = SystemColors.HighlightBrush
                    Case "SelectionBackgroundBrush"
                        resources(key) = SystemColors.WindowBrush
                    Case "TextSelectionForegroundBrush", "CheckmarkBrush"
                        resources(key) = SystemColors.HighlightTextBrush
                    Case "AeroReflectionBrush", "AeroProgressShineBrush"
                        resources(key) = Brushes.Transparent
                    Case Else
                        If key.Contains("DisabledForeground", StringComparison.Ordinal) Then
                            resources(key) = SystemColors.GrayTextBrush
                        ElseIf key.Contains("Foreground", StringComparison.Ordinal) OrElse
                               key.StartsWith("Text", StringComparison.Ordinal) OrElse key = "AeroLabelBrush" OrElse
                               key.Contains("Border", StringComparison.Ordinal) OrElse key.Contains("Focus", StringComparison.Ordinal) OrElse
                               key = "AeroFrameBrush" OrElse key = "AeroHighlightBrush" OrElse key = "ToolbarDividerBrush" OrElse key = "SplitterBrush" Then
                            resources(key) = SystemColors.WindowTextBrush
                        Else
                            resources(key) = SystemColors.WindowBrush
                        End If
                End Select
            Next
            resources("AeroDecorationVisibility") = Visibility.Collapsed
            resources("LegacyDecorationVisibility") = Visibility.Collapsed
            resources("AeroGlowEffect") = New DropShadowEffect With {.Opacity = 0, .BlurRadius = 0, .ShadowDepth = 0}
            Return resources
        End Function

        Private NotInheritable Class ResourceOverlay
            Private ReadOnly _resources As ResourceDictionary
            Private ReadOnly _originals As New Dictionary(Of Object, Object)()

            Public Sub New(resources As ResourceDictionary)
                _resources = resources
            End Sub

            Public ReadOnly Property Resources As ResourceDictionary
                Get
                    Return _resources
                End Get
            End Property

            Public Sub Apply(palette As ResourceDictionary)
                For Each key As Object In palette.Keys
                    If Not _originals.ContainsKey(key) Then
                        _originals(key) = If(_resources.Keys.Cast(Of Object)().Contains(key), _resources(key), DependencyProperty.UnsetValue)
                    End If
                    _resources(key) = palette(key)
                Next
            End Sub

            Public Sub Restore()
                For Each item In _originals
                    If item.Value Is DependencyProperty.UnsetValue Then
                        _resources.Remove(item.Key)
                    Else
                        _resources(item.Key) = item.Value
                    End If
                Next
                _originals.Clear()
            End Sub
        End Class

        Private NotInheritable Class ThemeState
            Implements IDisposable

            Private ReadOnly _window As Window
            Private ReadOnly _overlays As New List(Of ResourceOverlay)()
            Private ReadOnly _bindings As New List(Of CommandBinding)()
            Private ReadOnly _originalTemplate As Object
            Private ReadOnly _originalChrome As Object
            Private ReadOnly _originalRounding As Object
            Private ReadOnly _originalSnapping As Object
            Private ReadOnly _ownsForeground As Boolean
            Private _palette As ResourceDictionary
            Private _chromeActive As Boolean
            Private _disposed As Boolean

            Public Sub New(window As Window)
                _window = window
                _originalTemplate = window.ReadLocalValue(Control.TemplateProperty)
                _originalChrome = window.ReadLocalValue(WindowChrome.WindowChromeProperty)
                _originalRounding = window.ReadLocalValue(FrameworkElement.UseLayoutRoundingProperty)
                _originalSnapping = window.ReadLocalValue(UIElement.SnapsToDevicePixelsProperty)
                _ownsForeground = window.ReadLocalValue(Control.ForegroundProperty) Is DependencyProperty.UnsetValue
                _overlays.Add(New ResourceOverlay(window.Resources))
                AddHandler window.Loaded, AddressOf WindowLoaded
                AddHandler window.Closed, AddressOf WindowClosed
                AddHandler SystemParameters.StaticPropertyChanged, AddressOf PreferenceChanged
            End Sub

            Public Sub Refresh()
                If _disposed Then Return
                Dim highContrast = SystemParameters.HighContrast
                _palette = CreateResources(highContrast)
                DiscoverResources(_window)
                For Each overlay In _overlays
                    overlay.Apply(_palette)
                Next
                _window.SetCurrentValue(FrameworkElement.UseLayoutRoundingProperty, True)
                _window.SetCurrentValue(UIElement.SnapsToDevicePixelsProperty, True)
                If _ownsForeground Then _window.SetResourceReference(Control.ForegroundProperty, "TextPrimaryBrush")

                If Not highContrast AndAlso _window.WindowStyle <> WindowStyle.None AndAlso Not _window.AllowsTransparency Then
                    If Not _chromeActive Then
                        WindowChrome.SetWindowChrome(_window, New WindowChrome With {
                            .CaptionHeight = 32,
                            .ResizeBorderThickness = New Thickness(If(CanResize(), 6, 0)),
                            .GlassFrameThickness = New Thickness(0),
                            .CornerRadius = New CornerRadius(3),
                            .UseAeroCaptionButtons = False
                        })
                        AddCommands()
                        _chromeActive = True
                    End If
                    _window.SetResourceReference(Control.TemplateProperty, "AeroWindowTemplate")
                Else
                    RestoreChrome()
                End If
            End Sub

            Private Sub DiscoverResources(node As DependencyObject)
                Dim element = TryCast(node, FrameworkElement)
                If element IsNot Nothing AndAlso UsesSettingsStyles(element.Resources) AndAlso
                   Not _overlays.Any(Function(overlay) overlay.Resources Is element.Resources) Then
                    _overlays.Add(New ResourceOverlay(element.Resources))
                End If
                For Each child In LogicalTreeHelper.GetChildren(node).OfType(Of DependencyObject)()
                    DiscoverResources(child)
                Next
            End Sub

            Private Shared Function UsesSettingsStyles(resources As ResourceDictionary) As Boolean
                Return (resources.Source IsNot Nothing AndAlso resources.Source.OriginalString.EndsWith("SettingsStyles.xaml", StringComparison.OrdinalIgnoreCase)) OrElse
                       resources.MergedDictionaries.Any(Function(dictionary) UsesSettingsStyles(dictionary))
            End Function

            Private Sub WindowLoaded(sender As Object, e As RoutedEventArgs)
                If _disposed Then Return
                DiscoverResources(_window)
                For Each overlay In _overlays
                    overlay.Apply(_palette)
                Next
            End Sub

            Private Sub AddCommands()
                _bindings.Add(New CommandBinding(SystemCommands.CloseWindowCommand, AddressOf CloseWindow, AddressOf CloseCanExecute))
                _bindings.Add(New CommandBinding(SystemCommands.MinimizeWindowCommand, AddressOf MinimizeWindow, AddressOf MinimizeCanExecute))
                _bindings.Add(New CommandBinding(SystemCommands.MaximizeWindowCommand, AddressOf ToggleMaximize, AddressOf ResizeCanExecute))
                _bindings.Add(New CommandBinding(SystemCommands.RestoreWindowCommand, AddressOf RestoreWindow, AddressOf ResizeCanExecute))
                For Each binding In _bindings
                    _window.CommandBindings.Add(binding)
                Next
            End Sub

            Private Sub CloseWindow(sender As Object, e As ExecutedRoutedEventArgs)
                e.Handled = True
                SystemCommands.CloseWindow(_window)
            End Sub

            Private Sub MinimizeWindow(sender As Object, e As ExecutedRoutedEventArgs)
                e.Handled = True
                SystemCommands.MinimizeWindow(_window)
            End Sub

            Private Sub ToggleMaximize(sender As Object, e As ExecutedRoutedEventArgs)
                e.Handled = True
                If _window.WindowState = WindowState.Maximized Then
                    SystemCommands.RestoreWindow(_window)
                Else
                    SystemCommands.MaximizeWindow(_window)
                End If
            End Sub

            Private Sub RestoreWindow(sender As Object, e As ExecutedRoutedEventArgs)
                e.Handled = True
                SystemCommands.RestoreWindow(_window)
            End Sub

            Private Function CanResize() As Boolean
                Return _window.ResizeMode = ResizeMode.CanResize OrElse _window.ResizeMode = ResizeMode.CanResizeWithGrip
            End Function

            Private Sub CloseCanExecute(sender As Object, e As CanExecuteRoutedEventArgs)
                e.CanExecute = Not _disposed
                e.Handled = True
            End Sub

            Private Sub MinimizeCanExecute(sender As Object, e As CanExecuteRoutedEventArgs)
                e.CanExecute = _window.ResizeMode <> ResizeMode.NoResize
                e.Handled = True
            End Sub

            Private Sub ResizeCanExecute(sender As Object, e As CanExecuteRoutedEventArgs)
                e.CanExecute = CanResize()
                e.Handled = True
            End Sub

            Private Sub PreferenceChanged(sender As Object, e As PropertyChangedEventArgs)
                If e.PropertyName <> NameOf(SystemParameters.HighContrast) AndAlso
                   e.PropertyName <> NameOf(SystemParameters.ClientAreaAnimation) AndAlso Not String.IsNullOrEmpty(e.PropertyName) Then Return
                If _disposed OrElse _window.Dispatcher.HasShutdownStarted OrElse _window.Dispatcher.HasShutdownFinished Then Return
                _window.Dispatcher.BeginInvoke(Sub()
                                                   If Not _disposed Then Refresh()
                                               End Sub)
            End Sub

            Private Sub RestoreProperty(propertyId As DependencyProperty, value As Object)
                If value Is DependencyProperty.UnsetValue Then
                    _window.ClearValue(propertyId)
                Else
                    _window.SetValue(propertyId, value)
                End If
            End Sub

            Private Sub RestoreChrome()
                If Not _chromeActive Then Return
                For Each binding In _bindings
                    _window.CommandBindings.Remove(binding)
                Next
                _bindings.Clear()
                ' WindowChrome must undo its visual-tree fixups before the custom template is removed.
                _window.ApplyTemplate()
                RestoreProperty(WindowChrome.WindowChromeProperty, _originalChrome)
                RestoreProperty(Control.TemplateProperty, _originalTemplate)
                _chromeActive = False
            End Sub

            Private Sub WindowClosed(sender As Object, e As EventArgs)
                Release(False)
            End Sub

            Public Sub Dispose() Implements IDisposable.Dispose
                Release(True)
            End Sub

            Private Sub Release(restore As Boolean)
                If _disposed Then Return
                _disposed = True
                RemoveHandler _window.Loaded, AddressOf WindowLoaded
                RemoveHandler _window.Closed, AddressOf WindowClosed
                RemoveHandler SystemParameters.StaticPropertyChanged, AddressOf PreferenceChanged
                If restore Then
                    RestoreChrome()
                    For Each overlay In _overlays
                        overlay.Restore()
                    Next
                    RestoreProperty(FrameworkElement.UseLayoutRoundingProperty, _originalRounding)
                    RestoreProperty(UIElement.SnapsToDevicePixelsProperty, _originalSnapping)
                    If _ownsForeground Then _window.ClearValue(Control.ForegroundProperty)
                Else
                    For Each binding In _bindings
                        _window.CommandBindings.Remove(binding)
                    Next
                End If
                _bindings.Clear()
                _overlays.Clear()
                _palette = Nothing
                _window.ClearValue(StateProperty)
            End Sub
        End Class
    End Class
End Namespace
