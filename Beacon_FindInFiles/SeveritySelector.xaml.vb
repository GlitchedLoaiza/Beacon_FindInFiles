Imports System.Globalization
Imports System.Windows
Imports System.Windows.Controls

Namespace BeaconFilterControls
    Partial Public Class SeveritySelector
        Public Shared ReadOnly TextProperty As DependencyProperty = DependencyProperty.Register(NameOf(Text), GetType(String), GetType(SeveritySelector),
            New FrameworkPropertyMetadata("", FrameworkPropertyMetadataOptions.BindsTwoWayByDefault, AddressOf SelectionChanged))
        Private ReadOnly _checks As New List(Of CheckBox)()
        Private ReadOnly _labels As String() = {"Always", "Critical", "Error", "Warning", "Information", "Verbose"}
        Private _updating As Boolean

        Public Property Text As String
            Get
                Return CStr(GetValue(TextProperty))
            End Get
            Set(value As String)
                SetValue(TextProperty, If(value, ""))
            End Set
        End Property

        Public Sub New()
            InitializeComponent()
            For index = 0 To 5
                Dim check As New CheckBox With {.Content = $"{_labels(index)} ({index})"}
                AddHandler check.Checked, AddressOf CheckChanged
                AddHandler check.Unchecked, AddressOf CheckChanged
                _checks.Add(check)
                LevelChoices.Children.Add(check)
            Next
            RefreshSelection()
        End Sub

        Private Shared Sub SelectionChanged(sender As DependencyObject, e As DependencyPropertyChangedEventArgs)
            DirectCast(sender, SeveritySelector).RefreshSelection()
        End Sub

        Private Sub RefreshSelection()
            If OpenLevels Is Nothing OrElse _checks.Count <> 6 Then Return
            _updating = True
            Try
                Dim selected As New HashSet(Of String)(If(Text, "").Split(","c).Select(Function(value) value.Trim()))
                For index = 0 To 5
                    _checks(index).IsChecked = selected.Contains(index.ToString(CultureInfo.InvariantCulture))
                Next
                Dim labels = _checks.Where(Function(check) check.IsChecked.GetValueOrDefault()).Select(Function(check) _labels(_checks.IndexOf(check))).ToArray()
                Dim valid = String.IsNullOrWhiteSpace(Text) OrElse selected.All(Function(value) Enumerable.Range(0, 6).Any(Function(level) value = level.ToString(CultureInfo.InvariantCulture)))
                OpenLevels.Content = If(Not valid, "Invalid saved levels ▾", If(labels.Length = 0, "All levels ▾", If(labels.Length = 1, labels(0) & " ▾", $"{labels.Length} levels selected ▾")))
                OpenLevels.ToolTip = If(labels.Length = 0, "All severity levels", String.Join(", ", labels))
            Finally
                _updating = False
            End Try
        End Sub

        Private Sub CheckChanged(sender As Object, e As RoutedEventArgs)
            If _updating Then Return
            Text = String.Join(",", Enumerable.Range(0, 6).Where(Function(index) _checks(index).IsChecked.GetValueOrDefault()))
        End Sub
        Public Sub Clear()
            Text = ""
        End Sub
        Private Sub OpenLevels_Click(sender As Object, e As RoutedEventArgs)
            LevelsPopup.IsOpen = Not LevelsPopup.IsOpen
        End Sub
        Private Sub AllLevels_Click(sender As Object, e As RoutedEventArgs)
            Clear()
        End Sub
    End Class
End Namespace
