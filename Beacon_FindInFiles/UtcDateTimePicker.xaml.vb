Imports System.Globalization
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Input

Namespace BeaconFilterControls
    Partial Public Class UtcDateTimePicker
        Public Shared ReadOnly TextProperty As DependencyProperty = DependencyProperty.Register(NameOf(Text), GetType(String), GetType(UtcDateTimePicker),
            New FrameworkPropertyMetadata("", FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))
        Private ReadOnly _fields As New List(Of TextBox)()

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
            For Each label In {"Year", "Month", "Day", "Hour", "Minute", "Second"}
                Dim index = _fields.Count
                Dim panel As New StackPanel With {.Margin = New Thickness(2, 0, 2, 8)}
                panel.Children.Add(New TextBlock With {.Text = label, .HorizontalAlignment = HorizontalAlignment.Center})
                Dim up As New Button With {.Content = "▴", .Style = DirectCast(FindResource("ArrowButton"), Style)}
                Automation.AutomationProperties.SetName(up, "Increase " & label)
                AddHandler up.Click, Sub() StepComponent(index, 1)
                panel.Children.Add(up)
                Dim field As New TextBox With {.Margin = New Thickness(2), .Padding = New Thickness(3, 5, 3, 5), .TextAlignment = TextAlignment.Center, .MinWidth = 0}
                Automation.AutomationProperties.SetName(field, label & " UTC")
                AddHandler field.PreviewKeyDown, Sub(sender, e)
                                                    If e.Key = Key.Up OrElse e.Key = Key.Down Then
                                                        StepComponent(index, If(e.Key = Key.Up, 1, -1))
                                                        e.Handled = True
                                                    End If
                                                End Sub
                _fields.Add(field)
                panel.Children.Add(field)
                Dim down As New Button With {.Content = "▾", .Style = DirectCast(FindResource("ArrowButton"), Style)}
                Automation.AutomationProperties.SetName(down, "Decrease " & label)
                AddHandler down.Click, Sub() StepComponent(index, -1)
                panel.Children.Add(down)
                Components.Children.Add(panel)
            Next
        End Sub

        Friend Sub OpenEditor()
            Dim value As DateTime
            PickerError.Text = ""
            If String.IsNullOrWhiteSpace(Text) Then
                value = DateTime.UtcNow
            ElseIf Not DateTime.TryParseExact(Text.Trim(), {"yyyy-MM-dd HH:mm:ss", "yyyy-MM-ddTHH:mm:ss'Z'"}, CultureInfo.InvariantCulture,
                                             DateTimeStyles.AssumeUniversal Or DateTimeStyles.AdjustToUniversal, value) Then
                PickerError.Text = "Invalid typed date. Correct the fields below or cancel to keep your text."
                For Each field In _fields
                    field.Clear()
                Next
                PickerPopup.IsOpen = True
                Return
            End If
            Fill(value)
            PickerPopup.IsOpen = True
        End Sub

        Private Sub Fill(value As DateTime)
            Dim values = {value.Year, value.Month, value.Day, value.Hour, value.Minute, value.Second}
            For index = 0 To values.Length - 1
                _fields(index).Text = values(index).ToString(If(index = 0, "D4", "D2"), CultureInfo.InvariantCulture)
            Next
        End Sub

        Private Function ReadComponents() As DateTime
            Dim values(5) As Integer
            For index = 0 To 5
                If Not Integer.TryParse(_fields(index).Text, NumberStyles.None, CultureInfo.InvariantCulture, values(index)) Then
                    Throw New ArgumentException("Enter valid numbers in every date/time field.")
                End If
            Next
            Return New DateTime(values(0), values(1), values(2), values(3), values(4), values(5), DateTimeKind.Utc)
        End Function

        Friend Shared Function Adjust(value As DateTime, component As Integer, amount As Integer) As DateTime
            Select Case component
                Case 0 : Return value.AddYears(amount)
                Case 1 : Return value.AddMonths(amount)
                Case 2 : Return value.AddDays(amount)
                Case 3 : Return value.AddHours(amount)
                Case 4 : Return value.AddMinutes(amount)
                Case 5 : Return value.AddSeconds(amount)
                Case Else : Throw New ArgumentOutOfRangeException(NameOf(component))
            End Select
        End Function

        Friend Sub StepComponent(component As Integer, amount As Integer)
            Try
                Fill(Adjust(ReadComponents(), component, amount))
                PickerError.Text = ""
            Catch ex As ArgumentException
                PickerError.Text = "Enter a valid date/time (years 0001–9999). " & ex.Message.Split(ControlChars.Cr, ControlChars.Lf)(0)
            End Try
        End Sub

        Friend Sub CommitEditor()
            Try
                Text = ReadComponents().ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture)
                PickerError.Text = ""
                PickerPopup.IsOpen = False
            Catch ex As ArgumentException
                PickerError.Text = "Enter a valid date/time; check the day, month and time values."
            End Try
        End Sub

        Public Sub Clear()
            Text = ""
            PickerPopup.IsOpen = False
        End Sub

        Private Sub OpenPicker_Click(sender As Object, e As RoutedEventArgs)
            If PickerPopup.IsOpen Then
                PickerPopup.IsOpen = False
            Else
                OpenEditor()
            End If
        End Sub
        Private Sub Apply_Click(sender As Object, e As RoutedEventArgs)
            CommitEditor()
        End Sub
        Private Sub Cancel_Click(sender As Object, e As RoutedEventArgs)
            PickerPopup.IsOpen = False
        End Sub
        Private Sub Clear_Click(sender As Object, e As RoutedEventArgs)
            Clear()
        End Sub
        Private Sub Popup_KeyDown(sender As Object, e As KeyEventArgs)
            If e.Key = Key.Escape Then
                PickerPopup.IsOpen = False
                e.Handled = True
            ElseIf e.Key = Key.Enter Then
                CommitEditor()
                e.Handled = True
            End If
        End Sub
    End Class
End Namespace
