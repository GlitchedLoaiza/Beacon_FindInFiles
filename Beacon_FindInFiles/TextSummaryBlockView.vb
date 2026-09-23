Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Documents
Imports System.Windows.Media

Namespace Beacon
    Public NotInheritable Class TextSummaryBlockView
        Inherits ContentControl
        Private _renderVersion As Integer
        Public Shared ReadOnly BlockProperty As DependencyProperty = DependencyProperty.Register(
            NameOf(Block), GetType(Object), GetType(TextSummaryBlockView), New FrameworkPropertyMetadata(Nothing, AddressOf Rebuild))
        Public Shared ReadOnly ActiveOccurrenceProperty As DependencyProperty = DependencyProperty.Register(
            NameOf(ActiveOccurrence), GetType(Integer), GetType(TextSummaryBlockView), New FrameworkPropertyMetadata(-1, AddressOf Rebuild))
        Public Shared ReadOnly IsCurrentProperty As DependencyProperty = DependencyProperty.Register(
            NameOf(IsCurrent), GetType(Boolean), GetType(TextSummaryBlockView), New FrameworkPropertyMetadata(False, AddressOf Rebuild))
        Public Shared ReadOnly WrapLinesProperty As DependencyProperty = DependencyProperty.Register(
            NameOf(WrapLines), GetType(Boolean), GetType(TextSummaryBlockView), New FrameworkPropertyMetadata(False, AddressOf Rebuild))

        Public Property Block As Object
            Get
                Return GetValue(BlockProperty)
            End Get
            Set(value As Object)
                SetValue(BlockProperty, value)
            End Set
        End Property
        Public Property ActiveOccurrence As Integer
            Get
                Return CInt(GetValue(ActiveOccurrenceProperty))
            End Get
            Set(value As Integer)
                SetValue(ActiveOccurrenceProperty, value)
            End Set
        End Property
        Public Property IsCurrent As Boolean
            Get
                Return CBool(GetValue(IsCurrentProperty))
            End Get
            Set(value As Boolean)
                SetValue(IsCurrentProperty, value)
            End Set
        End Property
        Public Property WrapLines As Boolean
            Get
                Return CBool(GetValue(WrapLinesProperty))
            End Get
            Set(value As Boolean)
                SetValue(WrapLinesProperty, value)
            End Set
        End Property

        Public Sub New()
            Focusable = False
            IsTabStop = False
            HorizontalContentAlignment = HorizontalAlignment.Stretch
        End Sub

        Private Shared Sub Rebuild(owner As DependencyObject, e As DependencyPropertyChangedEventArgs)
            DirectCast(owner, TextSummaryBlockView).Render()
        End Sub

        Private Sub Render()
            _renderVersion += 1
            Dim version = _renderVersion
            Dim activeRun As Run = Nothing
            Dim block = TryCast(Me.Block, TextSummaryBlock)
            If block Is Nothing Then
                Content = Nothing
                Return
            End If
            Dim stack As New StackPanel()
            stack.Children.Add(Label(block.Location, False, True))
            If block.IsMetadata Then
                stack.Children.Add(Label("Name/path match, not file content.", True))
                stack.Children.Add(Label(block.Excerpt, False))
            ElseIf block.Lines.Count = 0 Then
                stack.Children.Add(Label("Full line context was not captured; showing the saved excerpt.", True))
                stack.Children.Add(Label(block.Excerpt, False))
            Else
                stack.Children.Add(Label(If(block.ContextKind, "Captured source lines") & " · up to five captured lines" &
                                         If(block.IsShortened, " · long lines shortened", ""), True))
                Dim rows As New Grid With {.Margin = New Thickness(0, 6, 0, 0)}
                rows.ColumnDefinitions.Add(New ColumnDefinition With {.Width = GridLength.Auto})
                rows.ColumnDefinitions.Add(New ColumnDefinition With {.Width = New GridLength(1, GridUnitType.Star)})
                For row = 0 To block.Lines.Count - 1
                    rows.RowDefinitions.Add(New RowDefinition With {.Height = GridLength.Auto})
                    Dim line = block.Lines(row)
                    Dim number = Label(line.LineNumber.ToString("N0"), True)
                    number.Margin = New Thickness(0, 2, 12, 2)
                    number.HorizontalAlignment = HorizontalAlignment.Right
                    Grid.SetRow(number, row)
                    rows.Children.Add(number)
                    Dim body As New TextBlock With {.TextWrapping = If(WrapLines, TextWrapping.Wrap, TextWrapping.NoWrap), .Padding = New Thickness(5, 2, 5, 2)}
                    body.SetResourceReference(TextBlock.ForegroundProperty, "TextPrimaryBrush")
                    If line.IsMatchLine Then body.SetResourceReference(TextBlock.BackgroundProperty, "SelectionBackgroundBrush")
                    Dim text = If(line.Text, "")
                    Dim position As Integer = 0
                    Dim current = If(line Is block.Focus, block.CurrentSpan, Nothing)
                    For Each span In line.Highlights.OrderBy(Function(item) item.Start)
                        If span.Start < position OrElse span.Start < 0 OrElse span.Length <= 0 OrElse span.Start >= text.Length OrElse span.Length > text.Length - span.Start Then Continue For
                        If span.Start > position Then body.Inlines.Add(New Run(text.Substring(position, span.Start - position)))
                        Dim isActive = current IsNot Nothing AndAlso current.Start = span.Start AndAlso current.Length = span.Length
                        Dim highlight As New Run(text.Substring(span.Start, span.Length)) With {
                            .Background = If(isActive, Brushes.Orange, Brushes.Yellow), .Foreground = Brushes.Black, .FontWeight = FontWeights.SemiBold}
                        body.Inlines.Add(highlight)
                        If isActive Then activeRun = highlight
                        position = span.Start + span.Length
                    Next
                    If position < text.Length Then body.Inlines.Add(New Run(text.Substring(position)))
                    If text.Length = 0 Then body.Inlines.Add(New Run(" "))
                    Grid.SetRow(body, row)
                    Grid.SetColumn(body, 1)
                    rows.Children.Add(body)
                Next
                stack.Children.Add(rows)
            End If
            Dim border As New Border With {.CornerRadius = New CornerRadius(4), .Padding = New Thickness(10),
                .BorderThickness = New Thickness(1), .Child = stack}
            border.SetResourceReference(Border.BackgroundProperty, "CardBackgroundBrush")
            border.SetResourceReference(Border.BorderBrushProperty, If(block.IsCurrent, "AccentBrush", "CardBorderBrush"))
            Content = border
            If activeRun IsNot Nothing Then
                Dispatcher.BeginInvoke(Sub()
                                           If IsLoaded AndAlso version = _renderVersion AndAlso block.IsCurrent Then activeRun.BringIntoView()
                                       End Sub, Threading.DispatcherPriority.Loaded)
            End If
        End Sub

        Private Shared Function Label(text As String, secondary As Boolean, Optional heading As Boolean = False) As TextBlock
            Dim result As New TextBlock With {.Text = If(text, ""), .TextWrapping = TextWrapping.Wrap, .Margin = New Thickness(0, 0, 0, 3)}
            result.SetResourceReference(TextBlock.ForegroundProperty, If(secondary, "TextSecondaryBrush", "TextPrimaryBrush"))
            If heading Then result.FontWeight = FontWeights.SemiBold
            Return result
        End Function
    End Class
End Namespace
