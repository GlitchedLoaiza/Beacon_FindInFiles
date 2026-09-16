Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Controls.Primitives
Imports System.Windows.Media

Namespace Beacon
    Public NotInheritable Class BeaconThemePalette
        Public Shared ReadOnly LogoRed As Color = Color.FromRgb(&HFF, 0, 0)
        Public Shared ReadOnly ButtonCrimson As Color = Color.FromRgb(&HB8, &H32, &H43)
        Public Shared ReadOnly ButtonText As Color = Color.FromRgb(&HFF, &HF5, &HF5)
        Private Const Marker As String = "UsesBeaconButtons"

        Private Sub New()
        End Sub

        Public Shared Function FollowsSystem(theme As AppTheme) As Boolean
            Return theme = AppTheme.System OrElse theme = AppTheme.Beacon
        End Function

        Public Shared Function UsesDarkBackground(theme As AppTheme, systemDark As Boolean) As Boolean
            Return theme = AppTheme.Dark OrElse (FollowsSystem(theme) AndAlso systemDark)
        End Function

        Public Shared Function IsEnabled(resources As ResourceDictionary) As Boolean
            Return resources.Contains(Marker) AndAlso TypeOf resources(Marker) Is Boolean AndAlso CBool(resources(Marker))
        End Function

        Public Shared Sub ApplyButtons(resources As ResourceDictionary, enabled As Boolean, Optional darkBackground As Boolean = False)
            resources(Marker) = enabled
            resources("ActionButtonDisabledOpacity") = If(enabled, 1.0, 0.5)
            resources("SettingsButtonDisabledOpacity") = If(enabled, 1.0, 0.55)
            If enabled Then
                SetBrush(resources, "AccentBrush", If(darkBackground, Color.FromRgb(&HE5, &H8E, &H9A), ButtonCrimson))
                SetBrush(resources, "SelectionBackgroundBrush", If(darkBackground, Color.FromRgb(&H48, &H2B, &H31), Color.FromRgb(&HF7, &HE6, &HE9)))
                SetBrush(resources, "InputBorderBrush", If(darkBackground, Color.FromRgb(&H7F, &H4A, &H55), Color.FromRgb(&HBA, &H7B, &H87)))
                SetBrush(resources, "ActionButtonBackgroundBrush", ButtonCrimson)
                SetBrush(resources, "ActionButtonHoverBrush", Color.FromRgb(&HC4, &H3D, &H4D))
                SetBrush(resources, "ActionButtonPressedBrush", Color.FromRgb(&H9E, &H29, &H39))
                SetBrush(resources, "ActionButtonForegroundBrush", ButtonText)
                resources("ActionButtonDisabledBrush") = resources("ButtonBackgroundBrush")
                resources("ActionButtonDisabledForegroundBrush") = resources("TextSecondaryBrush")
                SetBrush(resources, "ActionButtonBorderBrush", ButtonCrimson)
                SetBrush(resources, "ActionButtonFocusBrush", ButtonText)
                SetBrush(resources, "PrimaryButtonBrush", ButtonCrimson)
                SetBrush(resources, "PrimaryButtonHoverBrush", Color.FromRgb(&HC4, &H3D, &H4D))
                SetBrush(resources, "PrimaryButtonPressedBrush", Color.FromRgb(&H9E, &H29, &H39))
                SetBrush(resources, "PrimaryButtonForegroundBrush", ButtonText)
                resources("PrimaryButtonDisabledBrush") = resources("ButtonBackgroundBrush")
                resources("PrimaryButtonDisabledForegroundBrush") = resources("TextSecondaryBrush")
                resources("TextSelectionBrush") = resources("SelectionBackgroundBrush")
                resources("TextSelectionForegroundBrush") = resources("TextPrimaryBrush")
                resources("TextSelectionOpacity") = 1.0
            Else
                resources("ActionButtonBackgroundBrush") = resources("ButtonBackgroundBrush")
                resources("ActionButtonHoverBrush") = resources("ButtonHoverBrush")
                resources("ActionButtonPressedBrush") = resources("ButtonPressedBrush")
                resources("ActionButtonForegroundBrush") = resources("TextPrimaryBrush")
                resources("ActionButtonDisabledBrush") = resources("ButtonBackgroundBrush")
                resources("ActionButtonDisabledForegroundBrush") = resources("TextPrimaryBrush")
                resources("ActionButtonBorderBrush") = resources("CardBorderBrush")
                resources("ActionButtonFocusBrush") = resources("AccentBrush")
                SetBrush(resources, "PrimaryButtonBrush", Color.FromRgb(0, &H66, &HCC))
                SetBrush(resources, "PrimaryButtonHoverBrush", Color.FromRgb(0, &H5C, &HB8))
                SetBrush(resources, "PrimaryButtonPressedBrush", Color.FromRgb(0, &H4A, &H99))
                SetBrush(resources, "PrimaryButtonForegroundBrush", Colors.White)
                resources("PrimaryButtonDisabledBrush") = resources("PrimaryButtonBrush")
                resources("PrimaryButtonDisabledForegroundBrush") = resources("PrimaryButtonForegroundBrush")
                resources("TextSelectionBrush") = resources("AccentBrush")
                resources("TextSelectionForegroundBrush") = TextBoxBase.SelectionTextBrushProperty.GetMetadata(GetType(TextBox)).DefaultValue
                resources("TextSelectionOpacity") = TextBoxBase.SelectionOpacityProperty.GetMetadata(GetType(TextBox)).DefaultValue
            End If
        End Sub

        Public Shared Sub CopyOwnerColors(owner As Window, child As Window, Optional darkBackground As Boolean = False)
            For Each key In owner.Resources.Keys
                Dim brush = TryCast(owner.Resources(key), SolidColorBrush)
                If brush IsNot Nothing Then child.Resources(key) = brush.CloneCurrentValue()
            Next
            ApplyButtons(child.Resources, IsEnabled(owner.Resources), darkBackground)
        End Sub

        Private Shared Sub SetBrush(resources As ResourceDictionary, key As String, color As Color)
            Dim brush As New SolidColorBrush(color)
            brush.Freeze()
            resources(key) = brush
        End Sub
    End Class
End Namespace
