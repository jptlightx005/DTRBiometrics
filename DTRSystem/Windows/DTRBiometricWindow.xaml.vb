Imports ZKFPEngXControl
Public Class DTRBiometricWindow
    Dim fp As ZKFPEngX
    Public otemplate As Object
    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        fp = New ZKFPEngX
        
    End Sub

    Private Sub Window_Loaded(sender As Object, e As RoutedEventArgs)
        fp.SensorIndex = 0
        If (fp.InitEngine = 0) Then
            imgConnected.Visibility = Windows.Visibility.Visible
            lblConnected.Content = "Device Connected"
        Else
            imgConnected.Visibility = Windows.Visibility.Hidden
            lblConnected.Content = "Device Not Connected"
        End If
    End Sub

    Private Sub btnClose_Click(sender As Object, e As RoutedEventArgs) Handles btnClose.Click
        Me.Close()
        dtrMainWindow.Show()
        dtrMainWindow.Focus()
    End Sub

End Class
