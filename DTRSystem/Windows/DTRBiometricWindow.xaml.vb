Public Class DTRBiometricWindow
    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

    End Sub

    Private Sub btnClose_Click(sender As Object, e As RoutedEventArgs) Handles btnClose.Click
        Me.Close()
        dtrMainWindow.Show()
        dtrMainWindow.Focus()
    End Sub
End Class
