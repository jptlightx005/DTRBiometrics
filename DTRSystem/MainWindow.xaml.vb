Class MainWindow 

    Private Sub btnDTR_Click(sender As Object, e As RoutedEventArgs) Handles btnDTR.Click
        Dim dtrWindow As New DTRBiometricWindow
        dtrWindow.Show()
        Me.Hide()
    End Sub

    Private Sub btnClose_Click(sender As Object, e As RoutedEventArgs) Handles btnClose.Click
        Me.Close()
    End Sub
End Class
