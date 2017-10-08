Public Class LoginWindow

    Private Sub Button_Click(sender As Object, e As RoutedEventArgs)
        If Trim(txtUsername.Text).Count = 0 Or Trim(psswBox.Password).Count = 0 Then
            MsgBox("Please enter username and password!", vbExclamation)
            Return
        End If

        Dim usrn = txtUsername.Text
        Dim encr = HashPassword(psswBox.Password)

        If ValidateCredentials(usrn, encr) Then
            If MsgBox("Successfully logged in!", vbInformation) = vbOK Then
                dtrMainWindow = New MainWindow
                dtrMainWindow.Show()
                Me.Close()
            End If
        Else
            MsgBox("Invalid credentials!", vbExclamation)
        End If

    End Sub

    Private Sub psswBox_KeyDown(sender As Object, e As KeyEventArgs) Handles psswBox.KeyDown, txtUsername.KeyDown
        If e.Key = Key.Return Then
            Button_Click(sender, e)
        End If
    End Sub
End Class
