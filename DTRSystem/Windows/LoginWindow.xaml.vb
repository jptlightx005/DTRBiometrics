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
                If dtrMainWindow Is Nothing Then
                    dtrMainWindow = New MainWindow
                End If
                dtrMainWindow.Show()
                Me.Close()
            End If
        Else
            If MsgBox("Invalid credentials!", vbExclamation) = MsgBoxResult.Ok Then
                If psswBox.IsFocused Then
                    psswBox.SelectAll()
                Else
                    psswBox.Focus()
                End If
            End If
        End If

    End Sub

    Private Sub psswBox_KeyDown(sender As Object, e As KeyEventArgs) Handles psswBox.KeyDown, txtUsername.KeyDown
        If e.Key = Key.Return Then
            Button_Click(sender, e)
        End If
    End Sub

    Private Sub psswBox_GotFocus(sender As Object, e As RoutedEventArgs) Handles psswBox.GotFocus
        psswBox.SelectAll()
    End Sub

    Private Sub txtUsername_GotFocus(sender As Object, e As RoutedEventArgs) Handles txtUsername.GotFocus
        txtUsername.SelectAll()
    End Sub

    Private Sub Window_Loaded(sender As Object, e As RoutedEventArgs)
        txtUsername.Focus()
    End Sub
End Class
