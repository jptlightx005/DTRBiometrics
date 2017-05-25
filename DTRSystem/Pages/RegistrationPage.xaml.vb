Imports System.IO
Class RegistrationPage
    Dim hasFingerPrint = False
    Dim hasImage = False
    Private Sub btnEnroll_Click(sender As Object, e As RoutedEventArgs) Handles btnEnroll.Click
        Dim regFP As New RegFPWindow
        regFP.regPage = Me
        regFP.ShowDialog()

    End Sub

    Private Sub btnRegister_Click(sender As Object, e As RoutedEventArgs) Handles btnRegister.Click
        If ValidateFields() Then
            Dim a As New DTRBiometricDataSetTableAdapters.tbl_employeeTableAdapter
            Try
                a.Insert("", txtUsrn.Text, txtPssw.Password, File.ReadAllBytes(applicationPath & "\fptemp.tpl"), txtFName.Text, txtMName.Text, txtLName.Text, 0, txtDesignation.Text, txtCorporation.Text, "8:00 A.M.", "5:00 P.M.", File.ReadAllBytes(applicationPath & "\fptemp.tpl"), 25)
                MsgBox("Successfully registered!", vbInformation)
            Catch ex As Exception
                MsgBox("Registration failed!", vbInformation)
            End Try
        Else
            MsgBox("Registration Failed! One or more required fields is empty!", vbExclamation)
        End If
    End Sub

    Public Sub FingerprintEnrolled(success As Boolean)
        hasFingerPrint = success
        If success Then
            lblStatus.Content = "Fingerprint has been enrolled"
        Else
            lblStatus.Content = "Fingerprint enrollment failed"
        End If
    End Sub
    Function ValidateFields() As Boolean
        Dim isValid As Boolean = True
        Dim i As Integer = 0

        isValid = isValid And txtUsrn.Text.Length > 0
        isValid = isValid And txtPssw.Password.Length > 6
        isValid = isValid And txtFName.Text.Length > 0
        isValid = isValid And txtMName.Text.Length > 0
        isValid = isValid And txtLName.Text.Length > 0
        isValid = isValid And hasFingerPrint
        Return isValid
    End Function
End Class
