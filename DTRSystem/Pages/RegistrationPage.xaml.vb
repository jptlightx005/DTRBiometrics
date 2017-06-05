﻿Imports Microsoft.Win32
Imports System.Windows.Media.Imaging
Imports System.IO
Class RegistrationPage
    Dim hasFingerPrint = False
    Dim imageFileName As String
    Private Sub btnEnroll_Click(sender As Object, e As RoutedEventArgs) Handles btnEnroll.Click
        Dim regFP As New RegFPWindow
        regFP.regPage = Me
        regFP.ShowDialog()

    End Sub

    Private Sub btnRegister_Click(sender As Object, e As RoutedEventArgs) Handles btnRegister.Click
        If ValidateFields() Then
            Dim tblEmployee As New DTRBiometricDataSetTableAdapters.EmployeeTableAdapter
            Try
                Dim image As Byte()
                If imageFileName <> "" Then
                    image = File.ReadAllBytes(imageFileName)
                Else
                    Dim b As New System.Text.ASCIIEncoding
                    image = b.GetBytes(0)
                End If

                tblEmployee.Insert("",
                                   txtUsrn.Text,
                                   txtPssw.Password,
                                   File.ReadAllBytes(applicationPath & "\fptemp.tpl"),
                                   txtFName.Text,
                                   txtMName.Text,
                                   txtLName.Text,
                                   0,
                                   txtDesignation.Text,
                                   txtCorporation.Text,
                                   "8:00 A.M.",
                                   "5:00 P.M.",
                                   image, 25)

                File.Delete(applicationPath & "\fptemp.tpl")
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

    Private Sub btnBrowse_Click(sender As Object, e As RoutedEventArgs) Handles btnBrowse.Click
        Dim fileDialog As New OpenFileDialog
        fileDialog.InitialDirectory = My.Computer.FileSystem.SpecialDirectories.MyDocuments
        fileDialog.Filter = "Image files (*.png;*.jpg)|*.png;*.jpg|All files (*.*)|*.*"
        If fileDialog.ShowDialog Then
            imgEmpPicture.Source = New BitmapImage(New Uri(fileDialog.FileName))
            imageFileName = fileDialog.FileName
        End If
    End Sub
End Class
