Imports Microsoft.Win32
Imports System.Windows.Media.Imaging
Imports System.IO
Imports DTRSystem.DTRDataSetTableAdapters
Imports DTRSystem.DTRDataSet
Class RegistrationPage
    Dim hasFingerPrint = False
    Dim imageFileName As String

    Dim dataTable As New EmployeeTableDataTable
    Dim employeeRow As EmployeeTableRow

    'Custom Functions
    Sub EmptyBoxes()
        hasFingerPrint = False
        imgEmpPicture.Source = Nothing
        txtFName.Text = ""
        txtMName.Text = ""
        txtLName.Text = ""
        lblStatus.Content = "No Fingerprint Enrolled"
    End Sub

    'Page Events
    Private Sub regPage_Initialized(sender As Object, e As EventArgs) Handles regPage.Initialized
        employeeRow = dataTable.NewRow
        employeeRow.first_name = ""
        employeeRow.middle_name = ""
        employeeRow.last_name = ""
        employeeRow.work_timeb = 0
        employeeRow.work_timee = 0
        grdEmployee.DataContext = employeeRow

        'cmbDepartment.ItemsSource = tblDepartmentAdapter.GetData
    End Sub
    Private Sub btnEnroll_Click(sender As Object, e As RoutedEventArgs) Handles btnEnroll.Click
        Dim regFP As New RegFPWindow
        regFP.regPage = Me
        regFP.ShowDialog()

    End Sub

    Private Sub btnRegister_Click(sender As Object, e As RoutedEventArgs) Handles btnRegister.Click
        If ValidateFields() Then
            Try
                Dim image As Byte()
                If imageFileName <> "" Then
                    image = File.ReadAllBytes(imageFileName)
                Else
                    Dim b As New System.Text.ASCIIEncoding
                    image = b.GetBytes(0)
                End If

                employeeRow.biometric = File.ReadAllBytes(applicationPath & "\fptemp.tpl")
                File.Delete(applicationPath & "\fptemp.tpl")
                employeeRow.work_timeb = 8
                employeeRow.work_timee = 17
                employeeRow.picture = image

                dataTable.Rows.Add(employeeRow)

                If tblEmployeeAdapter.Update(dataTable) = 1 Then
                    MsgBox("Successfully registered!", vbInformation)
                    EmptyBoxes()
                Else
                    MsgBox("Registration failed!", vbInformation)
                End If
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

        isValid = isValid And txtFName.Text.Length > 0
        isValid = isValid And txtMName.Text.Length > 0
        isValid = isValid And txtLName.Text.Length > 0

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
