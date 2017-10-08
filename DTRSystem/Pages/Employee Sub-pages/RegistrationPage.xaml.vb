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
        imgEmpPicture.Source = New BitmapImage(New Uri("pack://siteoforigin:,,,/Resources/placeholder.png", UriKind.Absolute))
        txtFName.Text = ""
        txtMName.Text = ""
        txtLName.Text = ""
        cmbSalaryGrade.SelectedIndex = -1
        cmbStepGrade.SelectedIndex = -1
        lblStatus.Content = "No Fingerprint Enrolled"

        regPage_Initialized(Me, New EventArgs)
    End Sub

    'Page Events
    Private Sub regPage_Initialized(sender As Object, e As EventArgs) Handles regPage.Initialized
        employeeRow = dataTable.NewRow
        employeeRow.first_name = ""
        employeeRow.middle_name = ""
        employeeRow.last_name = ""
        employeeRow.emp_type = ""
        employeeRow.deptID = 0
        employeeRow.desgID = 0
        employeeRow.age = 0
        employeeRow.date_of_birth = Now
        employeeRow.address = ""
        employeeRow.work_timeb = 0
        employeeRow.work_timee = 0
        employeeRow.salarygrade = 0
        employeeRow.stepgrade = 0
        grdEmployee.DataContext = employeeRow

        cmbDepartment.ItemsSource = tblDeptAdapter.GetData
    End Sub
    Private Sub btnEnroll_Click(sender As Object, e As RoutedEventArgs) Handles btnEnroll.Click
        Dim regFP As New RegFPWindow
        regFP.regPage = Me
        isRegisteringFingerprint = True
        If isRegisteringFingerprint Then
            Debug.Print("Should be disabled")
        Else
            Debug.Print("dfq is going on")
        End If
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
                employeeRow.work_timeb = 8
                employeeRow.work_timee = 17
                employeeRow.picture = image
                employeeRow.salarygrade = cmbSalaryGrade.SelectedIndex + 1
                employeeRow.stepgrade = cmbStepGrade.SelectedIndex + 1
                dataTable.Rows.Add(employeeRow)
                If tblEmployeeAdapter.Update(dataTable) = 1 Then
                    File.Delete(applicationPath & "\fptemp.tpl")
                    MsgBox("Successfully registered!", vbInformation)
                    EmptyBoxes()
                Else
                    MsgBox("Registration failed! Error 2", vbInformation)
                End If
            Catch ex As Exception
                MsgBox("Registration failed! Error 1", vbInformation)
            End Try
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
        If txtFName.Text.Length = 0 Then
            MsgBox("Must fill up First Name!", vbExclamation)
            Return False
        End If

        If txtMName.Text.Length = 0 Then
            MsgBox("Must fill up Middle Name!", vbExclamation)
            Return False
        End If

        If txtLName.Text.Length = 0 Then
            MsgBox("Must fill up Last Name!", vbExclamation)
            Return False
        End If

        If cmbEmpType.Text.Length = 0 Then
            MsgBox("Must fill up Employment Type!", vbExclamation)
            Return False
        End If

        If cmbDepartment.SelectedIndex < 0 Then
            MsgBox("Must select Department!", vbExclamation)
            Return False
        End If

        If cmbDesignation.SelectedIndex < 0 Then
            MsgBox("Must select Designation!", vbExclamation)
            Return False
        End If

        If cmbSalaryGrade.SelectedIndex < 0 Then
            MsgBox("Must select Salary Grade!", vbExclamation)
            Return False
        End If

        If cmbStepGrade.SelectedIndex < 0 Then
            MsgBox("Must select Step Grade!", vbExclamation)
            Return False
        End If

        If Not hasFingerPrint Then
            MsgBox("Must register a fingerprint first!", vbExclamation)
            Return False
        End If
        Return True
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

    Private Sub cmbDepartment_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) Handles cmbDepartment.SelectionChanged
        Dim filter = String.Format("DeptID = {0} ", cmbDepartment.SelectedValue)
        cmbDesignation.ItemsSource = tblDesgAdapter.GetData.Select(filter)
    End Sub

    Sub txtBoxDidFocus(sender As Object, e As EventArgs) Handles txtFName.GotFocus, txtMName.GotFocus, txtLName.GotFocus, txtAge.GotFocus, txtAddress.GotFocus
        Dim txtBox As TextBox = sender
        txtBox.SelectAll()
    End Sub

    Private Sub birthDatePicker_SelectedDateChanged(sender As Object, e As SelectionChangedEventArgs) Handles birthDatePicker.SelectedDateChanged
        Dim age = Math.Floor(Now.Subtract(birthDatePicker.SelectedDate).TotalDays / 365.25)
        txtAge.Text = age
    End Sub
End Class
