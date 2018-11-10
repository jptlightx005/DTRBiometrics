Imports DTRSystem.DTRDataSet
Class EmployeeEditPage
    Public selectedEmployee As EmployeeTableRow

    Dim departmentList As DepartmentTableDataTable
    Dim selectedDepartment As DepartmentTableRow
    Private Sub Page_Loaded(sender As Object, e As RoutedEventArgs)
        departmentList = tblDeptAdapter.GetData
        cmbDepartment.ItemsSource = departmentList

        ChangeEmployee(selectedEmployee)
    End Sub
    Public Sub ChangeEmployee(emp As EmployeeTableRow)
        If Not emp Is Nothing Then
            selectedEmployee = emp

            grdEmpForm.DataContext = selectedEmployee

            cmbDepartment.SelectedIndex = -1
            cmbDepartment.SelectedValue = selectedEmployee.deptID

            'cmbSalaryGrade.SelectedIndex = selectedEmployee.salarygrade - 1
            'cmbStepGrade.SelectedIndex = selectedEmployee.stepgrade - 1

            imgEmpPicture.Source = DataToBitmap(selectedEmployee("picture"))
            
            btnSave.IsEnabled = True
        Else
            btnSave.IsEnabled = False
        End If
    End Sub
    Private Sub btnSave_Click(sender As Object, e As RoutedEventArgs) Handles btnSave.Click
        If ValidateFields() Then
            selectedEmployee.deptID = cmbDepartment.SelectedValue
            selectedEmployee.desgID = cmbDesignation.SelectedValue

            'selectedEmployee.salarygrade = cmbSalaryGrade.SelectedIndex + 1
            'selectedEmployee.stepgrade = cmbStepGrade.SelectedIndex + 1

            If tblEmployeeAdapter.Update(selectedEmployee) = 1 Then
                MsgBox("Successfully updated!", vbInformation)
            Else
                MsgBox("Failed to Update!", vbExclamation)
            End If
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

        'If cmbSalaryGrade.SelectedIndex < 0 Then
        '    MsgBox("Must select Salary Grade!", vbExclamation)
        '    Return False
        'End If

        'If cmbStepGrade.SelectedIndex < 0 Then
        '    MsgBox("Must select Step Grade!", vbExclamation)
        '    Return False
        'End If

        Return True
    End Function
    Private Sub cmbDepartment_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) Handles cmbDepartment.SelectionChanged
        Dim id = 0
        If cmbDepartment.SelectedIndex >= 0 Then
            selectedDepartment = departmentList(cmbDepartment.SelectedIndex)

            id = selectedDepartment.ID
        End If
        Dim filter = String.Format("DeptID = {0}", id)
        cmbDesignation.ItemsSource = tblDesgAdapter.GetData.Select(filter)
        cmbDesignation.SelectedValue = selectedEmployee.desgID
    End Sub

    Sub txtBoxDidFocus(sender As Object, e As EventArgs) Handles txtFName.GotFocus, txtMName.GotFocus, txtLName.GotFocus, txtAge.GotFocus, txtAddress.GotFocus
        Dim txtBox As TextBox = sender
        txtBox.SelectAll()
    End Sub
End Class
