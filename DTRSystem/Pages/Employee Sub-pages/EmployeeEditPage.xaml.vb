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

            cmbSalaryGrade.SelectedIndex = selectedEmployee.salarygrade - 1
            cmbStepGrade.SelectedIndex = selectedEmployee.stepgrade - 1



            imgEmpPicture.Source = DataToBitmap(selectedEmployee("picture"))
            
            btnSave.IsEnabled = True
        Else
            btnSave.IsEnabled = False
        End If

    End Sub
    Private Sub btnSave_Click(sender As Object, e As RoutedEventArgs) Handles btnSave.Click
        selectedEmployee.deptID = cmbDepartment.SelectedValue
        selectedEmployee.desgID = cmbDesignation.SelectedValue

        selectedEmployee.salarygrade = cmbSalaryGrade.SelectedIndex + 1
        selectedEmployee.stepgrade = cmbStepGrade.SelectedIndex + 1

        If tblEmployeeAdapter.Update(selectedEmployee) = 1 Then
            MsgBox("Successfully updated!", vbInformation)
        Else
            MsgBox("Failed to Update!", vbExclamation)
        End If
    End Sub

    
    Private Sub cmbDepartment_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) Handles cmbDepartment.SelectionChanged
        Dim id = 0
        If cmbDepartment.SelectedIndex >= 0 Then
            Debug.Print("Flag 1")
            selectedDepartment = departmentList(cmbDepartment.SelectedIndex)

            id = selectedDepartment.ID
        End If
        Debug.Print("Flag 2")
        Dim filter = String.Format("DeptID = {0}", id)
        Debug.Print("Firuta: {0}", filter)
        cmbDesignation.ItemsSource = tblDesgAdapter.GetData.Select(filter)
        Debug.Print("Flag 4")
        cmbDesignation.SelectedValue = selectedEmployee.desgID
        Debug.Print("Flag 5")
    End Sub
End Class
