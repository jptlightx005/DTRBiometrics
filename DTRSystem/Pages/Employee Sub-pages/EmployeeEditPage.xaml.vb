Imports DTRSystem.DTRDataSet
Class EmployeeEditPage
    Public selectedEmployee As EmployeeTableRow
    Private Sub Page_Loaded(sender As Object, e As RoutedEventArgs)
        ChangeEmployee(selectedEmployee)
        cmbDepartment.ItemsSource = tblDeptAdapter.GetData
        cmbDesignation.ItemsSource = tblDesgAdapter.GetData
    End Sub
    Public Sub ChangeEmployee(emp As EmployeeTableRow)
        If Not emp Is Nothing Then
            selectedEmployee = emp

            grdEmpForm.DataContext = selectedEmployee
            imgEmpPicture.Source = DataToBitmap(selectedEmployee("picture"))

            btnSave.IsEnabled = True
        Else
            btnSave.IsEnabled = False
        End If

    End Sub
    Private Sub btnSave_Click(sender As Object, e As RoutedEventArgs) Handles btnSave.Click
        If tblEmployeeAdapter.Update(selectedEmployee) = 1 Then
            MsgBox("Successfully updated!", vbInformation)
        Else
            MsgBox("Failed to Update!", vbExclamation)
        End If
    End Sub

    
End Class
