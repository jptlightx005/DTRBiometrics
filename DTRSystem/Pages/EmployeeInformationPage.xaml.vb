Imports DTRSystem.DTRBiometricDataSet
Class EmployeeInformationPage
    Dim selectedEmployee As EmployeeTableRow
    Private Sub employeePage_Initialized(sender As Object, e As EventArgs) Handles employeePage.Initialized
        cmbEmployees.ItemsSource = tblEmployeeAdapter.GetData
        'cmbDepartment.ItemsSource = tblDepartmentAdapter.GetData
    End Sub

    Private Sub cmbEmployees_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) Handles cmbEmployees.SelectionChanged
        Dim filterExpression = String.Format("ID = {0}", cmbEmployees.SelectedValue)
        Dim result = tblEmployeeAdapter.GetData.Select(filterExpression)
        If result.Count > 0 Then
            selectedEmployee = result(0)

            grdEmpForm.DataContext = selectedEmployee
        End If
    End Sub

    Private Sub btnSave_Click(sender As Object, e As RoutedEventArgs) Handles btnSave.Click
        If tblEmployeeAdapter.Update(selectedEmployee) = 1 Then
            MsgBox("Successfully updated!", vbInformation)
        Else
            MsgBox("Failed to Update!", vbExclamation)
        End If
    End Sub

    Private Sub btnAdd_Click(sender As Object, e As RoutedEventArgs) Handles btnAdd.Click

    End Sub
End Class
