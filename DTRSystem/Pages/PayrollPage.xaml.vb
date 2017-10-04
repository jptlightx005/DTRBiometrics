Imports DTRSystem.DTRDataSet
Class PayrollPage
    Dim salaryPage As SalaryPage
    'Dim payrollPage As PayrollPage
    Dim activePage As Page
    Private Sub ReplacePage(ByRef pg As Page)
        activePage = pg
        payrollFrame.Content = pg
    End Sub
    Private Sub cmbEmployees_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) Handles cmbEmployees.SelectionChanged
        If activePage Is Nothing Then
            'by default, if employee page is shown the first time and selected an empoyee
            'the employee editing page will show up
            If salaryPage Is Nothing Then
                salaryPage = New SalaryPage
            End If

            ReplacePage(salaryPage)
        End If

        Dim selectedEmployee As employeeTableRow
        If cmbEmployees.SelectedIndex >= 0 Then
            Dim filterExpression = String.Format("ID = {0}", cmbEmployees.SelectedValue)
            Dim result = tblEmployeeAdapter.GetData.Select(filterExpression)
            If result.Count > 0 Then
                selectedEmployee = result(0)
            End If
        Else
            selectedEmployee = Nothing
        End If

        cmbEmployees.Visibility = Windows.Visibility.Visible
        If activePage Is salaryPage Then
            salaryPage.ChangeEmployee(selectedEmployee)
        End If

    End Sub
    Private Sub treeSalary_Selected(sender As Object, e As RoutedEventArgs) Handles treeSalary.Selected
        If salaryPage Is Nothing Then
            salaryPage = New SalaryPage
        End If
        ReplacePage(salaryPage)
    End Sub

    Private Sub Page_Initialized(sender As Object, e As EventArgs)
        cmbEmployees.ItemsSource = tblEmployeeAdapter.GetData
    End Sub
End Class
