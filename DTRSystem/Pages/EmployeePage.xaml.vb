Imports DTRSystem.DTRDataSet
Class EmployeePage
    Dim regPage As RegistrationPage
    Dim editPage As EmployeeEditPage
    Dim leavePage As LeaveManagementpage
    Dim leaveApplicationPage As LeaveApplication
    Dim activePage As Page
    Private Sub employeePage_Initialized(sender As Object, e As EventArgs) Handles MyBase.Initialized
        cmbEmployees.ItemsSource = tblEmployeeAdapter.GetData
    End Sub
    Private Sub ReplacePage(ByRef pg As Page)
        activePage = pg
        employeeFrame.Content = pg
    End Sub

    Private Sub cmbEmployees_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) Handles cmbEmployees.SelectionChanged
        If activePage Is Nothing Then
            'by default, if employee page is shown the first time and selected an empoyee
            'the employee editing page will show up
            If editPage Is Nothing Then
                editPage = New EmployeeEditPage
            End If

            ReplacePage(editPage)
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
        If activePage Is editPage Then
            editPage.ChangeEmployee(selectedEmployee)

        ElseIf activePage Is leavePage Then
            leavePage.ChangeEmployee(selectedEmployee)
        ElseIf activePage Is leaveApplicationPage Then
            leaveApplicationPage.ChangeEmployee(selectedEmployee)
        End If
        
    End Sub
    Private Sub treeAdd_Selected(sender As Object, e As RoutedEventArgs) Handles treeAdd.Selected
        If regPage Is Nothing Then
            regPage = New RegistrationPage
        End If
        cmbEmployees.Visibility = Windows.Visibility.Hidden
        ReplacePage(regPage)
    End Sub

    Private Sub treeEdit_Selected(sender As Object, e As RoutedEventArgs) Handles treeEdit.Selected
        If editPage Is Nothing Then
            editPage = New EmployeeEditPage
        End If
        cmbEmployees.Visibility = Windows.Visibility.Visible
        ReplacePage(editPage)
    End Sub

    Private Sub treeLCAdd_Selected(sender As Object, e As RoutedEventArgs) Handles treeLCAdd.Selected
        If leavePage Is Nothing Then
            leavePage = New LeaveManagementpage
        End If

        If cmbEmployees.SelectedIndex >= 0 Then
            Dim filterExpression = String.Format("ID = {0}", cmbEmployees.SelectedValue)
            Dim result = tblEmployeeAdapter.GetData.Select(filterExpression)
            If result.Count > 0 Then
                leavePage.ChangeEmployee(result(0))
            End If
        Else
            leavePage.ChangeEmployee(Nothing)
        End If

        cmbEmployees.Visibility = Windows.Visibility.Visible
        ReplacePage(leavePage)
    End Sub

    Private Sub treeLCApply_Selected(sender As Object, e As RoutedEventArgs) Handles treeLCApply.Selected
        If leaveApplicationPage Is Nothing Then
            leaveApplicationPage = New LeaveApplication
        End If

        If cmbEmployees.SelectedIndex >= 0 Then
            Dim filterExpression = String.Format("ID = {0}", cmbEmployees.SelectedValue)
            Dim result = tblEmployeeAdapter.GetData.Select(filterExpression)
            If result.Count > 0 Then
                leaveApplicationPage.ChangeEmployee(result(0))
            End If
        Else
            leaveApplicationPage.ChangeEmployee(Nothing)
        End If

        cmbEmployees.Visibility = Windows.Visibility.Visible
        ReplacePage(leaveApplicationPage)
    End Sub
End Class
