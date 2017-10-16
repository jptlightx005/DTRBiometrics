Imports DTRSystem.DTRDataSet
Class LeavePage
    Dim leaveCreditPage As LeaveCreditsPage
    Dim leaveApplicationPage As LeaveApplicationPage
    Dim activePage As Page

    Private Sub ReplacePage(ByRef pg As Page)
        activePage = pg
        leaveFrame.Content = pg
    End Sub

    Private Sub cmbEmployees_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) Handles cmbEmployees.SelectionChanged
        
    End Sub

    Private Sub treeAdd_Selected(sender As Object, e As RoutedEventArgs) Handles treeAdd.Selected
        If leaveCreditPage Is Nothing Then
            leaveCreditPage = New LeaveCreditsPage
        End If

        If cmbEmployees.SelectedIndex >= 0 Then
            Dim filterExpression = String.Format("ID = {0}", cmbEmployees.SelectedValue)
            Dim result = tblEmployeeAdapter.GetData.Select(filterExpression)
            If result.Count > 0 Then
                leaveCreditPage.ChangeEmployee(result(0))
            End If
        Else
            leaveCreditPage.ChangeEmployee(Nothing)
        End If

        cmbEmployees.Visibility = Windows.Visibility.Visible
        ReplacePage(leaveCreditPage)
    End Sub

    Private Sub treeLeave_Selected(sender As Object, e As RoutedEventArgs) Handles treeLeave.Selected
        If leaveApplicationPage Is Nothing Then
            leaveApplicationPage = New LeaveApplicationPage
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

    Private Sub leaveMgmtPage_Initialized(sender As Object, e As EventArgs) Handles leaveMgmtPage.Initialized
        cmbEmployees.ItemsSource = tblEmployeeAdapter.GetData
    End Sub

    Private Sub cmbEmployees_PreviewTextInput(sender As Object, e As TextCompositionEventArgs) Handles cmbEmployees.PreviewTextInput
        Dim cmbBx As ComboBox = sender
        cmbBx.IsDropDownOpen = True
    End Sub

    Private Sub btnSearch_Click(sender As Object, e As RoutedEventArgs) Handles btnSearch.Click
        If activePage Is Nothing Then
            'by default, if employee page is shown the first time and selected an empoyee
            'the employee editing page will show up
            If leaveApplicationPage Is Nothing Then
                leaveApplicationPage = New LeaveApplicationPage
            End If

            ReplacePage(leaveApplicationPage)
        End If

        Dim selectedEmployee As EmployeeTableRow
        If cmbEmployees.SelectedIndex >= 0 Then
            Dim filterExpression = String.Format("ID = {0}", cmbEmployees.SelectedValue)
            Dim result = tblEmployeeAdapter.GetData.Select(filterExpression)
            If result.Count > 0 Then
                selectedEmployee = result(0)
            End If
        Else
            selectedEmployee = Nothing
            MsgBox("Name is not found!", vbExclamation)
        End If

        cmbEmployees.Visibility = Windows.Visibility.Visible
        If activePage Is leaveApplicationPage Then
            leaveApplicationPage.ChangeEmployee(selectedEmployee)
        ElseIf activePage Is leaveCreditPage Then
            leaveCreditPage.ChangeEmployee(selectedEmployee)
        End If
    End Sub
End Class
