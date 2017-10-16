Imports DTRSystem.DTRDataSet
Class EmployeePage
    Dim regPage As RegistrationPage
    Dim editPage As EmployeeEditPage
    
    Dim activePage As Page
    Private Sub employeePage_Initialized(sender As Object, e As EventArgs) Handles MyBase.Initialized
        cmbEmployees.ItemsSource = tblEmployeeAdapter.GetData

        For Each emp In tblEmployeeAdapter.GetData

        Next
    End Sub
    Private Sub ReplacePage(ByRef pg As Page)
        activePage = pg
        employeeFrame.Content = pg
    End Sub

    Private Sub cmbEmployees_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) Handles cmbEmployees.SelectionChanged
        
        
    End Sub
    Private Sub treeAdd_Selected(sender As Object, e As RoutedEventArgs) Handles treeAdd.Selected
        If regPage Is Nothing Then
            regPage = New RegistrationPage
        End If
        cmbEmployees.Visibility = Windows.Visibility.Hidden
        btnSearch.Visibility = cmbEmployees.Visibility
        ReplacePage(regPage)
    End Sub

    Private Sub treeEdit_Selected(sender As Object, e As RoutedEventArgs) Handles treeEdit.Selected
        If editPage Is Nothing Then
            editPage = New EmployeeEditPage
        End If
        cmbEmployees.Visibility = Windows.Visibility.Visible
        btnSearch.Visibility = cmbEmployees.Visibility
        ReplacePage(editPage)
    End Sub

    Private Sub cmbEmployees_PreviewTextInput(sender As Object, e As TextCompositionEventArgs) Handles cmbEmployees.PreviewTextInput
        Dim cmbBx As ComboBox = sender
        cmbBx.IsDropDownOpen = True
    End Sub

    Private Sub btnSearch_Click(sender As Object, e As RoutedEventArgs) Handles btnSearch.Click
        If activePage Is Nothing Then
            'by default, if employee page is shown the first time and selected an empoyee
            'the employee editing page will show up
            If editPage Is Nothing Then
                editPage = New EmployeeEditPage
            End If

            ReplacePage(editPage)
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
        If activePage Is editPage Then
            editPage.ChangeEmployee(selectedEmployee)
        End If
    End Sub
End Class
