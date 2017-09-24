Class EmployeePage
    Private Sub ReplacePage(ByRef pg As Page)
        employeeFrame.Content = pg
    End Sub

    Private Sub treeAdd_Selected(sender As Object, e As RoutedEventArgs) Handles treeAdd.Selected
        ReplacePage(New RegistrationPage)
    End Sub
End Class
