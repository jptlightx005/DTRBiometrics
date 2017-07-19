Imports DTRSystem.DTRDataSetTableAdapters
Imports DTRSystem.DTRDataSet
Class LeaveManagementpage

    Private Sub adminPage_Initialized(sender As Object, e As EventArgs) Handles MyBase.Initialized, MyBase.Initialized
        cmbEmployees.ItemsSource = tblEmployeeAdapter.GetData

    End Sub

    Private Sub btnAdd_Click(sender As Object, e As RoutedEventArgs) Handles btnAdd.Click

    End Sub
End Class
