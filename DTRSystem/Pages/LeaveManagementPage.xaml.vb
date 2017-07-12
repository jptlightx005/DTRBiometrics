Imports DTRSystem.DTRBiometricDataSetTableAdapters
Imports DTRSystem.DTRBiometricDataSet
Class LeaveManagementpage
    Dim dataTable As New LeaveCreditsTableDataTable
    Dim creditRow As LeaveCreditsTableRow

    Private Sub adminPage_Initialized(sender As Object, e As EventArgs) Handles MyBase.Initialized, MyBase.Initialized
        cmbEmployees.ItemsSource = tblEmployeeAdapter.GetData
        creditRow = dataTable.NewRow
    End Sub

    Private Sub btnAdd_Click(sender As Object, e As RoutedEventArgs) Handles btnAdd.Click

    End Sub
End Class
