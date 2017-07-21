Class ReportsPage

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        cmbEmployees.ItemsSource = tblEmployeeAdapter.GetData
    End Sub

    Private Sub btnSearch_Click(sender As Object, e As RoutedEventArgs) Handles btnSearch.Click
        timeLogDataGrid.ItemsSource = tblLogAdapter.GetEmployeeTableLog(cmbEmployees.SelectedValue)

    End Sub
End Class
