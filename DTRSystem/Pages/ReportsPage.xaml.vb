Class ReportsPage

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        cmbEmployees.ItemsSource = tblEmployeeAdapter.GetData


    End Sub

    Private Sub btnSearch_Click(sender As Object, e As RoutedEventArgs) Handles btnSearch.Click
        If fromPicker.SelectedDate Is Nothing Or toPicker.SelectedDate Is Nothing Then
            timeLogDataGrid.ItemsSource = tblLogAdapter.GetEmployeeTableLog(cmbEmployees.SelectedValue)
        Else
            Dim filterExpression As String = String.Format("DateOfTheDay >= #{0}# AND DateOfTheDay <= #{1}#", fromPicker.SelectedDate, toPicker.SelectedDate)
            timeLogDataGrid.ItemsSource = tblLogAdapter.GetEmployeeTableLog(cmbEmployees.SelectedValue).Select(filterExpression)
        End If
    End Sub
End Class
