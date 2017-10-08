Class LeaveReportPage

    Private Sub Page_Initialized(sender As Object, e As EventArgs)
        cmbEmployees.ItemsSource = tblEmployeeAdapter.GetData
    End Sub

    Private Sub btnSearch_Click(sender As Object, e As RoutedEventArgs) Handles btnSearch.Click
        btnPrint.IsEnabled = False
        If cmbEmployees.SelectedIndex >= 0 Then
            leaveDataGrid.ItemsSource = tblLeaveCreditsAdapter.GetData.Select("EmpID = " & cmbEmployees.SelectedValue)
            btnPrint.IsEnabled = True
        End If
    End Sub

    Private Sub btnPrint_Click(sender As Object, e As RoutedEventArgs) Handles btnPrint.Click
        Dim webReportWindow As New LeaveReportWebWindow
        Dim employees = tblEmployeeFullAdapter.GetData.Select("ID = " & cmbEmployees.SelectedValue)
        If employees.Count > 0 Then
            webReportWindow.employee = employees(0)

            webReportWindow.records = tblLeaveCreditsAdapter.GetData.Select("EmpID = " & cmbEmployees.SelectedValue)

            webReportWindow.ShowDialog()
        Else
            MsgBox("Employee is not found", vbExclamation)
        End If
    End Sub
End Class
