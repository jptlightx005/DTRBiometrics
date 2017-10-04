Class DTRReportsPage

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

    Private Sub btnPrint_Click(sender As Object, e As RoutedEventArgs) Handles btnPrint.Click
        Dim webReportWindow As New DTRReportWebWindow
        Dim employees = tblEmployeeFullAdapter.GetData.Select("ID = " & cmbEmployees.SelectedValue)
        If employees.Count > 0 Then
            webReportWindow.employee = employees(0)
            webReportWindow.ShowDialog()
        Else
            MsgBox("Employee is not found", vbExclamation)
        End If

        'webReportWindow.allMonth = chkAllMonths.IsChecked
        'If Not chkAllMonths.IsChecked Then
        '    If fromPicker.SelectedDate Is Nothing And toPicker.SelectedDate Is Nothing Then
        '        MsgBox("Please select dates!", vbExclamation)
        '        Return
        '    End If
        '    webReportWindow.fromMonth = fromPicker.SelectedDate.Value
        '    webReportWindow.toMonth = toPicker.SelectedDate.Value

        'End If

        'webReportWindow.selectedDep = department_list(cmb_dep.SelectedIndex)
    End Sub

    Private Sub cmbEmployees_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) Handles cmbEmployees.SelectionChanged
        btnPrint.IsEnabled = cmbEmployees.SelectedIndex >= 0
    End Sub
End Class
