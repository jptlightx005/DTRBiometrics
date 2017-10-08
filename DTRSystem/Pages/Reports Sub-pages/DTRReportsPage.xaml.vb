Imports SMSCSFuncs
Imports System.Data
Imports DTRSystem.DTRDataSetTableAdapters
Imports DTRSystem.DTRDataSet
Class DTRReportsPage

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        cmbEmployees.ItemsSource = tblEmployeeAdapter.GetData


    End Sub

    Private Sub btnSearch_Click(sender As Object, e As RoutedEventArgs) Handles btnSearch.Click
        If fromPicker.SelectedDate Is Nothing Or toPicker.SelectedDate Is Nothing Then
            timeLogDataGrid.ItemsSource = FilteredDataTable(tblLogAdapter.GetEmployeeTableLog(cmbEmployees.SelectedValue))
        Else
            Dim filterExpression As String = String.Format("DateOfTheDay >= #{0}# AND DateOfTheDay <= #{1}#", fromPicker.SelectedDate, toPicker.SelectedDate)
            Dim dataTable As New TimelogTableDataTable
            For Each row In tblLogAdapter.GetEmployeeTableLog(cmbEmployees.SelectedValue).Select(filterExpression)
                dataTable.ImportRow(row)
            Next
            timeLogDataGrid.ItemsSource = FilteredDataTable(dataTable)
        End If
    End Sub

    Function FilteredDataTable(timeLog As TimelogTableDataTable) As TimelogTableDataTable
        Dim newDataTable = timeLog
        newDataTable.Columns.Add("TimeInAMStr", GetType(String))
        newDataTable.Columns.Add("TimeOutAMStr", GetType(String))
        newDataTable.Columns.Add("TimeInPMStr", GetType(String))
        newDataTable.Columns.Add("TimeOutPMStr", GetType(String))
        newDataTable.Columns.Add("TotalHours", GetType(Decimal))
        For i = 0 To newDataTable.Rows.Count - 1
            Dim log As TimelogTableRow = newDataTable.Rows(i)
            Dim timeSTR = "--:-- --"
            If Not IsDBNull(log("TimeInAM")) Then
                timeSTR = log.TimeInAM.ToString("hh:mm tt")
            End If
            newDataTable.Rows(i)("TimeInAMStr") = timeSTR


            timeSTR = "--:-- --"
            If Not IsDBNull(log("TimeOutAM")) Then
                timeSTR = log.TimeOutAM.ToString("hh:mm tt")
            End If
            newDataTable.Rows(i)("TimeOutAMStr") = timeSTR


            timeSTR = "--:-- --"
            If Not IsDBNull(log("TimeInPM")) Then
                timeSTR = log.TimeInPM.ToString("hh:mm tt")
            End If
            newDataTable.Rows(i)("TimeInPMStr") = timeSTR


            timeSTR = "--:-- --"
            If Not IsDBNull(log("TimeOutPM")) Then
                timeSTR = log.TimeOutPM.ToString("hh:mm tt")
            End If
            newDataTable.Rows(i)("TimeOutPMStr") = timeSTR

            newDataTable.Rows(i)("TotalHours") = log.TotalTime / 60.0
        Next

        Return newDataTable
    End Function
    Private Sub btnPrint_Click(sender As Object, e As RoutedEventArgs) Handles btnPrint.Click
        Dim webReportWindow As New DTRReportWebWindow
        Dim employees = tblEmployeeFullAdapter.GetData.Select("ID = " & cmbEmployees.SelectedValue)
        If employees.Count > 0 Then
            webReportWindow.employee = employees(0)
            If fromPicker.SelectedDate Is Nothing Or toPicker.SelectedDate Is Nothing Then
                webReportWindow.records = FilteredDataTable(tblLogAdapter.GetEmployeeTableLog(cmbEmployees.SelectedValue))
            Else
                Dim filterExpression As String = String.Format("DateOfTheDay >= #{0}# AND DateOfTheDay <= #{1}#", fromPicker.SelectedDate, toPicker.SelectedDate)
                Dim dataTable As New TimelogTableDataTable
                For Each row In tblLogAdapter.GetEmployeeTableLog(cmbEmployees.SelectedValue).Select(filterExpression)
                    dataTable.ImportRow(row)
                Next
                webReportWindow.records = FilteredDataTable(dataTable)
            End If
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
