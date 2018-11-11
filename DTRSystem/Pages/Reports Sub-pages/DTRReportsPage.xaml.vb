Imports DTRFuncs
Imports System.Data
Imports DTRSystem.DTRDataSetTableAdapters
Imports DTRSystem.DTRDataSet
Imports DTRFuncs.DocumentExporter


Class DTRReportsPage

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        cmbEmployees.ItemsSource = tblEmployeeAdapter.GetData


    End Sub

    Private Sub btnSearch_Click(sender As Object, e As RoutedEventArgs) Handles btnSearch.Click
        If cmbEmployees.SelectedIndex >= 0 Then
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
        Else
            MsgBox("Name is not found!", vbExclamation)
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

    Private Sub cmbEmployees_PreviewTextInput(sender As Object, e As TextCompositionEventArgs) Handles cmbEmployees.PreviewTextInput
        Dim cmbBx As ComboBox = sender
        cmbBx.IsDropDownOpen = True
    End Sub

    Private Sub btn_Export_Click(sender As Object, e As RoutedEventArgs) Handles btn_Export.Click
        If cmbEmployees.SelectedIndex = -1 Then
            MsgBox("There are no employees added yet!", MsgBoxStyle.Exclamation, "Empty Records")
            Return
        End If
        Dim employees = tblEmployeeFullAdapter.GetData.Select("ID = " & cmbEmployees.SelectedValue)
        If employees.Count = 0 Then
            MsgBox("Employee not found!", MsgBoxStyle.Exclamation, "Empty Records")
            Return
        End If

        'webReportWindow.employee = employees(0)
        'If fromPicker.SelectedDate Is Nothing Or toPicker.SelectedDate Is Nothing Then
        '    webReportWindow.records = FilteredDataTable(tblLogAdapter.GetEmployeeTableLog(cmbEmployees.SelectedValue))
        'Else
        '    Dim filterExpression As String = String.Format("DateOfTheDay >= #{0}# AND DateOfTheDay <= #{1}#", fromPicker.SelectedDate, toPicker.SelectedDate)
        '    Dim dataTable As New TimelogTableDataTable
        '    For Each row In tblLogAdapter.GetEmployeeTableLog(cmbEmployees.SelectedValue).Select(filterExpression)
        '        dataTable.ImportRow(row)
        '    Next
        '    webReportWindow.records = FilteredDataTable(dataTable)
        'End If
        'webReportWindow.ShowDialog()

        Dim employee = employees.First
        Dim exportDate = exportMonthPicker.SelectedDate
        If exportDate Is Nothing Then
            exportDate = DateTime.Now
        End If
        Dim exportFrom = New DateTime(exportDate.Value.Year, exportDate.Value.Month, 1)
        Dim exportTo = exportFrom.AddMonths(1).AddSeconds(-1)

        Dim record = New Dictionary(Of String, Object)
        Dim period = New Dictionary(Of String, DateTime)
        period("From") = exportFrom
        period("To") = exportTo

        record("Period") = period
        record("Employee") = employee("first_name_first")

        Dim timeLogs = New List(Of Dictionary(Of String, Object))

        Dim filterExpression As String = String.Format("DateOfTheDay >= #{0}# AND DateOfTheDay <= #{1}#", exportFrom, exportTo)

        'Dim filterExpression As String = String.Format("DateOfTheDay >= #{0}# AND DateOfTheDay <= #{1}#", fromPicker.SelectedDate, toPicker.SelectedDate)
        Debug.Print("Filter Expression: {0}", filterExpression)

        'For Each rrow In tblLogAdapter.GetEmployeeTableLog(cmbEmployees.SelectedValue).Select(filterExpression)

        'Next
        For Each rrow In tblLogAdapter.GetEmployeeTableLog(cmbEmployees.SelectedValue).Select(filterExpression)

            Dim timeLog = New Dictionary(Of String, Object)
            Dim row = DirectCast(rrow, TimelogTableRow)
            timeLog("DateOfTheDay") = row.DateOfTheDay
            timeLog("TimeInAM") = row.TimeInAM.ToString("hh:mm tt")
            timeLog("TimeOutAM") = row.TimeOutAM.ToString("hh:mm tt")
            timeLog("TimeInPM") = row.TimeInPM.ToString("hh:mm tt")
            timeLog("TimeOutPM") = row.TimeOutPM.ToString("hh:mm tt")

            timeLogs.Add(timeLog)
        Next
        record("TimeLogs") = timeLogs
        Debug.Print(timeLogs.ToString())
        DocumentExporter.exportToExcel(record)

    End Sub
End Class
