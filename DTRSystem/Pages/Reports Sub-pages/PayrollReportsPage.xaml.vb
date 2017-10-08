﻿Imports System.Data
Imports DTRSystem.DTRDataSet
Imports DTRSystem.DTRDataSetTableAdapters
Class PayrollReportsPage
    Function GetPayrollTable(deptID As Integer, startDate As Date, endDate As Date) As DataTable
        Dim employees = tblEmployeeFullAdapter.GetData.Select("deptID = " & deptID)
        ' Create new DataTable instance.
        Dim table As New DataTable
        ' Create four typed columns in the DataTable.
        table.Columns.Add("full_name", GetType(String))
        table.Columns.Add("designation", GetType(String))
        table.Columns.Add("monthly", GetType(Decimal))
        table.Columns.Add("partial", GetType(Decimal))
        table.Columns.Add("tax", GetType(Decimal))
        table.Columns.Add("gsis", GetType(Decimal))
        table.Columns.Add("others", GetType(Decimal))
        table.Columns.Add("totaldeductions", GetType(Decimal))
        table.Columns.Add("netpay", GetType(Decimal))

        Dim tax = 1000
        Dim gsis = 300
        ' Add five rows with those columns filled in the DataTable.
        For Each employee As EmployeeFullRow In employees
            Dim row As DataRow = table.NewRow
            row("full_name") = employee.full_name
            row("designation") = employee.designation_name
            Dim salary = GetSalary(employee.salarygrade, employee.stepgrade)
            row("monthly") = salary
            row("partial") = salary / 2
            row("tax") = tax
            row("gsis") = gsis
            Dim others = GetLeaveDeductions(employee, salary, fromPicker.SelectedDate, toPicker.SelectedDate)
            row("others") = others
            Dim totalDeductions = tax + gsis + others
            row("totaldeductions") = totalDeductions
            row("netpay") = (salary / 2) - totalDeductions

            table.Rows.Add(row)
        Next
        Return table
    End Function

    Function GetSalary(salaryGradeInt As Integer, stepGradeInt As Integer) As Decimal
        Dim filter As String = "ID = " & salaryGradeInt
        Dim rows = tblSalaryGradeAdapter.GetData().Select(Filter)
        Dim salaryGrade As SalaryGradeTableRow = rows(0)
        Dim stepColumn As String = "Step" & stepGradeInt
        Return salaryGrade(stepColumn)
    End Function
    Function GetLeaveDeductions(employee As EmployeeFullRow, employeeSalary As Decimal, fromDate As Date, toDate As Date) As Double
        Dim filter As String = String.Format("DateOfTheDay >= #{0}# AND DateOfTheDay <= #{1}#", fromDate, toDate) 'filter to get logs between fromdate and todate
        Dim rows = tblLogAdapter.GetEmployeeTableLog(employee.ID).Select(filter) 'gets the rows using the above filter

        Dim totalMinutes As Integer = 0 'initialize total minute worked
        Dim totalDays As Integer = Weekdays(fromDate, toDate) 'assigns total working days between dates

        Dim presents As New List(Of Date) 'initializes present days array
        Dim wDays = WorkingDays(fromDate, toDate) 'assigns working days array between dates
        Dim absences As New List(Of Date) 'initializes absent days array


        Dim lates = 0 'initializes lates

        For Each row As TimelogTableRow In rows 'for each time log in imported logs
            totalMinutes += row.TotalTime 'adds total minutes worked from the log
            lates += (8 * 60) - row.TotalTime 'adds late from 480 - total time worked (if fully worked, 480 - 480 = 0)
            Debug.Print("Testing: {0} - {1} = {0}", (8 * 60), row.TotalTime, (8 * 60) - row.TotalTime)
            Debug.Print("lates test: {0}", lates)
            For Each d In wDays ' for each days in working days
                If d.Date = row.DateOfTheDay Then 'if the working day is the same as the log's date
                    presents.Add(d) 'adds the days to the present days array
                    wDays.Remove(d) 'removes the day from the working days array
                    Exit For 'stops the loop
                End If
            Next

        Next
        absences.AddRange(wDays) 'adds the working days left from the array into the absences array

        'leaves counting
        Dim leaves As New List(Of Date) 'initializes leaves array

        Dim leaveApp As LeaveApplicationsTableRow 'declaring leave application row
        For Each row As LeaveApplicationsTableRow In tblLeaveApplicationAdapter.GetData.Select("EmpID = " & employee.ID) 'for each application from the database
            If row.LeaveFrom.Date <= toDate.Date And row.LeaveTo >= fromDate Then 'if the application range overlaps with current leave
                leaveApp = row 'selects the application
                Exit For 'stops the loop
            End If
        Next

        If Not leaveApp Is Nothing Then 'if leave application is not empty, will start calculating
            Dim leaveStartDate = leaveApp.LeaveFrom 'assigns the leave application start date
            Dim leaveEndDate = leaveApp.LeaveTo 'assigns the leave application end date
            If leaveStartDate < fromDate Then 'if the start date was before the date from selected
                leaveStartDate = fromDate 'change the leave start to the selected start date
            End If
            If leaveEndDate > toDate Then 'if the end date was ater the date to selected
                leaveEndDate = toDate 'change the leave end to the selected end date
            End If
            Dim lDays = WorkingDays(leaveStartDate, leaveEndDate) 'initializes leave days array
            For Each ld In lDays 'for each day in leave days
                For Each d In absences ' for eac day in absent days
                    If d.Date = ld.Date Then 'if absent day is the same as leave day
                        presents.Add(d) 'the employee is paid , so the absent day is considered present
                        absences.Remove(d) 'removes the absent day from the absent array
                        Exit For 'stops the loop
                    End If
                Next
            Next
        End If

        'lates calculations
        Debug.Print("Lates before: {0}", lates)
        For Each leaveCreditLog As LeaveCreditsTableRow In tblLeaveCreditsAdapter.GetData.Select("EmpID = " & employee.ID)
            For Each d In presents
                If leaveCreditLog.DateOfTransaction.Date = d.Date And leaveCreditLog.VC_Used < 1.0 Then
                    Dim balance = leaveCreditLog.VC_Balance
                    Debug.Print("Balance left: {0}", balance)
                    If balance >= 0 Then
                        balance = 0
                    End If
                    Debug.Print("To be deducted: {0}", balance)
                    Dim minutes = Math.Round((leaveCreditLog.VC_Used + balance) * 480)
                    Debug.Print("Aaaa: {0}", (leaveCreditLog.VC_Used + balance))
                    Debug.Print("Bbbb: {0}", (leaveCreditLog.VC_Used + balance) * 480)
                    Debug.Print("CCCC: {0}", minutes)
                    lates -= minutes
                    Debug.Print("Lates ded: {0}", lates)
                    Exit For
                End If
            Next
        Next

        If lates < 0 Then
            lates = 0
        End If

        Debug.Print("Abscenses: {0} days", absences.Count)
        Debug.Print("Lates: {0} min", lates)
        Dim awols = absences.Count * 8 * 60 'calculating days of absent * hours per day * minutes per hour
        Dim workedTime = totalMinutes 'total worked time in minutes
        Dim totalTime = totalDays * 8 * 60 'total working time in minutes

        Dim partialSalary = employeeSalary / 2
        Dim ratePerMinute = partialSalary / totalTime


        Dim lateDeductions = lates * ratePerMinute 'to be calculated
        Dim leaveDeductions = awols * ratePerMinute

        Return lateDeductions + leaveDeductions
    End Function
    Private Sub btnSearch_Click(sender As Object, e As RoutedEventArgs) Handles btnSearch.Click
        If cmbDepartments.SelectedIndex >= 0 Then
            If fromPicker.SelectedDate Is Nothing And toPicker.SelectedDate Is Nothing Then
                MsgBox("Must select a start and end dates!", vbExclamation)
                Return
            End If

            Dim table = GetPayrollTable(cmbDepartments.SelectedValue, fromPicker.SelectedDate, toPicker.SelectedDate)
            payrollDataGrid.ItemsSource = table.DefaultView
        Else
            MsgBox("Must select a department!", vbExclamation)
        End If
    End Sub

    Private Sub Page_Initialized(sender As Object, e As EventArgs)
        cmbDepartments.ItemsSource = tblDeptAdapter.GetData
    End Sub
End Class
