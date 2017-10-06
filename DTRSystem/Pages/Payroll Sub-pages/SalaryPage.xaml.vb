Imports DTRSystem.DTRDataSetTableAdapters
Imports DTRSystem.DTRDataSet
Class SalaryPage
    Dim employee As EmployeeFullRow
    Dim employeeSalary As Decimal

    Dim workedTime As Double 'in minutes
    Dim totalTime As Double 'in minutes

    Dim awols As Double
    Dim lates As Double
    Public Sub ChangeEmployee(emp As EmployeeTableRow)
        If Not emp Is Nothing Then
            Dim employees = tblEmployeeFullAdapter.GetData.Select("ID = " & emp.ID)
            If employees.Count > 0 Then
                employee = employees(0)

                Dim filter As String = "ID = " & employee.salarygrade
                Dim rows = tblSalaryGradeAdapter.GetData().Select(filter)
                Dim salaryGrade As SalaryGradeTableRow = rows(0)
                Dim stepColumn As String = "Step" & employee.stepgrade
                employeeSalary = salaryGrade(stepColumn)

                lblMonthlyRate.Content = "P " & employeeSalary.ToString("N")
                lblKinsenas.Content = "P " & (employeeSalary / 2).ToString("N")
            Else
                MsgBox("Employee is not found", vbExclamation)
            End If
        End If
    End Sub

    Private Sub btnCalculate_Click(sender As Object, e As RoutedEventArgs) Handles btnCalculate.Click
        'awols = totalTime - workedTime
        Dim partialSalary = employeeSalary / 2
        Dim ratePerMinute = partialSalary / totalTime

        Dim payslipWindow = New PayslipWindow
        payslipWindow.fromDate = fromPicker.SelectedDate.Value
        payslipWindow.toDate = toPicker.SelectedDate.Value

        payslipWindow.basicPay = partialSalary

        payslipWindow.withholding = Double.Parse(txtWithholding.Text)
        payslipWindow.misc = Double.Parse(txtSSS.Text) + Double.Parse(txtPhilhealth.Text) + Double.Parse(txtPagibig.Text)
        payslipWindow.tardiness = lates * ratePerMinute 'to be calculated
        payslipWindow.leaves = awols * ratePerMinute

        payslipWindow.Show()
    End Sub

    Private Sub fromPicker_SelectedDateChanged(sender As Object, e As SelectionChangedEventArgs) Handles fromPicker.SelectedDateChanged
        toPicker.DisplayDateStart = fromPicker.SelectedDate
    End Sub

    Private Sub toPicker_SelectedDateChanged(sender As Object, e As SelectionChangedEventArgs) Handles toPicker.SelectedDateChanged
        Dim fromDate = fromPicker.SelectedDate.Value 'date from selected
        Dim toDate = toPicker.SelectedDate.Value 'date to selected
        Dim filter As String = String.Format("DateOfTheDay >= #{0}# AND DateOfTheDay <= #{1}#", fromDate, toDate) 'filter to get logs between fromdate and todate
        Dim rows = tblLogAdapter.GetEmployeeTableLog(employee.ID).Select(filter) 'gets the rows using the above filter

        Dim totalMinutes As Integer = 0 'initialize total minute worked
        Dim totalDays As Integer = Weekdays(fromDate, toDate) 'assigns total working days between dates

        Dim presents As New List(Of Date) 'initializes present days array
        Dim wDays = WorkingDays(fromDate, toDate) 'assigns working days array between dates
        Dim absences As New List(Of Date) 'initializes absent days array


        lates = 0 're-initializes lates
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
            If row.LeaveFrom.Date >= fromDate And row.LeaveFrom <= toDate Or _
                 row.LeaveTo.Date >= fromDate And row.LeaveTo <= toDate Then 'if the application's start or end date is between the selected dates
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
        awols = absences.Count * 8 * 60 'calculating days of absent * hours per day * minutes per hour
        workedTime = totalMinutes 'total worked time in minutes
        totalTime = totalDays * 8 * 60 'total working time in minutes

        'asigns values to the labels
        lblPeriod.Content = String.Format("{0} {1}-{2}", MonthName(fromDate.Month), fromDate.Day, toDate.Day)
        lblHoursWorked.Content = (workedTime / 60).ToString("0.00") & " hrs"
        lblHoursOfWork.Content = (totalTime / 60) & " hrs"


    End Sub

    Public Shared Function Weekdays(ByVal startDate As Date, ByVal endDate As Date) As Integer
        Dim numWeekdays As Integer
        Dim totalDays As Integer
        Dim WeekendDays As Integer
        numWeekdays = 0
        WeekendDays = 0

        totalDays = DateDiff(DateInterval.Day, startDate, endDate) + 1
        For i As Integer = 1 To totalDays
            If DatePart(DateInterval.Weekday, startDate) = 1 Then
                WeekendDays = WeekendDays + 1
            End If
            If DatePart(DateInterval.Weekday, startDate) = 7 Then
                WeekendDays = WeekendDays + 1
            End If
            startDate = DateAdd("d", 1, startDate)
        Next

        numWeekdays = totalDays - WeekendDays

        Return numWeekdays
    End Function

    Public Shared Function WorkingDays(ByVal startDate As Date, ByVal endDate As Date) As List(Of Date)
        Dim wDays As New List(Of Date)
        Dim totalDays As Integer

        totalDays = DateDiff(DateInterval.Day, startDate, endDate) + 1

        For i As Integer = 1 To totalDays
            Dim isWeekend As Boolean = False
            If startDate.DayOfWeek = DayOfWeek.Sunday Then
                isWeekend = True
            ElseIf startDate.DayOfWeek = DayOfWeek.Saturday Then
                isWeekend = True
            End If
            If Not isWeekend Then
                wDays.Add(startDate)
            End If
            startDate = DateAdd("d", 1, startDate)
        Next
        Return wDays
    End Function
End Class
