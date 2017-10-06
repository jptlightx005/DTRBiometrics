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
        Dim fromDate = fromPicker.SelectedDate.Value
        Dim toDate = toPicker.SelectedDate.Value
        Dim filter As String = String.Format("DateOfTheDay >= #{0}# AND DateOfTheDay <= #{1}#", fromDate, toDate)
        Dim rows = tblLogAdapter.GetEmployeeTableLog(employee.ID).Select(filter)

        Dim totalMinutes As Integer = 0
        Dim totalDays As Integer = Weekdays(fromDate, toDate)

        Dim presents As New List(Of Date)
        Dim wDays = WorkingDays(fromDate, toDate)
        Dim absences As New List(Of Date)
        lates = 0
        For Each row As TimelogTableRow In rows
            totalMinutes += row.TotalTime
            lates += (8 * 60) - row.TotalTime
            Debug.Print("Testing: {0} - {1} = {0}", (8 * 60), row.TotalTime, (8 * 60) - row.TotalTime)
            Debug.Print("lates test: {0}", lates)
            For Each d In wDays
                If d.Date = row.DateOfTheDay Then
                    presents.Add(d)
                    wDays.Remove(d)
                    Exit For
                End If
            Next

        Next
        absences.AddRange(wDays)
        Debug.Print("Abscenses: {0} days", absences.Count)
        Debug.Print("Lates: {0} min", lates)
        awols = absences.Count * 8 * 60
        workedTime = totalMinutes
        totalTime = totalDays * 8 * 60

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
