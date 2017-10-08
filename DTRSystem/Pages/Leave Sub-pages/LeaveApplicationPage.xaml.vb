Imports DTRSystem.DTRDataSetTableAdapters
Imports DTRSystem.DTRDataSet
Class LeaveApplicationPage
    Dim employee As EmployeeFullRow

    Dim vacleavecredits As Double = 0
    Dim sickleavecredits As Double = 0
    Public Sub ChangeEmployee(emp As EmployeeTableRow)
        If Not emp Is Nothing Then
            Dim employees = tblEmployeeFullAdapter.GetData.Select("ID = " & emp.ID)
            If employees.Count > 0 Then
                employee = employees(0)
                grdLeaveApp.DataContext = employee
                deparment.Text = IIf(Not IsDBNull(employee("dept_name")), employee("dept_name"), "")
                designation.Text = IIf(Not IsDBNull(employee("designation_name")), employee("designation_name"), "")

                GetTotalCredits(employee)
                leaveApplicationsDataGrid.ItemsSource = tblLeaveApplicationAdapter.GetData.Select("EmpID = " & employee.ID)
            Else
                MsgBox("Employee is not found", vbExclamation)
            End If
        End If
    End Sub

    Public Sub GetTotalCredits(emp As EmployeeFullRow)
        Dim leaveCreditTransactions = tblLeaveCreditsAdapter.GetData.Select("EmpID = " & emp.ID)
        Dim date_modified = "---- --, ----"
        vacleavecredits = 0
        sickleavecredits = 0
        For Each lctransac As LeaveCreditsTableRow In leaveCreditTransactions
            Dim vc_earned = lctransac.VC_Earned
            Dim vc_used = lctransac.VC_Used
            Dim vc_bal = lctransac.VC_Balance

            vacleavecredits += vc_earned - vc_used
            Debug.Print("Is v equal {0} == {1}", vacleavecredits, vc_bal)

            Dim sc_earned = lctransac.SC_Earned
            Dim sc_used = lctransac.SC_Used
            Dim sc_bal = lctransac.SC_Balance

            sickleavecredits += sc_earned - sc_used
            Debug.Print("Is s equal {0} == {1}", sickleavecredits, sc_bal)

            date_modified = lctransac.DateOfTransaction.ToString("MMMM dd, yyyy")
        Next

        vacationcredits.Text = Math.Round(vacleavecredits, 3)
        sickcredits.Text = Math.Round(sickleavecredits, 3)

    End Sub

    Private Sub btnSubmit_Click(sender As Object, e As RoutedEventArgs) Handles btnSubmit.Click
        If ValidateFields() Then
            Dim leaveToDate As Date
            If chkMultipleDays.IsChecked Then
                leaveToDate = leaveToPicker.SelectedDate.Value
            Else
                leaveToDate = leaveFromPicker.SelectedDate.Value
            End If

            Dim totalNumberOfDays = Weekdays(leaveFromPicker.SelectedDate, leaveToDate)
            Dim creditsToBeUsed = 0

            If cmbLeaveType.SelectedIndex = 0 Then 'Vacation Leave
                creditsToBeUsed = vacleavecredits
            ElseIf cmbLeaveType.SelectedIndex = 1 Then 'Sick Leave
                creditsToBeUsed = sickleavecredits
            ElseIf cmbLeaveType.SelectedIndex = 2 Then 'Forced Leave
                creditsToBeUsed = vacleavecredits + sickleavecredits
            End If

            Debug.Print("Total: {0} left: {1}", totalNumberOfDays, creditsToBeUsed)
            If totalNumberOfDays > creditsToBeUsed Then
                MsgBox("Insufficient credits", vbExclamation)
                Return
            End If

            Dim dataTable As New LeaveApplicationsTableDataTable
            Dim leaveapplication As LeaveApplicationsTableRow = dataTable.NewRow

            leaveapplication.EmpID = employee.ID

            leaveapplication.Reason = txtReason.Text
            leaveapplication.ContactDuringLeave = txtContact.Text

            leaveapplication.LeaveFrom = leaveFromPicker.SelectedDate.Value
            leaveapplication.FromComment = ""

            leaveapplication.LeaveTo = leaveToDate
            leaveapplication.ToComment = ""

            leaveapplication.LeaveType = cmbLeaveType.Text
            leaveapplication.NumberOfDays = totalNumberOfDays

            dataTable.Rows.Add(leaveapplication)
            If tblLeaveApplicationAdapter.Update(dataTable) = 1 Then

                Dim lcDataTable = New LeaveCreditsTableDataTable
                Dim lctransaction As LeaveCreditsTableRow = lcDataTable.NewRow
                lctransaction.EmpID = employee.ID

                If cmbLeaveType.SelectedIndex = 0 Then 'Vacation Leave
                    lctransaction.VC_Earned = 0
                    lctransaction.VC_Used = totalNumberOfDays
                    lctransaction.VC_Balance = vacleavecredits - totalNumberOfDays

                    lctransaction.SC_Earned = 0
                    lctransaction.SC_Used = 0
                    lctransaction.SC_Balance = 0
                ElseIf cmbLeaveType.SelectedIndex = 1 Then 'Sick Leave
                    lctransaction.VC_Earned = 0
                    lctransaction.VC_Used = 0
                    lctransaction.VC_Balance = 0

                    lctransaction.SC_Earned = 0
                    lctransaction.SC_Used = totalNumberOfDays
                    lctransaction.SC_Balance = sickleavecredits - totalNumberOfDays
                ElseIf cmbLeaveType.SelectedIndex = 2 Then 'Forced Leave
                    'creditsToBeUsed = vacleavecredits + sickleavecredits
                End If
                lctransaction.DateOfTransaction = Now
                lctransaction.Remarks = txtReason.Text
                lcDataTable.Rows.Add(lctransaction)

                If tblLeaveCreditsAdapter.Update(lcDataTable) = 1 Then
                    MsgBox("Successfully added application!", vbInformation)
                    GetTotalCredits(employee)
                    leaveApplicationsDataGrid.ItemsSource = tblLeaveApplicationAdapter.GetData.Select("EmpID = " & employee.ID)

                    txtReason.Text = ""
                    txtContact.Text = ""
                    leaveFromPicker.SelectedDate = Nothing
                    leaveFromPicker.DisplayDate = DateTime.Today

                    leaveToPicker.SelectedDate = Nothing
                    leaveToPicker.DisplayDate = DateTime.Today

                    chkMultipleDays.IsChecked = False
                    cmbLeaveType.SelectedIndex = 0
                Else
                    MsgBox("Failed to add!", vbInformation)
                End If

                'GetTotalCredits(employee)

            Else
                MsgBox("Failed to add!", vbInformation)
            End If
        End If
    End Sub

    Private Sub chkMultipleDays_Click(sender As Object, e As RoutedEventArgs) Handles chkMultipleDays.Click
        leaveToPicker.IsEnabled = chkMultipleDays.IsChecked
    End Sub

    Function ValidateFields() As Boolean
        If employee Is Nothing Then
            MsgBox("Select employee first!", vbExclamation)
            Return False
        End If
        If txtReason.Text = "" Then
            MsgBox("Must fill out the reason to apply for a leave!", vbExclamation)
            Return False
        End If
        If txtContact.Text = "" Then
            MsgBox("Must fill out the contact number upon leave!", vbExclamation)
            Return False
        End If
        If txtContact.Text = "" Then
            MsgBox("Must fill out the contact number upon leave!", vbExclamation)
            Return False
        End If
        If leaveFromPicker.SelectedDate Is Nothing Then
            MsgBox("Must select a start date!", vbExclamation)
            Return False
        End If

        If chkMultipleDays.IsChecked And leaveToPicker.SelectedDate Is Nothing Then
            MsgBox("Must select an end date if multiple days is checked!", vbExclamation)
            Return False
        End If

        Dim leaveToDate As Date
        If chkMultipleDays.IsChecked Then
            leaveToDate = leaveToPicker.SelectedDate.Value
        Else
            leaveToDate = leaveFromPicker.SelectedDate.Value
        End If

        If Weekdays(leaveFromPicker.SelectedDate, leaveToDate) = 0 Then
            MsgBox("Day(s) selected are weekends!", vbExclamation)
            Return False
        End If

        Dim leaves = CheckForExistingLeavesBetweenDates(leaveFromPicker.SelectedDate, leaveToDate)
        
        If leaves.Count > 0 Then
            Dim leavesStr = ""
            For Each days In leaves
                leavesStr = leavesStr & days.ToString("MMMM dd") & vbCrLf
            Next
            leavesStr = Trim(leavesStr)
            MsgBox("Leaves already exists in the following days: " & vbCrLf & leavesStr, vbExclamation)
            Return False
        End If
        Return True
    End Function

    Function CheckForExistingLeavesBetweenDates(fromDate As Date, toDate As Date) As List(Of Date)
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
                leaves.Add(ld)
            Next
        End If

        Return leaves
    End Function
    Private Sub leaveFromPicker_SelectedDateChanged(sender As Object, e As SelectionChangedEventArgs) Handles leaveFromPicker.SelectedDateChanged
        leaveToPicker.DisplayDateStart = leaveFromPicker.SelectedDate
    End Sub
End Class
