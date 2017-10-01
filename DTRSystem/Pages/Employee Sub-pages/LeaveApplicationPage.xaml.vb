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
        For Each lctransac As LeaveCreditsTableRow In leaveCreditTransactions
            Dim vc_earned = lctransac.VC_Earned
            Dim vc_used = lctransac.VC_Used
            Dim vc_bal = lctransac.VC_Balance

            vacleavecredits += vc_earned - vc_used
            Debug.Print("Is v equal {0} == {1}", vacleavecredits, vc_bal)
            If vacleavecredits <> vc_bal Then
                vacleavecredits -= vc_earned - vc_used
                Debug.Print("Deducted v = {0}", vacleavecredits)
            End If

            Dim sc_earned = lctransac.SC_Earned
            Dim sc_used = lctransac.SC_Used
            Dim sc_bal = lctransac.SC_Balance

            sickleavecredits += sc_earned - sc_used
            Debug.Print("Is s equal {0} == {1}", sickleavecredits, sc_bal)
            If sickleavecredits <> sc_bal Then
                sickleavecredits -= sc_earned - sc_used
                Debug.Print("Deducted s = {0}", sickleavecredits)
            End If

            date_modified = lctransac.DateOfTransaction.ToString("MMMM dd, yyyy")
        Next

        vacationcredits.Text = vacleavecredits
        sickcredits.Text = sickleavecredits

    End Sub

    Private Sub btnSubmit_Click(sender As Object, e As RoutedEventArgs) Handles btnSubmit.Click
        Dim dataTable As New LeaveApplicationsTableDataTable
        Dim leaveapplication As LeaveApplicationsTableRow = dataTable.NewRow
        leaveapplication.EmpID = employee.ID

        leaveapplication.Reason = txtReason.Text
        leaveapplication.ContactDuringLeave = txtContact.Text

        leaveapplication.LeaveFrom = leaveFromPicker.SelectedDate.Value
        leaveapplication.FromDuration = (cmbLeaveFromDuration.SelectedIndex + 1) * 0.5
        leaveapplication.FromComment = ""

        leaveapplication.LeaveTo = leaveToPicker.SelectedDate.Value
        leaveapplication.ToDuration = (cmbLeaveToDuration.SelectedIndex + 1) * 0.5
        leaveapplication.ToComment = ""

        leaveapplication.LeaveType = cmbLeaveType.Text
        leaveapplication.NumberOfDays = DateDiff(DateInterval.Day, leaveFromPicker.SelectedDate.Value, leaveToPicker.SelectedDate.Value) + 1

        dataTable.Rows.Add(leaveapplication)
        If tblLeaveApplicationAdapter.Update(dataTable) = 1 Then
            MsgBox("Successfully added application!", vbInformation)
            GetTotalCredits(employee)
            leaveApplicationsDataGrid.ItemsSource = tblLeaveApplicationAdapter.GetData.Select("EmpID = " & employee.ID)
        Else
            MsgBox("Failed to add!", vbInformation)
        End If
    End Sub
End Class
