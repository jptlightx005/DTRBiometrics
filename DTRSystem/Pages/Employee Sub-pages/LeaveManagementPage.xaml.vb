Imports DTRSystem.DTRDataSetTableAdapters
Imports DTRSystem.DTRDataSet
Class LeaveManagementpage
    Dim employee As EmployeeFullRow

    Dim vacleavecredits As Double = 0
    Dim sickleavecredits As Double = 0

    Public Sub ChangeEmployee(emp As EmployeeTableRow)
        If Not emp Is Nothing Then
            Dim employees = tblEmployeeFullAdapter.GetData.Select("ID = " & emp.ID)
            If employees.Count > 0 Then
                employee = employees(0)
                grdLeave.DataContext = employee
                deparment.Text = IIf(Not IsDBNull(employee("dept_name")), employee("dept_name"), "")
                designation.Text = IIf(Not IsDBNull(employee("designation_name")), employee("designation_name"), "")

                GetTotalCredits(employee)
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
        dateupdated.Text = date_modified
    End Sub

    Private Sub btnVCAdd_Click(sender As Object, e As RoutedEventArgs) Handles btnVCAdd.Click
        Dim dataTable = New LeaveCreditsTableDataTable
        Dim lctransaction As LeaveCreditsTableRow = dataTable.NewRow
        lctransaction.EmpID = employee.ID

        lctransaction.VC_Earned = txtVLCredits.Text
        lctransaction.VC_Used = 0
        lctransaction.VC_Balance = lctransaction.VC_Earned + vacleavecredits

        lctransaction.SC_Earned = 0
        lctransaction.SC_Used = 0
        lctransaction.SC_Balance = 0

        lctransaction.DateOfTransaction = Now

        dataTable.Rows.Add(lctransaction)

        If tblLeaveCreditsAdapter.Update(dataTable) = 1 Then
            MsgBox("Successfully added credits!", vbInformation)
            GetTotalCredits(employee)
        Else
            MsgBox("Failed to add!", vbInformation)
        End If

    End Sub

    Private Sub btnSCAdd_Click(sender As Object, e As RoutedEventArgs) Handles btnSCAdd.Click
        Dim dataTable = New LeaveCreditsTableDataTable
        Dim lctransaction As LeaveCreditsTableRow = dataTable.NewRow
        lctransaction.EmpID = employee.ID

        lctransaction.VC_Earned = 0
        lctransaction.VC_Used = 0
        lctransaction.VC_Balance = 0

        lctransaction.SC_Earned = txtSLCredits.Text
        lctransaction.SC_Used = 0
        lctransaction.SC_Balance = lctransaction.SC_Earned + sickleavecredits

        lctransaction.DateOfTransaction = Now

        dataTable.Rows.Add(lctransaction)

        If tblLeaveCreditsAdapter.Update(dataTable) = 1 Then
            MsgBox("Successfully added credits!", vbInformation)
            GetTotalCredits(employee)
        Else
            MsgBox("Failed to add!", vbInformation)
        End If
    End Sub
End Class
