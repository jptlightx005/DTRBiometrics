Imports DTRSystem.DTRDataSetTableAdapters
Imports DTRSystem.DTRDataSet
Class LeaveCreditsPage
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

        vcDataGrid.ItemsSource = tblLeaveCreditsAdapter.GetData.Select("EmpID = " & employee.ID & " AND (VC_Balance <> 0.00 OR VC_Earned <> 0.00 OR VC_Used <> 0.00)")
        scDataGrid.ItemsSource = tblLeaveCreditsAdapter.GetData.Select("EmpID = " & employee.ID & " AND (SC_Balance <> 0.00 OR SC_Earned <> 0.00 OR SC_Used <> 0.00)")
    End Sub

    Private Sub btnAddCredit_Click(sender As Object, e As RoutedEventArgs) Handles btnAddCredit.Click
        If ValidateFields() Then
            Dim dataTable = New LeaveCreditsTableDataTable
            Dim lctransaction As LeaveCreditsTableRow = dataTable.NewRow
            lctransaction.EmpID = employee.ID

            Dim credit = Decimal.Parse(txtCredits.Text)
            lctransaction.VC_Earned = credit
            lctransaction.VC_Used = 0
            lctransaction.VC_Balance = lctransaction.VC_Earned + vacleavecredits

            lctransaction.SC_Earned = credit
            lctransaction.SC_Used = 0
            lctransaction.SC_Balance = lctransaction.SC_Earned + sickleavecredits

            lctransaction.DateOfTransaction = Now
            lctransaction.Remarks = txtRemarks.Text

            Debug.Print("should appear: {0}", lctransaction.Remarks)
            dataTable.Rows.Add(lctransaction)

            If tblLeaveCreditsAdapter.Update(dataTable) = 1 Then
                MsgBox("Successfully added credits!", vbInformation)
                GetTotalCredits(employee)
                txtCredits.Text = 0
                txtRemarks.Text = ""
            Else
                MsgBox("Failed to add!", vbInformation)
            End If
        End If
    End Sub

    Function ValidateFields() As Boolean
        Dim d As Decimal

        If employee Is Nothing Then
            MsgBox("Select employee first!", vbExclamation)
            Return False
        End If
        If txtCredits.Text = "" Then
            MsgBox("Credits must not be empty!", vbExclamation)
            Return False
        End If
        If Not Decimal.TryParse(txtCredits.Text, d) Then
            MsgBox("Invalid decimal input!", vbExclamation)
            Return False
        End If

        If d <= 0 Then
            MsgBox("Invalid credit input! Must be more than 0", vbExclamation)
            Return False
        End If
        Return True
    End Function
End Class
