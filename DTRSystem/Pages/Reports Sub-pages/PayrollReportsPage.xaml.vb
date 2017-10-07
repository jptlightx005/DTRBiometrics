Imports System.Data
Imports DTRSystem.DTRDataSet
Imports DTRSystem.DTRDataSetTableAdapters
Class PayrollReportsPage
    '<DataGridTextColumn Binding="{Binding full_name}" Header="Name" />
    '<DataGridTextColumn Binding="{Binding designation}" Header="Designation" />
    '<DataGridTextColumn Binding="{Binding from, StringFormat=\{0:hh:mm tt\}}" Header="From Date" />
    '<DataGridTextColumn Binding="{Binding to, StringFormat=\{0:hh:mm tt\}}" Header="To Date" />
    '<DataGridTextColumn Binding="{Binding monthly, StringFormat=N2}" Header="Monthly Rate" />
    '<DataGridTextColumn Binding="{Binding partial, StringFormat=N2}" Header="Amount Accrued" />
    '<DataGridTextColumn Binding="{Binding tax, StringFormat=N2}" Header="Withholding Tax" />
    '<DataGridTextColumn Binding="{Binding gsis, StringFormat=N2}" Header="Government Service Insurance System" />
    '<DataGridTextColumn Binding="{Binding totaldeductions, StringFormat=N2}" Header="Total Deductions" />
    '<DataGridTextColumn Binding="{Binding netpay}" Header="Net Pay" />

    Function GetPayrollTable(deptID As Integer, startDate As Date, endDate As Date) As DataTable
        Dim employees = tblEmployeeFullAdapter.GetData.Select("dept_id = " & deptID)


        ' Create new DataTable instance.
        Dim table As New DataTable

        ' Create four typed columns in the DataTable.
        table.Columns.Add("full_name", GetType(String))
        table.Columns.Add("designation", GetType(String))
        table.Columns.Add("from", GetType(Date))
        table.Columns.Add("to", GetType(Date))
        table.Columns.Add("monthly", GetType(Decimal))
        table.Columns.Add("partial", GetType(Decimal))
        table.Columns.Add("tax", GetType(Decimal))
        table.Columns.Add("gsis", GetType(Decimal))
        table.Columns.Add("totaldeductions", GetType(Decimal))
        table.Columns.Add("netpay", GetType(Decimal))

        ' Add five rows with those columns filled in the DataTable.
        table.Rows.Add(25, "Indocin", "David", DateTime.Now)
        table.Rows.Add(50, "Enebrel", "Sam", DateTime.Now)
        table.Rows.Add(10, "Hydralazine", "Christoff", DateTime.Now)
        table.Rows.Add(21, "Combivent", "Janet", DateTime.Now)
        table.Rows.Add(100, "Dilantin", "Melanie", DateTime.Now)

        Return table
    End Function


    Private Sub btnSearch_Click(sender As Object, e As RoutedEventArgs) Handles btnSearch.Click
        If cmbDepartments.SelectedIndex >= 0 Then
            If fromPicker.SelectedDate Is Nothing And toPicker.SelectedDate Is Nothing Then
                MsgBox("Must select a start and end dates!", vbExclamation)
                Return
            End If

            GetPayrollTable(cmbDepartments.SelectedValue, fromPicker.SelectedDate, toPicker.SelectedDate)
        Else
            MsgBox("Must select a department!", vbExclamation)
        End If
    End Sub

    Private Sub Page_Initialized(sender As Object, e As EventArgs)
        cmbDepartments.ItemsSource = tblDeptAdapter.GetData
    End Sub
End Class
