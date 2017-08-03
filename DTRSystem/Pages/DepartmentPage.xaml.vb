Imports DTRSystem.DTRDataSetTableAdapters
Imports DTRSystem.DTRDataSet
Class DepartmentPage
    Dim dataTable As New DepartmentTableDataTable
    Dim departmentRow As DepartmentTableRow
    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        departmentRow = dataTable.NewRow
        departmentRow.dept_name = ""
        gridDepartment.DataContext = departmentRow

        departmentDataGrid.ItemsSource = tblDeptAdapter.GetData
    End Sub


    Private Sub txtDepartmentName_TextChanged(sender As Object, e As TextChangedEventArgs) Handles txtDepartmentName.TextChanged
        btnAdd.IsEnabled = txtDepartmentName.Text.Length > 0
    End Sub

    Private Sub btnAdd_Click(sender As Object, e As RoutedEventArgs) Handles btnAdd.Click
        dataTable.Rows.Add(departmentRow)
        txtDepartmentName.Text = ""
    End Sub
End Class
