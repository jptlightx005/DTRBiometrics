Imports DTRSystem.DTRDataSetTableAdapters
Imports DTRSystem.DTRDataSet
Class DepartmentPage
    Dim departmentRow As DepartmentTableRow
    Dim dataSet As New DTRSystem.DTRDataSet
    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        departmentDataGrid.DataContext = dataSet.DepartmentTable.DefaultView
        tblDeptAdapter.Fill(dataSet.DepartmentTable)

        departmentRow = dataSet.DepartmentTable.NewRow
        departmentRow.dept_name = ""
        gridDepartment.DataContext = departmentRow
    End Sub


    Private Sub txtDepartmentName_TextChanged(sender As Object, e As TextChangedEventArgs) Handles txtDepartmentName.TextChanged
        btnAdd.IsEnabled = txtDepartmentName.Text.Length > 0
    End Sub

    Private Sub btnAdd_Click(sender As Object, e As RoutedEventArgs) Handles btnAdd.Click
        Try
            dataSet.DepartmentTable.Rows.Add(departmentRow)
            If tblDeptAdapter.Update(departmentRow) = 1 Then
                departmentDataGrid.DataContext = dataSet.DepartmentTable.DefaultView
                tblDeptAdapter.Fill(dataSet.DepartmentTable)
                MsgBox("Successfully added!", vbInformation)
            Else
                MsgBox("Failed to add!", vbInformation)
            End If
        Catch ex As Exception

        End Try
        departmentRow = dataSet.DepartmentTable.NewRow
        departmentRow.dept_name = ""
        gridDepartment.DataContext = departmentRow
    End Sub

    Private Sub departmentDataGrid_RowEditEnding(sender As Object, e As DataGridRowEditEndingEventArgs) Handles departmentDataGrid.RowEditEnding
        btnSave.IsEnabled = True
        btnCancel.IsEnabled = True
    End Sub

    Private Sub departmentDataGrid_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) Handles departmentDataGrid.SelectionChanged
        btnRemove.IsEnabled = departmentDataGrid.SelectedIndex >= 0
    End Sub

    Private Sub btnSave_Click(sender As Object, e As RoutedEventArgs) Handles btnSave.Click
        tblDeptAdapter.Update(dataSet.DepartmentTable)
        btnSave.IsEnabled = False
        btnCancel.IsEnabled = False
    End Sub

    Private Sub btnRemove_Click(sender As Object, e As RoutedEventArgs) Handles btnRemove.Click
        Dim selectedRow As DepartmentTableRow = dataSet.DepartmentTable.Rows(departmentDataGrid.SelectedIndex)
        tblDeptAdapter.Delete(selectedRow.ID, selectedRow.dept_name)
        departmentDataGrid.DataContext = dataSet.DepartmentTable.DefaultView
        tblDeptAdapter.Fill(dataSet.DepartmentTable)
        MsgBox("Successfully removed!", vbInformation)
    End Sub

    Private Sub btnCancel_Click(sender As Object, e As RoutedEventArgs) Handles btnCancel.Click
        departmentDataGrid.DataContext = dataSet.DepartmentTable.DefaultView
        tblDeptAdapter.Fill(dataSet.DepartmentTable)
        btnSave.IsEnabled = False
        btnCancel.IsEnabled = False
    End Sub
End Class
