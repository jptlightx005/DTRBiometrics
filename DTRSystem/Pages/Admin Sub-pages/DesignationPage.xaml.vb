Imports DTRSystem.DTRDataSetTableAdapters
Imports DTRSystem.DTRDataSet
Class DesignationPage
    Dim designationRow As DesignationTableRow
    Dim dataSet As New DTRSystem.DTRDataSet
    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        designationDataGrid.DataContext = dataSet.DesignationFullTable.DefaultView
        tblDesgFullAdapter.Fill(dataSet.DesignationFullTable)

        designationRow = dataSet.DesignationTable.NewRow
        designationRow.designation_name = ""
        gridDesignation.DataContext = designationRow
    End Sub

    Private Sub btnAdd_Click(sender As Object, e As RoutedEventArgs) Handles btnAdd.Click
        If cmbDepartment.SelectedIndex >= 0 Then
            Try
                designationRow.DeptID = cmbDepartment.SelectedValue
                dataSet.DesignationTable.Rows.Add(designationRow)
                If tblDesgAdapter.Update(designationRow) = 1 Then
                    tblDesgFullAdapter.Fill(dataSet.DesignationFullTable)
                    MsgBox("Successfully added!", vbInformation)
                Else
                    MsgBox("Failed to add!", vbInformation)
                End If
            Catch ex As Exception

            End Try
            designationRow = dataSet.DesignationTable.NewRow
            designationRow.designation_name = ""
            gridDesignation.DataContext = designationRow
        Else
            MsgBox("Select a department first!", MsgBoxStyle.Exclamation)
        End If
        
    End Sub

    Private Sub btnSave_Click(sender As Object, e As RoutedEventArgs) Handles btnSave.Click
        tblDesgAdapter.Update(dataSet.DesignationTable)
        btnSave.IsEnabled = False
        btnCancel.IsEnabled = False
    End Sub

    Private Sub btnRemove_Click(sender As Object, e As RoutedEventArgs) Handles btnRemove.Click
        Dim selectedRow As DesignationTableRow = dataSet.DesignationTable.Rows(designationDataGrid.SelectedIndex)
        tblDesgAdapter.Delete(selectedRow.ID, selectedRow.DeptID, selectedRow.designation_name)
        designationDataGrid.DataContext = dataSet.DesignationTable.DefaultView
        tblDesgAdapter.Fill(dataSet.DesignationTable)
        MsgBox("Successfully removed!", vbInformation)
    End Sub

    Private Sub btnCancel_Click(sender As Object, e As RoutedEventArgs) Handles btnCancel.Click
        designationDataGrid.DataContext = dataSet.DesignationTable.DefaultView
        tblDesgAdapter.Fill(dataSet.DesignationTable)
        btnSave.IsEnabled = False
        btnCancel.IsEnabled = False
    End Sub

    Private Sub designationDataGrid_RowEditEnding(sender As Object, e As DataGridRowEditEndingEventArgs) Handles designationDataGrid.RowEditEnding
        btnSave.IsEnabled = True
        btnCancel.IsEnabled = True
    End Sub

    Private Sub designationDataGrid_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) Handles designationDataGrid.SelectionChanged
        btnRemove.IsEnabled = designationDataGrid.SelectedIndex >= 0
    End Sub

    Private Sub txtDesignationName_TextChanged(sender As Object, e As TextChangedEventArgs) Handles txtDesignationName.TextChanged
        btnAdd.IsEnabled = txtDesignationName.Text.Length > 0
    End Sub

    Private Sub desgPage_Initialized(sender As Object, e As EventArgs) Handles desgPage.Initialized
        cmbDepartment.ItemsSource = tblDeptAdapter.GetData
    End Sub
End Class
