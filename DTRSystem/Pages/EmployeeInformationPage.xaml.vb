Imports DTRSystem.DTRBiometricDataSetTableAdapters
Imports DTRSystem.DTRBiometricDataSet
Class EmployeeInformationPage
    Dim tblAdapter As New EmployeeTableAdapter
    Dim selectedEmployee As EmployeeTableRow
    Private Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        Me.Resources.Add("selectedEmployee", selectedEmployee)
    End Sub
    Private Sub employeePage_Initialized(sender As Object, e As EventArgs) Handles employeePage.Initialized
        cmbEmployees.ItemsSource = tblAdapter.GetData
        cmbEmployees.DisplayMemberPath = "first_name"
        cmbEmployees.SelectedValuePath = "ID"
    End Sub

    Private Sub cmbEmployees_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) Handles cmbEmployees.SelectionChanged
        selectedEmployee = tblAdapter.GetData.Select(String.Format("ID = {0}", cmbEmployees.SelectedValue))(0)
        grdEmpForm.DataContext = selectedEmployee
    End Sub

    Private Sub btnSave_Click(sender As Object, e As RoutedEventArgs) Handles btnSave.Click
        If tblAdapter.Update(selectedEmployee) = 1 Then
            MsgBox("Successfully updated!", vbInformation)
        Else
            MsgBox("Failed to Update!", vbExclamation)
        End If
    End Sub
End Class
