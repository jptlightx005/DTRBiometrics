Imports System.Data
Class SalaryGradePage
    Dim dataSet As New DTRSystem.DTRDataSet
    Private Sub btnSave_Click(sender As Object, e As RoutedEventArgs) Handles btnSave.Click

        If tblSalaryGradeAdapter.Update(dataSet.SalaryGradeTable) = 1 Then
            MsgBox("Succesfully saved, motherfather!", MsgBoxStyle.Information)
        End If
    End Sub

    Private Sub DidChangeDataTable(ByVal sender As Object, ByVal e As DataRowChangeEventArgs)
        Debug.Print("Changed...")

    End Sub
    Private Sub Page_Initialized(sender As Object, e As EventArgs)
        dataGridSalaryGrade.DataContext = dataSet.SalaryGradeTable.DefaultView
        tblSalaryGradeAdapter.Fill(dataSet.SalaryGradeTable)

        AddHandler dataSet.SalaryGradeTable.RowChanged, AddressOf DidChangeDataTable
    End Sub
End Class
