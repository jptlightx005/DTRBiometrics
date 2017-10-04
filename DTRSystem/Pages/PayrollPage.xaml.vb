Class PayrollPage
    Dim salaryPage As SalaryPage
    'Dim payrollPage As PayrollPage
    Dim activePage As Page
    Private Sub ReplacePage(ByRef pg As Page)
        activePage = pg
        payrollFrame.Content = pg
    End Sub
    Private Sub treeSalary_Selected(sender As Object, e As RoutedEventArgs) Handles treeSalary.Selected
        If salaryPage Is Nothing Then
            salaryPage = New SalaryPage
        End If
        ReplacePage(salaryPage)
    End Sub

End Class
