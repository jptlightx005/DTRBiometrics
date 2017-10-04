Class ReportsPage

    Dim dtrPage As DTRReportsPage
    'Dim payrollPage As PayrollPage
    Dim activePage As Page
    Private Sub ReplacePage(ByRef pg As Page)
        activePage = pg
        reportsFrame.Content = pg
    End Sub
    Private Sub treeDTR_Selected(sender As Object, e As RoutedEventArgs) Handles treeDTR.Selected
        If dtrPage Is Nothing Then
            dtrPage = New DTRReportsPage
        End If
        ReplacePage(dtrPage)
    End Sub

    Private Sub treePayroll_Selected(sender As Object, e As RoutedEventArgs) Handles treePayroll.Selected

    End Sub
End Class
