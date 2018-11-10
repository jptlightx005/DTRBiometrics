Class AdminPage
    Dim deptPage As DepartmentPage
    Dim desgPage As DesignationPage
    Dim salGPage As SalaryGradePage
    Dim activePage As Page
    Private Sub ReplacePage(ByRef pg As Page)
        activePage = pg
        adminFrame.Content = pg
    End Sub

    Public Sub setAccess()
        treeDept.IsEnabled = isHR()
        treeDesg.IsEnabled = isHR()
        treeDept.IsEnabled = isHR()
    End Sub
    Private Sub treeDept_Selected(sender As Object, e As RoutedEventArgs) Handles treeDept.Selected
        If deptPage Is Nothing Then
            deptPage = New DepartmentPage
        End If

        ReplacePage(deptPage)
    End Sub

    Private Sub treeDesg_Selected(sender As Object, e As RoutedEventArgs) Handles treeDesg.Selected
        If desgPage Is Nothing Then
            desgPage = New DesignationPage
        End If

        desgPage.PageDidAppear()
        ReplacePage(desgPage)
    End Sub

    'Private Sub treeSalG_Selected(sender As Object, e As RoutedEventArgs) Handles treeSalG.Selected
    '    If salGPage Is Nothing Then
    '        salGPage = New SalaryGradePage
    '    End If

    '    ReplacePage(salGPage)
    'End Sub
End Class
