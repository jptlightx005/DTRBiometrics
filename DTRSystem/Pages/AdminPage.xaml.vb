Class AdminPage
    Dim deptPage As DepartmentPage
    Dim desgPage As DesignationPage

    Dim activePage As Page
    Private Sub ReplacePage(ByRef pg As Page)
        activePage = pg
        employeeFrame.Content = pg
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

        ReplacePage(desgPage)
    End Sub
End Class
