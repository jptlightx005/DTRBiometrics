Class MainWindow 
    Dim registrationTab As TabItem
    Dim employeeTab As TabItem
    Dim reportsTab As TabItem
    Dim departmentTab As TabItem
    Dim designationTab As TabItem

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

    End Sub
    Private Sub btnDTR_Click(sender As Object, e As RoutedEventArgs) Handles btnDTR.Click
        Dim dtrWindow As New DTRBiometricWindow
        dtrWindow.Show()
        Me.Hide()
    End Sub

    Private Sub btnClose_Click(sender As Object, e As RoutedEventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub mnu_register_Click(sender As Object, e As RoutedEventArgs) Handles mnu_register.Click
        If registrationTab Is Nothing Then
            registrationTab = NewTab(New RegistrationPage, "Register")
        End If
        tab_panels.SelectedItem = registrationTab
    End Sub

    Private Sub tlbrEmployee_Click(sender As Object, e As RoutedEventArgs) Handles tlbrEmployee.Click
        If employeeTab Is Nothing Then
            Dim empPage As New EmployeeInformationPage
            AddHandler empPage.btnAdd.Click, AddressOf mnu_register_Click

            employeeTab = NewTab(empPage, "Employee")
        End If
        tab_panels.SelectedItem = employeeTab
    End Sub

    Private Sub tlbrReports_Click(sender As Object, e As RoutedEventArgs)
        If reportsTab Is Nothing Then
            employeeTab = NewTab(New ReportsPage, "Reports")
        End If
        tab_panels.SelectedItem = employeeTab
    End Sub
    Private Sub mnu_departments_Click(sender As Object, e As RoutedEventArgs) Handles mnu_departments.Click
        If departmentTab Is Nothing Then
            departmentTab = NewTab(New DepartmentPage, "Departments")
        End If
        tab_panels.SelectedItem = departmentTab
    End Sub

    Private Sub mnu_designation_Click(sender As Object, e As RoutedEventArgs) Handles mnu_designation.Click
        If designationTab Is Nothing Then
            designationTab = NewTab(New DesignationPage, "Designations")
        End If
        tab_panels.SelectedItem = designationTab
    End Sub
    Private Function NewTab(ByRef pg As Page, header As String) As TabItem
        Dim tab As New TabItem
        tab.Header = header

        Dim frm As New Frame
        frm.Content = pg

        tab.Content = frm

        tab_panels.Items.Add(tab)

        Return tab
    End Function
End Class
