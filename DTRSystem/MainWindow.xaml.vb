Class MainWindow 
    Dim registrationTab As TabItem
    Dim employeeTab As TabItem

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

    Private Function NewTab(ByRef pg As Page, header As String) As TabItem
        Dim tab As New TabItem
        tab.Header = header

        Dim frm As New Frame
        frm.Content = pg

        tab.Content = frm

        tab_panels.Items.Add(tab)

        Return tab
    End Function

    Private Sub tlbrEmployee_Click(sender As Object, e As RoutedEventArgs) Handles tlbrEmployee.Click
        If employeeTab Is Nothing Then
            Dim empPage As New EmployeeInformationPage
            AddHandler empPage.btnAdd.Click, AddressOf mnu_register_Click

            employeeTab = NewTab(empPage, "Employee")
        End If
        tab_panels.SelectedItem = employeeTab
    End Sub
End Class
