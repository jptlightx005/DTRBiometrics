Class MainWindow 
    Dim registrationTab As TabItem
    Dim employeeTab As TabItem
    Dim reportsTab As TabItem
    Dim leaveTab As TabItem
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

    Private Sub tlbrEmployee_Click(sender As Object, e As RoutedEventArgs) Handles tlbrEmployee.Click
        If employeeTab Is Nothing Then
            Dim empPage As New EmployeePage
            employeeTab = NewTab(empPage, "Employee")
        End If
        tab_panels.SelectedItem = employeeTab
    End Sub

    Private Sub tlbrReports_Click(sender As Object, e As RoutedEventArgs)
        If reportsTab Is Nothing Then
            reportsTab = NewTab(New ReportsPage, "Reports")
        End If
        tab_panels.SelectedItem = reportsTab
    End Sub


    Private Sub tblrLeave_Click(sender As Object, e As RoutedEventArgs)
        If leaveTab Is Nothing Then
            leaveTab = NewTab(New LeaveManagementpage, "Leave")
        End If
        tab_panels.SelectedItem = leaveTab
    End Sub
    Private Function NewTab(ByRef pg As Page, header As String) As TabItem
        Dim tab As New TabItem
        tab.Header = header

        Dim frm As New Frame
        frm.Content = pg
        frm.NavigationUIVisibility = NavigationUIVisibility.Hidden
        tab.Content = frm

        tab_panels.Items.Add(tab)

        Return tab
    End Function
End Class
