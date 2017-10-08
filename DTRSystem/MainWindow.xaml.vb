
Imports SMSCSFuncs
Imports System.Windows.Forms

Class MainWindow
    Dim employeeTab As TabItem
    Dim payrollTab As TabItem
    Dim reportsTab As TabItem
    Dim leaveTab As TabItem
    Dim adminTab As TabItem

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

    End Sub
    Private Sub btnDTR_Click(sender As Object, e As RoutedEventArgs) Handles btnDTR.Click
        Dim dtrWindow As New DTRBiometricWindow

        If System.Windows.Forms.Screen.AllScreens.Length > 1 Then
            Dim s2 = Screen.AllScreens(1)
            Dim r2 = s2.WorkingArea
            dtrWindow.Top = r2.Top
            dtrWindow.Left = r2.Left
            dtrWindow.Width = r2.Width
            dtrWindow.Height = r2.Height
        End If

        If System.Windows.Forms.Screen.AllScreens.Length = 1 Then
            dtrWindow.WindowStyle = Windows.WindowStyle.SingleBorderWindow
            dtrWindow.Width = Screen.PrimaryScreen.WorkingArea.Width
            dtrWindow.Height = Screen.PrimaryScreen.WorkingArea.Height
        End If

        dtrWindow.Show()
        dtrWindow.WindowState = Windows.WindowState.Maximized
        
        'Me.Hide()
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
            Dim leavePage As New LeavePage
            leaveTab = NewTab(leavePage, "Leave")
        End If
        tab_panels.SelectedItem = leaveTab
    End Sub

    Private Sub tlbrAdmin_Click(sender As Object, e As RoutedEventArgs)
        If adminTab Is Nothing Then
            Dim adminPage As New AdminPage
            adminTab = NewTab(adminPage, "Admin")
        End If
        tab_panels.SelectedItem = adminTab
    End Sub

    Private Sub tlbrPayroll_Click(sender As Object, e As RoutedEventArgs)
        If payrollTab Is Nothing Then
            payrollTab = NewTab(New PayrollPage, "Payroll")
        End If
        tab_panels.SelectedItem = payrollTab
    End Sub
    Private Function NewTab(ByRef pg As Page, header As String) As ClosableTab
        Dim tab As New SMSCSFuncs.ClosableTab
        tab.Title = header
        AddHandler tab.OnTabClose, AddressOf Me.Tab_Closed

        Dim frm As New Frame
        frm.Content = pg
        frm.NavigationUIVisibility = NavigationUIVisibility.Hidden
        tab.Content = frm

        tab_panels.Items.Add(tab)

        Return tab
    End Function
    Sub Tab_Closed(sender As Object, e As EventArgs)
        If sender Is employeeTab Then
            employeeTab = Nothing
        ElseIf sender Is payrollTab Then
            payrollTab = Nothing
        ElseIf sender Is reportsTab Then
            reportsTab = Nothing
        ElseIf sender Is leaveTab Then
            leaveTab = Nothing
        ElseIf sender Is adminTab Then
            adminTab = Nothing
        Else
            Debug.Print("Tab not found!")
        End If
     
    End Sub

    Private Sub Window_Loaded(sender As Object, e As RoutedEventArgs)
        Debug.Print("Number of screens: " & System.Windows.Forms.Screen.AllScreens.Length)

    End Sub
End Class
