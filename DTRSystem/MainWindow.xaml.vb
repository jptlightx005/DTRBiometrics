Imports SMSCSFuncs
Imports System.Windows.Forms

Class MainWindow
    Dim employeeTab As TabItem
    Dim payrollTab As TabItem
    Dim reportsTab As TabItem
    Dim leaveTab As TabItem
    Dim adminTab As TabItem

    Dim reportsPage As ReportsPage
    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        isRegisteringFingerprint = False

    End Sub

    Public Sub SetAccess()
        tlbrAdmin.IsEnabled = isHR()
        tlbrEmployee.IsEnabled = isHR()
        tlbrLeave.IsEnabled = isHR()

        Me.Title = acctLogged.type
    End Sub
    Private Sub btnDTR_Click(sender As Object, e As RoutedEventArgs) Handles btnDTR.Click
        If dtrBioWindow Is Nothing Then
            dtrBioWindow = New DTRBiometricWindow
        End If

        If System.Windows.Forms.Screen.AllScreens.Length > 1 Then
            Dim s2 = Screen.AllScreens(1)
            Dim r2 = s2.WorkingArea
            dtrBioWindow.Top = r2.Top
            dtrBioWindow.Left = r2.Left
            dtrBioWindow.Width = r2.Width
            dtrBioWindow.Height = r2.Height
        End If

        If System.Windows.Forms.Screen.AllScreens.Length = 1 Then
            dtrBioWindow.WindowStyle = Windows.WindowStyle.SingleBorderWindow
            dtrBioWindow.Width = Screen.PrimaryScreen.WorkingArea.Width
            dtrBioWindow.Height = Screen.PrimaryScreen.WorkingArea.Height
        End If

        dtrBioWindow.Show()
        dtrBioWindow.WindowState = Windows.WindowState.Maximized

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
        Debug.Print("reports tab is nothing: {0}", reportsTab)
        If reportsTab Is Nothing Then
            Debug.Print("reports tab is nothing: YES")
            If reportsPage Is Nothing Then
                reportsPage = New ReportsPage
            End If
            reportsTab = NewTab(reportsPage, "Reports")
        End If
        Debug.Print("reports tab is nothing: {0}", reportsTab)
        reportsPage.SetAccess()
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

    Private Sub Window_Closed(sender As Object, e As EventArgs)
        If Not dtrBioWindow Is Nothing Then
            Debug.Print("will close..")
            dtrBioWindow.Close()
            End
        End If
    End Sub

    Private Sub mnu_settings_Click(sender As Object, e As RoutedEventArgs) Handles mnu_settings.Click
        Dim settingsWindow As New SettingsWindow
        settingsWindow.ShowDialog()

    End Sub

    Private Sub mnu_close_Click(sender As Object, e As RoutedEventArgs) Handles mnu_close.Click
        Me.Close()
        End
    End Sub


    Private Sub mnu_logout_Click(sender As Object, e As RoutedEventArgs) Handles mnu_logout.Click
        acctLogged = Nothing
        dtrLoginWindow = New LoginWindow

        If tab_panels.Items.Count > 0 Then
            For i = tab_panels.Items.Count - 1 To 0 Step -1
                Dim tabItem = tab_panels.Items(i)
                Tab_Closed(tabItem, e)
                tab_panels.Items.RemoveAt(i)
            Next
        End If

        Me.Hide()
        dtrLoginWindow.Show()
    End Sub
End Class
