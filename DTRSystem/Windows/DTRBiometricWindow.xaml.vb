Imports ZKFPEngXControl
Imports System.IO
Imports System.Media
Imports System.Threading
Imports System.Windows.Threading

Imports DTRSystem.DTRDataSet
Imports DTRSystem.DTRDataSetTableAdapters
Public Class DTRBiometricWindow
    Dim WithEvents fpscanner As ZKFPEngX
    Dim fpHandle As Integer
    Dim idList As List(Of Integer)
    Public otemplate As Object

    Dim employeeFound As EmployeeFullRow

    Dim dateTimer As DispatcherTimer
    Dim resetTimer As DispatcherTimer

    Dim vacleavecredits As Double = 0
    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        fpscanner = New ZKFPEngX

        idList = New List(Of Integer)
        dateTimer = New DispatcherTimer
        AddHandler dateTimer.Tick, AddressOf dateTimer_Tick
        dateTimer.Interval = New TimeSpan(0, 0, 1)
        dateTimer.Start()

        resetTimer = New DispatcherTimer
        AddHandler resetTimer.Tick, AddressOf resetTimer_Tick
        resetTimer.Interval = New TimeSpan(0, 0, 5)

        lblEmpName.Content = ""
        lblDepartment.Content = ""
        lblDesignation.Content = ""
        lblMessage.Content = ""
    End Sub

    Private Sub Window_Loaded(sender As Object, e As RoutedEventArgs)
        Dim a = fpscanner.InitEngine
        If a = 0 Then
            imgConnected.Visibility = Windows.Visibility.Visible
            lblConnected.Content = "Device Connected"
        Else
            imgConnected.Visibility = Windows.Visibility.Hidden
            lblConnected.Content = "Device Not Connected"
        End If

        

        If fpscanner.IsRegister Then
            Debug.Print("Is a register")
            fpscanner.CancelEnroll()
        End If

        fpHandle = fpscanner.CreateFPCacheDB

        Dim i = 0
        For Each row In tblEmployeeFullAdapter.GetData

            Dim fileName = String.Format(applicationPath & "\fptemp{0}.tpl", row.ID)
            File.WriteAllBytes(fileName, row.biometric)

            If fpscanner.AddRegTemplateFileToFPCacheDB(fpHandle, i, fileName) = 1 Then
                Debug.Print("Succesfully loaded {0}'s biometric...", row.first_name)
                idList.Add(row.ID)
                i += 1
            End If
            File.Delete(fileName)
        Next
    End Sub
    Public Function GetTotalCredits(emp As EmployeeFullRow) As Double
        Dim leaveCreditTransactions = tblLeaveCreditsAdapter.GetData.Select("EmpID = " & emp.ID)
        vacleavecredits = 0
        For Each lctransac As LeaveCreditsTableRow In leaveCreditTransactions
            Dim vc_earned = lctransac.VC_Earned
            Dim vc_used = lctransac.VC_Used
            Dim vc_bal = lctransac.VC_Balance

            vacleavecredits += vc_earned - vc_used
            Debug.Print("Is v equal {0} == {1}", vacleavecredits, vc_bal)
        Next
        Return vacleavecredits
    End Function
    Private Sub btnClose_Click(sender As Object, e As RoutedEventArgs) Handles btnClose.Click
        Me.Hide()
        dtrMainWindow.Show()
        dtrMainWindow.Focus()
    End Sub

    Private Sub fp_OnCapture(ByVal ActionResult As Boolean, ByVal atemplate As Object) Handles fpscanner.OnCapture
        Debug.Print("is capturing from dtr biometrics window")
        If isRegisteringFingerprint Then
            Debug.Print("is registering")
            Return
        End If
        Dim sTemp As Object
        Dim ProcessNum As Long
        sTemp = fpscanner.GetTemplate

        Dim score = 8
        'Selects the ID of the Employee in the ID Lists
        Dim fi = fpscanner.IdentificationInFPCacheDB(fpHandle, sTemp, score, ProcessNum)

        Dim success As Boolean = (fi <> -1)

        If fi = -1 Then 'if employee is not found
            lblEmpName.Content = ""
            lblMessage.Content = ""
            lblDepartment.Content = ""
            lblDesignation.Content = ""
            txbStatus.Text = "Not registered"
            lblMessage.Content = "You are not registered!"
            imgEmployee.Source = New BitmapImage(New Uri("pack://siteoforigin:,,,/Resources/placeholder.png", UriKind.Absolute))
        Else
            Dim filter = String.Format("ID = {0}", idList(fi)) 'actual ID of the employee
            Dim rows = tblEmployeeFullAdapter.GetData().Select(filter) '

            If rows.Count > 0 Then
                employeeFound = rows(0)
            End If

            Debug.Print("Found {0}'s record!", employeeFound.first_name)

            If Not employeeFound Is Nothing Then
                lblMessage.Content = String.Format("You have logged in at {0}", DateTime.Now.ToString("hh:mm:ss tt"))

                Dim logRows = tblLogAdapter.GetTimeLog(employeeFound.ID, Now.Date).Rows 'looks for employee's log on the day
                Dim timeLogFound As DTRDataSet.TimelogTableRow
                If logRows.Count <= 0 Then
                    tblLogAdapter.Insert(employeeFound.ID, Now.Date, Nothing, Nothing, Nothing, Nothing, 0)
                    timeLogFound = tblLogAdapter.GetTimeLog(employeeFound.ID, Now.Date).Rows(0)
                Else
                    timeLogFound = logRows(0)
                End If

                'AM IN 7AM-12PM
                If Not timeLogFound Is Nothing Then


                    If Now.TimeOfDay >= New TimeSpan(6, 0, 0) And Now.TimeOfDay < New TimeSpan(12, 30, 0) Then
                        If IsDBNull(timeLogFound("TimeInAM")) Then
                            timeLogFound.TimeInAM = DateTime.Now
                        Else
                            If Date.Compare(Now, timeLogFound.TimeInAM.AddMinutes(30)) < 0 Then
                                success = False
                                lblMessage.Content = String.Format("You have just logged in at {0}", timeLogFound("TimeInAM"))
                                lblMessage.Content &= vbCrLf & "Wait for 30 minutes."
                            Else
                                If IsDBNull(timeLogFound("TimeOutAM")) Then
                                    timeLogFound.TimeOutAM = DateTime.Now

                                    'TimeCalculation
                                    If Not IsDBNull(timeLogFound("TimeInAM")) Then
                                        Dim timeBegin = timeLogFound.TimeInAM
                                        If timeBegin.TimeOfDay < New TimeSpan(8, 0, 0) Then
                                            Dim a = Now.ToString("dd/MM/yyyy")
                                            a = a + " 08:00 AM" ' 10/04/2017 08:00 AM
                                            timeBegin = DateTime.Parse(a)
                                        End If

                                        Dim timeEnd = Now
                                        If Now.TimeOfDay > New TimeSpan(17, 0, 0) Then
                                            Dim a = Now.ToString("dd/MM/yyyy")
                                            a = a + " 12:00 PM" ' 10/04/2017 12:00 PM
                                            timeEnd = DateTime.Parse(a)
                                        End If
                                        Debug.Print("AM: Begin: {0}", timeBegin.ToString)
                                        Debug.Print("AM: End: {0}", timeEnd.ToString)
                                        Dim total As TimeSpan = timeEnd - timeBegin
                                        timeLogFound.TotalTime += total.TotalMinutes

                                        If total.TotalMinutes < 240 Then
                                            'calculate leave
                                            Dim lateMinutes As TimeSpan = New TimeSpan(0, 240, 0) - total

                                            Dim lcDataTable = New LeaveCreditsTableDataTable
                                            Dim lctransaction As LeaveCreditsTableRow = lcDataTable.NewRow

                                            Dim creditsToBeFuckingDeducted = Math.Round(lateMinutes.TotalMinutes) / 480
                                            Debug.Print("{0} / 480", Math.Round(lateMinutes.TotalMinutes))
                                            Debug.Print("Credits to be deducted {0}", creditsToBeFuckingDeducted)
                                            lctransaction.EmpID = employeeFound.ID
                                            lctransaction.Remarks = "Late"
                                            lctransaction.VC_Earned = 0
                                            lctransaction.VC_Used = creditsToBeFuckingDeducted
                                            lctransaction.VC_Balance = GetTotalCredits(employeeFound) - creditsToBeFuckingDeducted

                                            lctransaction.SC_Earned = 0
                                            lctransaction.SC_Used = 0
                                            lctransaction.SC_Balance = 0

                                            lctransaction.DateOfTransaction = Now

                                            lcDataTable.Rows.Add(lctransaction)

                                            If tblLeaveCreditsAdapter.Update(lcDataTable) = 1 Then
                                                Debug.Print("Successfully added credits!", vbInformation)
                                            Else
                                                Debug.Print("Failed to add!", vbInformation)
                                            End If
                                        End If
                                    End If
                                Else
                                    success = False
                                    lblMessage.Content = String.Format("You have have already logged out at {0}", timeLogFound("TimeOutAM"))
                                End If
                            End If
                        End If

                        'PM IN 1PM-5PM
                    ElseIf Now.TimeOfDay >= New TimeSpan(12, 30, 0) And Now.TimeOfDay < New TimeSpan(21, 0, 0) Then
                        If IsDBNull(timeLogFound("TimeInPM")) Then
                            timeLogFound.TimeInPM = DateTime.Now
                        Else
                            If Date.Compare(Now, timeLogFound.TimeInPM.AddMinutes(30)) < 0 Then
                                success = False
                                lblMessage.Content = String.Format("You have just logged in at {0}", timeLogFound("TimeInAM"))
                                lblMessage.Content &= vbCrLf & "Wait for 30 minutes."
                            Else
                                If IsDBNull(timeLogFound("TimeOutPM")) Then
                                    timeLogFound.TimeOutPM = DateTime.Now

                                    'TimeCalculation
                                    If Not IsDBNull(timeLogFound("TimeInPM")) Then
                                        Dim timeBegin = timeLogFound.TimeInPM
                                        If timeBegin.TimeOfDay < New TimeSpan(13, 0, 0) Then
                                            Dim a = Now.ToString("dd/MM/yyyy")
                                            a = a + " 01:00 PM" ' 10/04/2017 01:00 PM
                                            timeBegin = DateTime.ParseExact(a, "dd/MM/yyyy hh:mm tt", Nothing)
                                        End If

                                        Dim timeEnd = Now
                                        If Now.TimeOfDay > New TimeSpan(17, 0, 0) Then
                                            Dim a = Now.ToString("dd/MM/yyyy")
                                            a = a + " 05:00 PM" ' 10/04/2017 05:00 PM

                                            Debug.Print("Nice try: {0}", a)
                                            timeEnd = DateTime.ParseExact(a, "dd/MM/yyyy hh:mm tt", Nothing)
                                        End If
                                        Debug.Print("PM: Begin: {0}", timeBegin.ToString)
                                        Debug.Print("PM: End: {0}", timeEnd.ToString)
                                        Dim total As TimeSpan = timeEnd - timeBegin
                                        Debug.Print("Test: {0}", total.TotalMinutes)
                                        Debug.Print("Meh: {0}", total)

                                        Dim minutes = 240
                                        If timeLogFound.TotalTime = 0 Then
                                            minutes = 480
                                        End If
                                        timeLogFound.TotalTime += total.TotalMinutes

                                        If total.TotalMinutes < 240 Then
                                            'calculate leave

                                            Dim lateMinutes As TimeSpan = New TimeSpan(0, minutes, 0) - total

                                            Dim lcDataTable = New LeaveCreditsTableDataTable
                                            Dim lctransaction As LeaveCreditsTableRow = lcDataTable.NewRow

                                            Dim creditsToBeFuckingDeducted = Math.Round(lateMinutes.TotalMinutes) / 480
                                            Debug.Print("{0} / 480", Math.Round(lateMinutes.TotalMinutes))
                                            Debug.Print("Credits to be deducted {0}", creditsToBeFuckingDeducted)
                                            Debug.Print("Employee has: {0} credits", GetTotalCredits(employeeFound))
                                            lctransaction.EmpID = employeeFound.ID
                                            lctransaction.Remarks = "Late"
                                            lctransaction.VC_Earned = 0
                                            lctransaction.VC_Used = creditsToBeFuckingDeducted
                                            lctransaction.VC_Balance = GetTotalCredits(employeeFound) - creditsToBeFuckingDeducted

                                            lctransaction.SC_Earned = 0
                                            lctransaction.SC_Used = 0
                                            lctransaction.SC_Balance = 0

                                            lctransaction.DateOfTransaction = Now

                                            lcDataTable.Rows.Add(lctransaction)

                                            If tblLeaveCreditsAdapter.Update(lcDataTable) = 1 Then
                                                Debug.Print("Successfully added credits!", vbInformation)
                                            Else
                                                Debug.Print("Failed to add!", vbInformation)
                                            End If
                                        End If
                                    End If
                                Else
                                    success = False
                                    lblMessage.Content = String.Format("You have have already logged out at {0}", timeLogFound("TimeOutPM"))
                                End If
                            End If
                        End If

                    End If
                    tblLogAdapter.Update(timeLogFound)
                End If

                txbStatus.Text = "Record Found"
                lblEmpName.Content = Coalesce(employeeFound("full_name"))
                lblDepartment.Content = Coalesce(employeeFound("dept_name"))
                lblDesignation.Content = Coalesce(employeeFound("designation_name"))
                imgEmployee.Source = DataToBitmap(employeeFound.picture)

                File.WriteAllBytes(applicationPath & "\employee.jpg", employeeFound.picture)
            End If
        End If

        Dim beep As New Thread(Sub()
                                   Console.Beep(My.Settings.BeepFrequency, My.Settings.BeepDuration)
                                   If success Then
                                       My.Computer.Audio.Play(My.Settings.SuccessMessageAudioFile, AudioPlayMode.Background)
                                   Else
                                       My.Computer.Audio.Play(My.Settings.FailedMessageAudioFile, AudioPlayMode.Background)
                                   End If

                               End Sub)
        beep.Start()

        resetTimer.Stop()
        resetTimer.Start()
    End Sub
    Private Sub dateTimer_Tick(sender As Object, e As EventArgs)
        lblTime.Content = DateTime.Now.ToString("MMMM dd, yyyy hh:mm:ss tt")
    End Sub

    Private Sub resetTimer_Tick(sender As Object, e As EventArgs)
        lblEmpName.Content = ""
        txbStatus.Text = "Waiting"
        lblMessage.Content = ""
        lblDepartment.Content = ""
        lblDesignation.Content = ""
        imgEmployee.Source = New BitmapImage(New Uri("pack://siteoforigin:,,,/Resources/placeholder.png", UriKind.Absolute))
        resetTimer.Stop()
    End Sub

    Private Sub Window_Closed(sender As Object, e As EventArgs)
        fpscanner.CancelCapture()
        fpscanner.EndEngine()
    End Sub

    Private Sub Window_MouseDown(sender As Object, e As MouseButtonEventArgs)
        If e.ChangedButton = MouseButton.Left Then
            Me.DragMove()
        End If
    End Sub

    Private Sub Window_Closing(sender As Object, e As ComponentModel.CancelEventArgs)
        e.Cancel = True
        Me.Hide()
    End Sub
End Class
