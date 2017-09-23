Imports ZKFPEngXControl
Imports System.IO
Imports System.Media
Imports System.Threading
Imports System.Windows.Threading

Imports DTRSystem.DTRDataSet
Public Class DTRBiometricWindow
    Dim WithEvents fp As ZKFPEngX
    Dim fpHandle As Integer
    Dim idList As List(Of Integer)
    Public otemplate As Object

    Dim employeeFound As EmployeeFullRow

    Dim dateTimer As DispatcherTimer
    Dim resetTimer As DispatcherTimer

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        fp = New ZKFPEngX
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
        fp.SensorIndex = 0
        If (fp.InitEngine = 0) Then
            imgConnected.Visibility = Windows.Visibility.Visible
            lblConnected.Content = "Device Connected"
        Else
            imgConnected.Visibility = Windows.Visibility.Hidden
            lblConnected.Content = "Device Not Connected"
        End If
        If fp.IsRegister Then
            fp.CancelEnroll()
        End If

        fpHandle = fp.CreateFPCacheDB

        Dim i = 0
        For Each row In tblEmployeeFullAdapter.GetData

            Dim fileName = String.Format(applicationPath & "\fptemp{0}.tpl", row.ID)
            File.WriteAllBytes(fileName, row.biometric)

            If fp.AddRegTemplateFileToFPCacheDB(fpHandle, i, fileName) = 1 Then
                Debug.Print("Succesfully loaded {0}'s biometric...", row.first_name)
                idList.Add(row.ID)
                i += 1
            End If
            File.Delete(fileName)
        Next
    End Sub

    Private Sub btnClose_Click(sender As Object, e As RoutedEventArgs) Handles btnClose.Click
        Me.Close()
        dtrMainWindow.Show()
        dtrMainWindow.Focus()
    End Sub

    Private Sub fp_OnCapture(ByVal ActionResult As Boolean, ByVal atemplate As Object) Handles fp.OnCapture
        Dim sTemp As Object
        Dim ProcessNum As Long
        sTemp = fp.GetTemplate

        Dim score = 8
        'Selects the ID of the Employee in the ID Lists
        Dim fi = fp.IdentificationInFPCacheDB(fpHandle, sTemp, score, ProcessNum)

        Dim beep As New Thread(Sub()
                                   Console.Beep(750, 500)
                                   If fi = -1 Then
                                       My.Computer.Audio.Play("Resources\pls_try_again.wav", AudioPlayMode.Background)
                                   Else
                                       My.Computer.Audio.Play("Resources\thank_you.wav", AudioPlayMode.Background)
                                   End If

                               End Sub)
        beep.Start()

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


                    If Now.TimeOfDay >= New TimeSpan(6, 0, 0) And Now.TimeOfDay < New TimeSpan(13, 0, 0) Then
                        If IsDBNull(timeLogFound("TimeInAM")) Then
                            timeLogFound.TimeInAM = DateTime.Now
                        Else
                            If Date.Compare(Now, timeLogFound.TimeInAM.AddMinutes(30)) < 0 Then
                                lblMessage.Content = String.Format("You have just logged in at {0}", timeLogFound("TimeInAM"))
                                lblMessage.Content &= vbCrLf & "Wait for 30 minutes."
                            Else
                                If IsDBNull(timeLogFound("TimeOutAM")) Then
                                    timeLogFound.TimeOutAM = DateTime.Now

                                    'TimeCalculation
                                    'If Not IsDBNull(timeLogFound("timeinpm")) Then
                                    '    Dim total As TimeSpan = timeLogFound.TimeOutPM - timeLogFound.TimeInPM
                                    '    timeLogFound.TotalTime += total.TotalHours
                                    'End If
                                Else
                                    lblMessage.Content = String.Format("You have have already logged out at {0}", timeLogFound("TimeOutAM"))
                                End If

                            End If
                        End If

                        'PM IN 1PM-5PM
                    ElseIf Now.TimeOfDay >= New TimeSpan(13, 0, 0) And Now.TimeOfDay < New TimeSpan(21, 0, 0) Then
                        If IsDBNull(timeLogFound("TimeInPM")) Then
                            timeLogFound.TimeInPM = DateTime.Now
                        Else
                            If Date.Compare(Now, timeLogFound.TimeInPM.AddMinutes(30)) < 0 Then
                                lblMessage.Content = String.Format("You have just logged in at {0}", timeLogFound("TimeInAM"))
                                lblMessage.Content &= vbCrLf & "Wait for 30 minutes."
                            Else
                                If IsDBNull(timeLogFound("TimeOutPM")) Then
                                    timeLogFound.TimeOutPM = DateTime.Now

                                    'TimeCalculation
                                    'If Not IsDBNull(timeLogFound("timeinpm")) Then
                                    '    Dim total As TimeSpan = timeLogFound.TimeOutPM - timeLogFound.TimeInPM
                                    '    timeLogFound.TotalTime += total.TotalHours
                                    'End If
                                Else
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
        fp.CancelCapture()
        fp.EndEngine()
    End Sub
End Class
