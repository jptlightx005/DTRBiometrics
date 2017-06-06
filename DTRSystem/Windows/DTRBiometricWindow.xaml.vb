Imports ZKFPEngXControl
Imports System.IO
Imports System.Media
Imports System.Threading
Public Class DTRBiometricWindow
    Dim WithEvents fp As ZKFPEngX
    Dim fpHandle As Integer
    Dim idList As List(Of Integer)
    Public otemplate As Object

    Dim employeeFound As DTRBiometricDataSet.EmployeeTableRow


    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        fp = New ZKFPEngX
        idList = New List(Of Integer)
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
        For Each row In tblEmployeeAdapter.GetData
            idList.Add(row.ID)
            Dim fileName = String.Format(applicationPath & "\fptemp{0}.tpl", row.ID)
            File.WriteAllBytes(fileName, row.biometric)

            fp.AddRegTemplateFileToFPCacheDB(fpHandle, i, fileName)
            File.Delete(fileName)
            i += 1
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

        If fi = -1 Then
            txtEmpName.Text = ""
            txbStatus.Text = "Not registered"
            imgEmployee.Source = New BitmapImage(New Uri("pack://siteoforigin:,,,/Resources/placeholder.png", UriKind.Absolute))
        Else
            Dim filter = String.Format("ID = {0}", idList(fi))
            Dim rows = tblEmployeeAdapter.GetData().Select(filter)
            If rows.Count > 0 Then
                employeeFound = rows(0)
            End If

            If Not employeeFound Is Nothing Then
                Dim logRows = tblLogAdapter.GetTimeLog(employeeFound.ID, Now.Date).Rows
                Dim timeLogFound As DTRBiometricDataSet.TimeLogTableRow
                If logRows.Count <= 0 Then
                    tblLogAdapter.Insert(employeeFound.ID, Now.Date, Nothing, Nothing, Nothing, Nothing)
                    timeLogFound = tblLogAdapter.GetTimeLog(employeeFound.ID, Now.Date).Rows(0)
                Else
                    timeLogFound = logRows(0)
                End If

                'AM IN 7AM-12PM
                If Not timeLogFound Is Nothing Then
                    If Now.TimeOfDay >= New TimeSpan(7, 0, 0) And Now.TimeOfDay < New TimeSpan(12, 0, 0) Then
                        If IsDBNull(timeLogFound("TimeInAM")) Then
                            tblLogAdapter.UpdateInAM(Now, timeLogFound.ID)
                        End If
                        'AM Out 12PM-1PM
                    ElseIf Now.TimeOfDay >= New TimeSpan(12, 0, 0) And Now.TimeOfDay < New TimeSpan(13, 0, 0) Then
                        If IsDBNull(timeLogFound("TimeOutAM")) Then
                            tblLogAdapter.UpdateOutAM(Now, timeLogFound.ID)
                        End If
                        'PM IN 1PM-5PM
                    ElseIf Now.TimeOfDay >= New TimeSpan(13, 0, 0) And Now.TimeOfDay < New TimeSpan(17, 0, 0) Then
                        If IsDBNull(timeLogFound("TimeInPM")) Then
                            tblLogAdapter.UpdateInPM(Now, timeLogFound.ID)
                        End If
                        'PM OUT 5PM-8PM
                    ElseIf Now.TimeOfDay >= New TimeSpan(17, 0, 0) And Now.TimeOfDay < New TimeSpan(20, 0, 0) Then
                        If IsDBNull(timeLogFound("TimeOutPM")) Then
                            tblLogAdapter.UpdateOutPM(Now, timeLogFound.ID)
                        End If
                    End If

                End If

                txbStatus.Text = "Record Found"
                Dim firstName = employeeFound.first_name
                Dim middleInitial = IIf(employeeFound.middle_name.Length > 0, employeeFound.middle_name(0) & ". ", "")
                Dim lastName = employeeFound.last_name
                Dim fullName = String.Format("{0} {1}{2}", firstName, middleInitial, lastName)
                txtEmpName.Text = fullName

                imgEmployee.Source = DataToBitmap(employeeFound.picture)


            End If
        End If
    End Sub

    Function DataToBitmap(data As Byte()) As BitmapImage
        Try
            Dim image = New BitmapImage()
            Using mem = New MemoryStream(data)
                mem.Position = 0
                image.BeginInit()
                image.CreateOptions = BitmapCreateOptions.PreservePixelFormat
                image.CacheOption = BitmapCacheOption.OnLoad
                image.UriSource = Nothing
                image.StreamSource = mem
                image.EndInit()
            End Using

            image.Freeze()
            Return image
        Catch ex As Exception
            Return New BitmapImage(New Uri("pack://siteoforigin:,,,/Resources/placeholder.png", UriKind.Absolute))
        End Try
    End Function
End Class
