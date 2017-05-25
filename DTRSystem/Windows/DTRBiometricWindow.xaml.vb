Imports ZKFPEngXControl
Imports System.IO
Public Class DTRBiometricWindow
    Dim WithEvents fp As ZKFPEngX
    Dim fpHandle As Integer
    Dim idList As List(Of Integer)
    Public otemplate As Object
    Dim a As New DTRBiometricDataSetTableAdapters.tbl_employeeTableAdapter

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
        For Each row In a.GetData
            idList.Add(row.ID)
            Dim fileName = String.Format(applicationPath & "\fptemp{0}.tpl", row.ID)
            File.WriteAllBytes(fileName, row.biometric)
            fp.AddRegTemplateFileToFPCacheDB(fpHandle, i, fileName)
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
        If fi = -1 Then
            txbStatus.Text = "Not registered"
        Else
            Dim emp As DTRBiometricDataSet.tbl_employeeRow
            For Each row In a.GetData()
                If row.ID = idList(fi) Then
                    emp = row
                End If
            Next
            If Not emp Is Nothing Then
                txtEmpName.Text = emp.first_name
            End If

        End If
    End Sub
End Class
