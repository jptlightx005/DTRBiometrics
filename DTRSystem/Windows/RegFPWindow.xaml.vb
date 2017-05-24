Imports ZKFPEngXControl
Imports System.Drawing
Imports System.Windows.Interop
Public Class RegFPWindow
    Dim WithEvents fp As ZKFPEngX
    Public otemplate As Object
    Dim FTempLen As Integer
    Dim FRegTemplate As String
    Dim FRegTemp As Object
    Dim fpcHandle As Long
    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        fp = New ZKFPEngX

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

        fp.BeginEnroll()
        txbStatus.Text = "Press with your finger 3 times."
    End Sub

    Private Sub fp_OnFeatureInfo(ByVal AQuality As Long) Handles fp.OnFeatureInfo
        Dim sTemp As String

        sTemp = ""
        If fp.IsRegister Then
            If fp.EnrollIndex - 1 > 0 Then
                sTemp = "Press with your finger " & fp.EnrollIndex - 1 & " times"
            Else
                sTemp = ""
            End If
        End If

        txbStatus.Text = sTemp
    End Sub

    Private Sub fp_OnEnroll(ByVal ActionResult As Boolean, ByVal aTemplate As Object) Handles fp.OnEnroll
        Dim i As Long

        If Not ActionResult Then
            MsgBox("Registration failed!", vbExclamation)
        Else
            MsgBox("Regsitration success!", vbInformation)

            FRegTemplate = fp.GetTemplateAsString()
            FRegTemp = fp.GetTemplate()
            fp.SaveTemplateStr(applicationPath & "\fptemp.txt", FRegTemplate)
            fp.SaveTemplate(applicationPath & "\fptemp.tpl", FRegTemp)

            'To Save to Database

            '  ZKFPEngX1.AddRegTemplateStrToFPCacheDB fpcHandle, FingerCount, FRegTemplate
            'fp.AddRegTemplateToFPCacheDB(fpcHandle, FingerCount, FRegTemp)

            Me.Close()
        End If
    End Sub

    Private Sub fp_OnImageReceived(ByRef AImageValid As Boolean) Handles fp.OnImageReceived
        Dim myHandle As IntPtr = New WindowInteropHelper(Me).Handle
        Dim myGraphics As Graphics = Graphics.FromHwnd(myHandle)
        Dim x = (Me.Width / 2) - (fp.ImageWidth / 2)
        fp.PrintImageAt(myGraphics.GetHdc, x, 10, fp.ImageWidth, fp.ImageHeight)

    End Sub
End Class
