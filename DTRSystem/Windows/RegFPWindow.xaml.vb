Imports ZKFPEngXControl

Imports System.Drawing
Imports System.Windows.Interop
Public Class RegFPWindow
    Dim FTempLen As Integer
    Dim FRegTemplate As String
    Dim FRegTemp As Object
    Dim fpcHandle As Long

    Public regPage As RegistrationPage
    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

        AddHandler fprintscanner.OnFeatureInfo, AddressOf Me.fp_OnFeatureInfo
        AddHandler fprintscanner.OnEnroll, AddressOf Me.fp_OnEnroll
        AddHandler fprintscanner.OnImageReceived, AddressOf Me.fp_OnImageReceived

        fprintscanner.s
    End Sub

    Private Sub Window_Loaded(sender As Object, e As RoutedEventArgs)
        fprintscanner.SensorIndex = 0
        If (fprintscanner.InitEngine = 0) Then
            imgConnected.Visibility = Windows.Visibility.Visible
            lblConnected.Content = "Device Connected"
        Else
            imgConnected.Visibility = Windows.Visibility.Hidden
            lblConnected.Content = "Device Not Connected"
        End If

        fprintscanner.BeginEnroll()
        txbStatus.Text = "Press with your finger 3 times."
    End Sub

    Private Sub fp_OnFeatureInfo(ByVal AQuality As Long)
        Dim sTemp As String

        sTemp = ""
        If fprintscanner.IsRegister Then
            If fprintscanner.EnrollIndex - 1 > 0 Then
                sTemp = "Press with your finger " & fprintscanner.EnrollIndex - 1 & " times"
            Else
                sTemp = ""
            End If
        End If

        txbStatus.Text = sTemp
    End Sub

    Private Sub fp_OnEnroll(ByVal ActionResult As Boolean, ByVal aTemplate As Object)
        If Not ActionResult Then
            MsgBox("Registration failed!", vbExclamation)
        Else
            MsgBox("Regsitration success!", vbInformation)

            FRegTemplate = fprintscanner.GetTemplateAsString()
            FRegTemp = fprintscanner.GetTemplate()
            fprintscanner.SaveTemplate(applicationPath & "\fptemp.tpl", FRegTemp)
        End If
        regPage.FingerprintEnrolled(ActionResult)
        Me.Close()
    End Sub

    Private Sub fp_OnImageReceived(ByRef AImageValid As Boolean)
        Dim myHandle As IntPtr = New WindowInteropHelper(Me).Handle
        Dim myGraphics As Graphics = Graphics.FromHwnd(myHandle)
        Dim x = (Me.Width / 2) - (fprintscanner.ImageWidth / 2)
        Dim i As IntPtr = myGraphics.GetHdc
        fprintscanner.PrintImageAt(i, x, 10, fprintscanner.ImageWidth, fprintscanner.ImageHeight)
        myGraphics.Dispose()
    End Sub

    Private Sub Window_Closed(sender As Object, e As EventArgs)
        
    End Sub

    Private Sub Window_Closing(sender As Object, e As ComponentModel.CancelEventArgs)
        e.Cancel = True
        If fprintscanner.IsRegister Then
            fprintscanner.CancelEnroll()
        End If
        isRegisteringFingerprint = False
        Me.Hide()

        RemoveHandler fprintscanner.OnFeatureInfo, AddressOf Me.fp_OnFeatureInfo
        AddHandler fprintscanner.OnEnroll, AddressOf Me.fp_OnEnroll
        AddHandler fprintscanner.OnImageReceived, AddressOf Me.fp_OnImageReceived
    End Sub
End Class
