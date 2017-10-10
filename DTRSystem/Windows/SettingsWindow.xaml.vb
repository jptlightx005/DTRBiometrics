Imports System.IO
Imports WMPLib

Public Class SettingsWindow
    Dim successMessageFile As String
    Dim failedMessageFile As String
    Private Sub Grid_Loaded(sender As Object, e As RoutedEventArgs)
        successMessageFile = My.Settings.SuccessMessageAudioFile
        failedMessageFile = My.Settings.FailedMessageAudioFile

        txtSuccessAudio.Text = Path.GetFileName(successMessageFile)
        txtFailedAudio.Text = Path.GetFileName(failedMessageFile)

        txtSuccessAudio.ToolTip = successMessageFile
        txtFailedAudio.ToolTip = failedMessageFile

        txtDuration.Text = My.Settings.BeepDuration
        txtFrequency.Text = My.Settings.BeepFrequency
    End Sub

    Private Sub btnBrowseSuccess_Click(sender As Object, e As RoutedEventArgs) Handles btnBrowseSuccess.Click
        Dim dlg As New Microsoft.Win32.OpenFileDialog()

        dlg.DefaultExt = ".wav"
        dlg.Filter = "Audio Files|*.wav;*.mp3;*.mp4"

        Dim result As Nullable(Of Boolean) = dlg.ShowDialog()

        If result = True Then
            Dim infoReader As System.IO.FileInfo
            infoReader = My.Computer.FileSystem.GetFileInfo(dlg.FileName)
            Debug.Print("File is " & infoReader.Length & " bytes.")

            Dim filename As String = dlg.FileName
            successMessageFile = filename
            txtSuccessAudio.Text = Path.GetFileName(successMessageFile)
            txtSuccessAudio.ToolTip = successMessageFile
        End If
    End Sub

    Private Sub btnTestSuccess_Click(sender As Object, e As RoutedEventArgs) Handles btnTestSuccess.Click
        'My.Computer.Audio.Play(successMessageFile)
        Dim audioPlayer As New WindowsMediaPlayer
        audioPlayer.URL = successMessageFile
        audioPlayer.controls.play()
    End Sub

    Private Sub btnBrowseFailed_Click(sender As Object, e As RoutedEventArgs) Handles btnBrowseFailed.Click
        Dim dlg As New Microsoft.Win32.OpenFileDialog()

        dlg.DefaultExt = ".wav"
        dlg.Filter = "Audio Files|*.wav;*.mp3;*.mp4"

        Dim result As Nullable(Of Boolean) = dlg.ShowDialog()

        If result = True Then
            Dim infoReader As System.IO.FileInfo
            infoReader = My.Computer.FileSystem.GetFileInfo(dlg.FileName)
            Debug.Print("File is " & infoReader.Length & " bytes.")

            Dim filename As String = dlg.FileName
            failedMessageFile = filename
            txtFailedAudio.Text = Path.GetFileName(failedMessageFile)
            txtFailedAudio.ToolTip = failedMessageFile
        End If
    End Sub

    Private Sub btnTestFailed_Click(sender As Object, e As RoutedEventArgs) Handles btnTestFailed.Click
        Dim audioPlayer As New WindowsMediaPlayer
        audioPlayer.URL = failedMessageFile
        audioPlayer.controls.play()
    End Sub

    Private Sub btnTestBeep_Click(sender As Object, e As RoutedEventArgs) Handles btnTestBeep.Click
        Dim freq As Integer
        Dim duration As Integer
        If Not Integer.TryParse(txtFrequency.Text, freq) Then
            MsgBox("Invalid frequency value!", vbExclamation)
        End If
        If Not Integer.TryParse(txtDuration.Text, duration) Then
            MsgBox("Invalid duration value!", vbExclamation)
        End If

        Console.Beep(freq, duration)
    End Sub

    Private Sub btnSave_Click(sender As Object, e As RoutedEventArgs) Handles btnSave.Click
        Dim freq As Integer
        Dim duration As Integer
        If Not Integer.TryParse(txtFrequency.Text, freq) Then
            MsgBox("Invalid frequency value!", vbExclamation)
        End If
        If Not Integer.TryParse(txtDuration.Text, duration) Then
            MsgBox("Invalid duration value!", vbExclamation)
        End If

        My.Settings.SuccessMessageAudioFile = successMessageFile
        My.Settings.FailedMessageAudioFile = failedMessageFile

        My.Settings.BeepDuration = duration
        My.Settings.BeepFrequency = freq

        My.Settings.Save()

        If MsgBox("Settings successfully saved!", vbInformation) = vbOK Then
            Me.Close()
        End If
    End Sub

    Private Sub btnCancel_Click(sender As Object, e As RoutedEventArgs) Handles btnCancel.Click
        Me.Close()
    End Sub

    Private Sub btnDefault_Click(sender As Object, e As RoutedEventArgs) Handles btnDefault.Click
        My.Settings.Reset()

        successMessageFile = My.Settings.SuccessMessageAudioFile
        failedMessageFile = My.Settings.FailedMessageAudioFile

        txtSuccessAudio.Text = Path.GetFileName(successMessageFile)
        txtFailedAudio.Text = Path.GetFileName(failedMessageFile)

        txtSuccessAudio.ToolTip = successMessageFile
        txtFailedAudio.ToolTip = failedMessageFile

        txtDuration.Text = My.Settings.BeepDuration
        txtFrequency.Text = My.Settings.BeepFrequency
    End Sub
End Class
