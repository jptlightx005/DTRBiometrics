Imports System.IO
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
        ' Create OpenFileDialog 
        Dim dlg As New Microsoft.Win32.OpenFileDialog()

        ' Set filter for file extension and default file extension 
        dlg.DefaultExt = ".wav"
        dlg.Filter = "Audio Files|*.wav;*.mp3;*.mp4"

        ' Display OpenFileDialog by calling ShowDialog method 
        Dim result As Nullable(Of Boolean) = dlg.ShowDialog()

        ' Get the selected file name and display in a TextBox 
        If result = True Then
            Dim infoReader As System.IO.FileInfo
            infoReader = My.Computer.FileSystem.GetFileInfo(dlg.FileName)
            Debug.Print("File is " & infoReader.Length & " bytes.")

            ' Open document 
            Dim filename As String = dlg.FileName
            successMessageFile = filename
            txtSuccessAudio.Text = Path.GetFileName(successMessageFile)
            txtSuccessAudio.ToolTip = successMessageFile

        End If
    End Sub

    Private Sub btnTestSuccess_Click(sender As Object, e As RoutedEventArgs) Handles btnTestSuccess.Click
        My.Computer.Audio.Play(successMessageFile)

    End Sub
End Class
