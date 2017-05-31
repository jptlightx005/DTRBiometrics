Module DTRBiometricSystemModule
    Public applicationPath As String
    Public dtrMainWindow As MainWindow
    Sub Main()
        applicationPath = AppDomain.CurrentDomain.BaseDirectory

        Dim app As New System.Windows.Application
        dtrMainWindow = New MainWindow
        app.Run(dtrMainWindow)
    End Sub
End Module
