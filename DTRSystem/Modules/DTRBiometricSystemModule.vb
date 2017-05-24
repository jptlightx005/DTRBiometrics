Module DTRBiometricSystemModule
    Dim applicationPath As String
    Sub Main()
        applicationPath = AppDomain.CurrentDomain.BaseDirectory

        Dim app As New System.Windows.Application
        app.Run(New MainWindow)
    End Sub
End Module
