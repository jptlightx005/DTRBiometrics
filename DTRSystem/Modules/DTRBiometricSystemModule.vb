Imports DTRSystem.DTRBiometricDataSetTableAdapters
Imports DTRSystem.DTRBiometricDataSet
Module DTRBiometricSystemModule
    Public applicationPath As String
    Public dtrMainWindow As MainWindow

    Public tblEmployeeAdapter As New EmployeeTableAdapter
    Public tblLogAdapter As New DTRBiometricDataSetTableAdapters.TimeLogTableAdapter

    Sub Main()
        applicationPath = AppDomain.CurrentDomain.BaseDirectory

        Dim app As New System.Windows.Application
        dtrMainWindow = New MainWindow
        app.Run(dtrMainWindow)
    End Sub
End Module
