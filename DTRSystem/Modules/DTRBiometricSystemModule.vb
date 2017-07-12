Imports DTRSystem.DTRBiometricDataSetTableAdapters
Imports DTRSystem.DTRBiometricDataSet
Module DTRBiometricSystemModule
    Public applicationPath As String
    Public dtrMainWindow As MainWindow

    Public tblEmployeeAdapter As New EmployeeTableAdapter
    Public tblEmployeeROAdapter As New EmployeeROTableAdapter
    Public tblLogAdapter As New TimeLogTableAdapter
    Public tblLeaveCreditsAdapter As New LeaveCreditsTableAdapter
    Public tblDepartmentAdapter As New DepartmentTableAdapter

    Sub Main()
        applicationPath = AppDomain.CurrentDomain.BaseDirectory

        Dim app As New System.Windows.Application
        dtrMainWindow = New MainWindow
        app.Run(dtrMainWindow)
    End Sub
End Module
