Imports System.IO.File
Imports DTRSystem.DTRDataSet
Imports DTRSystem.DTRDataSetTableAdapters
Imports SMSCSFuncs
Public Class LeaveReportWebWindow
    Public records As LeaveCreditsTableRow()
    Public employee As EmployeeFullRow
    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        AddHandler wb_report.Navigated, AddressOf wb_report_navigated
    End Sub
    Sub wb_report_navigated(ByVal sender As Object, e As NavigationEventArgs)
        webbrowser_extension.SetSilent(wb_report, True)

    End Sub
    Private Sub Window_Loaded(sender As Object, e As RoutedEventArgs)
        Dim htmlTable As String = ""
        For Each leaveRow In records

            Dim vcearnedSTR As String = IIf(leaveRow.VC_Earned = 0, "", leaveRow.VC_Earned.ToString("N2"))
            Dim vcusedSTR As String = IIf(leaveRow.VC_Used = 0, "", leaveRow.VC_Used.ToString("N2"))
            Dim vcbalanceSTR As String = IIf(leaveRow.VC_Balance = 0, "", leaveRow.VC_Balance.ToString("N2"))

            Dim scearnedSTR As String = IIf(leaveRow.SC_Earned = 0, "", leaveRow.SC_Earned.ToString("N2"))
            Dim scusedSTR As String = IIf(leaveRow.SC_Used = 0, "", leaveRow.SC_Used.ToString("N2"))
            Dim scbalanceSTR As String = IIf(leaveRow.SC_Balance = 0, "", leaveRow.SC_Balance.ToString("N2"))

            htmlTable += "<tr align='center'>" & vbCrLf
            htmlTable += String.Format("<td>{0}</td>", leaveRow.DateOfTransaction.ToString("MMMM dd, yyyy")) & vbCrLf
            htmlTable += String.Format("<td>{0}</td>", leaveRow.RemarksV2) & vbCrLf

            htmlTable += String.Format("<td>{0}</td>", vcearnedSTR) & vbCrLf
            htmlTable += String.Format("<td>{0}</td>", vcusedSTR) & vbCrLf
            htmlTable += String.Format("<td>{0}</td>", vcbalanceSTR) & vbCrLf
            htmlTable += String.Format("<td></td>") & vbCrLf

            htmlTable += String.Format("<td>{0}</td>", scearnedSTR) & vbCrLf
            htmlTable += String.Format("<td>{0}</td>", scusedSTR) & vbCrLf
            htmlTable += String.Format("<td>{0}</td>", scbalanceSTR) & vbCrLf
            htmlTable += String.Format("<td></td>") & vbCrLf

            htmlTable += String.Format("<td></td>") & vbCrLf

            htmlTable += "</tr>" & vbCrLf
        Next

        Dim HTMLReportPage As String = ReadAllText(applicationPath & "\MonthlyReports\LeaveCreditsReportPage.html")
        Dim monthStr As String = "ALL"
        Dim cleanHTML1 = HTMLReportPage.Replace("{0}", htmlTable).Replace("{1}", applicationPath.Replace("\", "/") & "MonthlyReports/").Replace("{2}", monthStr)
        Dim cleanHTML2 = cleanHTML1.Replace("{3}", employee.full_name).Replace("{4}", Coalesce(employee("dept_name")))
        Dim finalHTML = cleanHTML2
        Debug.Print("Test: {0}", finalHTML)
        wb_report.NavigateToString(finalHTML)
    End Sub

    Private Sub btnPrint_Click(sender As Object, e As RoutedEventArgs) Handles btnPrint.Click
        webbrowser_extension.PrintDocument(wb_report)
    End Sub

    Private Sub btnClose_Click(sender As Object, e As RoutedEventArgs) Handles btnClose.Click
        Close()
    End Sub
End Class
