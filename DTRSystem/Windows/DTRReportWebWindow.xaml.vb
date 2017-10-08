Imports System.IO.File
Imports SMSCSFuncs
Imports DTRSystem.DTRDataSet
Public Class DTRReportWebWindow
    Public employee As EmployeeFullRow
    Public records As TimelogTableDataTable
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
        For Each timeLog In records

            Dim timeInAM As String = ReturnFormattedDateTimeIfNotNull(timeLog("TimeInAM"))
            Dim timeOutAM As String = ReturnFormattedDateTimeIfNotNull(timeLog("TimeOutAM"))
            Dim timeInPM As String = ReturnFormattedDateTimeIfNotNull(timeLog("TimeInPM"))
            Dim timeOutPM As String = ReturnFormattedDateTimeIfNotNull(timeLog("TimeOutPM"))


            htmlTable += "<tr>" & vbCrLf
            htmlTable += String.Format("<td>{0}</td>", timeLog.DateOfTheDay.ToString("MMMM dd, yyyy")) & vbCrLf
            htmlTable += String.Format("<td>{0}</td>", timeInAM) & vbCrLf
            htmlTable += String.Format("<td>{0}</td>", timeOutAM) & vbCrLf
            htmlTable += String.Format("<td>{0}</td>", timeInPM) & vbCrLf
            htmlTable += String.Format("<td>{0}</td>", timeOutPM) & vbCrLf
            htmlTable += String.Format("<td>{0}</td>", (timeLog.TotalTime / 60.0).ToString("N2")) & vbCrLf
            htmlTable += "</tr>" & vbCrLf
        Next

        Dim HTMLReportPage As String = ReadAllText(applicationPath & "\MonthlyReports\MonthlyReportPage.html")
        Dim monthStr As String = "ALL"
        Dim cleanHTML1 = HTMLReportPage.Replace("{0}", htmlTable).Replace("{1}", applicationPath.Replace("\", "/") & "MonthlyReports/").Replace("{2}", monthStr)
        Dim cleanHTML2 = cleanHTML1.Replace("{3}", employee.full_name).Replace("{4}", Coalesce(employee("designation_name")))
        Dim finalHTML = cleanHTML2
        Debug.Print("Test: {0}", finalHTML)
        wb_report.NavigateToString(finalHTML)
    End Sub

    Function ReturnFormattedDateTimeIfNotNull(obj As Object) As String
        If IsDBNull(obj) Then
            Return "--:-- --"
        Else
            Dim newDate As DateTime = obj
            Return newDate.ToString("hh:mm tt")
        End If
    End Function

    Private Sub btnPrint_Click(sender As Object, e As RoutedEventArgs) Handles btnPrint.Click
        webbrowser_extension.PrintDocument(wb_report)
    End Sub

    Private Sub btnClose_Click(sender As Object, e As RoutedEventArgs) Handles btnClose.Click
        Close()
    End Sub
End Class
