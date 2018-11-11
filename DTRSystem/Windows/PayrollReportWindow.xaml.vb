Imports System.IO.File
Imports DTRFuncs
Imports DTRSystem.DTRDataSet
Imports System.Data
Public Class PayrollReportWindow
    Public department As DepartmentTableRow
    Public records As DataTable

    Public fromDate As Date
    Public toDate As Date

    Private Sub Window_Loaded(sender As Object, e As RoutedEventArgs)
        Dim htmlTable As String = ""
        Dim i = 1
        For Each row In records.Rows

            htmlTable += "<tr>" & vbCrLf
            htmlTable += String.Format("<td>{0}</td>", i & vbCrLf)
            htmlTable += String.Format("<td>{0}</td>", row("full_name") & vbCrLf)
            htmlTable += String.Format("<td>{0}</td>", row("designation") & vbCrLf)

            Dim fromToStr = String.Format("{0}-{1}, {2}", fromDate.ToString("MMMM dd"), toDate.ToString("MMMM dd"), toDate.Year)
            htmlTable += String.Format("<td colspan=2>{0}</td>", fromToStr)

            htmlTable += String.Format("<td>{0}</td>", row("monthly")) & vbCrLf
            htmlTable += String.Format("<td>{0}</td>", row("partial")) & vbCrLf

            htmlTable += String.Format("<td>{0}</td>", ReturnFormatN2(row("tax"))) & vbCrLf
            htmlTable += String.Format("<td>{0}</td>", ReturnFormatN2(row("gsis"))) & vbCrLf
            htmlTable += String.Format("<td>{0}</td>", ReturnFormatN2(row("others"))) & vbCrLf
            htmlTable += String.Format("<td>{0}</td>", ReturnFormatN2(row("totaldeductions"))) & vbCrLf

            htmlTable += String.Format("<td>{0}</td>", ReturnFormatN2(row("netpay"))) & vbCrLf
            htmlTable += String.Format("<td></td>")
            htmlTable += String.Format("<td></td>")

            htmlTable += "</tr>" & vbCrLf
        Next

        Dim HTMLReportPage As String = ReadAllText(applicationPath & "\MonthlyReports\PayrollReportPage.html")
        Dim monthStr As String = "ALL"
        Dim cleanHTML1 = HTMLReportPage.Replace("{0}", htmlTable).Replace("{1}", applicationPath.Replace("\", "/") & "MonthlyReports/").Replace("{2}", monthStr)
        Dim cleanHTML2 = cleanHTML1.Replace("{3}", department.dept_name)
        Dim finalHTML = cleanHTML2
        Debug.Print("Test: {0}", finalHTML)
        wb_report.NavigateToString(finalHTML)
    End Sub

    Function ReturnFormatN2(obj As Object) As String
        If IsDBNull(obj) Then
            Return "0.00"
        Else
            Dim newNumber As Decimal = obj
            Return newNumber.ToString("N2")
        End If
    End Function

    Private Sub btnPrint_Click(sender As Object, e As RoutedEventArgs) Handles btnPrint.Click
        webbrowser_extension.PrintDocument(wb_report)
    End Sub

    Private Sub btnClose_Click(sender As Object, e As RoutedEventArgs) Handles btnClose.Click
        Close()
    End Sub
End Class
