Public Class PayslipWindow
    Public fromDate As Date
    Public toDate As Date

    Public basicPay As Double

    Public withholding As Double
    Public misc As Double
    Public tardiness As Double
    Public leaves As Double

    Private subtotal As Double

    Private netpay As Double

    Private Sub Window_Loaded(sender As Object, e As RoutedEventArgs)

        subtotal = withholding + misc + tardiness + leaves

        netpay = basicPay - subtotal

        lblFromDate.Content = fromDate.ToString("MMMM dd, yyyy")
        lblToDate.Content = toDate.ToString("MMMM dd, yyyy")

        lblBasicPay.Content = "P " & basicPay.ToString("N")

        lblWithholding.Content = "P " & withholding.ToString("N")
        lblMisc.Content = "P " & misc.ToString("N")
        lblTardiness.Content = "P " & tardiness.ToString("N")
        lblLeaves.Content = "P " & leaves.ToString("N")

        lblSubtotal.Content = "P " & subtotal.ToString("N")

        lblNetPay.Content = "P " & netpay.ToString("N")
    End Sub
End Class
