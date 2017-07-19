Imports DTRSystem.DTRDataSetTableAdapters
Imports DTRSystem.DTRDataSet

Imports System.IO

Module DTRBiometricSystemModule
    Public applicationPath As String
    Public dtrMainWindow As MainWindow

    Public tblEmployeeAdapter As New EmployeeTableAdapter
    Public tblLogAdapter As New TimeLogTableAdapter

    Sub Main()
        applicationPath = AppDomain.CurrentDomain.BaseDirectory

        Dim app As New System.Windows.Application
        dtrMainWindow = New MainWindow
        app.Run(dtrMainWindow)
    End Sub

    Public Function DataToBitmap(data As Object) As BitmapImage
        Try
            Dim image = New BitmapImage()
            Using mem = New MemoryStream(TryCast(data, Byte()))
                mem.Position = 0
                image.BeginInit()
                image.CreateOptions = BitmapCreateOptions.PreservePixelFormat
                image.CacheOption = BitmapCacheOption.OnLoad
                image.UriSource = Nothing
                image.StreamSource = mem
                image.EndInit()
            End Using

            image.Freeze()
            Return image
        Catch ex As Exception
            Return New BitmapImage(New Uri("pack://siteoforigin:,,,/Resources/placeholder.png", UriKind.Absolute))
        End Try
    End Function
End Module
