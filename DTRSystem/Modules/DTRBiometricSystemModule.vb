Imports DTRSystem.DTRDataSetTableAdapters
Imports DTRSystem.DTRDataSet

Imports System.IO

Module DTRBiometricSystemModule
    Public applicationPath As String
    Public myDocumentsFolder As String
    Public dtrSystemFolder As String
    Public dtrSystemDB As String

    Public dtrMainWindow As MainWindow

    Public tblEmployeeAdapter As New EmployeeTableAdapter
    Public tblLogAdapter As New TimelogTableAdapter
    Public tblDeptAdapter As New DepartmentTableAdapter

    Sub Main()
        applicationPath = AppDomain.CurrentDomain.BaseDirectory
        myDocumentsFolder = My.Computer.FileSystem.SpecialDirectories.MyDocuments
        dtrSystemFolder = Path.Combine(myDocumentsFolder, "DTRBiometricSystem")
        dtrSystemDB = Path.Combine(dtrSystemFolder, "dtr-biometric.accdb")
        AppDomain.CurrentDomain.SetData("DataDirectory", dtrSystemFolder)

        InitializeDBIfNotExist()

        Dim app As New System.Windows.Application
        dtrMainWindow = New MainWindow
        app.Run(dtrMainWindow)
    End Sub

    Public Sub InitializeDBIfNotExist()
        If Not Directory.Exists(dtrSystemFolder) Then
            Directory.CreateDirectory(dtrSystemFolder)
        End If

        If Not File.Exists(dtrSystemDB) Then
            File.Copy(Path.Combine(applicationPath, "dtr-biometric.accdb"), dtrSystemDB)
        End If
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
