Imports DTRSystem.DTRDataSetTableAdapters
Imports DTRSystem.DTRDataSet

Imports System.IO

Module DTRBiometricSystemModule
    Public applicationPath As String
    Public myDocumentsFolder As String
    Public dtrSystemFolder As String
    Public dtrSystemDB As String

    Public dtrMainWindow As MainWindow

    Public tblAdminAdapter As New AdminTableAdapter
    Public tblEmployeeAdapter As New EmployeeTableAdapter
    Public tblLogAdapter As New TimelogTableAdapter
    Public tblDeptAdapter As New DepartmentTableAdapter
    Public tblDesgAdapter As New DesignationTableAdapter
    Public tblDesgFullAdapter As New DesignationFullTableAdapter
    Public tblEmployeeFullAdapter As New EmployeeFullTableAdapter
    Public tblLeaveCreditsAdapter As New LeaveCreditsTableAdapter
    Public tblLeaveApplicationAdapter As New LeaveApplicationsTableAdapter
    Public tblSalaryGradeAdapter As New SalaryGradeTableAdapter

    Sub Main()
        applicationPath = AppDomain.CurrentDomain.BaseDirectory
        myDocumentsFolder = My.Computer.FileSystem.SpecialDirectories.MyDocuments
        dtrSystemFolder = Path.Combine(myDocumentsFolder, "DTRBiometricSystem")
        dtrSystemDB = Path.Combine(dtrSystemFolder, "dtr-biometric.accdb")
        AppDomain.CurrentDomain.SetData("DataDirectory", dtrSystemFolder)

        InitializeDBIfNotExist()

        Dim app As New System.Windows.Application
        app.Run(New LoginWindow)
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
    Function Coalesce(obj As Object)
        If IsDBNull(obj) Then
            If obj.GetType() = GetType(String) Then
                Return ""
            ElseIf obj.GetType() = GetType(Integer) Then
                Return 0
            Else
                Return Nothing
            End If
        Else
            Return obj
        End If
    End Function

    Public Function Weekdays(ByVal startDate As Date, ByVal endDate As Date) As Integer
        Dim numWeekdays As Integer
        Dim totalDays As Integer
        Dim WeekendDays As Integer
        numWeekdays = 0
        WeekendDays = 0

        totalDays = DateDiff(DateInterval.Day, startDate, endDate) + 1
        For i As Integer = 1 To totalDays
            If DatePart(DateInterval.Weekday, startDate) = 1 Then
                WeekendDays = WeekendDays + 1
            End If
            If DatePart(DateInterval.Weekday, startDate) = 7 Then
                WeekendDays = WeekendDays + 1
            End If
            startDate = DateAdd("d", 1, startDate)
        Next

        numWeekdays = totalDays - WeekendDays

        Return numWeekdays
    End Function

    Public Function WorkingDays(ByVal startDate As Date, ByVal endDate As Date) As List(Of Date)
        Dim wDays As New List(Of Date)
        Dim totalDays As Integer

        totalDays = DateDiff(DateInterval.Day, startDate, endDate) + 1

        For i As Integer = 1 To totalDays
            Dim isWeekend As Boolean = False
            If startDate.DayOfWeek = DayOfWeek.Sunday Then
                isWeekend = True
            ElseIf startDate.DayOfWeek = DayOfWeek.Saturday Then
                isWeekend = True
            End If
            If Not isWeekend Then
                wDays.Add(startDate)
            End If
            startDate = DateAdd("d", 1, startDate)
        Next
        Return wDays
    End Function
End Module
