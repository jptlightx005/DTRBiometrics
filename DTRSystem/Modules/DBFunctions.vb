Imports System.Data.SqlClient
Module DBFunctions
    Sub Test()
        Dim connection As SqlConnection = New SqlConnection()
        connection.ConnectionString = "Data Source=KABIR-DESKTOP;Initial Catalog=testDB;Integrated Security=True"
        connection.Open()

    End Sub
End Module
