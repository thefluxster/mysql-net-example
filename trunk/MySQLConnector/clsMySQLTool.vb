Imports System.Data
Imports MySql.Data.MySqlClient
Imports ErrorHandlerNS
Imports System.Web
Public MustInherit Class clsMySQLTool
    Protected MyServer As String
    Protected MyDatabase As String
    Protected MyUser As String
    Protected MyConnection As New MySqlConnection
    Protected Overridable ReadOnly Property strModuleName() As String
        Get
            strModuleName = "clsMySQLTool"
        End Get
    End Property
    Protected Function ConnnectToMySQLDB(ByVal Server As String, ByVal Database As String, ByVal Username As String, ByVal Password As String) As Boolean
        Try
            MyConnection = New MySqlConnection("server=" & Server & "; database=" & Database & "; uid=" & Username & "; pwd=" & Password & "; pooling=false;")
            'test the connection:
            MyConnection.Open()
            MyConnection.Close()
            'Remember the settings for this connection in case you need to refer to them in the future for some reason...
            MyServer = Server
            MyUser = Username
            MyDatabase = Database
            'Everything worked, return true:
            Return True
        Catch ex As Exception
            modErrorHandler.Errors.PrintMessage(1, "Unable to connect to DataBase: " & ex.Message, strModuleName)
            Return False
        End Try
    End Function
    'Public Function InsertData(ByVal SQLcommand As String) As Boolean
    '    Try
    '        'make sure the connection is open...
    '        If Not MyConnection.State = ConnectionState.Open Then
    '            MyConnection.Open()
    '        End If

    '        MyConnection.Close()
    '    Catch ex As Exception

    '    End Try
    'End Function
    Public Function ExecuteNonQuery(ByVal SQLQuery As String) As Boolean
        If MyConnection.ConnectionString.Length < 1 Then
            modErrorHandler.Errors.PrintMessage(2, "Connection must first be made to a valid MySQL server before issuing this command.", strModuleName)
            Return False
        Else
            Try
                'Dim myData As New DataTable
                'Setup the query:
                Dim SQLCommand As String = SQLQuery
                Dim myDataAdapter As New MySqlDataAdapter
                Dim myCommand As New MySqlCommand()

                myCommand.Connection = MyConnection
                myCommand.CommandText = SQLCommand
                'make sure the connection is open...
                If Not MyConnection.State = ConnectionState.Open Then
                    MyConnection.Open()
                End If
                myCommand.ExecuteNonQuery()
                'Close the connection
                MyConnection.Close()

                Return True
            Catch ex As Exception
                modErrorHandler.Errors.PrintMessage(2, "Failed to execute MySQL command. Message: " + ex.Message.ToString, strModuleName + "-LinkOutputToJob")
                MyConnection.Close()
                Return False
            End Try
        End If
    End Function
    Public Function QueryDB(ByVal SQLCommand As String, ByVal KeyFieldName As String) As DataTable
        If MyConnection.ConnectionString.Length < 1 Then
            modErrorHandler.Errors.PrintMessage(2, "Connection must first be made to a valid MySQL server before issuing this command.", strModuleName)
            Return Nothing
        Else
            Try
                'make sure the connection is open...
                If Not MyConnection.State = ConnectionState.Open Then
                    MyConnection.Open()
                End If

                Dim myDataAdapter As New MySqlDataAdapter
                'Dim myDataReader As MySqlDataReader
                Dim myCommand As New MySqlCommand()
                Dim myData As New DataTable

                myCommand.Connection = MyConnection
                myCommand.CommandText = SQLCommand
                myDataAdapter.SelectCommand = myCommand
                myDataAdapter.Fill(myData)

                Return myData

                'myDataReader = myCommand.ExecuteReader(CommandBehavior.CloseConnection)
                ''create and then populate Object array:
                'Dim objCells(0) As Object
                'Dim myResults As New Hashtable
                'With myDataReader
                '    If .HasRows Then
                '        ReDim objCells(.FieldCount - 1)
                '        While .Read
                '            Dim keycolnum As Integer = 0
                '            'add rows from the returned query result to the Object Array
                '            .GetValues(objCells)
                '            myResults.Add(.GetValue(0), objCells)
                '        End While
                '    End If
                'End With
                'myDataReader.Close()

                'MyConnection.Close()
                'Return myResults
            Catch ex As Exception
                Return Nothing
            End Try
        End If


    End Function

    Public Sub New()

    End Sub
End Class
