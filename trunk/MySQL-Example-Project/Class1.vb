Imports MySql.Data.MySqlClient
Imports ErrorHandlerNS
Imports MySQLConnector
Imports System.Text


Public Class Class1
    Inherits clsMySQLTool
    Protected Overrides ReadOnly Property strModuleName() As String
        Get
            strModuleName = "clsMySQLTool-AutoPublishingDB"
        End Get
    End Property

    Public DatabaseConnected As Boolean = False

    Public Function OutputEntryExists(ByVal strEnvironment As String, ByVal PUBGUID As String, ByVal PUBVER As String, ByVal OutLang As String, ByVal OUTTYPE As String, Optional ByVal strUserName As String = "Anonymous") As String
        If Me.DatabaseConnected = True Then
            Try
                Dim myData As New DataTable
                'Switch based on content of string
                Dim SQLCommand As String = "SELECT OUTID FROM autosrv_autopub_OutputList where PUBGUID = '" & PUBGUID & "' AND JOBENV = '" & strEnvironment & "' AND PUBVER = '" & PUBVER & "' AND OUTLANG = '" & OutLang & "' AND OUTTYPE = '" & OUTTYPE & "'"
                Dim myDataAdapter As New MySqlDataAdapter
                Dim myCommand As New MySqlCommand()

                myCommand.Connection = MyConnection
                myCommand.CommandText = SQLCommand
                myDataAdapter.SelectCommand = myCommand
                'make sure the connection is open...
                If Not MyConnection.State = ConnectionState.Open Then
                    MyConnection.Open()
                End If
                myDataAdapter.Fill(myData)

                MyBase.MyConnection.Close()
                'Check to see if we found any matches.
                If myData.Rows.Count > 0 Then
                    Dim jobids As New Hashtable
                    For Each JobEntry As DataRow In myData.Select()
                        If Not jobids.ContainsKey(JobEntry.Item(0).ToString) Then
                            jobids.Add(JobEntry.Item(0).ToString, JobEntry.Item(0).ToString)
                        End If
                    Next
                    If jobids.Count > 1 Then
                        'Not sure why, but there appear to be multiple output entries matching this setup in the DB...
                        'Report the issue
                        modErrorHandler.Errors.PrintMessage(2, "[" + strUserName + "] There are multiple OUTIDs in the DB matching this criteria: JobEnv: " + strEnvironment + ", PubGUID: " + PUBGUID + ", PubVer: " + PUBVER + ", OutType: " + OUTTYPE, strModuleName + "-OutputEntryExists")
                    End If
                    'Return the first one:
                    For Each mykey As DictionaryEntry In jobids
                        Return mykey.Value
                    Next
                Else
                    Return Nothing
                End If

            Catch ex As Exception
                modErrorHandler.Errors.PrintMessage(2, "[" + strUserName + "] Failed to search for entry in DB. Message: " + ex.Message.ToString, strModuleName + "-OutputEntryExists")
                MyBase.MyConnection.Close()
                Return Nothing
            End Try
        Else
            modErrorHandler.Errors.PrintMessage(2, "[" + strUserName + "] Skipping function - no valid database connection.", strModuleName + "-OutputEntryExists")
            Return Nothing
        End If
    End Function
    Public Function DeleteOutputEntrybyID(ByVal OutID As String, Optional ByVal strUserName As String = "Anonymous") As Boolean
        If Me.DatabaseConnected = True Then
            Dim SQLCommand As String = "DELETE FROM autosrv_autopub_OutputList WHERE OUTID = '" & OutID & "' LIMIT 1"
            If ExecuteNonQuery(SQLCommand) = True Then
                Return True
            Else
                modErrorHandler.Errors.PrintMessage(2, "[" + strUserName + "] Failed to delete entry from DB for " + OutID, strModuleName + "-DeleteJobEntry")
                Return False
            End If
        Else
            modErrorHandler.Errors.PrintMessage(2, "[" + strUserName + "] Skipping function - no valid database connection.", strModuleName + "-DeleteEntrybyGUID")
            Return False
        End If
    End Function
    Public Function DeleteJobEntrybyID(ByVal JobID As String, Optional ByVal strUserName As String = "Anonymous") As Boolean
        If Me.DatabaseConnected = True Then
            Try
                If JobEntryExists(JobID, strUserName) Then
                    'Get a list of all outputs.  We need to remove them if they are no longer used by any jobs.
                    Dim OutIDs As Hashtable = GetOutIDsFromJob(JobID, strUserName)

                    'Deactivate the job entry in the joblist:
                    Dim SQLCommand As String = "DELETE FROM autosrv_autopub_JobList WHERE JOBID = '" & JobID & "' LIMIT 1"
                    If ExecuteNonQuery(SQLCommand) = True Then
                        Return True
                    Else
                        modErrorHandler.Errors.PrintMessage(2, "[" + strUserName + "] Attempted to delete " + JobID + " from database but it doesn't exist.", strModuleName + "-DeleteJobEntry")
                        Return False
                    End If
                    'now remove the linked job/outputs
                    For Each OUTID As DictionaryEntry In OutIDs
                        SQLCommand = "DELETE FROM autosrv_autopub_Job_Output_Links WHERE JOBID = '" + JobID + "' AND OUTID = '" + OUTID.Value.ToString + "'"
                        If ExecuteNonQuery(SQLCommand) = True Then
                            'Return True
                        Else
                            modErrorHandler.Errors.PrintMessage(2, "[" + strUserName + "] Attempted to delete link between " + JobID + " and " + OUTID.Value.ToString + " from database but it doesn't exist.", strModuleName + "-DeleteJobEntry")
                            'Return False
                        End If
                        'Now check to see if any of the outputs are now orphaned. Delete if they are.
                        Dim jobids As Hashtable = GetJobIDsFromOutput(OUTID.Value, strUserName)
                        If jobids.Count = 0 Then
                            'no Jobs for the current output. Delete the output.
                            SQLCommand = "DELETE FROM autosrv_autopub_OutputList WHERE OUTID = '" + OUTID.Value + "'"
                            If ExecuteNonQuery(SQLCommand) = True Then
                                'Return True
                            Else
                                modErrorHandler.Errors.PrintMessage(2, "[" + strUserName + "] Attempted to delete orphaned output " + OUTID.Value.ToString + " from database but it doesn't exist.", strModuleName + "-DeleteJobEntry")
                                'Return False
                            End If
                        End If
                    Next


                Else
                    modErrorHandler.Errors.PrintMessage(2, "[" + strUserName + "] Attempted to delete " + JobID + " from database but it doesn't exist.", strModuleName + "-DeleteJobEntry")
                    Return False
                End If
            Catch ex As Exception
                modErrorHandler.Errors.PrintMessage(2, "[" + strUserName + "] Failed to delete entry from DB for " + JobID + "Message: " + ex.Message.ToString, strModuleName + "-DeleteJobEntry")
                Return False
            End Try
        Else
            modErrorHandler.Errors.PrintMessage(2, "[" + strUserName + "] Skipping function - no valid database connection.", strModuleName + "-DeleteEntrybyGUID")
            Return False
        End If
    End Function

    ''' <summary>
    ''' Returns a list of JobIDs that reference a given OutID.
    ''' </summary>
    ''' <param name="OutID"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetJobIDsFromOutput(ByVal OutID As String, Optional ByVal strUserName As String = "Anonymous") As Hashtable
        If Me.DatabaseConnected = True Then
            Try
                Dim myData As New DataTable
                Dim SQLCommand As String = "SELECT JOBID FROM autosrv_autopub_Job_Output_Links where OUTID = '" & OutID & "'"
                Dim myDataAdapter As New MySqlDataAdapter
                Dim myCommand As New MySqlCommand()

                myCommand.Connection = MyConnection
                myCommand.CommandText = SQLCommand
                myDataAdapter.SelectCommand = myCommand

                'make sure the connection is open...
                If Not MyConnection.State = ConnectionState.Open Then
                    MyConnection.Open()
                End If
                myDataAdapter.Fill(myData)
                Dim jobids As New Hashtable
                'For each ID in items, add to the hash.                
                For Each JobEntry As DataRow In myData.Select()
                    If Not jobids.ContainsKey(JobEntry.Item(0).ToString) Then
                        jobids.Add(JobEntry.Item(0).ToString, JobEntry.Item(0).ToString)
                    End If
                Next
                MyBase.MyConnection.Close()
                Return jobids
            Catch ex As Exception
                modErrorHandler.Errors.PrintMessage(2, "[" + strUserName + "] Failed to find JobID. Message: " + ex.Message.ToString, strModuleName + "-GetGUID")
                MyBase.MyConnection.Close()
                Return Nothing
            End Try
        Else
            modErrorHandler.Errors.PrintMessage(2, "[" + strUserName + "] Skipping function - no valid database connection.", strModuleName + "-GetGUID")
            Return Nothing
        End If
    End Function
    Public Function GetJobEnv(ByVal strJobID As String, Optional ByVal strUserName As String = "Anonymous") As String
        If Me.DatabaseConnected = True Then
            Try
                Dim myData As New DataTable
                Dim SQLCommand As String = "SELECT JOBENV FROM autosrv_autopub_JobList where JOBID = '" & strJobID & "'"
                Dim myDataAdapter As New MySqlDataAdapter
                Dim myCommand As New MySqlCommand()

                myCommand.Connection = MyConnection
                myCommand.CommandText = SQLCommand
                myDataAdapter.SelectCommand = myCommand

                'make sure the connection is open...
                If Not MyConnection.State = ConnectionState.Open Then
                    MyConnection.Open()
                End If
                myDataAdapter.Fill(myData)
                'For each ID in items, add to the hash.
                Dim JobEnv As String = myData.Rows.Item(0).Item(0)
                MyBase.MyConnection.Close()
                Return JobEnv
            Catch ex As Exception
                modErrorHandler.Errors.PrintMessage(2, "[" + strUserName + "] Failed to find JobEnv. Message: " + ex.Message.ToString, strModuleName + "-GetJobEnv")
                MyBase.MyConnection.Close()
                Return ""
            End Try
        Else
            modErrorHandler.Errors.PrintMessage(2, "[" + strUserName + "] Skipping function - no valid database connection.", strModuleName + "-GetJobDesc")
            Return ""
        End If
    End Function

    ''' <summary>
    ''' Returns a Job Description given a JobID.
    ''' </summary>
    ''' 
    ''' <param name="JobID"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Function GetJobDesc(ByVal JobID As String, Optional ByVal strUserName As String = "Anonymous") As String
        If Me.DatabaseConnected = True Then
            Try
                Dim myData As New DataTable
                Dim SQLCommand As String = "SELECT JOBDESC FROM autosrv_autopub_JobList where JOBID = '" & JobID & "'"
                Dim myDataAdapter As New MySqlDataAdapter
                Dim myCommand As New MySqlCommand()

                myCommand.Connection = MyConnection
                myCommand.CommandText = SQLCommand
                myDataAdapter.SelectCommand = myCommand

                'make sure the connection is open...
                If Not MyConnection.State = ConnectionState.Open Then
                    MyConnection.Open()
                End If
                myDataAdapter.Fill(myData)
                'For each ID in items, add to the hash.
                Dim JobDesc As String = myData.Rows.Item(0).Item(0)
                MyBase.MyConnection.Close()
                Return JobDesc
            Catch ex As Exception
                modErrorHandler.Errors.PrintMessage(2, "[" + strUserName + "] Failed to find JobDesc. Message: " + ex.Message.ToString, strModuleName + "-GetJobDesc")
                MyBase.MyConnection.Close()
                Return ""
            End Try
        Else
            modErrorHandler.Errors.PrintMessage(2, "[" + strUserName + "] Skipping function - no valid database connection.", strModuleName + "-GetJobDesc")
            Return ""
        End If
    End Function
    Public Function FillComboBox(ByVal query As String, ByRef ddl As Global.System.Web.UI.WebControls.DropDownList)
        MyConnection.Open()
        Dim myCommand As New MySqlCommand(query, MyConnection)
        'Dim dr As MySqlDataReader
        Dim mydataAdapter As New MySqlDataAdapter
        Dim dsToBeFilled As New DataSet




        myCommand.Connection = MyConnection
        myCommand.CommandText = query
        mydataAdapter.SelectCommand = myCommand
        mydataAdapter.Fill(dsToBeFilled)
        Dim dt As New DataTable
        dt = dsToBeFilled.Tables(0)
        If dsToBeFilled.Tables(0).Rows.Count = 0 Then
            'No data was returned by our query. Populate ddl with "No items available"
            ddl.Items.Add("")
            ddl.Items(0).Text = "No items available"
            ddl.Items(0).Value = "0"
            Return True
        End If
        'otherwise, continue

        Try

            ddl.Items.Clear()
            ddl.DataSource = dt
            ddl.DataTextField = dt.Columns(0).ToString
            ddl.DataValueField = dt.Columns(1).ToString
            ddl.DataBind()
            '   dr = cmd.ExecuteReader
            '  While (dr.Read = True) ' listitem
            'ddl.DataTextField
            'ddl.Items.Add(dr.GetValue(0))
            'End While
            'dr.Close()
        Catch ex As System.Exception
            Return False
        Finally
            myCommand.Dispose()
            MyConnection.Close()
        End Try
        Return True
    End Function
    ''' <summary>
    ''' Returns a JobName given a JobID.
    ''' </summary>
    ''' 
    ''' <param name="JobID"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetJobName(ByVal JobID As String, Optional ByVal strUserName As String = "Anonymous") As String
        If Me.DatabaseConnected = True Then
            Try
                Dim myData As New DataTable
                Dim SQLCommand As String = "SELECT JOBNAME FROM autosrv_autopub_JobList where JOBID = '" & JobID & "'"
                Dim myDataAdapter As New MySqlDataAdapter
                Dim myCommand As New MySqlCommand()

                myCommand.Connection = MyConnection
                myCommand.CommandText = SQLCommand
                myDataAdapter.SelectCommand = myCommand

                'make sure the connection is open...
                If Not MyConnection.State = ConnectionState.Open Then
                    MyConnection.Open()
                End If
                myDataAdapter.Fill(myData)
                'For each ID in items, add to the hash.
                Dim JobName As String = myData.Rows.Item(0).Item(0)
                MyBase.MyConnection.Close()
                Return JobName
            Catch ex As Exception
                modErrorHandler.Errors.PrintMessage(2, "[" + strUserName + "] Failed to find JobName. Message: " + ex.Message.ToString, strModuleName + "-GetJobName")
                MyBase.MyConnection.Close()
                Return ""
            End Try
        Else
            modErrorHandler.Errors.PrintMessage(2, "[" + strUserName + "] Skipping function - no valid database connection.", strModuleName + "-GetJobName")
            Return ""
        End If
    End Function

    ''' <summary>
    ''' Returns a ServerEnv given an OutID.
    ''' </summary>
    ''' 
    ''' <param name="OutID"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetOutputServEnv(ByVal OutID As String, Optional ByVal strUserName As String = "Anonymous") As String
        If Me.DatabaseConnected = True Then
            Try
                Dim myData As New DataTable
                Dim SQLCommand As String = "SELECT JOBENV FROM autosrv_autopub_OutputList where OUTID = '" & OutID & "'"
                Dim myDataAdapter As New MySqlDataAdapter
                Dim myCommand As New MySqlCommand()

                myCommand.Connection = MyConnection
                myCommand.CommandText = SQLCommand
                myDataAdapter.SelectCommand = myCommand

                'make sure the connection is open...
                If Not MyConnection.State = ConnectionState.Open Then
                    MyConnection.Open()
                End If
                myDataAdapter.Fill(myData)
                'For each ID in items, add to the hash.
                Dim JobName As String = myData.Rows.Item(0).Item(0)
                MyBase.MyConnection.Close()
                Return JobName
            Catch ex As Exception
                modErrorHandler.Errors.PrintMessage(2, "[" + strUserName + "] Failed to find OutID. Message: " + ex.Message.ToString, strModuleName + "-GetOutputServEnv")
                MyBase.MyConnection.Close()
                Return ""
            End Try
        Else
            modErrorHandler.Errors.PrintMessage(2, "[" + strUserName + "] Skipping function - no valid database connection.", strModuleName + "-GetOutputServEnv")
            Return ""
        End If
    End Function

    Public Function GetOutputTaskDataTable(ByVal OutID As String, Optional ByVal strUserName As String = "Anonymous") As DataTable
        If Me.DatabaseConnected = True Then
            Try
                'make sure the connection is open...
                If Not MyConnection.State = ConnectionState.Open Then
                    MyConnection.Open()
                End If

                Dim myDataAdapter As New MySqlDataAdapter
                'Dim myDataReader As MySqlDataReader
                Dim myCommand As New MySqlCommand()
                Dim myData As New DataTable
                Dim SQLCommand As String = "SELECT JOBENV, PUBGUID, PUBVER, OUTLANG, OUTTYPE FROM autosrv_autopub_OutputList where OUTID = '" & OutID & "'"

                myCommand.Connection = MyConnection
                myCommand.CommandText = SQLCommand
                myDataAdapter.SelectCommand = myCommand
                myDataAdapter.Fill(myData)



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

                MyConnection.Close()
                Return myData

            Catch ex As Exception
                Return Nothing
            End Try
        Else
            modErrorHandler.Errors.PrintMessage(2, "[" + strUserName + "] Skipping function - no valid database connection.", strModuleName + "-GetGUID")
            Return Nothing
        End If
    End Function
    Public Function GetJobData(ByVal JobID As String, Optional ByVal strUserName As String = "Anonymous") As Dictionary(Of String, String)
        If Me.DatabaseConnected = True Then
            Try
                'make sure the connection is open...
                If Not MyConnection.State = ConnectionState.Open Then
                    MyConnection.Open()
                End If

                Dim myDataAdapter As New MySqlDataAdapter
                'Dim myDataReader As MySqlDataReader
                Dim myCommand As New MySqlCommand()
                Dim myData As New DataTable
                Dim SQLCommand As String = "SELECT * FROM autosrv_autopub_JobList where JOBID = '" & JobID & "'"

                myCommand.Connection = MyConnection
                myCommand.CommandText = SQLCommand
                myDataAdapter.SelectCommand = myCommand
                myDataAdapter.Fill(myData)



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

                MyConnection.Close()
                'Return myData
                Dim Datahash As New Dictionary(Of String, String)
                Datahash.Add("JOBID", myData.Rows(0).Item(0).ToString)
                Datahash.Add("CREATEDBY", myData.Rows(0).Item(1).ToString)
                Datahash.Add("MODIFIEDBY", myData.Rows(0).Item(2).ToString)
                Datahash.Add("JOBNAME", myData.Rows(0).Item(3).ToString)
                Datahash.Add("JOBDESC", myData.Rows(0).Item(4).ToString)
                Datahash.Add("JOBENV", myData.Rows(0).Item(5).ToString)
                Datahash.Add("ACTIVE", myData.Rows(0).Item(6).ToString)
                'Datahash.Add("CAN_DEL_ZIP", myData.Rows(0).Item(7).ToString)
                Return Datahash
            Catch ex As Exception
                Return Nothing
            End Try
        Else
            modErrorHandler.Errors.PrintMessage(2, "[" + strUserName + "] Skipping function - no valid database connection.", strModuleName + "-GetGUID")
            Return Nothing
        End If
    End Function
    Public Sub New()
        MyBase.New()
        If Me.ConnnectToMySQLDB("techpubs.eng.vmware.com", "CMSData", "ditatools", "ditatools") Then
            'Note that this DB can be managed graphically at http://techpubs.eng.vmware.com/phpmyadmin Root login is 'root':monetary bovine
            'connection worked just fine
            Me.DatabaseConnected = True
        Else
            'uh oh - connection failed.
            modErrorHandler.Errors.PrintMessage(3, "Error while connecting to techpubs Database. Tracking and lookup on GUIDs is disabled.", strModuleName)
            Me.DatabaseConnected = False
        End If
    End Sub
    Function GetOutIDsFromJob(ByVal JobID As String, Optional ByVal strUsername As String = "Anonymous") As Hashtable
        If Me.DatabaseConnected = True Then
            Try
                'make sure the connection is open...
                If Not MyConnection.State = ConnectionState.Open Then
                    MyConnection.Open()
                End If

                Dim myDataAdapter As New MySqlDataAdapter
                'Dim myDataReader As MySqlDataReader
                Dim myCommand As New MySqlCommand()
                Dim myData As New DataTable
                Dim SQLCommand As String = "SELECT OUTID FROM autosrv_autopub_Job_Output_Links WHERE JOBID = '" & JobID & "'"

                myCommand.Connection = MyConnection
                myCommand.CommandText = SQLCommand
                myDataAdapter.SelectCommand = myCommand
                myDataAdapter.Fill(myData)
                Dim outids As New Hashtable

                For Each OutputEntry As DataRow In myData.Select()
                    If Not outids.ContainsKey(OutputEntry.Item(0).ToString) Then
                        outids.Add(OutputEntry.Item(0).ToString, OutputEntry.Item(0).ToString)
                    End If
                Next
                'don't need this dataset content any more. Reset.
                myData.Reset()
                Return outids
            Catch ex As Exception
                Return Nothing
            End Try
        Else
            modErrorHandler.Errors.PrintMessage(2, "[" + strUsername + "] Skipping function - no valid database connection.", strModuleName + "-GetOutIDsFromJob")
            Return Nothing
        End If

    End Function
    Function GetJobsOutputData(ByVal JobID As String, Optional ByVal strUserName As String = "Anonymous") As DataTable
        If Me.DatabaseConnected = True Then
            Try
                'make sure the connection is open...
                If Not MyConnection.State = ConnectionState.Open Then
                    MyConnection.Open()
                End If

                Dim myDataAdapter As New MySqlDataAdapter
                'Dim myDataReader As MySqlDataReader
                Dim myCommand As New MySqlCommand()
                Dim myData As New DataTable
                Dim SQLCommand As String = "SELECT OUTID FROM autosrv_autopub_Job_Output_Links WHERE JOBID = '" & JobID & "'"

                myCommand.Connection = MyConnection
                myCommand.CommandText = SQLCommand
                myDataAdapter.SelectCommand = myCommand
                myDataAdapter.Fill(myData)
                Dim outids As New Hashtable

                For Each OutputEntry As DataRow In myData.Select()
                    If Not outids.ContainsKey(OutputEntry.Item(0).ToString) Then
                        outids.Add(OutputEntry.Item(0).ToString, OutputEntry.Item(0).ToString)
                    End If
                Next
                'don't need this dataset content any more. Reset.
                myData.Reset()
                Dim strOUTIDs As New StringBuilder
                Dim i As Integer = 0
                For Each outID As DictionaryEntry In outids
                    i += 1
                    strOUTIDs.Append("'" + outID.Value + "'")
                    If i < outids.Count Then
                        strOUTIDs.Append(", ")
                    End If
                Next
                'Find all the outputs by each OutID
                SQLCommand = "SELECT PUBNAME, PUBDESC, PUBGUID, PUBVER, OUTTYPE, OUTLANG, OUTID, OUTSTATE FROM autosrv_autopub_OutputList WHERE OUTID IN (" & strOUTIDs.ToString & ") AND ACTIVE = 1 ORDER BY PUBNAME, OUTTYPE, OUTLANG"

                myCommand.CommandText = SQLCommand
                myDataAdapter.SelectCommand = myCommand
                'refill the dataset with the new content.
                myDataAdapter.Fill(myData)

                MyConnection.Close()
                Return myData

            Catch ex As Exception
                Return Nothing
            End Try
        Else
            modErrorHandler.Errors.PrintMessage(2, "[" + strUserName + "] Skipping function - no valid database connection.", strModuleName + "-GetJobsOutputData")
            Return Nothing
        End If
    End Function
    ''' <summary>
    ''' Creates a new job in the DB with the specified Title and Description. Returns the resulting JOBID.    '''
    ''' </summary>
    ''' <param name="JobTitle"></param>
    ''' <param name="JobDesc"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function CreateJob(ByVal JobTitle As String, ByVal JobDesc As String, ByVal strEnvironment As String, Optional ByVal strUserName As String = "Anonymous") As String
        If Me.DatabaseConnected = True Then
            Try
                If JobTitleIsUnique(JobTitle, strEnvironment, strUserName) Then
                    'TODO: rewrite this to do stuff for autopublishing instead of what it was doing before...
                    'create a timestamp for today
                    Dim strDateTime As String = Date.Now.ToString("u")
                    'Create a GUID for this JobID
                    Dim GUID As String = "JOBID-" + System.Guid.NewGuid().ToString.ToUpper
                    JobTitle = JobTitle.Replace("\", "\\")
                    JobTitle = JobTitle.Replace("""", "\""")
                    JobTitle = JobTitle.Replace("'", "\'")
                    JobDesc = JobDesc.Replace("\", "\\")
                    JobDesc = JobDesc.Replace("""", "\""")
                    JobDesc = JobDesc.Replace("'", "\'")
                    'Create the MySQL query string
                    Dim SQLCommand As String = "INSERT INTO autosrv_autopub_JobList " & _
                    "(JOBID, CREATEDBY, MODIFIEDBY, JOBNAME, JOBDESC, JOBENV, ACTIVE) " & _
                    "VALUES('" & GUID & "',  '" & strUserName & "',  '" & strUserName & "', '" & JobTitle & "', '" & JobDesc & "', '" & strEnvironment & "', 1)"
                    Dim myDataAdapter As New MySqlDataAdapter
                    Dim myCommand As New MySqlCommand()
                    Dim myData As New DataTable

                    myCommand.Connection = MyConnection
                    myCommand.CommandText = SQLCommand
                    'make sure connection is open before we go forward...
                    If Not MyBase.MyConnection.State = ConnectionState.Open Then
                        MyBase.MyConnection.Open()
                    End If
                    myCommand.ExecuteNonQuery()
                    modErrorHandler.Errors.PrintMessage(1, "[" + strUserName + "] Added entry to DB for " + JobTitle, strModuleName + "-CreateJob")
                    MyBase.MyConnection.Close()
                    'Returns the JOBID.
                    Return GUID
                Else
                    modErrorHandler.Errors.PrintMessage(3, "[" + strUserName + "] Entry already exists in the DB for " + JobTitle, strModuleName + "-CreateJob")
                    MyBase.MyConnection.Close()
                    Return ""
                End If
            Catch ex As Exception
                modErrorHandler.Errors.PrintMessage(3, "[" + strUserName + "] Failed to add entry to DB for " + JobTitle + ". Message: " + ex.Message.ToString, strModuleName + "-CreateJob")
                MyBase.MyConnection.Close()
                Return ""
            End Try
        Else
            modErrorHandler.Errors.PrintMessage(2, "[" + strUserName + "] Skipping function - no valid database connection.", strModuleName + "-CreateJob")
            Return ""
        End If
    End Function

    Private Function JobTitleIsUnique(ByVal JobTitle As String, ByVal strEnvironment As String, Optional ByVal strUserName As String = "Anonymous") As Boolean
        If Me.DatabaseConnected = True Then
            Try
                Dim myData As New DataTable
                'Switch based on content of string
                Dim SQLCommand As String = "SELECT JOBID FROM autosrv_autopub_JobList where JOBNAME = '" & JobTitle & "' AND JOBENV = '" & strEnvironment & "'"
                Dim myDataAdapter As New MySqlDataAdapter
                Dim myCommand As New MySqlCommand()

                myCommand.Connection = MyConnection
                myCommand.CommandText = SQLCommand
                myDataAdapter.SelectCommand = myCommand
                'make sure the connection is open...
                If Not MyConnection.State = ConnectionState.Open Then
                    MyConnection.Open()
                End If
                If Not myDataAdapter.TableMappings.Count = 0 Then
                    myDataAdapter.Fill(myData)
                End If


                MyBase.MyConnection.Close()
                'Check to see if we found any matches.
                If myData.Rows.Count > 0 Then
                    Return False
                Else
                    Return True
                End If

            Catch ex As Exception
                modErrorHandler.Errors.PrintMessage(2, "[" + strUserName + "] Failed to search for entry in DB. Message: " + ex.Message.ToString, strModuleName + "-JobTitleIsUnique")
                MyBase.MyConnection.Close()
                Return False
            End Try
        Else
            modErrorHandler.Errors.PrintMessage(2, "[" + strUserName + "] Skipping function - no valid database connection.", strModuleName + "-JobTitleIsUnique")
            Return False
        End If
    End Function

    Private Function JobEntryExists(ByVal JobID As String, ByVal strUserName As String) As Boolean
        If Me.DatabaseConnected = True Then
            Try
                Dim myData As New DataTable
                'Switch based on content of string
                Dim SQLCommand As String = "SELECT JOBID FROM autosrv_autopub_JobList where JOBID = '" & JobID & "'"
                Dim myDataAdapter As New MySqlDataAdapter
                Dim myCommand As New MySqlCommand()

                myCommand.Connection = MyConnection
                myCommand.CommandText = SQLCommand
                myDataAdapter.SelectCommand = myCommand
                'make sure the connection is open...
                If Not MyConnection.State = ConnectionState.Open Then
                    MyConnection.Open()
                End If
                myDataAdapter.Fill(myData)

                MyBase.MyConnection.Close()
                'Check to see if we found any matches.
                If myData.Rows.Count > 0 Then
                    Return True
                Else
                    Return False
                End If

            Catch ex As Exception
                modErrorHandler.Errors.PrintMessage(2, "[" + strUserName + "] Failed to search for entry in DB. Message: " + ex.Message.ToString, strModuleName + "-JobTitleIsUnique")
                MyBase.MyConnection.Close()
                Return False
            End Try
        Else
            modErrorHandler.Errors.PrintMessage(2, "[" + strUserName + "] Skipping function - no valid database connection.", strModuleName + "-JobTitleIsUnique")
            Return False
        End If
    End Function

    Function CreateOutput(ByVal PubEnv As String, ByVal PubGUID As String, ByVal PubVer As String, ByVal PubName As String, ByVal outtype As String, ByVal lng As String, Optional ByVal strUserName As String = "Anonymous") As String
        'Check to see if the output exists before creating a new one.
        Dim OutID As String = OutputEntryExists(PubEnv, PubGUID, PubVer, lng, outtype, strUserName)
        If Not OutID Is Nothing Then
            'There is an outID in the DB that matches our criteria. Use it!
            Return OutID
        Else
            If Me.DatabaseConnected = True Then
                Try
                    'create a timestamp for today
                    'Dim strDateTime As String = Date.Now.ToString("u")
                    'Create a GUID for this JobID
                    Dim GUID As String = "OUTID-" + System.Guid.NewGuid().ToString.ToUpper
                    'DB Structure: OUTID JOBENV PUBNAME PUBDESC PUBGUID PUBVER OUTLANG OUTTYPE ACTIVE
                    'Create the MySQL query string
                    PubName = PubName.Replace("\", "\\")
                    PubName = PubName.Replace("""", "\""")
                    PubName = PubName.Replace("'", "\'")

                    Dim SQLCommand As String = "INSERT INTO autosrv_autopub_OutputList " & _
                    "(OUTID, JOBENV, PUBNAME, PUBGUID, PUBVER, OUTLANG, OUTTYPE, ACTIVE) " & _
                    "VALUES('" & GUID & "', '" & PubEnv & "', '" & PubName & "', '" & PubGUID & "', '" & PubVer & "', '" & lng & "', '" & outtype & "', 1)"
                    Dim myDataAdapter As New MySqlDataAdapter
                    Dim myCommand As New MySqlCommand()
                    Dim myData As New DataTable

                    myCommand.Connection = MyConnection
                    myCommand.CommandText = SQLCommand
                    'make sure connection is open before we go forward...
                    If Not MyBase.MyConnection.State = ConnectionState.Open Then
                        MyBase.MyConnection.Open()
                    End If
                    myCommand.ExecuteNonQuery()
                    modErrorHandler.Errors.PrintMessage(1, "[" + strUserName + "] Added entry to DB for " + GUID, strModuleName + "-CreateOutput")
                    MyBase.MyConnection.Close()
                    'Returns the OutID.
                    Return GUID
                Catch ex As Exception
                    modErrorHandler.Errors.PrintMessage(3, "[" + strUserName + "] Failed to add entry to DB for " + PubGUID + ", v" + PubVer + ", " + outtype + ". Message: " + ex.Message.ToString, strModuleName + "-CreateOutput")
                    MyBase.MyConnection.Close()
                    Return ""
                End Try
            Else
                modErrorHandler.Errors.PrintMessage(2, "[" + strUserName + "] Skipping function - no valid database connection.", strModuleName + "-CreateOutput")
                Return ""
            End If
        End If

    End Function

    Public Function LinkOutputToJob(ByVal JobID As String, ByVal OutID As String, Optional ByVal strUsername As String = "Anonymous") As Boolean
        If Me.DatabaseConnected = True Then
            Try
                'Dim myData As New DataTable
                'Setup the query:
                Dim SQLCommand As String = "REPLACE INTO autosrv_autopub_Job_Output_Links " & _
                    "(JOBID, OUTID) " & _
                    "VALUES('" & JobID & "',  '" & OutID & "')"
                ExecuteNonQuery(SQLCommand)
                'Update who modified this job:
                UpdateModifiedBy(JobID, strUsername)

                Return True
            Catch ex As Exception
                modErrorHandler.Errors.PrintMessage(2, "[" + strUsername + "] Failed to insert linked entry in DB. Message: " + ex.Message.ToString, strModuleName + "-LinkOutputToJob")
                MyBase.MyConnection.Close()
                Return False
            End Try
        Else
            modErrorHandler.Errors.PrintMessage(2, "[" + strUsername + "] Skipping function - no valid database connection.", strModuleName + "-LinkOutputToJob")
            Return False
        End If
    End Function
    Public Function RenameJob(ByVal strEnvironment As String, ByVal JobID As String, ByVal strNewName As String, ByVal strNewDesc As String, ByVal strUserName As String) As Boolean
        If JobTitleIsUnique(strNewName, strEnvironment, strUserName) Then
            Dim SQLQuery As String = "UPDATE autosrv_autopub_JobList SET JOBNAME='" + strNewName + "', JOBDESC='" + strNewDesc + "' WHERE JOBID='" + JobID + "'"
            ExecuteNonQuery(SQLQuery)
            'Update who modified this job:
            UpdateModifiedBy(JobID, strUserName)
            Return True

        End If

    End Function
    Public Function UnlinkOutputfromJob(ByVal strJobID As String, ByVal strOutID As String, ByVal strUsername As String) As Boolean
        If Me.DatabaseConnected = True Then
            Try
                'Dim myData As New DataTable
                'Setup the query:
                Dim SQLCommand As String = "DELETE FROM autosrv_autopub_Job_Output_Links where JOBID = '" & strJobID & "' AND OUTID = '" & strOutID & "'"
                ExecuteNonQuery(SQLCommand)
                Dim linkedjobs As Hashtable = GetJobIDsFromOutput(strOutID, strUsername)
                If linkedjobs.Count > 0 Then
                    'There are still some jobs linked to this output. Don't delete it.
                Else
                    'No jobs reference this output. Let's clean it up.
                    DeleteOutputEntrybyID(strOutID, strUsername)
                End If
                'Update who modified this job:
                UpdateModifiedBy(strJobID, strUsername)
                Return True
            Catch ex As Exception
                modErrorHandler.Errors.PrintMessage(2, "[" + strUsername + "] Failed to unlink entry in DB. Message: " + ex.Message.ToString, strModuleName + "-UnlinkOutputfromJob")
                MyBase.MyConnection.Close()
                Return False
            End Try
        Else
            modErrorHandler.Errors.PrintMessage(2, "[" + strUsername + "] Skipping function - no valid database connection.", strModuleName + "-UnlinkOutputfromJob")
            Return False
        End If
    End Function

    Private Sub UpdateModifiedBy(ByVal JobID As String, ByVal strUsername As String)
        Dim SQLCommand As String = "UPDATE autosrv_autopub_JobList SET MODIFIEDBY = '" + strUsername + "' WHERE JOBID = '" + JobID + "' LIMIT 1"
        ExecuteNonQuery(SQLCommand)
    End Sub


    Public Sub UpdateJobStates(ByVal JobID As String)
        'Build a dictionary of objects containing all of the outputs in the DB
        Dim OutputDict As Dictionary(Of String, objOutput)
        OutputDict = GetOutListFromJob(JobID)
        If IsNothing(OutputDict) Then
            Return
        End If
        'iterate through the dictionary
        'Dim counter As Integer = 0
        Dim ServerEnv As String = GetJobEnv(JobID)
        Dim objIshObj As VMwareISHModulesNS.clsISHObj = modSettings.GetIshObj(ServerEnv)
        For Each output As KeyValuePair(Of String, objOutput) In OutputDict
            'For each entry, if the last known state was not "Released", get the current state

            If Not output.Value.OutState = "Released" Then
                'Get latest state of the output ID:

                Dim state As String = objIshObj.oPubOutput.GetOutputState(output.Value.PubGUID, output.Value.PubVer, output.Value.OutLang, output.Value.OutType)
                If IsNothing(state) Then
                    state = "UNKNOWN"
                End If
                UpdateState(output.Value.OutID, state)

            End If
        Next
        objIshObj = Nothing
    End Sub
    Private Function GetOutListFromJob(ByVal JobID As String) As Dictionary(Of String, objOutput)
        Dim myDict As New Dictionary(Of String, objOutput)
        'Query the DB to get a table of all the output entries
        Dim OutputTable As DataTable = GetJobsOutputData(JobID)
        'TODO: Rather than relying on the data structure in the DB, it would be better to find each value by a name rather than a column number...
        'Convert each row to an object and add it to our dictionary
        Try
            For Each outrow As DataRow In OutputTable.Select()
                Dim newobj As objOutput
                'PUBNAME, PUBDESC, PUBGUID, PUBVER, OUTTYPE, OUTLANG, OUTID, OUTSTATE
                newobj.PubGUID = outrow.Item(2).ToString
                newobj.PubVer = outrow.Item(3).ToString
                newobj.OutType = outrow.Item(4).ToString
                newobj.OutLang = outrow.Item(5).ToString
                newobj.OutID = outrow.Item(6).ToString
                newobj.OutState = outrow.Item(7).ToString
                newobj.ServerEnv = GetOutputServEnv(newobj.OutID)
                myDict.Add(newobj.OutID, newobj)
            Next
            'return the dictionary
            Return myDict
        Catch ex As Exception
            Return Nothing
        End Try

    End Function

    Public Sub UpdateState(ByVal OutID As String, ByVal state As String)
        Dim querystring As String = "UPDATE autosrv_autopub_OutputList SET OUTSTATE = '" + state + "' WHERE OUTID='" + OutID + "'"
        ExecuteNonQuery(querystring)
    End Sub
    Public Sub UpdateCanDelZip(ByVal JobID As String, ByVal BoolSetting As String)
        Dim mybool As Integer = 0

        If BoolSetting = "True" Then
            mybool = 1
        End If
        If BoolSetting = "False" Then
            mybool = 0
        End If
        Dim querystring As String = "UPDATE autosrv_autopub_JobList SET CAN_DEL_ZIP = " & mybool & " WHERE JOBID='" & JobID & "'"
        ExecuteNonQuery(querystring)
    End Sub

    Function DuplicateJob(ByVal strJobID As String, Optional ByVal strUserName As String = "Anonymous") As String
        If Me.DatabaseConnected = True Then
            Dim jobdesc As String = GetJobDesc(strJobID)
            Dim jobtitle As String = GetJobName(strJobID) + " - Copy"
            Dim strEnvironment As String = GetJobEnv(strJobID)
            Try

                Dim strDateTime As String = Date.Now.ToString("u")
                'Create a GUID for this JobID
                Dim GUID As String = "JOBID-" + System.Guid.NewGuid().ToString.ToUpper
                'JobTitle = JobTitle.Replace("\", "\\")
                'JobTitle = JobTitle.Replace("""", "\""")
                'JobTitle = JobTitle.Replace("'", "\'")
                'JobDesc = JobDesc.Replace("\", "\\")
                'JobDesc = JobDesc.Replace("""", "\""")
                'JobDesc = JobDesc.Replace("'", "\'")
                'Create the MySQL query string
                Dim SQLCommand As String = "INSERT INTO autosrv_autopub_JobList " & _
                "(JOBID, CREATEDBY, MODIFIEDBY, JOBNAME, JOBDESC, JOBENV, ACTIVE) " & _
                "VALUES('" & GUID & "',  '" & strUserName & "',  '" & strUserName & "', '" & jobtitle & "', '" & jobdesc & "', '" & strEnvironment & "', 1)"
                Dim myDataAdapter As New MySqlDataAdapter
                Dim myCommand As New MySqlCommand()
                Dim myData As New DataTable
                myCommand.Connection = MyConnection
                myCommand.CommandText = SQLCommand
                'make sure connection is open before we go forward...
                If Not MyBase.MyConnection.State = ConnectionState.Open Then
                    MyBase.MyConnection.Open()
                End If
                myCommand.ExecuteNonQuery()
                modErrorHandler.Errors.PrintMessage(1, "[" + strUserName + "] Added entry to DB for " + jobtitle, strModuleName + "-DuplicateJob")
                MyBase.MyConnection.Close()

                'Add the outputs from the old job to the newly created job.
                CopyOutputs(strJobID, GUID)

                'Returns the JOBID.
                Return GUID
            Catch ex As Exception
                modErrorHandler.Errors.PrintMessage(3, "[" + strUserName + "] Failed to add entry to DB for " + jobtitle + ". Message: " + ex.Message.ToString, strModuleName + "-DuplicateJob")
                MyBase.MyConnection.Close()
                Return ""
            End Try
        Else
            modErrorHandler.Errors.PrintMessage(2, "[" + strUserName + "] Skipping function - no valid database connection.", strModuleName + "-DuplicateJob")
            Return ""
        End If


    End Function

    Public Sub CopyOutputs(ByVal strSourceJobID As String, ByVal strTargetJobID As String)
        'Get a list of OutIDs associated with the source job.
        Dim outids As New Hashtable
        outids = GetOutIDsFromJob(strSourceJobID)
        'for each outid gathered, create a link with the target job.
        For Each outid As String In outids.Values
            LinkOutputToJob(strTargetJobID, outid)
        Next
    End Sub


    Public Structure objOutput
        Public ServerEnv As String
        Public PubGUID As String
        Public PubVer As String
        Public OutLang As String
        Public OutType As String
        Public OutState As String
        Public OutID As String
    End Structure

End Class
