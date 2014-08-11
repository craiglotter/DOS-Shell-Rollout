Imports System.IO
Imports Microsoft.Win32

Public Class Worker

    Inherits System.ComponentModel.Component

    ' Declares the variables you will use to hold your thread objects.

    Public WorkerThread As System.Threading.Thread

    Public dbuser As String
    Public dbpassword As String
    Public dbtable As String
    Public dbserver As String

    Public SQLstatement As String

    Public error_level As String = "none"
    Public DOSObject As DOS_Object
    Public serverstatus As Windows.Forms.TextBox

    Public result As String = ""

    Private blocklistvalue As ArrayList
    Private blocklistcount As ArrayList


    Public Event WorkerErrorEncountered(ByVal ex As Exception, ByVal descriptor As String)
    Public Event WorkerComplete(ByVal Result As String)
    Public Event WorkerStatusMessageUpdate(ByVal message As String)



#Region " Component Designer generated code "

    Public Sub New(ByVal Container As System.ComponentModel.IContainer)
        MyClass.New()

        'Required for Windows.Forms Class Composition Designer support
        Container.Add(Me)
    End Sub

    Public Sub New()
        MyBase.New()

        'This call is required by the Component Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    
        error_level = "none"
        DOSObject = New DOS_Object
        serverstatus = New Windows.Forms.TextBox
        dbuser = ""
        dbpassword = "commerceitstaff"
        dbtable = ""
        dbserver = ""
        SQLstatement = ""
        blocklistvalue = New ArrayList
        blocklistcount = New ArrayList
    End Sub

    'Component overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Component Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Component Designer
    'It can be modified using the Component Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        components = New System.ComponentModel.Container
    End Sub

#End Region

    Private Sub Error_Handler(ByVal exc As Exception, Optional ByVal descriptor As String = "")
        Try
            RaiseEvent WorkerErrorEncountered(exc, descriptor)
        Catch ex As Exception
            MsgBox("An error occurred in DOS Shell Rollout's error handling routine. The application will try to recover from this serious error.", MsgBoxStyle.Critical, "Critical Error Encountered")
        End Try
    End Sub


    Public Sub ChooseThreads(ByVal threadNumber As Integer)
        Try
            ' Determines which thread to start based on the value it receives.

            Select Case threadNumber
                Case 1
                    ' Sets the thread using the AddressOf the subroutine where
                    ' the thread will start.
                    WorkerThread = New System.Threading.Thread(AddressOf WorkerExecute)
                    ' Starts the thread.
                    WorkerThread.Start()
                Case 3
                    WorkerThread = New System.Threading.Thread(AddressOf WorkerConnectionTest)
                    WorkerThread.Start()
            End Select
        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub



    Public Sub WorkerExecute()
        Dim result As String = "No Active Updates Required"
        Try
            WorkerConnectionTest()
            If serverstatus.Text.IndexOf("Offline") = -1 Then

                Dim oldcmds, newcmds, executecmds As ArrayList
                oldcmds = New ArrayList
                newcmds = New ArrayList
                executecmds = New ArrayList
                oldcmds.Clear()
                newcmds.Clear()
                executecmds.Clear()



                Dim filewriter As StreamWriter
                If File.Exists((Application.StartupPath & "\").Replace("\\", "\") & "activity.log") = False Then
                    Try
                        filewriter = New StreamWriter((Application.StartupPath & "\").Replace("\\", "\") & "activity.log", False)
                        filewriter.Write("")
                        filewriter.Close()
                    Catch ex As Exception
                        Error_Handler(ex)
                        result = "Failure. Unable to create new Activity Log"
                    End Try
                End If

                If File.Exists((Application.StartupPath & "\").Replace("\\", "\") & "activity.log") = True Then
                    Try


                        Dim filereader As IO.StreamReader = New StreamReader((Application.StartupPath & "\").Replace("\\", "\") & "activity.log")
                        While filereader.Peek > -1
                            Try
                                oldcmds.Add(filereader.ReadLine.Trim())
                            Catch ex As Exception
                                Error_Handler(ex)
                                result = "Failure. Unable to read a specified line in existing activity Activity Log"
                            End Try

                        End While
                        filereader.Close()
                    Catch ex As Exception
                        Error_Handler(ex)
                        result = "Failure. Unable to open existing activity Activity Log"
                    End Try
                End If

                '  oldcmds.Sort()
                Dim conn As OleDb.OleDbConnection
                Try
                    conn = Get_Connection()
                    conn.Open()

                    Dim counter As Integer = 0
                    Dim sql As OleDb.OleDbCommand = New OleDb.OleDbCommand
                    Dim datareader As OleDb.OleDbDataReader


                    sql = New OleDb.OleDbCommand
                    sql.CommandText = "Select * from [DOS_Transactions] where [Transaction_Active] = 'Active' order by [Transaction_Ranking] asc, [Transaction_ID] asc"
                    sql.Connection = conn
                    datareader = sql.ExecuteReader(CommandBehavior.Default)
                    If datareader.HasRows = True Then
                        While datareader.Read = True
                            Try


                                'Dim dos As New DOS_Object
                                'dos.Transaction_ID = datareader.Item("Transaction_ID").ToString.Trim
                                'dos.Transaction_Command = datareader.Item("Transaction_Command").ToString.Trim
                                'dos.Transaction_Ranking = datareader.Item("Transaction_Ranking").ToString.Trim
                                'dos.Transaction_Modifier = datareader.Item("Transaction_Modifier").ToString.Trim
                                'dos.Transaction_Date = datareader.Item("Transaction_Date").ToString.Trim
                                'dos.Transaction_Time = datareader.Item("Transaction_Time").ToString.Trim
                                'dos.Transaction_Active = datareader.Item("Transaction_Active").ToString.Trim
                                newcmds.Add(datareader.Item("Transaction_ID").ToString.Trim)
                                'executecmds.Add(dos)
                            Catch ex As Exception
                                Error_Handler(ex)
                                result = "Failure. Unable to read specific data result"
                            End Try
                        End While
                    End If
                    datareader.Close()
                    sql.Dispose()

                Catch ex As Exception
                    Error_Handler(ex)
                    result = "Failure. Unable to retrieve transaction list from Database"
                Finally
                    conn.Close()
                End Try

                'newcmds.Sort()
                'Dim it As Integer = 0

                For Each obj As Object In oldcmds
                    If newcmds.IndexOf(obj.ToString.Trim) > -1 Then
                        newcmds.Remove(obj)
                    End If
                    '   it = it + 1
                Next

                ' MsgBox(newcmds.Count)

                Try
                    For Each obj As Object In newcmds

                        conn = Get_Connection()
                        conn.Open()

                        Dim counter As Integer = 0
                        Dim sql As OleDb.OleDbCommand = New OleDb.OleDbCommand
                        Dim datareader As OleDb.OleDbDataReader


                        sql = New OleDb.OleDbCommand
                        sql.CommandText = "Select * from [DOS_Transactions] where [Transaction_ID] = '" & obj.ToString.Trim & "'"
                        sql.Connection = conn
                        datareader = sql.ExecuteReader(CommandBehavior.Default)
                        If datareader.HasRows = True Then
                            While datareader.Read = True
                                Try


                                    Dim dos As New DOS_Object
                                    dos.Transaction_ID = datareader.Item("Transaction_ID").ToString.Trim
                                    dos.Transaction_Command = datareader.Item("Transaction_Command").ToString.Trim
                                    dos.Transaction_Ranking = datareader.Item("Transaction_Ranking").ToString.Trim
                                    dos.Transaction_Modifier = datareader.Item("Transaction_Modifier").ToString.Trim
                                    dos.Transaction_Date = datareader.Item("Transaction_Date").ToString.Trim
                                    dos.Transaction_Time = datareader.Item("Transaction_Time").ToString.Trim
                                    dos.Transaction_Active = datareader.Item("Transaction_Active").ToString.Trim

                                    executecmds.Add(dos)
                                Catch ex As Exception
                                    Error_Handler(ex)
                                    result = "Failure. Unable to read specific data result"
                                End Try
                            End While
                        End If
                        datareader.Close()
                        sql.Dispose()


                    Next
                Catch ex As Exception
                    Error_Handler(ex)
                    result = "Failure. Unable to retrieve transaction list from Database"
                Finally
                    conn.Close()
                End Try
                filewriter = New StreamWriter((Application.StartupPath & "\").Replace("\\", "\") & "activity.log", True)
                For Each obj As DOS_Object In executecmds
                    If obj Is Nothing = False Then

                        Dim proceed As Boolean = False
                        Try
                            'MsgBox(blocklistvalue.Count & " " & blocklistcount.Count)
                            'If blocklistvalue.IndexOf(obj.Transaction_ID) > -1 Then
                            'MsgBox(blocklistvalue.Item(blocklistvalue.IndexOf(obj.Transaction_ID)) & " " & blocklistcount.Item(blocklistvalue.IndexOf(obj.Transaction_ID)))
                            'End If
                            If blocklistvalue.IndexOf(obj.Transaction_ID) = -1 Then
                                proceed = True
                            Else
                                If CInt(blocklistcount.Item(blocklistvalue.IndexOf(obj.Transaction_ID))) < 5 Then
                                    proceed = True
                                Else
                                    proceed = False
                                End If
                            End If
                            If proceed = True Then
                                Dim apptorun As String
                                apptorun = obj.Transaction_Command.ToString.Trim
                                Try
                                    result = ApplicationLauncher(apptorun)
                                Catch ex As Exception
                                    Error_Handler(ex, "Application Launcher")
                                    result = "Fail"
                                End Try

                                If Not result = "Fail" Then
                                    filewriter.WriteLine(obj.Transaction_ID)
                                    RaiseEvent WorkerStatusMessageUpdate(Format(Now(), "dd/MM/yyyy HH:mm:ss") & " - [" & obj.Transaction_ID & "] " & apptorun & " (Pass)")
                                    result = "Success. Active Update Successfully Applied"
                                Else
                                    RaiseEvent WorkerStatusMessageUpdate(Format(Now(), "dd/MM/yyyy HH:mm:ss") & " - [" & obj.Transaction_ID & "] " & apptorun & " (Fail)")
                                    result = "Failure. Unable to Apply Active Update"
                                    If blocklistvalue.IndexOf(obj.Transaction_ID) = -1 Then
                                        blocklistvalue.Add(obj.Transaction_ID)
                                        blocklistcount.Add(1)
                                    Else
                                        blocklistcount.Item(blocklistvalue.IndexOf(obj.Transaction_ID)) = CInt(blocklistcount.Item(blocklistvalue.IndexOf(obj.Transaction_ID))) + 1
                                        If CInt(blocklistcount.Item(blocklistvalue.IndexOf(obj.Transaction_ID))) = 5 Then
                                            RaiseEvent WorkerStatusMessageUpdate(Format(Now(), "dd/MM/yyyy HH:mm:ss") & " - [" & obj.Transaction_ID & "] " & apptorun & " (Blocked this Session)")
                                        End If

                                    End If
                                End If
                            End If
                        Catch ex As Exception
                            Error_Handler(ex)
                            result = "Failure. DOS Command Error"
                        End Try
                    End If
                Next
                filewriter.Close()

            Else
                result = "Failure. Database Server cannot be accessed."
            End If
        Catch ex As Exception
            Error_Handler(ex)
            result = "Failure. Unknown Reason."
        Finally
            RaiseEvent WorkerComplete(result)
        End Try
    End Sub




    Protected Friend Function Get_Connection() As OleDb.OleDbConnection
        'Standard(Security)
        '"Provider=sqloledb;Data Source=Aron1;Initial Catalog=pubs;User Id=sa;Password=asdasd;" 

        'Trusted(Connection)
        '"Provider=sqloledb;Data Source=Aron1;Initial Catalog=pubs;Integrated Security=SSPI;" 
        '(use serverName\instanceName as Data Source to use an specifik SQLServer instance, only SQLServer2000)

        'Prompt for username and password:
        'oConn.Provider = "sqloledb"
        'oConn.Properties("Prompt") = adPromptAlways
        'oConn.Open("Data Source=Aron1;Initial Catalog=pubs;")

        'Connect via an IP address:
        '"Provider=sqloledb;Data Source=190.190.200.100,1433;Network Library=DBMSSOCN;Initial Catalog=pubs;User ID=sa;Password=asdasd;" 
        '(DBMSSOCN=TCP/IP instead of Named Pipes, at the end of the Data Source is the port to use (1433 is the default))

        Dim connection_string As String
        If dbserver.IndexOf(".") = -1 Then
            connection_string = "Provider=sqloledb;Data Source=" & dbserver & ";Initial Catalog=" & dbtable & ";User Id=" & dbuser & ";Password=" & dbpassword & ";"
        Else
            connection_string = "Provider=sqloledb;Data Source=" & dbserver & ",1433;Network Library=DBMSSOCN;Initial Catalog=" & dbtable & ";User Id=" & dbuser & ";Password=" & dbpassword & ";"
        End If
        'Dim connection_string As String = "User ID=" & dbuser & ";password=" & dbpassword & ";Data Source=" & dbserver & ";Tag with column collation when possible=False;Initial Catalog=" & dbtable & ";Use Procedure for Prepare=1;Auto Translate=True;Persist Security Info=False;Provider=""SQLOLEDB.1"";Use Encryption for Data=False;Packet Size=4096"

        Dim conn As OleDb.OleDbConnection = New OleDb.OleDbConnection(connection_string)
        Return conn
    End Function


    Private Function ApplicationLauncher(ByVal apptoRun As String) As String
        Dim sresult As String = ""
        Try
            Dim myProcess As Process = New Process

            Dim executable, arguments As String
            Dim str As String()
            executable = ""
            arguments = ""
            If apptoRun.StartsWith("""") = True Then
                Dim endpos As Integer = apptoRun.IndexOf("""", apptoRun.IndexOf("""", 0) + 1)
                executable = apptoRun.Substring(0, endpos + 1)
                If apptoRun.Length >= (endpos + 3) Then
                    arguments = apptoRun.Substring(endpos + 2)
                End If
            Else
                str = apptoRun.Split(" ")
                For i As Integer = 0 To str.Length - 1
                    If i = 0 Then
                        executable = str(i)
                    Else
                        arguments = arguments & str(i) & " "
                    End If
                Next
                arguments = arguments.Remove(arguments.Length - 1, 1)
            End If
            Activity_Logger("LAUNCH ATTEMPT INITIATED")
            
            myProcess.StartInfo.FileName = executable.Replace("""", "")
            myProcess.StartInfo.Arguments = arguments
            Activity_Logger("Executable: " & myProcess.StartInfo.FileName)
            Activity_Logger("Arguments: " & myProcess.StartInfo.Arguments)
            myProcess.StartInfo.UseShellExecute = True

            myProcess.StartInfo.CreateNoWindow = False
            myProcess.StartInfo.WindowStyle = ProcessWindowStyle.Normal

            myProcess.StartInfo.RedirectStandardInput = False
            myProcess.StartInfo.RedirectStandardOutput = False
            myProcess.StartInfo.RedirectStandardError = False
            myProcess.Start()
            sresult = "Success"
            Return sresult

        Catch ex As Exception
            Error_Handler(ex, "Executing: " & apptoRun)
            sresult = "Fail"
        End Try
        Return sresult
    End Function

    'Private Function DosShellCommand(ByVal AppToRun As String) As String
    '    Dim s As String = ""
    '    Try
    '        Dim myProcess As Process = New Process

    '        myProcess.StartInfo.FileName = "cmd.exe"
    '        myProcess.StartInfo.UseShellExecute = False

    '        Dim sErr As StreamReader
    '        Dim sOut As StreamReader
    '        Dim sIn As StreamWriter


    '        myProcess.StartInfo.CreateNoWindow = True

    '        myProcess.StartInfo.RedirectStandardInput = True
    '        myProcess.StartInfo.RedirectStandardOutput = True
    '        myProcess.StartInfo.RedirectStandardError = True

    '        'myProcess.StartInfo.FileName = AppToRun
    '        'myProcess.Start()

    '        myProcess.Start()
    '        sIn = myProcess.StandardInput
    '        sIn.WriteLine(AppToRun)
    '        sIn.AutoFlush = True

    '        sOut = myProcess.StandardOutput()
    '        sErr = myProcess.StandardError

    '        sIn.Write(AppToRun & System.Environment.NewLine)
    '        sIn.Write("exit" & System.Environment.NewLine)
    '        s = sOut.ReadToEnd()

    '        If Not myProcess.HasExited Then
    '            myProcess.Kill()
    '        End If



    '        sIn.Close()
    '        sOut.Close()
    '        sErr.Close()
    '        myProcess.Close()


    '    Catch ex As Exception
    '        Error_Handler(ex, "Executing: " & AppToRun)
    '        s = "Fail"
    '    End Try
    '    Return s
    'End Function

    Private Sub WorkerConnectionTest()

        Dim conn As OleDb.OleDbConnection = Get_Connection()

        Try
            serverstatus.ForeColor = Color.Orange
            serverstatus.Text = "SQL Server Status: Checking"
            conn.Open()
            conn.Close()
            conn.Dispose()
            serverstatus.ForeColor = Color.Green
            serverstatus.Text = "SQL Server Status: Online"
            result = "Server Check Succeeded"
        Catch ex As Exception
            serverstatus.ForeColor = Color.Red
            serverstatus.Text = "SQL Server Status: Offline"
            Error_Handler(ex)
            result = "Server Check Failed"
        Finally

            conn.Dispose()
            ' RaiseEvent WorkerComplete(result)
        End Try

    End Sub

    Private Sub Activity_Logger(ByVal identifier_msg As String)
        Try
            Dim dir As DirectoryInfo = New DirectoryInfo((Application.StartupPath & "\").Replace("\\", "\") & "Activity Logs")
            If dir.Exists = False Then
                dir.Create()
            End If
            Dim filewriter As StreamWriter = New StreamWriter((Application.StartupPath & "\").Replace("\\", "\") & "Activity Logs\" & Format(Now(), "yyyyMMdd") & "_Activity_Log.txt", True)

            filewriter.WriteLine("#" & Format(Now(), "dd/MM/yyyy HH:mm:ss") & " - " & identifier_msg)


            filewriter.Flush()
            filewriter.Close()

        Catch exc As Exception
            Error_Handler(exc, "Activity_Logger")
        End Try
    End Sub


End Class
