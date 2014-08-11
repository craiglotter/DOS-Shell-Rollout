Imports System.IO

Public Class DOS_Object
    Public Transaction_ID As String
    Public Transaction_Command As String
    Public Transaction_Ranking As String
    Public Transaction_Modifier As String
    Public Transaction_Date As String
    Public Transaction_Time As String
    Public Transaction_Active As String

    Public Sub New()
        Try
            Transaction_ID = ""
            Transaction_Command = ""
            Transaction_Ranking = ""
            Transaction_Modifier = ""
            Transaction_Date = ""
            Transaction_Time = ""
            Transaction_Active = ""
        Catch ex As Exception
            Error_Handler(ex, "New Complaint Declaration")
        End Try
    End Sub

    Private Sub Error_Handler(ByVal ex As Exception, Optional ByVal identifier_msg As String = "")
        Try
            If ex.Message.IndexOf("Thread was being aborted") < 0 Then
                Dim Display_Message1 As New Display_Message("The Application encountered the following problem: " & vbCrLf & identifier_msg & ":" & ex.ToString)
                Display_Message1.ShowDialog()
                Dim dir As DirectoryInfo = New DirectoryInfo((Application.StartupPath & "\").Replace("\\", "\") & "Error Logs")
                If dir.Exists = False Then
                    dir.Create()
                End If
                Dim filewriter As StreamWriter = New StreamWriter((Application.StartupPath & "\").Replace("\\", "\") & "Error Logs\" & Format(Now(), "yyyyMMdd") & "_Error_Log.txt", True)
                filewriter.WriteLine("#" & Format(Now(), "dd/MM/yyyy hh:mm:ss tt") & " - " & identifier_msg & ":" & ex.ToString)
                filewriter.Flush()
                filewriter.Close()
            End If
        Catch exc As Exception
            MsgBox("An error occurred in DOS Shell Rollout's error handling routine. The application will try to recover from this serious error.", MsgBoxStyle.Critical, "Critical Error Encountered")
        End Try
    End Sub

    


    Protected Friend Sub Clear_Data()
        Try
            Transaction_ID = ""
            Transaction_Command = ""
            Transaction_Ranking = ""
            Transaction_Modifier = ""
            Transaction_Date = ""
            Transaction_Time = ""
            Transaction_Active = ""
        Catch ex As Exception
            Error_Handler(ex, "Clear Complaint")
        End Try
    End Sub

    
End Class


