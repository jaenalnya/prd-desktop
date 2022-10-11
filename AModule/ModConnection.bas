Attribute VB_Name = "ModConnection"
Option Explicit

Public CN                       As New ADODB.Connection
Public CN_SIP                   As New ADODB.Connection

Public Const APPNAME = "TSSI-PRD"
Public Const APP_CATEGORY = "Application"
Public Const APP_MSG_NAME = "TSSI-PRD"

Public Function Connected2DB() As Boolean
Dim isOpen                  As Boolean
Dim Reply                   As VbMsgBoxResult


Dim db_server As String
Dim db_port As String
isOpen = False
On Error GoTo ERR_CONNECTION

db_server = ReadINI("Server", "DBSERVER", App.Path & "\Database.ini")
db_port = ReadINI("Server", "DBPORT", App.Path & "\Database.ini")

Do Until isOpen = True
    Set CN = New ADODB.Connection
    CN.CursorLocation = adUseClient
    CN.ConnectionString = "Driver={MySQL ODBC 3.51 Driver}; " & _
                        "SERVER= " & db_server & "; " & _
                        "PWD= sip_production_passwd; " & _
                        "UID= sip_production_user; " & _
                        "PORT= " & db_port & "; " & _
                        "DATABASE= sip_production;OPTION=3;"
    CN.Open

    isOpen = True
                
Loop
    Connected2DB = isOpen

Exit Function

ERR_CONNECTION:
    Reply = MsgBox("Error Number:" & Err.Number & vbNewLine & "Description:" & Err.Description, vbExclamation + vbRetryCancel, "COnnection Failure")
    If Reply = vbCancel Then
        Connected2DB = False
        End
    ElseIf Reply = vbRetry Then
        End
    End If
End Function

Public Function Connected2SIP() As Boolean
Dim isOpen                  As Boolean
Dim Reply                   As VbMsgBoxResult


Dim db_server_2 As String
db_server_2 = ReadINI("Server", "DBSERVER2", App.Path & "\Database.ini")

isOpen = False
On Error GoTo ERR_CONNECTION

Do Until isOpen = True
    Set CN_SIP = New ADODB.Connection
    CN_SIP.CursorLocation = adUseClient
    CN_SIP.ConnectionString = "Driver={MySQL ODBC 3.51 Driver}; " & _
                        "SERVER=" & db_server_2 & "; " & _
                        "PWD= Kmzway87aa; " & _
                        "UID= aden; " & _
                        "PORT= 3306 ; " & _
                        "DATABASE= sip_234;OPTION=3;"
    CN_SIP.Open

    isOpen = True
                
Loop
    Connected2SIP = isOpen

Exit Function

ERR_CONNECTION:
    Reply = MsgBox("Error Number:" & Err.Number & vbNewLine & "Description:" & Err.Description, vbExclamation + vbRetryCancel, "COnnection Failure")
    If Reply = vbCancel Then
        Connected2SIP = False
    ElseIf Reply = vbRetry Then
        Resume
    End If
End Function

Public Sub CloseMySQL()
CN.Close
Set CN = Nothing
End Sub

