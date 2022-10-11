Attribute VB_Name = "ModCamera"
Public Const WS_CHILD As Long = &H40000000
Public Const WS_VISIBLE As Long = &H10000000
Public Const WM_USER As Long = &H400
Public Const WM_CAP_START As Long = WM_USER
Public Const WM_CAP_DRIVER_CONNECT As Long = WM_CAP_START + 10
Public Const WM_CAP_DRIVER_DISCONNECT As Long = WM_CAP_START + 11
Public Const WM_CAP_SET_PREVIEW As Long = WM_CAP_START + 50
Public Const WM_CAP_SET_PREVIEWRATE As Long = WM_CAP_START + 52
Public Const WM_CAP_DLG_VIDEOFORMAT As Long = WM_CAP_START + 41
Public Const WM_CAP_FILE_SAVEDIB As Long = WM_CAP_START + 25
Public Declare Function capCreateCaptureWindow _
Lib "avicap32.dll" Alias "capCreateCaptureWindowA" _
(ByVal lpszWindowName As String, ByVal dwStyle As Long _
, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long _
, ByVal nHeight As Long, ByVal hwndParent As Long _
, ByVal nID As Long) As Long
Public Declare Function SendMessage Lib "user32" _
Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long _
, ByVal wParam As Long, ByRef lParam As Any) As Long

Public eTo As String, eCC As String, eSubject As String, eFromName As String, eFromEmail As String, eMsg As String, eServer As String
Public ePort As Integer, eUsername As String, ePassword As String, eSSL As Boolean


Public Function SendMail(sTo As String, sSubject As String, sFrom As String, _
    sBody As String, sSmtpServer As String, iSmtpPort As Integer, _
    sSmtpUser As String, sSmtpPword As String, _
    bSmtpSSL As Boolean, sCC As String, Optional sAttach As String) As String
      
    On Error GoTo SendMail_Error:
    Dim lobj_cdomsg      As CDO.Message
    Set lobj_cdomsg = New CDO.Message
    lobj_cdomsg.Configuration.Fields(cdoSMTPServer) = sSmtpServer
    lobj_cdomsg.Configuration.Fields(cdoSMTPServerPort) = iSmtpPort
    lobj_cdomsg.Configuration.Fields(cdoSMTPUseSSL) = bSmtpSSL
    lobj_cdomsg.Configuration.Fields(cdoSMTPAuthenticate) = cdoBasic
    lobj_cdomsg.Configuration.Fields(cdoSendUserName) = sSmtpUser
    lobj_cdomsg.Configuration.Fields(cdoSendPassword) = sSmtpPword
    lobj_cdomsg.Configuration.Fields(cdoSMTPConnectionTimeout) = 30
    lobj_cdomsg.Configuration.Fields(cdoSendUsingMethod) = cdoSendUsingPort
    lobj_cdomsg.Configuration.Fields.Update
    lobj_cdomsg.To = sTo
    lobj_cdomsg.CC = sCC
    lobj_cdomsg.From = sFrom
    lobj_cdomsg.Subject = sSubject
    lobj_cdomsg.HTMLBody = sBody
    lobj_cdomsg.AddAttachment sAttach
    lobj_cdomsg.Send
    Set lobj_cdomsg = Nothing
    SendMail = "ok"
    Exit Function
          
SendMail_Error:
    SendMail = Err.Description
End Function

'Procedure save log
Public Sub Save_log(ByVal sLog As String)
    Open App.Path & "\log\" & Format(Now, "YYYYMMDD") & ".DAT" For Append As #1
        Print #1, Time & " > " & sLog
    Close #1
End Sub


