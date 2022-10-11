Attribute VB_Name = "ModVarTypes"
Public Declare Function InitCommonControlsEx Lib "comctl32.dll" (iccex As tagInitCommonControlsEx) As Boolean

Public Const ICC_USEREX_CLASSES        As Long = &H200

Public Type tagInitCommonControlsEx
   lngSize As Long
   lngICC As Long
End Type


Public Type USER_INFO
    SYSID                               As String
    KODEUSER                            As String
    KODEPIN                             As String
    KODENIK                             As String
    USERNAME                            As String
    PASSWORD                            As String
    FULLNAME                            As String
    USERTYPE                            As String
    USER_ISADMIN                        As Boolean
    ISADMIN                             As Byte
End Type

Public Enum FORM_STATE
    AddStateMode = 0
    EditStateMode = 1
End Enum

Public Type COMPANY_INFO
    IDPerusahaan                        As String
    Perusahaan                          As String
    Alamat                              As String
    NoTelepon                           As String
    NoFax                               As String
    Email                               As String
    Catatan                             As String
    FooterPrint                         As String
End Type

