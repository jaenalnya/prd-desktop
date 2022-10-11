Attribute VB_Name = "ModVariables"
Option Explicit


Public PortAddress      As String
Public NomorPort           As Byte

'-----------------------------------------------------------------
Public Declare Function Inp Lib "inpout32.dll" Alias "Inp32" _
    (ByVal PortAddress As Integer) _
    As Integer
    
Public Declare Sub Out Lib "inpout32.dll" Alias "Out32" _
    (ByVal PortAddress As Integer, _
    ByVal Value As Integer)
'-----------------------------------------------------------------

Public Declare Function GetTickCount Lib "kernel32" () As Long
'-----------------------------------------------------------------
Public Declare Sub PortOut Lib "io.dll" (ByVal Port As Integer, ByVal Value As Byte)
Public Declare Function PortIn Lib "io.dll" (ByVal Port As Integer) As Byte
'-----------------------------------------------------------------


Global END_APP                              As Boolean

Public LOG_APP                              As Boolean

Public Rs_search                            As New ADODB.Recordset
'RS_PRINT
Public RS_PRINT                             As New ADODB.Recordset
Public RS_USER                              As New ADODB.Recordset
Public RS_USER2                             As New ADODB.Recordset
Public RS_ADJUST                            As New ADODB.Recordset
Public RS_ADJUSTNG                          As New ADODB.Recordset
Public RS_LABELOK                           As New ADODB.Recordset
Public RS_PRODBYCUST                        As New ADODB.Recordset
Public RS_PRODBYDATE                        As New ADODB.Recordset
Public RS_UNLOCK                            As New ADODB.Recordset
Public RS_USERTYPE                          As New ADODB.Recordset
Public RS_COMPANY                           As New ADODB.Recordset
Public RS_PARAM                             As New ADODB.Recordset
Public RS_LOG                               As New ADODB.Recordset
Public RS_DATA                              As New ADODB.Recordset
Public RS_IDLE                              As New ADODB.Recordset
Public RS_MONITOR                           As New ADODB.Recordset
Public RS_PARAMETER                         As New ADODB.Recordset
Public RS_RPTLOGIN                          As New ADODB.Recordset
Public RS_RPTCALL                           As New ADODB.Recordset
Public RS_ABSENSI                           As New ADODB.Recordset

Public RS_MONITORING                        As New ADODB.Recordset
Public RS_TROUBLE                           As New ADODB.Recordset


Public ACTIVE_USER                          As USER_INFO
Public ACTIVE_ADMIN                         As USER_INFO

Public ACTIVE_USER_2                        As USER_INFO

Public ACTIVE_COMPANY                       As COMPANY_INFO
Public RS_PRODUCT                           As ADODB.Recordset
Public RS_NG                                As ADODB.Recordset

Public XLSFILENAME                          As String

Public COMMAND_INSERT                       As New ADODB.Command
Public COMMAND_UPDATE                       As New ADODB.Command
Public COMMAND_DELETE                       As New ADODB.Command

Public sPerusahaan                          As String
Public sPilihPer                            As String
Public sPrint                               As Integer
Public xls                                  As ActiveReportsExcelExport.ARExportExcel
Public pdf                                  As ActiveReportsPDFExport.ARExportPDF
Public sFile                                As String

Public bIdle                                As Boolean
Public bformIdle                            As Boolean

Public formNG                               As Boolean
Public formIdle                             As Boolean


Public PortOn                               As String
Public IdleOn                               As Boolean
Public LinkUpdate                           As String
Public MaxShot                              As String
Public ShowSensor                           As Boolean, ShowWI  As Boolean, ShowToolbar As Boolean, eMachine As Boolean, ShowInfo As Boolean

Public Display_Mesin                        As Boolean
Public Display_Setting                      As Boolean
