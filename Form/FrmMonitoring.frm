VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form FrmMonitoring 
   Caption         =   "Monitoring Machine"
   ClientHeight    =   10200
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   16380
   Icon            =   "FrmMonitoring.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10200
   ScaleWidth      =   16380
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   60000
      Left            =   450
      Top             =   8910
   End
   Begin VB.Frame Frame3 
      Caption         =   "Dashboard"
      Height          =   4200
      Left            =   8145
      TabIndex        =   4
      Top             =   4500
      Width           =   7935
      Begin MSComctlLib.ListView lvList3 
         Height          =   3420
         Left            =   135
         TabIndex        =   7
         Top             =   585
         Width           =   3870
         _ExtentX        =   6826
         _ExtentY        =   6033
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   16777215
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin MSComctlLib.ListView lvList4 
         Height          =   3420
         Left            =   4140
         TabIndex        =   8
         Top             =   585
         Width           =   3600
         _ExtentX        =   6350
         _ExtentY        =   6033
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   16777215
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin VB.Label lblTopNG 
         Alignment       =   2  'Center
         BackColor       =   &H000040C0&
         Caption         =   "TOP 15 NG PRODUCT"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   4140
         TabIndex        =   10
         Top             =   270
         Width           =   3570
      End
      Begin VB.Label lblTopIdle 
         Alignment       =   2  'Center
         BackColor       =   &H000000FF&
         Caption         =   "TOP 15 IDLE MACHINE"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   135
         TabIndex        =   9
         Top             =   270
         Width           =   3750
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Dashboard"
      Height          =   4290
      Left            =   8145
      TabIndex        =   3
      Top             =   90
      Width           =   7935
      Begin VB.ComboBox cboHour 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   720
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   270
         Width           =   915
      End
      Begin MSComctlLib.ListView lvList2 
         Height          =   3555
         Left            =   135
         TabIndex        =   5
         Top             =   630
         Width           =   7605
         _ExtentX        =   13414
         _ExtentY        =   6271
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   16777215
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin lvButton.lvButtons_H cmdRefresh1 
         Height          =   360
         Left            =   1665
         TabIndex        =   14
         Top             =   225
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   635
         CapAlign        =   1
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFHover         =   4210752
         LockHover       =   3
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         Image           =   "FrmMonitoring.frx":617A
         cBack           =   16119285
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Hour"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   180
         TabIndex        =   15
         Top             =   270
         Width           =   555
      End
      Begin VB.Label lblRunning 
         Alignment       =   2  'Center
         BackColor       =   &H00FF0000&
         Caption         =   "DATA RUNNING PRODUCT"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   135
         TabIndex        =   11
         Top             =   270
         Width           =   7575
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Dashboard "
      Height          =   8610
      Left            =   45
      TabIndex        =   0
      Top             =   90
      Width           =   8025
      Begin MSComCtl2.DTPicker DtDate 
         Height          =   375
         Left            =   135
         TabIndex        =   13
         Top             =   225
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   125566979
         CurrentDate     =   43570
      End
      Begin MSComctlLib.ListView lvList 
         Height          =   7740
         Left            =   135
         TabIndex        =   1
         Top             =   630
         Width           =   7695
         _ExtentX        =   13573
         _ExtentY        =   13653
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   16777215
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin lvButton.lvButtons_H cmdRefresh 
         Height          =   360
         Left            =   1800
         TabIndex        =   2
         Top             =   225
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   635
         CapAlign        =   1
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFHover         =   4210752
         LockHover       =   3
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         Image           =   "FrmMonitoring.frx":68F4
         cBack           =   16119285
      End
      Begin VB.Label lblSetRunning 
         Alignment       =   2  'Center
         BackColor       =   &H0000C000&
         Caption         =   "DATA  RUNNING MACHINE"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   135
         TabIndex        =   12
         Top             =   270
         Width           =   7665
      End
   End
End
Attribute VB_Name = "FrmMonitoring"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim srcItem                         As ListItem
Dim srcRecord                       As String
Dim srcMonitor                      As Variant
Dim srcSQL                          As String
Dim RecordPage                      As New clsPaging
Dim SQLParser                       As New clsSQLSelectParser
Dim RS_tYIELD                       As New ADODB.Recordset
Dim RS_dt1YIELD                     As New ADODB.Recordset
Dim RS_dt3YIELD                     As New ADODB.Recordset
Dim RS_dt4YIELD                     As New ADODB.Recordset



Private Sub cmdRefresh_Click()
    CommandPass "Refresh"
End Sub



Private Sub cmdUpdate_Click()
    CommandPass "Update"
End Sub

Private Sub cmdRefresh1_Click()
    CommandPass "Refresh"
End Sub

Private Sub Form_Activate()
On Error Resume Next
With MAIN
    Me.BackColor = .ACPMenu.BackColor
    Frame1.BackColor = .ACPMenu.BackColor
    Frame2.BackColor = .ACPMenu.BackColor
    Frame3.BackColor = .ACPMenu.BackColor
    Me.Picture = .ACPMenu.LoadBackground

End With

Call lvSizeColumns(lvList)
Call lvSizeColumns(lvList2)
Call lvSizeColumns(lvList3)
Call lvSizeColumns(lvList4)
    
End Sub

Private Sub Form_Load()
On Error GoTo ErrHandler

With cboHour
    Dim i As Integer
    i = 0
    For i = 0 To 23
        If i >= 0 And i <= 9 Then
            .AddItem "0" & i
        Else
            .AddItem i
        End If
    Next i
End With

cboHour.Text = Format(Now, "hh")


If Format(Now, "HH") >= 0 And Format(Now, "HH") <= 7 Then
    DtDate.Value = Format(DateAdd("d", -1, Format(Now, "yyyy-mm-dd")), "yyyy-mm-dd")
Else
    DtDate.Value = Format(Now, "yyyy-mm-dd")
End If

With lvList
    .GridLines = True
    .View = lvwReport

    .ColumnHeaders.Clear
    .ColumnHeaders.Add , , "M/C"
    .ColumnHeaders.Add , , "STATUS"
    .ColumnHeaders.Add , , "MC NAME"
    .ColumnHeaders.Add , , "TONNAGE"
    .ColumnHeaders.Add , , "PRD YIELD"
    .ColumnHeaders.Add , , "TTL YIELD"
    .ColumnHeaders.Add , , "% TARGET"

    .ListItems.Clear
End With

With lvList2
    .GridLines = True
    .View = lvwReport

    .ColumnHeaders.Clear
    .ColumnHeaders.Add , , "M/C"
    .ColumnHeaders.Add , , "PROD. NAME"
    .ColumnHeaders.Add , , "HOUR"
    .ColumnHeaders.Add , , "TARGET"
    .ColumnHeaders.Add , , "SHOT"
    .ColumnHeaders.Add , , "CAVITY"
    .ColumnHeaders.Add , , "GROSS "
    .ColumnHeaders.Add , , "NG "
    .ColumnHeaders.Add , , "NET "
    .ColumnHeaders.Add , , "TTL YIELD"
    .ColumnHeaders.Add , , "% TARGET"
    .ColumnHeaders.Add , , "IDLE "
    .ListItems.Clear
End With

With lvList3
    .GridLines = True
    .View = lvwReport
    
    .ColumnHeaders.Clear
    .ColumnHeaders.Add , , "M/C"
    .ColumnHeaders.Add , , "PRODCT NAME 1"
    .ColumnHeaders.Add , , "IDLE TIME"
    .ListItems.Clear
End With

With lvList4
    .GridLines = True
    .View = lvwReport
    
    .ColumnHeaders.Clear
    .ColumnHeaders.Add , , "M/C"
    .ColumnHeaders.Add , , "PRODUCT"
    .ColumnHeaders.Add , , "NG NAME"
    .ColumnHeaders.Add , , "TOTAL"
    .ListItems.Clear
End With

Call FillListview

'AltLVBackground lvList, vbWhite, &HC0FFFF, FrmMonitoring
'AltLVBackground lvList2, vbWhite, &HFFFFC0, FrmMonitoring
'AltLVBackground lvList3, vbWhite, &HC0FFC0, FrmMonitoring
'AltLVBackground lvList4, vbWhite, &HC0E0FF, FrmMonitoring

Exit Sub
ErrHandler:
    MsgBox "Error Number: " & Err.Number & vbCrLf & "Description: " & Err.Description, vbExclamation
End Sub

Private Sub Form_Resize()
On Error Resume Next
    If WindowState <> vbMinimized Then
        If Me.Width < 9195 Then Me.Width = 9195
        If Me.Height < 4500 Then Me.Height = 4500
        
        Frame1.Width = (Me.ScaleWidth / 2) - 2600
        Frame2.Left = Frame1.Width + 100
        Frame3.Left = Frame1.Width + 100
        Frame2.Width = (Me.ScaleWidth / 2) + 2300
        Frame3.Width = (Me.ScaleWidth / 2) + 2300
        Frame1.Height = Me.ScaleHeight - 500
        Frame2.Height = Me.ScaleHeight / 2
        Frame3.Top = Frame2.Height + 100
        Frame3.Height = (Me.ScaleHeight / 2) - 500
        
        lvList.Height = Frame1.Height - 800
        lvList.Width = Frame1.Width - 300
        lblSetRunning.Width = lvList.Width

        lvList2.Height = Frame2.Height - 800
        lvList2.Width = Frame2.Width - 300
        lblRunning.Width = lvList2.Width
'
        lvList3.Height = Frame3.Height - 800
        lvList3.Width = (Frame3.Width / 2) - 1500
        lblTopIdle.Width = lvList3.Width
        
        
        lvList4.Left = lvList3.Width + 300
        lvList4.Height = Frame3.Height - 800
        lvList4.Width = (Frame3.Width / 2) + 1000
        lblTopNG.Left = lvList4.Left
        lblTopNG.Width = lvList4.Width
    End If
    

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set FrmMonitoring = Nothing
    Set RS_tYIELD = Nothing
End Sub

Private Sub lvList_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    If lvList.Sorted And _
        ColumnHeader.Index - 1 = lvList.SortKey Then
        lvList.SortOrder = 1 - lvList.SortOrder
    Else
        lvList.SortOrder = lvwAscending
        lvList.SortKey = ColumnHeader.Index - 1
    End If
    lvList.Sorted = True
End Sub

Private Sub LvList_DblClick()
    RptData DtDate, lvList.SelectedItem.Text
End Sub

Private Sub LvList_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error Resume Next
    srcMonitor = lvList.SelectedItem.Index
    srcRecord = lvList.ListItems.Item(srcMonitor).Text
End Sub


Public Sub CommandPass(ByVal srcPerformWhat As String)
On Error GoTo errPerformWhat
Select Case srcPerformWhat

    Case "Refresh" 'Refresh
           FillListview
            
    Case "Close" 'Close
            Unload Me
End Select
Exit Sub
errPerformWhat:
     MsgBox "Error Number:" & Err.Number & vbNewLine & _
            "Description:" & Err.Description, vbExclamation
End Sub

Private Sub FillListview()
On Error Resume Next
Dim RS_tYIELD As New ADODB.Recordset
Dim sSQL As String
Dim DetSQL As String
Dim Det3SQL As String
Dim Det4SQL As String


sSQL = "SELECT X.plant_mark,X.prod_machine_id,X.machine_status,X.`status`,Y.NUMBER AS machine_no, "
sSQL = sSQL & " Y.NAME AS  machine_name, Y.tonnage, z.prod_yield,"
sSQL = sSQL & " z.total_yield , z.percent_target, z.period_shift"
sSQL = sSQL & " FROM sip_234.prod_running_products X"
sSQL = sSQL & " LEFT JOIN sip_234.prod_machines Y ON X.prod_machine_id = Y.id"
sSQL = sSQL & " LEFT JOIN (SELECT a.prod_machine_id,a.machine_no,a.machine_name,a.tonnage,AVG(a.prod_yield) AS prod_yield,"
            sSQL = sSQL & " AVG(a.total_yield) AS total_yield,AVG(a.percent_target) AS percent_target, a.period_shift"
            sSQL = sSQL & " FROM sip_production.prod_running_results a"
            sSQL = sSQL & " WHERE a.period_shift = '" & Format(DtDate.Value, "yyyy-mm-dd") & "'"
            sSQL = sSQL & " GROUP BY  a.prod_machine_id,a.machine_no,a.machine_name,a.tonnage, a.period_shift) z"
sSQL = sSQL & " ON X.prod_machine_id = z.prod_machine_id"
sSQL = sSQL & " WHERE X.`status` = 'active'"
sSQL = sSQL & " ORDER BY Y.NUMBER ASC"

Set RS_tYIELD = New ADODB.Recordset
If RS_tYIELD.State = adStateOpen Then RS_tYIELD.Close
RS_tYIELD.Open sSQL, CN, adOpenDynamic, adLockPessimistic

With lvList
    .ListItems.Clear
    Do While Not RS_tYIELD.EOF
    Set srcItem = .ListItems.Add(, , RS_tYIELD.Fields("machine_no"))
        srcItem.SubItems(1) = RS_tYIELD.Fields("machine_status")
        srcItem.SubItems(2) = RS_tYIELD.Fields("machine_name")
        srcItem.SubItems(3) = RS_tYIELD.Fields("tonnage")
        srcItem.SubItems(4) = IIf(IsNull(Round(RS_tYIELD.Fields("prod_yield"), 2)), "", Round(RS_tYIELD.Fields("prod_yield"), 2)) & " %"
        srcItem.SubItems(5) = IIf(IsNull(Round(RS_tYIELD.Fields("total_yield"), 2)), "", Round(RS_tYIELD.Fields("total_yield"), 2)) & " %"
        srcItem.SubItems(6) = IIf(IsNull(Round(RS_tYIELD.Fields("percent_target"), 2)), "", Round(RS_tYIELD.Fields("percent_target"), 2)) & " %"

        RS_tYIELD.MoveNext
    Loop
End With

DetSQL = "SELECT a.machine_no,a.product_name,a.period_shift,a.period_hour,a.target_hour,a.shot, a.cavity,a.gross_prod,a.total_ng,a.net_prod,a.total_yield,a.percent_target,a.jumlah_idle"
DetSQL = DetSQL & " FROM sip_production.prod_running_results a"
DetSQL = DetSQL & " WHERE a.period_shift = '" & Format(DtDate.Value, "yyyy-mm-dd") & "' AND a.period_hour = '" & cboHour.Text & "'"
DetSQL = DetSQL & " ORDER BY  a.period_hour DESC , a.machine_no ASC"

Set RS_dt1YIELD = New ADODB.Recordset
If RS_dt1YIELD.State = adStateOpen Then RS_dt1YIELD.Close
RS_dt1YIELD.Open DetSQL, CN, adOpenDynamic, adLockPessimistic

With lvList2
    .ListItems.Clear
    Do While Not RS_dt1YIELD.EOF
    Set srcItem = .ListItems.Add(, , RS_dt1YIELD.Fields("machine_no"))
        srcItem.SubItems(1) = RS_dt1YIELD.Fields("product_name")
        srcItem.SubItems(2) = RS_dt1YIELD.Fields("period_hour")
        srcItem.SubItems(3) = Round(RS_dt1YIELD.Fields("target_hour"), 0)
        srcItem.SubItems(4) = RS_dt1YIELD.Fields("shot")
        srcItem.SubItems(5) = RS_dt1YIELD.Fields("cavity")
        srcItem.SubItems(6) = RS_dt1YIELD.Fields("gross_prod")
        srcItem.SubItems(7) = RS_dt1YIELD.Fields("total_ng")
        srcItem.SubItems(8) = RS_dt1YIELD.Fields("net_prod")
        srcItem.SubItems(9) = IIf(IsNull(Round(RS_dt1YIELD.Fields("total_yield"), 2)), "", Round(RS_dt1YIELD.Fields("total_yield"), 2)) & " %"
        srcItem.SubItems(10) = IIf(IsNull(Round(RS_dt1YIELD.Fields("percent_target"), 2)), "", Round(RS_dt1YIELD.Fields("percent_target"), 2)) & " %"
        srcItem.SubItems(11) = RS_dt1YIELD.Fields("jumlah_idle")
        RS_dt1YIELD.MoveNext
    Loop
End With

Det3SQL = "SELECT a.plant_mark,a.prod_machine_id,a.machine_no,a.product_name_1,a.period_shift,"
Det3SQL = Det3SQL & " SEC_TO_TIME(Sum(TIME_TO_SEC(a.jumlah_idle)))  As jumlah_idle"
Det3SQL = Det3SQL & " FROM sip_production.prod_idle_results a"
Det3SQL = Det3SQL & " WHERE a.period_shift = '" & Format(DtDate.Value, "yyyy-mm-dd") & "'"
Det3SQL = Det3SQL & " GROUP BY a.plant_mark,a.prod_machine_id,a.machine_no,a.product_name_1,a.period_shift"
Det3SQL = Det3SQL & " ORDER BY jumlah_idle DESC, a.machine_no asc  limit 15"

Set RS_dt3YIELD = New ADODB.Recordset
If RS_dt3YIELD.State = adStateOpen Then RS_dt3YIELD.Close
RS_dt3YIELD.Open Det3SQL, CN, adOpenDynamic, adLockPessimistic

With lvList3
    .ListItems.Clear
    Do While Not RS_dt3YIELD.EOF
    Set srcItem = .ListItems.Add(, , RS_dt3YIELD.Fields("machine_no"))
        srcItem.SubItems(1) = RS_dt3YIELD.Fields("product_name_1")
        srcItem.SubItems(2) = Format(RS_dt3YIELD.Fields("jumlah_idle"), "hh:mm:ss")

        RS_dt3YIELD.MoveNext
    Loop
End With

Det4SQL = "SELECT a.plant_mark,a.prod_machine_id,a.machine_no,a.product_name,a.period_shift,a.ng_name,sum(a.jumlah_ng) AS total_ng "
Det4SQL = Det4SQL & " FROM sip_production.prod_ng_results a"
Det4SQL = Det4SQL & " WHERE a.period_shift = '" & Format(DtDate.Value, "yyyy-mm-dd") & "'"
Det4SQL = Det4SQL & " GROUP BY a.plant_mark,a.prod_machine_id,a.machine_no,a.product_name,a.period_shift,a.ng_name"
Det4SQL = Det4SQL & " ORDER BY total_ng DESC, a.machine_no ASC limit 15"

Set RS_dt4YIELD = New ADODB.Recordset
If RS_dt4YIELD.State = adStateOpen Then RS_dt4YIELD.Close
RS_dt4YIELD.Open Det4SQL, CN, adOpenDynamic, adLockPessimistic

With lvList4
    .ListItems.Clear
    Do While Not RS_dt4YIELD.EOF
    Set srcItem = .ListItems.Add(, , RS_dt4YIELD.Fields("machine_no"))
        srcItem.SubItems(1) = RS_dt4YIELD.Fields("product_name")
        srcItem.SubItems(2) = RS_dt4YIELD.Fields("ng_name")
        srcItem.SubItems(3) = RS_dt4YIELD.Fields("total_ng")
        RS_dt4YIELD.MoveNext
    Loop
End With


End Sub


Private Sub lvList_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then PopupMenu MAIN.mnuAction
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
    Select Case KeyCode
    Case vbKeyF5
        CommandPass "Refresh"
    Case vbKeyEscape
        CommandPass "Close"
    End Select
End Sub

Private Sub Timer1_Timer()



If Format(Now, "HH") >= 0 And Format(Now, "HH") <= 7 Then
    DtDate.Value = Format(DateAdd("d", -1, Format(Now, "yyyy-mm-dd")), "yyyy-mm-dd")
Else
    DtDate.Value = Format(Now, "yyyy-mm-dd")
End If

cboHour.Text = Format(Now, "hh")

FillListview

End Sub


Private Sub RptData(sDate As Date, sM_number As String)
On Error GoTo ErrHandler

Dim qSQL As String

qSQL = "select a.plant_mark, a.prod_machine_id, b.number, b.name as machine_name, b.tonnage, a.mkt_customer_id, c.name as customer_name,"
qSQL = qSQL & " a.eng_product_id, d.internal_part_id, d.product_name, d.customer_part_number, d.customer_part_name, d.prod_yield,"
qSQL = qSQL & " d.material_name , d.color_name, d.cavity, d.weight_gr, d.weight_runner_gr, d.cycle_time_ia, a.period_shift,e.nik, e.name as employee,"
qSQL = qSQL & " Round(((3600 / d.cycle_time_ia) * d.cavity),0) as target_shot"
qSQL = qSQL & " from sip_production.prod_runnings a"
qSQL = qSQL & " left join sip_234.prod_machines b on a.prod_machine_id = b.id"
qSQL = qSQL & " left join sip_234.mkt_customers c on a.mkt_customer_id = c.id"
qSQL = qSQL & " left join (select x.id, x.internal_part_id, x.name as product_name, x.customer_part_number, x.customer_part_name,"
qSQL = qSQL & " y.name as material_name, z.name as color_name,x.cavity, x.weight_gr, x.weight_runner_gr, x.cycle_time_ia, x.prod_yield"
qSQL = qSQL & " from sip_234.eng_products x left join sip_234.eng_materials y on x.eng_material_id = y.id"
qSQL = qSQL & " left join sip_234.eng_colors z on x.eng_color_id = z.id) d on a.eng_product_id = d.id"
qSQL = qSQL & " left join sip_234.hrd_employees e on a.created_by = e.id"
qSQL = qSQL & " where a.plant_mark = '" & p_plant_mark & "'"
qSQL = qSQL & " and b.number = '" & sM_number & "'"
qSQL = qSQL & " and a.period_shift = '" & Format(sDate, "yyyy-mm-dd") & "'"
qSQL = qSQL & " group by a.period_shift, a.plant_mark, a.prod_machine_id, a.mkt_customer_id, a.eng_product_id, a.period_shift, a.created_by"

Set RS_PRINT = New ADODB.Recordset
If RS_PRINT.State = adStateOpen Then RS_PRINT.Close
RS_PRINT.Open qSQL, CN, adOpenDynamic, adLockPessimistic

    With RptProduksi
        .DTRpt.Recordset = RS_PRINT
        If sPrint = 0 Then
            .Show 1
        End If
    
    End With
    
Exit Sub
ErrHandler:
    MsgBox "Error Number: " & Err.Number & vbCrLf & "Description: " & Err.Description, vbExclamation, Me.Caption
     
End Sub


