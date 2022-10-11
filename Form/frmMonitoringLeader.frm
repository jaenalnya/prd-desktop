VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "LvButton.ocx"
Begin VB.Form frmMonitoringLeader 
   Caption         =   "Monitoring Leaders"
   ClientHeight    =   9810
   ClientLeft      =   3855
   ClientTop       =   3015
   ClientWidth     =   14385
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9810
   ScaleWidth      =   14385
   Begin VB.OptionButton optData 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Summary Data"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   1
      Left            =   11475
      TabIndex        =   21
      Top             =   90
      Width           =   1815
   End
   Begin VB.OptionButton optData 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Detail Data"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   0
      Left            =   9675
      TabIndex        =   20
      Top             =   90
      Value           =   -1  'True
      Width           =   1815
   End
   Begin VB.ComboBox cboMesin 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      ItemData        =   "frmMonitoringLeader.frx":0000
      Left            =   2745
      List            =   "frmMonitoringLeader.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   14
      Top             =   90
      Width           =   1365
   End
   Begin VB.PictureBox picFooter 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   405
      Left            =   0
      ScaleHeight     =   405
      ScaleWidth      =   14385
      TabIndex        =   0
      Top             =   9405
      Width           =   14385
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   345
         Left            =   5175
         ScaleHeight     =   345
         ScaleWidth      =   4155
         TabIndex        =   1
         Top             =   45
         Width           =   4150
         Begin VB.CommandButton btnNext 
            Height          =   315
            Left            =   3390
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Next 250"
            Top             =   10
            Width           =   315
         End
         Begin VB.CommandButton btnLast 
            Height          =   315
            Left            =   3705
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Last 250"
            Top             =   10
            Width           =   315
         End
         Begin VB.CommandButton btnPrev 
            Height          =   315
            Left            =   3075
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Previous 250"
            Top             =   10
            Width           =   315
         End
         Begin VB.CommandButton btnFirst 
            Height          =   315
            Left            =   2760
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "First 250"
            Top             =   10
            Width           =   315
         End
         Begin VB.Label lblPageInfo 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "0 - 0 of 0"
            Height          =   255
            Left            =   120
            TabIndex        =   6
            Top             =   60
            Width           =   2535
         End
      End
      Begin PRD.Liner Liner1 
         Height          =   30
         Left            =   0
         TabIndex        =   7
         Top             =   0
         Width           =   7035
         _ExtentX        =   12409
         _ExtentY        =   53
      End
      Begin VB.Label lblRecSum 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Records"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   120
         Width           =   690
      End
   End
   Begin lvButton.lvButtons_H cmdSearch 
      Height          =   390
      Left            =   45
      TabIndex        =   9
      Top             =   1845
      Width           =   1290
      _ExtentX        =   2275
      _ExtentY        =   688
      Caption         =   "&Cari [F6]"
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
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      Image           =   "frmMonitoringLeader.frx":0004
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdRefresh 
      Height          =   405
      Left            =   45
      TabIndex        =   10
      Top             =   1350
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   714
      Caption         =   "&Refresh [F5]"
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
      Image           =   "frmMonitoringLeader.frx":077E
      cBack           =   16119285
   End
   Begin lvButton.lvButtons_H cmdClose 
      Height          =   405
      Left            =   45
      TabIndex        =   11
      Top             =   2790
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   714
      Caption         =   "&Close [ESC]"
      CapAlign        =   2
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
      Image           =   "frmMonitoringLeader.frx":0EF8
      cBack           =   16119285
   End
   Begin lvButton.lvButtons_H cmdExport 
      Height          =   405
      Left            =   45
      TabIndex        =   12
      Top             =   2295
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   714
      Caption         =   "&Export [F8]"
      CapAlign        =   2
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
      Image           =   "frmMonitoringLeader.frx":415A
      cBack           =   16119285
   End
   Begin lvButton.lvButtons_H cmdMasuk 
      Height          =   675
      Left            =   45
      TabIndex        =   13
      Top             =   585
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   1191
      Caption         =   "Check Leaders"
      CapAlign        =   2
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
      cBhover         =   16119285
      LockHover       =   3
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      Image           =   "frmMonitoringLeader.frx":44F4
      cBack           =   16777215
   End
   Begin MSComCtl2.DTPicker DTAwal 
      Height          =   330
      Left            =   5355
      TabIndex        =   15
      Top             =   90
      Width           =   1590
      _ExtentX        =   2805
      _ExtentY        =   582
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
      CustomFormat    =   "dd/MMM/yyyy"
      Format          =   87949315
      CurrentDate     =   43642
   End
   Begin MSComCtl2.DTPicker DTAkhir 
      Height          =   330
      Left            =   7515
      TabIndex        =   16
      Top             =   90
      Width           =   1590
      _ExtentX        =   2805
      _ExtentY        =   582
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
      CustomFormat    =   "dd/MMM/yyyy"
      Format          =   87949315
      CurrentDate     =   43642
   End
   Begin MSComctlLib.ListView lvList 
      Height          =   7335
      Left            =   1530
      TabIndex        =   22
      Top             =   585
      Width           =   11790
      _ExtentX        =   20796
      _ExtentY        =   12938
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "To"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   7110
      TabIndex        =   19
      Top             =   135
      Width           =   510
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Date :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4725
      TabIndex        =   18
      Top             =   135
      Width           =   600
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Machine No :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1440
      TabIndex        =   17
      Top             =   135
      Width           =   1185
   End
End
Attribute VB_Name = "frmMonitoringLeader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim srcItem                         As ListItem
Dim srcRecord                       As String
Dim srcMonitoring                           As Variant
Dim srcSQL                          As String
Dim RecordPage                      As New clsPaging
Dim SQLParser                       As New clsSQLSelectParser




Private Sub DTAwal_Click()
    Call LoadData

End Sub

Private Sub DTAwal_CloseUp()
    Call LoadData

End Sub

Private Sub LoadData()

If optData(0).Value = True Then
    With SQLParser
        .Fields = "*"
        .Tables = "(select A.plant_mark,A.prod_machine_id, B.number as machine_no, B.name as machine_name, " & _
                        "A.mkt_customer_id,C.name as customer_name, A.eng_product_id, D.name as product_name, " & _
                        "A.period_shift,A.date,A.check_kesesuaian, " & _
                        "A.check_material , A.check_abnormality, A.target_yield, A.cycle_time, A.hrd_employee_id,E.name as leader_name ,E.nik, " & _
                        "A.start_check,A.stop_check,A.time_check " & _
                        "from prod_monitoring_leaders A " & _
                        "INNER JOIN prod_machines B on A.prod_machine_id = B.id " & _
                        "INNER JOIN mkt_customers C on A.mkt_customer_id = C.id " & _
                        "INNER JOIN eng_products D on A.eng_product_id = D.id " & _
                        "INNER JOIN hrd_employees E on A.hrd_employee_id = E.id) as XX "
        
            If cboMesin.text = "All" Then
                .wCondition = "period_shift between '" & Format(DTAwal.Value, "yyyy-mm-dd") & "' and '" & Format(DTAkhir.Value, "yyyy-mm-dd") & "' and plant_mark = '" & p_plant_mark & "'"
            Else
                .wCondition = "period_shift between '" & Format(DTAwal.Value, "yyyy-mm-dd") & "' and '" & Format(DTAkhir.Value, "yyyy-mm-dd") & "' and machine_no = '" & cboMesin.text & "' and plant_mark = '" & p_plant_mark & "'"
            End If
            
        .SortOrder = "period_shift,leader_name,machine_no"
        .SaveStatement
    End With

ElseIf optData(1).Value = True Then
    With SQLParser
        .Fields = "*"
        .Tables = "(SELECT a.plant_mark, a.period_shift,  b.number AS machine_no, b.name AS machine_name, " & _
                    "c.name AS product_name,  d.nik, d.name AS leader_name, " & _
                    "COUNT(prod_machine_id) As jumlah_check " & _
                    "FROM prod_monitoring_leaders a " & _
                    "LEFT JOIN prod_machines b ON a.prod_machine_id = b.id " & _
                    "LEFT JOIN eng_products c ON a.eng_product_id = c.id " & _
                    "LEFT JOIN hrd_employees d ON a.hrd_employee_id = d.id " & _
                    "GROUP BY a.plant_mark,a.prod_machine_id,a.eng_product_id,a.period_shift,a.hrd_employee_id) as XX "
        
            If cboMesin.text = "All" Then
                .wCondition = "period_shift between '" & Format(DTAwal.Value, "yyyy-mm-dd") & "' and '" & Format(DTAkhir.Value, "yyyy-mm-dd") & "' and plant_mark = '" & p_plant_mark & "'"
            Else
                .wCondition = "period_shift between '" & Format(DTAwal.Value, "yyyy-mm-dd") & "' and '" & Format(DTAkhir.Value, "yyyy-mm-dd") & "' and machine_no = '" & cboMesin.text & "' and plant_mark = '" & p_plant_mark & "'"
            End If

        .SortOrder = "period_shift,leader_name,machine_no"
        .SaveStatement
    End With
End If



Set RS_MONITORING = New ADODB.Recordset
If RS_MONITORING.State = adStateOpen Then RS_MONITORING.Close
RS_MONITORING.Open SQLParser.SQLStatement, CN, adOpenDynamic, adLockPessimistic


With RecordPage
    .Start RS_MONITORING, 100
End With

FillListview 1

srcMonitoring = "NONE"
srcRecord = vbNullString

End Sub

Private Sub cboMesin_Click()
    LoadData
End Sub

Private Sub cmdClose_Click()
    CommandPass "Close"
End Sub

Private Sub cmdExport_Click()
    CommandPass "Export"
End Sub

Private Sub cmdKeluar_Click()
    CommandPass "Keluar"
End Sub

Private Sub cmdMasuk_Click()
    CommandPass "Masuk"
End Sub

Private Sub cmdRefresh_Click()
    CommandPass "Refresh"
End Sub

Private Sub CmdSearch_Click()
    CommandPass "Search"
End Sub

Private Sub Form_Activate()
On Error Resume Next
With MAIN
    Me.BackColor = .ACPMenu.BackColor
    Me.Picture = .ACPMenu.LoadBackground
    picFooter.BackColor = .ACPMenu.BackColor
    Picture2.BackColor = .ACPMenu.BackColor
End With
lvList.FlatScrollBar = False
'AltLVBackground lvList, vbWhite, &H8000000F, frmMonitoringLeader
MAIN.ActivateChild Me


End Sub

Private Sub Form_Load()
On Error GoTo ErrHandler

ButtonList lvList, btnFirst, btnPrev, btnNext, btnLast

Dim i As Integer
For i = 0 To 45
    If i = 0 Then
        cboMesin.AddItem "All"
    Else
        cboMesin.AddItem i
    End If
    
Next i
cboMesin.ListIndex = 0

DTAwal.Value = Format(Now, "dd/MMM/yyyy")
DTAkhir.Value = Format(Now, "dd/MMM/yyyy")

Call LoadData

Exit Sub
ErrHandler:
    MsgBox "Error Number: " & Err.Number & vbCrLf & "Description: " & Err.Description, vbExclamation
End Sub

Private Sub Form_Resize()
On Error Resume Next
    If WindowState <> vbMinimized Then
        If Me.Width < 9195 Then Me.Width = 9195
        If Me.Height < 4500 Then Me.Height = 4500

        Liner1.Width = ScaleWidth
        lvList.Width = Me.ScaleWidth - 1550
        lvList.Height = Me.ScaleHeight - 1000

    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
    MAIN.RemoveChild Me.Name
    Set frmMonitoringLeader = Nothing
    Set RS_MONITORING = Nothing
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


Private Sub LvList_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error Resume Next
    srcMonitoring = lvList.SelectedItem.Index
    srcRecord = lvList.ListItems.Item(srcMonitoring).text
    Call RefreshRecSum
End Sub


Public Sub CommandPass(ByVal srcPerformWhat As String)
On Error GoTo errPerformWhat
Select Case srcPerformWhat
 
    Case "Masuk" 'Refresh
           With frmMonitoringLeaderAE
                .Show 1
           End With

    Case "Refresh" 'Refresh
           RefreshRecords
    
    Case "Search" 'Search
            With frmSearch
                Set .srcForm = Me
                Set .srcColumnHeaders = lvList.ColumnHeaders
                .Show vbModal
            End With
           
    Case "Export" 'Preview
            With lvList
                If .ListItems.Count = 0 Then
                    MsgBox "There's no records to export!Please check it.", vbExclamation
                    Exit Sub
                End If
            End With
                         
            XLSFILENAME = ""
            
            With MAIN.CDExporter
                .Filter = "Excel Files(*.xls)|*.xls"
                .ShowSave
            XLSFILENAME = .FileName
            End With
            
            If XLSFILENAME = "" Then
            Exit Sub
            End If
            
            
            Call ExportListview(lvList, XLSFILENAME)
            MsgBox "Records successfully exported!", vbInformation
            XLSFILENAME = ""
            RefreshRecords
            
    Case "Close" 'Close
            Unload Me
End Select
Exit Sub
errPerformWhat:
     MsgBox "Error Number:" & Err.Number & vbNewLine & _
            "Description:" & Err.Description, vbExclamation
End Sub

Private Sub FillListview(ByVal whichPage As Long)
On Error Resume Next
Dim i As Integer
i = 1
RecordPage.CurrentPosition = whichPage
'RS_MONITORING.AbsolutePosition = RecordPage.PageStart
RecordPage.PageEnd

If optData(0).Value = True Then
    With lvList
        .GridLines = True
        .View = lvwReport
    
        .ColumnHeaders.Clear
        .ColumnHeaders.Add , , "No."
        .ColumnHeaders.Add , , "machine_no"
        .ColumnHeaders.Add , , "machine_name"
        .ColumnHeaders.Add , , "customer_name"
        .ColumnHeaders.Add , , "product_name"
        .ColumnHeaders.Add , , "period_shift"
        .ColumnHeaders.Add , , "date"
        .ColumnHeaders.Add , , "check_kesesuaian"
        .ColumnHeaders.Add , , "check_material"
        .ColumnHeaders.Add , , "check_abnormality"
        .ColumnHeaders.Add , , "target_yield"
        .ColumnHeaders.Add , , "cycle_time"
        .ColumnHeaders.Add , , "nik"
        .ColumnHeaders.Add , , "leader_name"
        .ColumnHeaders.Add , , "start_check"
        .ColumnHeaders.Add , , "stop_check"
        .ColumnHeaders.Add , , "time_check"
        
    
        .ListItems.Clear
        Do While Not RS_MONITORING.EOF
        Set srcItem = .ListItems.Add(, , i, 1, 1)
            
            srcItem.SubItems(1) = RS_MONITORING.Fields("machine_no")
            srcItem.SubItems(2) = RS_MONITORING.Fields("machine_name")
            srcItem.SubItems(3) = RS_MONITORING.Fields("customer_name")
            srcItem.SubItems(4) = RS_MONITORING.Fields("product_name")
            srcItem.SubItems(5) = Format(RS_MONITORING.Fields("period_shift"), "yyyy-mm-dd")
            srcItem.SubItems(6) = Format(RS_MONITORING.Fields("date"), "yyyy-mm-dd hh:mm:ss")
            srcItem.SubItems(7) = RS_MONITORING.Fields("check_kesesuaian")
            srcItem.SubItems(8) = RS_MONITORING.Fields("check_material")
            srcItem.SubItems(9) = RS_MONITORING.Fields("check_abnormality")
            srcItem.SubItems(10) = RS_MONITORING.Fields("target_yield")
            srcItem.SubItems(11) = RS_MONITORING.Fields("cycle_time")
            srcItem.SubItems(12) = RS_MONITORING.Fields("nik")
            srcItem.SubItems(13) = RS_MONITORING.Fields("leader_name")
            srcItem.SubItems(14) = Format(RS_MONITORING.Fields("start_check"), "dd-MMM-yyyy hh:mm:ss")
            srcItem.SubItems(15) = Format(RS_MONITORING.Fields("stop_check"), "dd-MMM-yyyy hh:mm:ss")
            srcItem.SubItems(16) = Format(RS_MONITORING.Fields("time_check"), "hh:mm:ss")
    
            If RS_MONITORING.AbsolutePosition >= RecordPage.PageEnd Then
                Exit Do
            Else
                RS_MONITORING.MoveNext
            End If
            i = i + 1
        Loop
    End With

ElseIf optData(1).Value = True Then
    With lvList
        .GridLines = True
        .View = lvwReport
    
        .ColumnHeaders.Clear
        .ColumnHeaders.Add , , "No."
        .ColumnHeaders.Add , , "machine_no"
        .ColumnHeaders.Add , , "machine_name"
        .ColumnHeaders.Add , , "product_name"
        .ColumnHeaders.Add , , "period_shift"
        .ColumnHeaders.Add , , "nik"
        .ColumnHeaders.Add , , "leader_name"
        .ColumnHeaders.Add , , "jumlah_check"
    
        .ListItems.Clear
        Do While Not RS_MONITORING.EOF
        Set srcItem = .ListItems.Add(, , i, 1, 1)
            
            srcItem.SubItems(1) = RS_MONITORING.Fields("machine_no")
            srcItem.SubItems(2) = RS_MONITORING.Fields("machine_name")
            srcItem.SubItems(3) = RS_MONITORING.Fields("product_name")
            srcItem.SubItems(4) = Format(RS_MONITORING.Fields("period_shift"), "yyyy-mm-dd")
            srcItem.SubItems(5) = RS_MONITORING.Fields("nik")
            srcItem.SubItems(6) = RS_MONITORING.Fields("leader_name")
            srcItem.SubItems(7) = RS_MONITORING.Fields("jumlah_check")
    
            If RS_MONITORING.AbsolutePosition >= RecordPage.PageEnd Then
                Exit Do
            Else
                RS_MONITORING.MoveNext
            End If
            i = i + 1
        Loop
    End With
End If


SetNavigation btnFirst, btnPrev, btnNext, btnLast
'Display the page information
lblPageInfo.Caption = "Record " & RecordPage.PageInfo
'Display the selected record
Call RefreshRecSum
Call lvSizeColumns(lvList)

End Sub


Private Sub RefreshRecSum()
    lblRecSum.Caption = "Record: " & srcMonitoring & " of " & lvList.ListItems.Count
End Sub


Private Sub btnFirst_Click()
    If RecordPage.PAGE_CURRENT <> 1 Then FillListview 1
End Sub

Private Sub btnLast_Click()
    If RecordPage.PAGE_CURRENT <> RecordPage.PAGE_TOTAL Then FillListview RecordPage.PAGE_TOTAL
End Sub

Private Sub btnNext_Click()
    If RecordPage.PAGE_CURRENT <> RecordPage.PAGE_TOTAL Then FillListview RecordPage.PAGE_NEXT
End Sub

Private Sub btnPrev_Click()
    If RecordPage.PAGE_CURRENT <> 1 Then FillListview RecordPage.PAGE_PREVIOUS
End Sub

Public Sub FilterRecord(ByVal srcCondition As String)
    SQLParser.RestoreStatement
    SQLParser.wCondition = SQLParser.wCondition & " and " & srcCondition
    ReloadRecords SQLParser.SQLStatement
End Sub
'Public Sub FilterRecord(ByVal srcCondition As String)
'    SQLParser.RestoreStatement
'    SQLParser.wCondition = srcCondition
'    ReloadRecords SQLParser.SQLStatement
'End Sub

'Procedure for reloadingrecords
Public Sub ReloadRecords(ByVal srcSQL As String)
On Error Resume Next
    
    Set RS_MONITORING = New ADODB.Recordset
    If RS_MONITORING.State = adStateOpen Then RS_MONITORING.Close
    RS_MONITORING.Open srcSQL, CN, adOpenDynamic, adLockPessimistic

With RecordPage
    .Start RS_MONITORING, 100
End With

FillListview 1

End Sub
Public Sub RefreshRecords()
    SQLParser.RestoreStatement
    ReloadRecords SQLParser.SQLStatement
End Sub


Private Sub optData_Click(Index As Integer)
    Call LoadData
End Sub

Private Sub picFooter_Resize()
    Picture2.Left = picFooter.ScaleWidth - Picture2.ScaleWidth
End Sub

Private Sub lvList_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then PopupMenu MAIN.mnuAction
End Sub





