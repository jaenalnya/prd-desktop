VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "LvButton.ocx"
Begin VB.Form frmRptHarian 
   Caption         =   "Data Harian Produksi"
   ClientHeight    =   5340
   ClientLeft      =   8130
   ClientTop       =   6195
   ClientWidth     =   9510
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5340
   ScaleWidth      =   9510
   Begin MSComCtl2.DTPicker DTAwal 
      Height          =   330
      Left            =   5535
      TabIndex        =   18
      Top             =   45
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
      Format          =   89325571
      CurrentDate     =   43642
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
      ItemData        =   "frmRptHarian.frx":0000
      Left            =   2925
      List            =   "frmRptHarian.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   16
      Top             =   45
      Width           =   1365
   End
   Begin lvButton.lvButtons_H cmdPrint 
      Height          =   465
      Left            =   45
      TabIndex        =   14
      Top             =   1800
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   820
      Caption         =   "Cetak Lap."
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
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
      Image           =   "frmRptHarian.frx":0004
      cBack           =   -2147483633
   End
   Begin VB.PictureBox picFooter 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   405
      Left            =   0
      ScaleHeight     =   405
      ScaleWidth      =   9510
      TabIndex        =   0
      Top             =   4935
      Width           =   9510
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
   Begin MSComctlLib.ListView lvList 
      Height          =   4095
      Left            =   1440
      TabIndex        =   9
      Top             =   495
      Width           =   7830
      _ExtentX        =   13811
      _ExtentY        =   7223
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
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
   Begin lvButton.lvButtons_H cmdSearch 
      Height          =   480
      Left            =   45
      TabIndex        =   10
      Top             =   630
      Width           =   1290
      _ExtentX        =   2275
      _ExtentY        =   847
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
      Image           =   "frmRptHarian.frx":055E
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdRefresh 
      Height          =   495
      Left            =   45
      TabIndex        =   11
      Top             =   45
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   873
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
      Image           =   "frmRptHarian.frx":0CD8
      cBack           =   16119285
   End
   Begin lvButton.lvButtons_H cmdClose 
      Height          =   495
      Left            =   45
      TabIndex        =   12
      Top             =   2385
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   873
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
      Image           =   "frmRptHarian.frx":1452
      cBack           =   16119285
   End
   Begin lvButton.lvButtons_H cmdExport 
      Height          =   495
      Left            =   45
      TabIndex        =   13
      Top             =   1215
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   873
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
      Image           =   "frmRptHarian.frx":46B4
      cBack           =   16119285
   End
   Begin MSComCtl2.DTPicker DTAkhir 
      Height          =   330
      Left            =   7695
      TabIndex        =   19
      Top             =   45
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
      Format          =   89325571
      CurrentDate     =   43642
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
      Left            =   7290
      TabIndex        =   20
      Top             =   90
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
      Left            =   4905
      TabIndex        =   17
      Top             =   90
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
      Left            =   1620
      TabIndex        =   15
      Top             =   90
      Width           =   1185
   End
End
Attribute VB_Name = "frmRptHarian"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim srcItem                         As ListItem
Dim srcRecord                       As String
Dim srcProduct                      As Variant
Dim srcSQL                          As String
Dim RecordPage                      As New clsPaging
Dim SQLParser                       As New clsSQLSelectParser



Private Sub cboMesin_Click()
Call LoadData
End Sub

Private Sub cmdClose_Click()
    CommandPass "Close"
End Sub

Private Sub cmdExport_Click()
    CommandPass "Export"
End Sub

Private Sub cmdPrint_Click()
    CommandPass "Print"
    
End Sub

Private Sub cmdRefresh_Click()
    CommandPass "Refresh"
End Sub

Private Sub CmdSearch_Click()
    CommandPass "Search"
End Sub

Private Sub cmdUpdate_Click()
    CommandPass "Update"
End Sub



Private Sub DTAwal_Click()
    Call LoadData
End Sub

Private Sub DTAwal_CloseUp()
    Call LoadData
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
cboMesin.Text = ReadINI("SETTING", "MACHINE", App.Path & "\Settings.ini")

DTAwal.Value = Format(Now, "dd/MMM/yyyy")
DTAkhir.Value = Format(Now, "dd/MMM/yyyy")

Call LoadData

srcProduct = "NONE"
srcRecord = vbNullString

Exit Sub
ErrHandler:
    MsgBox "Error Number: " & Err.Number & vbCrLf & "Description: " & Err.Description, vbExclamation
End Sub

Private Sub LoadData()


With SQLParser
    .Fields = "a.plant_mark,a.prod_machine_id,c.number AS machine_no,c.name AS machine_name, " & _
            "e.name as customer_name,a.eng_product_id,b.internal_part_id,b.name AS product_name, " & _
            "b.cavity,b.prod_yield AS target_yield,b.weight_gr,b.cycle_time_ia as cycle_time_ia,a.period_shift, " & _
            "sum(a.counter_ok) jumlah_shot, " & _
            "sum(a.counter_ok) * b.cavity AS gross_produksi, ifnull(data_ngs.jumlah_ng,0) AS jumlah_ng, " & _
            "((sum(a.counter_ok) * b.cavity) - ifnull(data_ngs.jumlah_ng,0)) AS net_produksi, " & _
            "ifnull(ROUND((((sum(a.counter_ok) * b.cavity) - ifnull(data_ngs.jumlah_ng,0)) / (sum(a.counter_ok) * b.cavity)) * 100,2),0) AS prod_yield, " & _
            "ifnull(f.t_hour_idle,0) as t_hour_idle, " & _
            "ifnull(f.t_sec_idle,0) as t_sec_idle, " & _
            "count(a.period_hour) AS jumlah_hour, " & _
            "count(a.period_hour) * 3600 AS jumlah_second, " & _
            "(count(a.period_hour) * 3600) - ifnull(f.t_sec_idle,0) as running_hour, " & _
            "ROUND(((count(a.period_hour) * 3600) - ifnull(f.t_sec_idle,0)) / b.cycle_time_ia) as running_target, " & _
            "Round(sum(a.counter_ok) / ROUND(((count(a.period_hour) * 3600) - ifnull(f.t_sec_idle,0)) / b.cycle_time_ia) * 100,2) as persen_target"

    .Tables = "prod_runnings a " & _
            "LEFT JOIN eng_products b ON a.eng_product_id = b.id " & _
            "LEFT JOIN prod_machines c ON a.prod_machine_id = c.id " & _
            "LEFT JOIN (SELECT d.plant_mark,d.prod_machine_id,d.eng_product_id,d.period_shift, sum(d.counter_ng) jumlah_ng " & _
            "FROM prod_data_ngs d " & _
            "GROUP BY d.plant_mark,d.prod_machine_id,d.eng_product_id,d.period_shift) AS data_ngs ON a.plant_mark = data_ngs.plant_mark " & _
            "AND a.prod_machine_id = data_ngs.prod_machine_id AND a.eng_product_id = data_ngs.eng_product_id " & _
            "AND a.period_shift = data_ngs.period_shift " & _
            "LEFT JOIN mkt_customers e ON a.mkt_customer_id = e.id " & _
            "left join (select prod_machine_idles.plant_mark, " & _
                "prod_machine_idles.prod_machine_id ,prod_machine_idles.eng_product_1 , " & _
                "prod_machine_idles.period_shift , " & _
                "sum(time_to_sec(prod_machine_idles.idle_time)) AS t_sec_idle, " & _
                "round(sum(time_to_sec(prod_machine_idles.idle_time)) / 3600, 1) As t_hour_idle " & _
                "from prod_machine_idles group by prod_machine_idles.plant_mark, " & _
                "prod_machine_idles.prod_machine_id,prod_machine_idles.mkt_customer_id, " & _
                "prod_machine_idles.eng_product_1,prod_machine_idles.period_shift) f " & _
                "on a.plant_mark = f.plant_mark and a.prod_machine_id = f.prod_machine_id " & _
                "and a.eng_product_id = f.eng_product_1 and a.period_shift = f.period_shift"

        If cboMesin.Text = "All" Then
            .wCondition = "a.period_shift between '" & Format(DTAwal.Value, "yyyy-mm-dd") & "' and '" & Format(DTAkhir.Value, "yyyy-mm-dd") & "' and a.plant_mark = '" & p_plant_mark & "'"
        Else
            .wCondition = "a.period_shift between '" & Format(DTAwal.Value, "yyyy-mm-dd") & "' and '" & Format(DTAkhir.Value, "yyyy-mm-dd") & "' and c.number = '" & cboMesin.Text & "' and a.plant_mark = '" & p_plant_mark & "'"
        End If

    .GroupOrder = "a.plant_mark,a.prod_machine_id,a.mkt_customer_id,a.eng_product_id,a.period_shift"
    .SortOrder = "a.period_shift, c.number ASC"
    .SaveStatement
End With


Set RS_PRODUCT = New ADODB.Recordset
If RS_PRODUCT.State = adStateOpen Then RS_PRODUCT.Close
RS_PRODUCT.Open SQLParser.SQLStatement, CN, adOpenDynamic, adLockPessimistic
 
 

With RecordPage
    .Start RS_PRODUCT, 500
End With

FillListview 1

End Sub
Private Sub Form_Resize()
On Error Resume Next
    If WindowState <> vbMinimized Then
        If Me.Width < 9195 Then Me.Width = 9195
        If Me.Height < 4500 Then Me.Height = 4500

        Liner1.Width = ScaleWidth
        lvList.Width = Me.ScaleWidth - 1550
        lvList.Height = Me.ScaleHeight - 1300
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
    MAIN.RemoveChild Me.Name
    Set frmRptHarian = Nothing
    Set RS_PRODUCT = Nothing
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
    CommandPass "Update"
End Sub

Private Sub LvList_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error Resume Next
    srcProduct = lvList.SelectedItem.Index
    srcRecord = lvList.ListItems.Item(srcProduct).Text
    Call RefreshRecSum
End Sub


Public Sub CommandPass(ByVal srcPerformWhat As String)
On Error GoTo errPerformWhat
Select Case srcPerformWhat

    Case "Update" 'Refresh
           frmAdjustshotAE.Show 1

    Case "Refresh" 'Refresh
           RefreshRecords
    
    Case "Search" 'Search
            With frmSearch
                Set .srcForm = Me
                Set .srcColumnHeaders = lvList.ColumnHeaders
                .Show vbModal
            End With

     Case "Print" 'Preview
            With lvList
                If .ListItems.count = 0 Then
                    MsgBox "There's no records to export!Please check it.", vbExclamation
                    Exit Sub
                End If
            End With
            
            RptData lvList.SelectedItem.SubItems(5), lvList.SelectedItem.SubItems(1), lvList.SelectedItem.SubItems(2)
            
    Case "Export" 'Preview
            With lvList
                If .ListItems.count = 0 Then
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
'RS_PRODUCT.AbsolutePosition = RecordPage.PageStart
RecordPage.PageEnd
With lvList
    .GridLines = True
    .View = lvwReport

    .ColumnHeaders.Clear
    .ColumnHeaders.Add , , "No."
    .ColumnHeaders.Add , , "period_shift"
    .ColumnHeaders.Add , , "machine_no"
    .ColumnHeaders.Add , , "machine_name"
    .ColumnHeaders.Add , , "customer_name"
    .ColumnHeaders.Add , , "internal_part_id"
    .ColumnHeaders.Add , , "product_name"
    .ColumnHeaders.Add , , "cycle_time_ia"
    .ColumnHeaders.Add , , "cavity"
    .ColumnHeaders.Add , , "running_target"
    .ColumnHeaders.Add , , "jumlah_shot"
    .ColumnHeaders.Add , , "gross_produksi"
    .ColumnHeaders.Add , , "total_ng"
    .ColumnHeaders.Add , , "net_produksi"
    .ColumnHeaders.Add , , "prod_yield"
    .ColumnHeaders.Add , , "persen_target"


    .ListItems.Clear
    Do While Not RS_PRODUCT.EOF
    Set srcItem = .ListItems.Add(, , i, 1, 1)
        srcItem.SubItems(1) = Format(RS_PRODUCT.Fields("period_shift"), "yyyy-mm-dd")
        srcItem.SubItems(2) = RS_PRODUCT.Fields("machine_no")
        srcItem.SubItems(3) = RS_PRODUCT.Fields("machine_name")
        srcItem.SubItems(4) = RS_PRODUCT.Fields("customer_name")
        srcItem.SubItems(5) = RS_PRODUCT.Fields("internal_part_id")
        srcItem.SubItems(6) = RS_PRODUCT.Fields("product_name")
        srcItem.SubItems(7) = RS_PRODUCT.Fields("cycle_time_ia")
        srcItem.SubItems(8) = RS_PRODUCT.Fields("cavity")
        srcItem.SubItems(9) = RS_PRODUCT.Fields("running_target")
        srcItem.SubItems(10) = RS_PRODUCT.Fields("jumlah_shot")
        srcItem.SubItems(11) = RS_PRODUCT.Fields("gross_produksi")
        srcItem.SubItems(12) = RS_PRODUCT.Fields("jumlah_ng")
        srcItem.SubItems(13) = RS_PRODUCT.Fields("net_produksi")
        srcItem.SubItems(14) = RS_PRODUCT.Fields("prod_yield")
        srcItem.SubItems(15) = RS_PRODUCT.Fields("persen_target")

        If RS_PRODUCT.AbsolutePosition >= RecordPage.PageEnd Then
            Exit Do
        Else
            RS_PRODUCT.MoveNext
        End If
        i = i + 1
    Loop
End With


SetNavigation btnFirst, btnPrev, btnNext, btnLast
'Display the page information
lblPageInfo.Caption = "Record " & RecordPage.PageInfo
'Display the selected record
Call RefreshRecSum
Call lvSizeColumns(lvList)

'AltLVBackground lvList, vbWhite, &H8000000F, frmRptHarian

End Sub


Private Sub RefreshRecSum()
    lblRecSum.Caption = "Record: " & srcProduct & " of " & lvList.ListItems.count
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

'Procedure for reloadingrecords
Public Sub ReloadRecords(ByVal srcSQL As String)
On Error Resume Next
    
    Set RS_PRODUCT = New ADODB.Recordset
    If RS_PRODUCT.State = adStateOpen Then RS_PRODUCT.Close
    RS_PRODUCT.Open srcSQL, CN, adOpenDynamic, adLockPessimistic

With RecordPage
    .Start RS_PRODUCT, 500
End With

FillListview 1

End Sub
Public Sub RefreshRecords()
    SQLParser.RestoreStatement
    ReloadRecords SQLParser.SQLStatement
End Sub

Private Sub picFooter_Resize()
    Picture2.Left = picFooter.ScaleWidth - Picture2.ScaleWidth
End Sub

Private Sub lvList_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then PopupMenu MAIN.mnuAction
End Sub




Private Sub RptData(sProd As String, sDate As Date, sM_number As String)
On Error GoTo ErrHandler

Dim qSQL As String

qSQL = "select a.plant_mark, a.prod_machine_id, b.number, b.name as machine_name, b.tonnage, a.mkt_customer_id, c.name as customer_name,"
qSQL = qSQL & " a.eng_product_id, d.internal_part_id, d.product_name, d.customer_part_number, d.customer_part_name, d.prod_yield,"
qSQL = qSQL & " d.material_name , d.color_name, d.cavity, d.weight_gr, d.weight_runner_gr, d.cycle_time_ia, a.period_shift,e.nik, e.name as employee,f.nik as nik_2, f.name as employee_2,"
qSQL = qSQL & " Round(((3600 / d.cycle_time_ia) * d.cavity),0) as target_shot"
qSQL = qSQL & " from sip_production.prod_runnings a"
qSQL = qSQL & " left join sip_production.prod_machines b on a.prod_machine_id = b.id"
qSQL = qSQL & " left join sip_production.mkt_customers c on a.mkt_customer_id = c.id"
qSQL = qSQL & " left join (select x.id, x.internal_part_id, x.name as product_name, x.customer_part_number, x.customer_part_name,"
qSQL = qSQL & " y.name as material_name, z.name as color_name,x.cavity, x.weight_gr, x.weight_runner_gr, x.cycle_time_ia, x.prod_yield"
qSQL = qSQL & " from sip_production.eng_products x left join sip_production.eng_materials y on x.eng_material_id = y.id"
qSQL = qSQL & " left join sip_production.eng_colors z on x.eng_color_id = z.id) d on a.eng_product_id = d.id"
qSQL = qSQL & " left join sip_production.hrd_employees e on a.operator_1 = e.id"
qSQL = qSQL & " left join sip_production.hrd_employees f on a.operator_2 = f.id"
qSQL = qSQL & " where a.plant_mark = '" & p_plant_mark & "'"
qSQL = qSQL & " and b.number = '" & sM_number & "'"
qSQL = qSQL & " and d.internal_part_id = '" & sProd & "'"
qSQL = qSQL & " and a.period_shift = '" & Format(sDate, "yyyy-mm-dd") & "'"
qSQL = qSQL & " group by a.period_shift, a.plant_mark, a.prod_machine_id, a.mkt_customer_id, a.eng_product_id, a.period_shift"
qSQL = qSQL & " limit 1"

Set RS_PRINT = New ADODB.Recordset
If RS_PRINT.State = adStateOpen Then RS_PRINT.Close
RS_PRINT.Open qSQL, CN, adOpenDynamic, adLockPessimistic

    With RptProduksi
        .DTRpt.Recordset = RS_PRINT
        '.txtFrom.Text = Format(DTDari.Value, "dd/MMM/yyyy")
        '.txtTo.Text = Format(DTSampai.Value, "dd/MMM/yyyy")
        If sPrint = 0 Then
            .Show 1
'        ElseIf sPrint = 1 Then
'            xls.FileName = sFileName
'            xls.Version = 8
'            .Run False
'            xls.Export .Pages
'        ElseIf sPrint = 2 Then
'            pdf.FileName = sFileName
'            .Run False
'            pdf.Export .Pages
        End If
    
    End With
    
Exit Sub
ErrHandler:
    MsgBox "Error Number: " & Err.Number & vbCrLf & "Description: " & Err.Description, vbExclamation, Me.Caption
     
End Sub



