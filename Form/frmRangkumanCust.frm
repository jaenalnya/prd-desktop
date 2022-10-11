VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "LvButton.ocx"
Begin VB.Form frmRangkumanCust 
   Caption         =   "Rangkuman Customer"
   ClientHeight    =   5235
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10155
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5235
   ScaleWidth      =   10155
   Begin VB.PictureBox picFooter 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   405
      Left            =   0
      ScaleHeight     =   405
      ScaleWidth      =   10155
      TabIndex        =   0
      Top             =   4830
      Width           =   10155
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
   Begin MSComCtl2.DTPicker DTAwal 
      Height          =   330
      Left            =   2160
      TabIndex        =   9
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
      Format          =   85327875
      CurrentDate     =   43642
   End
   Begin MSComctlLib.ListView lvList 
      Height          =   4095
      Left            =   1440
      TabIndex        =   10
      Top             =   495
      Width           =   8550
      _ExtentX        =   15081
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
      Height          =   435
      Left            =   45
      TabIndex        =   11
      Top             =   1170
      Width           =   1290
      _ExtentX        =   2275
      _ExtentY        =   767
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
      Image           =   "frmRangkumanCust.frx":0000
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdRefresh 
      Height          =   450
      Left            =   45
      TabIndex        =   12
      Top             =   630
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   794
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
      Image           =   "frmRangkumanCust.frx":077A
      cBack           =   16119285
   End
   Begin lvButton.lvButtons_H cmdClose 
      Height          =   450
      Left            =   45
      TabIndex        =   13
      Top             =   2295
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   794
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
      Image           =   "frmRangkumanCust.frx":0EF4
      cBack           =   16119285
   End
   Begin lvButton.lvButtons_H cmdExport 
      Height          =   450
      Left            =   45
      TabIndex        =   14
      Top             =   1710
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   794
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
      Image           =   "frmRangkumanCust.frx":4156
      cBack           =   16119285
   End
   Begin MSComCtl2.DTPicker DTAkhir 
      Height          =   330
      Left            =   4230
      TabIndex        =   15
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
      Format          =   85327875
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
      Left            =   3870
      TabIndex        =   17
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
      Left            =   1485
      TabIndex        =   16
      Top             =   90
      Width           =   600
   End
End
Attribute VB_Name = "frmRangkumanCust"
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


Private Sub cmdClose_Click()
    CommandPass "Close"
End Sub

Private Sub cmdExport_Click()
    CommandPass "Export"
End Sub


Private Sub cmdRefresh_Click()
    CommandPass "Refresh"
End Sub

Private Sub CmdSearch_Click()
    CommandPass "Search"
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
    .Fields = "plant_mark,mkt_customer_id,customer_name,target_yield,sum(gross_produksi) as part_keluar_dr_mc,sum(net_produksi) barang_ok, " & _
              "round(AVG(prod_yield_persen),1) as prod_yield, round(AVG(persen_target),1) as persen_target, sum(run_hours) as Running_hr"
    .Tables = "(SELECT a.plant_mark,a.mkt_customer_id,e.name AS customer_name, " & _
        "a.eng_product_id,b.internal_part_id,b.name AS product_name,  b.cavity,b.prod_yield AS target_yield,b.weight_gr  AS part_weight, " & _
        "b.cycle_time_ia,a.period_shift, count(a.period_hour) AS jumlah_hour, (floor(3600/b.cycle_time_ia) * count(a.period_hour))  AS target_shot, sum(a.counter_ok) jumlah_shot, " & _
        "sum(a.counter_ok) * b.cavity AS gross_produksi, ifnull(data_ngs.jumlah_ng,0) AS jumlah_ng, " & _
        "ifnull(((sum(a.counter_ok) * b.cavity) - ifnull(data_ngs.jumlah_ng,0)),0) AS net_produksi, " & _
        "ifnull(round((((sum(a.counter_ok) * b.cavity) - ifnull(data_ngs.jumlah_ng,0)) / (sum(a.counter_ok) * b.cavity)),2),0) AS prod_yield, " & _
        "ifnull(round((((sum(a.counter_ok) * b.cavity) - ifnull(data_ngs.jumlah_ng,0)) / (sum(a.counter_ok) * b.cavity)) * 100,2),0) AS prod_yield_persen, " & _
        "ifnull(round(sum(a.counter_ok) / floor((3600/b.cycle_time_ia) * round(sum(a.counter_ok) / floor(3600/b.cycle_time_ia),1)) * 100,2),0) AS persen_target, " & _
        "ifnull(round(Sum(a.counter_ok) / floor(3600 / b.cycle_time_ia), 1), 0) As run_hours " & _
        "FROM prod_runnings a " & _
        "LEFT JOIN eng_products b ON a.eng_product_id = b.id " & _
        "LEFT JOIN (SELECT d.plant_mark,d.prod_machine_id, " & _
                    "d.eng_product_id,d.period_shift, sum(d.counter_ng) jumlah_ng FROM prod_data_ngs d GROUP BY d.plant_mark,d.prod_machine_id, " & _
                    "d.eng_product_id,d.period_shift) AS data_ngs ON a.plant_mark = data_ngs.plant_mark AND a.prod_machine_id = data_ngs.prod_machine_id " & _
                    "AND a.eng_product_id = data_ngs.eng_product_id AND a.period_shift = data_ngs.period_shift " & _
        "LEFT JOIN mkt_customers e ON a.mkt_customer_id = e.id GROUP BY a.plant_mark,a.mkt_customer_id,a.eng_product_id,a.period_shift) XX"

    .wCondition = "period_shift between '" & Format(DTAwal.Value, "yyyy-mm-dd") & "' and '" & Format(DTAkhir.Value, "yyyy-mm-dd") & "' and plant_mark = '" & p_plant_mark & "'"
    .GroupOrder = "plant_mark,mkt_customer_id"
    .SortOrder = "customer_name ASC"
    .SaveStatement
End With


Set RS_PRODBYCUST = New ADODB.Recordset
If RS_PRODBYCUST.State = adStateOpen Then RS_PRODBYCUST.Close
RS_PRODBYCUST.Open SQLParser.SQLStatement, CN, adOpenDynamic, adLockPessimistic
 
 

With RecordPage
    .Start RS_PRODBYCUST, 500
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
    Set frmRangkumanCust = Nothing
    Set RS_PRODBYCUST = Nothing
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
    srcProduct = lvList.SelectedItem.Index
    srcRecord = lvList.ListItems.Item(srcProduct).Text
    Call RefreshRecSum
End Sub


Public Sub CommandPass(ByVal srcPerformWhat As String)
On Error GoTo errPerformWhat
Select Case srcPerformWhat

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
'RS_PRODBYCUST.AbsolutePosition = RecordPage.PageStart
RecordPage.PageEnd
With lvList


    .GridLines = True
    .View = lvwReport

    .ColumnHeaders.Clear
    .ColumnHeaders.Add , , "No."
    .ColumnHeaders.Add , , "Customer name"
    .ColumnHeaders.Add , , "Target Yield"
    .ColumnHeaders.Add , , "Prod Yield"
    .ColumnHeaders.Add , , "Persen target"
    .ColumnHeaders.Add , , "Run hours"
    .ColumnHeaders.Add , , "Barang ok"
    .ColumnHeaders.Add , , "Part keluar MC"


    .ListItems.Clear
    Do While Not RS_PRODBYCUST.EOF
    Set srcItem = .ListItems.Add(, , i, 1, 1)
        srcItem.SubItems(1) = RS_PRODBYCUST.Fields("customer_name")
        srcItem.SubItems(2) = "95 %"
        srcItem.SubItems(3) = RS_PRODBYCUST.Fields("prod_yield") & " %"
        srcItem.SubItems(4) = RS_PRODBYCUST.Fields("persen_target") & " %"
        srcItem.SubItems(5) = RS_PRODBYCUST.Fields("Running_hr")
        srcItem.SubItems(6) = RS_PRODBYCUST.Fields("barang_ok")
        srcItem.SubItems(7) = RS_PRODBYCUST.Fields("part_keluar_dr_mc")

        
        If RS_PRODBYCUST.AbsolutePosition >= RecordPage.PageEnd Then
            Exit Do
        Else
            RS_PRODBYCUST.MoveNext
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
    lblRecSum.Caption = "Record: " & srcProduct & " of " & lvList.ListItems.Count
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
    
    Set RS_PRODBYCUST = New ADODB.Recordset
    If RS_PRODBYCUST.State = adStateOpen Then RS_PRODBYCUST.Close
    RS_PRODBYCUST.Open srcSQL, CN, adOpenDynamic, adLockPessimistic

With RecordPage
    .Start RS_PRODBYCUST, 500
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


