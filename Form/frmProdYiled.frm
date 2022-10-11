VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{9EDDC69F-10E8-4DE7-9648-EC8A45F005C0}#1.0#0"; "b8Controls4.ocx"
Begin VB.Form frmProdYiled 
   ClientHeight    =   7245
   ClientLeft      =   120
   ClientTop       =   120
   ClientWidth     =   10695
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmProdYiled.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7245
   ScaleWidth      =   10695
   Begin VB.PictureBox picFooter 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   405
      Left            =   0
      ScaleHeight     =   405
      ScaleWidth      =   10695
      TabIndex        =   0
      Top             =   6840
      Width           =   10695
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
      Height          =   3960
      Left            =   1485
      TabIndex        =   9
      Top             =   450
      Width           =   7560
      _ExtentX        =   13335
      _ExtentY        =   6985
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
   Begin lvButton.lvButtons_H cmdSearch 
      Height          =   480
      Left            =   90
      TabIndex        =   10
      Top             =   1035
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
      Image           =   "frmProdYiled.frx":617A
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdRefresh 
      Height          =   495
      Left            =   90
      TabIndex        =   11
      Top             =   450
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
      Image           =   "frmProdYiled.frx":68F4
      cBack           =   16119285
   End
   Begin lvButton.lvButtons_H cmdClose 
      Height          =   495
      Left            =   90
      TabIndex        =   12
      Top             =   2205
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
      Image           =   "frmProdYiled.frx":706E
      cBack           =   16119285
   End
   Begin lvButton.lvButtons_H cmdExport 
      Height          =   495
      Left            =   90
      TabIndex        =   13
      Top             =   1620
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
      Image           =   "frmProdYiled.frx":A2D0
      cBack           =   16119285
   End
   Begin b8Controls4.b8TitleBar b8TB 
      Height          =   375
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   10680
      _ExtentX        =   18838
      _ExtentY        =   661
      Caption         =   "Data Harian Total Yield"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16777215
      ShadowColor     =   12632319
      BackColor       =   8421504
   End
End
Attribute VB_Name = "frmProdYiled"
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

Private Sub cmdUpdate_Click()
    CommandPass "Update"
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

End Sub

Private Sub Form_Load()
On Error GoTo ErrHandler
Dim p_shift As Date
ButtonList lvList, btnFirst, btnPrev, btnNext, btnLast

If Format(Now, "HH") >= 0 And Format(Now, "HH") <= 7 Then
    p_shift = Format(DateAdd("d", -1, Format(Now, "yyyy-mm-dd")), "yyyy-mm-dd")
Else
    p_shift = Format(Now, "yyyy-mm-dd")
End If
    
With SQLParser
    .Fields = "xx.plant_mark,xx.prod_machine_id,xx.number,xx.machine_name,xx.tonnage, xx.mkt_customer_id, " & _
         "xx.eng_product_id, xx.internal_part_id,xx.product_name, " & _
         "xx.period_shift,xx.jumlah_jam, xx.shot,xx.cavity,(xx.shot * xx.cavity) as gross_produksi, " & _
         "ifnull(xx.ng,0) as ng,((xx.shot * xx.cavity)-ifnull(xx.ng,0)) as net_produksi, " & _
         "xx.cycle_time_ia,(((xx.shot * xx.cavity)-ifnull(xx.ng,0)) / ((3600 / xx.cycle_time_ia) * xx.jumlah_jam)) * 100 as Total_yiled "
    .Tables = "(select a.plant_mark,a.prod_machine_id, c.number, c.name as machine_name,c.tonnage, " & _
            "a.mkt_customer_id,a.eng_product_id,b.internal_part_id,b.name as product_name, a.period_shift, " & _
            "count(a.period_hour) as jumlah_jam, sum(a.counter_ok) as shot,b.cavity,b.cycle_time_ia,d.ng " & _
            "from sip_production.prod_runnings a " & _
            "left join sip_234.eng_products b on a.eng_product_id = b.id " & _
            "left join sip_234.prod_machines c on a.prod_machine_id = c.id " & _
            "left join (select plant_mark,prod_machine_id,mkt_customer_id,eng_product_id,period_shift,count(prod_ng_id) as ng " & _
            "from sip_production.prod_ng_logs group by plant_mark,prod_machine_id,mkt_customer_id,eng_product_id,period_shift) d " & _
            "on a.plant_mark = d.plant_mark and a.prod_machine_id = d.prod_machine_id and a.mkt_customer_id = d.mkt_customer_id and " & _
            "a.eng_product_id = d.eng_product_id and a.period_shift = d.period_shift " & _
            "Group by a.plant_mark,a.prod_machine_id,a.mkt_customer_id,a.eng_product_id,a.period_shift) xx"
    .wCondition = "period_shift = '" & Format(p_shift, "yyyy-mm-dd") & "'"
    .SortOrder = "period_shift, number Asc"
    .SaveStatement
End With

Set RS_MONITOR = New ADODB.Recordset
If RS_MONITOR.State = adStateOpen Then RS_MONITOR.Close
RS_MONITOR.Open SQLParser.SQLStatement, CN, adOpenDynamic, adLockPessimistic
 

With RecordPage
    .Start RS_MONITOR, 500
End With

FillListview 1

srcMonitor = "NONE"
srcRecord = vbNullString

Exit Sub
ErrHandler:
    MsgBox "Error Number: " & Err.Number & vbCrLf & "Description: " & Err.Description, vbExclamation
End Sub

Private Sub Form_Resize()
On Error Resume Next
    If WindowState <> vbMinimized Then
        If Me.Width < 9195 Then Me.Width = 9195
        If Me.Height < 4500 Then Me.Height = 4500
        
        'b8TB.Width = ScaleWidth
        Liner1.Width = ScaleWidth
        lvList.Width = Me.ScaleWidth - 1550
        lvList.Height = Me.ScaleHeight - 1000
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmProdYiled = Nothing
    Set RS_MONITOR = Nothing
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
    srcMonitor = lvList.SelectedItem.Index
    srcRecord = lvList.ListItems.Item(srcMonitor).Text
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
'RS_MONITOR.AbsolutePosition = RecordPage.PageStart
RecordPage.PageEnd
With lvList
    .GridLines = True
    .View = lvwReport

    .ColumnHeaders.Clear
    .ColumnHeaders.Add , , "period_shift"
    .ColumnHeaders.Add , , "machine_no"
    .ColumnHeaders.Add , , "machine_name"
    .ColumnHeaders.Add , , "tonnage"
    .ColumnHeaders.Add , , "internal_part_id"
    .ColumnHeaders.Add , , "product_name"
    .ColumnHeaders.Add , , "jumlah_jam"
    .ColumnHeaders.Add , , "shot"
    .ColumnHeaders.Add , , "cavity"
    .ColumnHeaders.Add , , "gross_produksi"
    .ColumnHeaders.Add , , "ng"
    .ColumnHeaders.Add , , "net_produksi"
    .ColumnHeaders.Add , , "cycle_time_ia"
    .ColumnHeaders.Add , , "Total_Yield"


    .ListItems.Clear
    Do While Not RS_MONITOR.EOF
    Set srcItem = .ListItems.Add(, , Format(RS_MONITOR.Fields("period_shift"), "yyyy-mm-dd"), 1, 1)
        srcItem.SubItems(1) = RS_MONITOR.Fields("number")
        srcItem.SubItems(2) = RS_MONITOR.Fields("machine_name")
        srcItem.SubItems(3) = RS_MONITOR.Fields("tonnage")
        srcItem.SubItems(4) = RS_MONITOR.Fields("internal_part_id")
        srcItem.SubItems(5) = RS_MONITOR.Fields("product_name")
        srcItem.SubItems(6) = RS_MONITOR.Fields("jumlah_jam")
        srcItem.SubItems(7) = RS_MONITOR.Fields("shot")
        srcItem.SubItems(8) = RS_MONITOR.Fields("cavity")
        srcItem.SubItems(9) = RS_MONITOR.Fields("gross_produksi") '(xx.shot * xx.cavity) as gross_produksi
        srcItem.SubItems(10) = RS_MONITOR.Fields("ng")
        srcItem.SubItems(11) = RS_MONITOR.Fields("net_produksi") '((xx.shot * xx.cavity)-ifnull(xx.ng,0)) as net_produksi
        srcItem.SubItems(12) = RS_MONITOR.Fields("cycle_time_ia")
        srcItem.SubItems(13) = Round(RS_MONITOR.Fields("Total_yiled"), 2) & " %"
        '(((xx.shot * xx.cavity)-ifnull(xx.ng,0)) / ((3600 / xx.cycle_time_ia) * xx.jumlah_jam)) * 100 as Total_yiled
        
        If RS_MONITOR.AbsolutePosition >= RecordPage.PageEnd Then
            Exit Do
        Else
            RS_MONITOR.MoveNext
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

AltLVBackground lvList, vbWhite, &HFFFFC0, frmProdYiled

End Sub


Private Sub RefreshRecSum()
    lblRecSum.Caption = "Record: " & srcMonitor & " of " & lvList.ListItems.Count
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
    SQLParser.wCondition = srcCondition
    ReloadRecords SQLParser.SQLStatement
End Sub

'Procedure for reloadingrecords
Public Sub ReloadRecords(ByVal srcSQL As String)
On Error Resume Next
    
    Set RS_MONITOR = New ADODB.Recordset
    If RS_MONITOR.State = adStateOpen Then RS_MONITOR.Close
    RS_MONITOR.Open srcSQL, CN, adOpenDynamic, adLockPessimistic

With RecordPage
    .Start RS_MONITOR, 500
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

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
    Select Case KeyCode
    Case vbKeyF5
        CommandPass "Refresh"
    Case vbKeyF6
        CommandPass "Search"
    Case vbKeyF8
        CommandPass "Export"
    Case vbKeyEscape
        CommandPass "Close"
    End Select
End Sub



