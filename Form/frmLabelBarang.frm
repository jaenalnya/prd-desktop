VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{9EDDC69F-10E8-4DE7-9648-EC8A45F005C0}#1.0#0"; "B8CONT~2.OCX"
Begin VB.Form frmLabelBarang 
   Caption         =   "Data Label Barang OK"
   ClientHeight    =   5250
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9435
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5250
   ScaleWidth      =   9435
   Begin VB.PictureBox picFooter 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   405
      Left            =   0
      ScaleHeight     =   405
      ScaleWidth      =   9435
      TabIndex        =   0
      Top             =   4845
      Width           =   9435
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
   Begin b8Controls4.b8TitleBar b8TitleBar1 
      Height          =   375
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   9105
      _ExtentX        =   16060
      _ExtentY        =   661
      Caption         =   "Data Label Barang OK"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16777215
      BackColor       =   8421504
   End
   Begin MSComctlLib.ListView lvList 
      Height          =   3960
      Left            =   1530
      TabIndex        =   10
      Top             =   450
      Width           =   7560
      _ExtentX        =   13335
      _ExtentY        =   6985
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
   Begin lvButton.lvButtons_H cmdSearch 
      Height          =   390
      Left            =   135
      TabIndex        =   11
      Top             =   1440
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
      Image           =   "frmLabelBarang.frx":0000
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdRefresh 
      Height          =   405
      Left            =   135
      TabIndex        =   12
      Top             =   945
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
      Image           =   "frmLabelBarang.frx":077A
      cBack           =   16119285
   End
   Begin lvButton.lvButtons_H cmdClose 
      Height          =   405
      Left            =   135
      TabIndex        =   13
      Top             =   2385
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
      Image           =   "frmLabelBarang.frx":0EF4
      cBack           =   16119285
   End
   Begin lvButton.lvButtons_H cmdExport 
      Height          =   405
      Left            =   135
      TabIndex        =   14
      Top             =   1890
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
      Image           =   "frmLabelBarang.frx":4156
      cBack           =   16119285
   End
   Begin lvButton.lvButtons_H cmdNew 
      Height          =   405
      Left            =   135
      TabIndex        =   15
      Top             =   450
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   714
      Caption         =   "&New [F3]"
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
      cBhover         =   16119285
      LockHover       =   3
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      Image           =   "frmLabelBarang.frx":44F0
      cBack           =   16777215
   End
End
Attribute VB_Name = "frmLabelBarang"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim srcItem                         As ListItem
Dim srcRecord                       As String
Dim srcLabelBarang                  As Variant
Dim srcSQL                          As String
Dim RecordPage                      As New clsPaging
Dim SQLParser                       As New clsSQLSelectParser
Private Sub cmdClose_Click()
    CommandPass "Close"
End Sub

Private Sub cmdExport_Click()
    CommandPass "Export"
End Sub



Private Sub cmdNew_Click()
    CommandPass "New"
End Sub

Private Sub cmdRefresh_Click()
    CommandPass "Refresh"
End Sub

Private Sub CmdSearch_Click()
    CommandPass "Search"
End Sub

Private Sub Form_Activate()
'On Error Resume Next
With MAIN
    Me.BackColor = .ACPMenu.BackColor
    Me.Picture = .ACPMenu.LoadBackground
    picFooter.BackColor = .ACPMenu.BackColor
    Picture2.BackColor = .ACPMenu.BackColor
End With
lvList.FlatScrollBar = False
'AltLVBackground lvList, vbWhite, &H8000000F, frmAdjustment

End Sub

Private Sub Form_Load()
'On Error GoTo ErrHandler

ButtonList lvList, btnFirst, btnPrev, btnNext, btnLast

With SQLParser
    .Fields = "*"
    .Tables = "(select a.plant_mark,a.prod_machine_id,b.number as machine_no,b.name as machine_name, a.mkt_customer_id,c.name as customer_name, " & _
                "a.eng_product_id, d.internal_part_id,d.name as product_name, " & _
                "a.period_shift , a.period_hour, a.adjust, a.qty, a.Description, a.created_at, a.created_by,e.name as account_name " & _
                "from prod_adjustment_shots a " & _
                "left join prod_machines b on a.prod_machine_id = b.id " & _
                "left join mkt_customers c on a.mkt_customer_id = c.id " & _
                "left join sys_accounts e on a.created_by = e.id " & _
                "left join eng_products d on a.eng_product_id = d.id) as Adjustment"
    .wCondition = "period_shift = '" & Format(Now, "yyyy-mm-dd") & "'"
    .SaveStatement
End With

Set RS_LABELOK = New ADODB.Recordset
If RS_LABELOK.State = adStateOpen Then RS_LABELOK.Close
RS_LABELOK.Open SQLParser.SQLStatement, CN, adOpenDynamic, adLockPessimistic


With RecordPage
    .Start RS_LABELOK, 500
End With

FillListview 1

srcLabelBarang = "NONE"
srcRecord = vbNullString

Exit Sub
ErrHandler:
    MsgBox "Error Number: " & Err.Number & vbCrLf & "Description: " & Err.Description, vbExclamation
End Sub

Private Sub Form_Resize()
'On Error Resume Next
    If WindowState <> vbMinimized Then
        If Me.Width < 9195 Then Me.Width = 9195
        If Me.Height < 4500 Then Me.Height = 4500
        
        b8TitleBar1.Width = ScaleWidth
        Liner1.Width = ScaleWidth
        lvList.Width = Me.ScaleWidth - 1550
        lvList.Height = Me.ScaleHeight - 1000
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmAdjustment = Nothing
    Set RS_LABELOK = Nothing
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
'On Error Resume Next
    srcLabelBarang = lvList.SelectedItem.Index
    srcRecord = lvList.ListItems.Item(srcLabelBarang).Text
    Call RefreshRecSum
End Sub


Public Sub CommandPass(ByVal srcPerformWhat As String)
'On Error GoTo errPerformWhat
Select Case srcPerformWhat
 
    Case "New" 'Refresh
           With frmAdjustmentAE
                .State = AddStateMode
                .Show vbModal
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
'On Error Resume Next
Dim i As Integer
i = 1
RecordPage.CurrentPosition = whichPage
'RS_LABELOK.AbsolutePosition = RecordPage.PageStart
RecordPage.PageEnd
With lvList
    .GridLines = True
    .View = lvwReport

    .ColumnHeaders.Clear
    .ColumnHeaders.Add , , "No."
    .ColumnHeaders.Add , , "Plant Mark"
    .ColumnHeaders.Add , , "Machine No"
    .ColumnHeaders.Add , , "Internal Part Number"
    .ColumnHeaders.Add , , "Product Name"
    .ColumnHeaders.Add , , "Period Shift"
    .ColumnHeaders.Add , , "Period Hour"
    .ColumnHeaders.Add , , "Adjust"
    .ColumnHeaders.Add , , "Qty"
    .ColumnHeaders.Add , , "Description"
    .ColumnHeaders.Add , , "Created At"
    .ColumnHeaders.Add , , "Created By"

    .ListItems.Clear
    Do While Not RS_LABELOK.EOF
    Set srcItem = .ListItems.Add(, , i, 1, 1)
        srcItem.SubItems(1) = RS_LABELOK.Fields("plant_mark")
        srcItem.SubItems(2) = RS_LABELOK.Fields("machine_no")
        srcItem.SubItems(3) = RS_LABELOK.Fields("internal_part_id")
        srcItem.SubItems(4) = RS_LABELOK.Fields("product_name")
        srcItem.SubItems(5) = RS_LABELOK.Fields("period_shift")
        srcItem.SubItems(6) = RS_LABELOK.Fields("period_hour")
        srcItem.SubItems(7) = RS_LABELOK.Fields("adjust")
        srcItem.SubItems(8) = RS_LABELOK.Fields("qty")
        srcItem.SubItems(9) = RS_LABELOK.Fields("description")
        srcItem.SubItems(10) = RS_LABELOK.Fields("created_at")
        srcItem.SubItems(11) = RS_LABELOK.Fields("account_name")

        If RS_LABELOK.AbsolutePosition >= RecordPage.PageEnd Then
            Exit Do
        Else
            RS_LABELOK.MoveNext
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

End Sub


Private Sub RefreshRecSum()
    lblRecSum.Caption = "Record: " & srcLabelBarang & " of " & lvList.ListItems.Count
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
'On Error Resume Next
    
    Set RS_LABELOK = New ADODB.Recordset
    If RS_LABELOK.State = adStateOpen Then RS_LABELOK.Close
    RS_LABELOK.Open srcSQL, CN, adOpenDynamic, adLockPessimistic

With RecordPage
    .Start RS_LABELOK, 500
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
'On Error Resume Next
    Select Case KeyCode
    Case vbKeyF3
        CommandPass "New"
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






