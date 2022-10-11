VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "LvButton.ocx"
Begin VB.Form frmParameter 
   Caption         =   "Parameter Setting  Mesin"
   ClientHeight    =   5130
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9315
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5130
   ScaleWidth      =   9315
   Begin VB.PictureBox picFooter 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   405
      Left            =   0
      ScaleHeight     =   405
      ScaleWidth      =   9315
      TabIndex        =   0
      Top             =   4725
      Width           =   9315
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   345
         Left            =   5175
         ScaleHeight     =   345
         ScaleWidth      =   4155
         TabIndex        =   1
         Top             =   45
         Width           =   4150
         Begin VB.CommandButton btnFirst 
            Height          =   315
            Left            =   2760
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "First 250"
            Top             =   10
            Width           =   315
         End
         Begin VB.CommandButton btnPrev 
            Height          =   315
            Left            =   3075
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Previous 250"
            Top             =   10
            Width           =   315
         End
         Begin VB.CommandButton btnLast 
            Height          =   315
            Left            =   3705
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Last 250"
            Top             =   10
            Width           =   315
         End
         Begin VB.CommandButton btnNext 
            Height          =   315
            Left            =   3390
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Next 250"
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
      Top             =   45
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
      Left            =   90
      TabIndex        =   10
      Top             =   1035
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
      Image           =   "frmParameter.frx":0000
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdRefresh 
      Height          =   405
      Left            =   90
      TabIndex        =   11
      Top             =   540
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
      Image           =   "frmParameter.frx":077A
      cBack           =   16119285
   End
   Begin lvButton.lvButtons_H cmdClose 
      Height          =   405
      Left            =   90
      TabIndex        =   12
      Top             =   2070
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
      Image           =   "frmParameter.frx":0EF4
      cBack           =   16119285
   End
   Begin lvButton.lvButtons_H cmdExport 
      Height          =   405
      Left            =   90
      TabIndex        =   13
      Top             =   1485
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
      Image           =   "frmParameter.frx":4156
      cBack           =   16119285
   End
   Begin lvButton.lvButtons_H cmdView 
      Height          =   405
      Left            =   90
      TabIndex        =   14
      Top             =   45
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   714
      Caption         =   "&View [F4]"
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
      Image           =   "frmParameter.frx":44F0
      cBack           =   16119285
   End
End
Attribute VB_Name = "frmParameter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim srcItem                         As ListItem
Dim srcRecord                       As String
Dim srcNg                           As Variant
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


Private Sub cmdView_Click()
    On Error Resume Next
    Dim qSQL As String
    
    qSQL = "SELECT a.*,b.name AS plant_name,c.name AS customer_name,d.number AS machine_number, "
    qSQL = qSQL & " d.name AS machine_name,d.tonnage,e.internal_part_id,e.name AS product_name,"
    qSQL = qSQL & " e.customer_part_number, e.customer_part_name,e.cavity,e.eng_material_id,"
    qSQL = qSQL & " e.eng_color_id,e.weight_gr,e.weight_shot_gr, e.weight_runner_gr ,"
    qSQL = qSQL & " e.prod_yield,e.material_name,e.color_name FROM eng_paramset_standards a"
    qSQL = qSQL & " LEFT JOIN sys_plants b ON a.sys_plant_id = b.id"
    qSQL = qSQL & " LEFT JOIN mkt_customers c ON a.mkt_customer_id = c.id"
    qSQL = qSQL & " LEFT JOIN prod_machines d ON a.sys_plant_id = d.id"
    qSQL = qSQL & " LEFT JOIN (SELECT a.*,b.name AS material_name, c.name AS color_name FROM eng_products a"
                    qSQL = qSQL & " LEFT JOIN eng_materials b ON a.eng_material_for_label_id = b.id"
                    qSQL = qSQL & " LEFT JOIN eng_colors c ON a.eng_color_id = c.id) e ON a.eng_product_id = e.id "
    qSQL = qSQL & " WHERE a.mkt_customer_id = '" & lvList.SelectedItem.SubItems(11) & "' AND"
    qSQL = qSQL & " a.prod_machine_id = '" & lvList.SelectedItem.SubItems(10) & "' AND"
    qSQL = qSQL & " a.eng_product_id = '" & lvList.SelectedItem.SubItems(9) & "'"
    qSQL = qSQL & " ORDER BY a.rev DESC LIMIT 1"
    
    
    Set RS_PRINT = New ADODB.Recordset
    If RS_PRINT.State = adStateOpen Then RS_PRINT.Close
    RS_PRINT.Open qSQL, CN, adOpenDynamic, adLockPessimistic
    With RptStdParameter
        .DTRpt.Recordset = RS_PRINT
        .lblDate.Caption = Now
        .lblCompany.Caption = ACTIVE_COMPANY.Perusahaan
        .lblAlamat.Caption = ACTIVE_COMPANY.Alamat
        .Show
    End With
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
'AltLVBackground lvList, vbWhite, &H8000000F, frmParameter
MAIN.ActivateChild Me

End Sub

Private Sub Form_Load()
On Error GoTo ErrHandler

ButtonList lvList, btnFirst, btnPrev, btnNext, btnLast

With SQLParser
    .Fields = "*"
    .Tables = "(SELECT a.id,a.sys_plant_id,a.rev,a.rev_date,a.mkt_customer_id,b.NAME AS customer_name, a.prod_machine_id, c.NUMBER AS machine_no, " & _
                "c.NAME AS machine_name,c.tonnage,a.eng_product_id,d.internal_part_id,d.NAME AS product_name FROM eng_paramset_standards a " & _
                "LEFT JOIN sip_production.mkt_customers b ON a.mkt_customer_id = b.id " & _
                "LEFT JOIN sip_production.prod_machines c ON a.prod_machine_id = c.id " & _
                "LEFT JOIN sip_production.eng_products d ON a.eng_product_id = d.id) parameter"
    .wCondition = "machine_no = '" & p_machine_no & "' and sys_plant_id IN (" & p_sys_plant & ")"
    .SortOrder = "machine_no ASC"
    .SaveStatement
End With

Set RS_PARAMETER = New ADODB.Recordset
If RS_PARAMETER.State = adStateOpen Then RS_PARAMETER.Close
RS_PARAMETER.Open SQLParser.SQLStatement, CN, adOpenDynamic, adLockPessimistic


With RecordPage
    .Start RS_PARAMETER, 500
End With

FillListview 1

srcNg = "NONE"
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

        Liner1.Width = ScaleWidth
        lvList.Width = Me.ScaleWidth - 1550
        lvList.Height = Me.ScaleHeight - 800
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
MAIN.RemoveChild Me.Name
    Set frmParameter = Nothing
    Set RS_PARAMETER = Nothing
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
    srcNg = lvList.SelectedItem.Index
    srcRecord = lvList.ListItems.Item(srcNg).text
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
'RS_PARAMETER.AbsolutePosition = RecordPage.PageStart
RecordPage.PageEnd
With lvList
    .GridLines = True
    .View = lvwReport

    .ColumnHeaders.Clear
    .ColumnHeaders.Add , , "No."
    .ColumnHeaders.Add , , "machine_no"
    .ColumnHeaders.Add , , "machine_name"
    .ColumnHeaders.Add , , "tonnage"
    .ColumnHeaders.Add , , "customer_name"
    .ColumnHeaders.Add , , "internal_part_id"
    .ColumnHeaders.Add , , "product_name"
    .ColumnHeaders.Add , , "rev"
    .ColumnHeaders.Add , , "rev_date"
    .ColumnHeaders.Add , , "eng_product_id"
    .ColumnHeaders.Add , , "prod_machine_id"
    .ColumnHeaders.Add , , "mkt_customer_id"

    .ListItems.Clear
    Do While Not RS_PARAMETER.EOF
    Set srcItem = .ListItems.Add(, , i, 1, 1)
        srcItem.SubItems(1) = RS_PARAMETER.Fields("machine_no")
        srcItem.SubItems(2) = RS_PARAMETER.Fields("machine_name")
        srcItem.SubItems(3) = RS_PARAMETER.Fields("tonnage")
        srcItem.SubItems(4) = RS_PARAMETER.Fields("customer_name")
        srcItem.SubItems(5) = RS_PARAMETER.Fields("internal_part_id")
        srcItem.SubItems(6) = RS_PARAMETER.Fields("product_name")
        srcItem.SubItems(7) = RS_PARAMETER.Fields("rev")
        srcItem.SubItems(8) = Format(RS_PARAMETER.Fields("rev_date"), "yyyy-mm-dd")
        srcItem.SubItems(9) = RS_PARAMETER.Fields("eng_product_id")
        srcItem.SubItems(10) = RS_PARAMETER.Fields("prod_machine_id")
        srcItem.SubItems(11) = RS_PARAMETER.Fields("mkt_customer_id")
        

        
        If RS_PARAMETER.AbsolutePosition >= RecordPage.PageEnd Then
            Exit Do
        Else
            RS_PARAMETER.MoveNext
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

lvList.ColumnHeaders(10).Width = 0
lvList.ColumnHeaders(11).Width = 0
lvList.ColumnHeaders(12).Width = 0
End Sub


Private Sub RefreshRecSum()
    lblRecSum.Caption = "Record: " & srcNg & " of " & lvList.ListItems.Count
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
    
    Set RS_PARAMETER = New ADODB.Recordset
    If RS_PARAMETER.State = adStateOpen Then RS_PARAMETER.Close
    RS_PARAMETER.Open srcSQL, CN, adOpenDynamic, adLockPessimistic

With RecordPage
    .Start RS_PARAMETER, 500
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










