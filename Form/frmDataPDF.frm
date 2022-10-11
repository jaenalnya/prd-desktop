VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "LvButton.ocx"
Begin VB.Form frmDataPDF 
   Caption         =   "Data PDF Product"
   ClientHeight    =   5445
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9525
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5445
   ScaleWidth      =   9525
   Begin VB.Frame Frame1 
      Height          =   600
      Left            =   1485
      TabIndex        =   14
      Top             =   90
      Width           =   5865
      Begin lvButton.lvButtons_H cmdClear 
         Height          =   330
         Left            =   5355
         TabIndex        =   17
         Top             =   180
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   582
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
         Image           =   "frmDataPDF.frx":0000
         cBack           =   -2147483633
      End
      Begin VB.ComboBox cbocustomer 
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
         Left            =   1125
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   180
         Width           =   4155
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Customer"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   135
         TabIndex        =   15
         Top             =   225
         Width           =   1320
      End
   End
   Begin VB.PictureBox picFooter 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   405
      Left            =   0
      ScaleHeight     =   405
      ScaleWidth      =   9525
      TabIndex        =   0
      Top             =   5040
      Width           =   9525
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
      Left            =   1440
      TabIndex        =   9
      Top             =   765
      Width           =   7785
      _ExtentX        =   13732
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
      Left            =   45
      TabIndex        =   10
      Top             =   675
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
      Image           =   "frmDataPDF.frx":618A
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdRefresh 
      Height          =   405
      Left            =   45
      TabIndex        =   11
      Top             =   180
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
      Image           =   "frmDataPDF.frx":6904
      cBack           =   16119285
   End
   Begin lvButton.lvButtons_H cmdClose 
      Height          =   405
      Left            =   45
      TabIndex        =   12
      Top             =   1665
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
      Image           =   "frmDataPDF.frx":707E
      cBack           =   16119285
   End
   Begin lvButton.lvButtons_H cmdExport 
      Height          =   405
      Left            =   45
      TabIndex        =   13
      Top             =   1170
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
      Image           =   "frmDataPDF.frx":A2E0
      cBack           =   16119285
   End
End
Attribute VB_Name = "frmDataPDF"
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



Private Sub cboCustomer_Click()
Dim sSQL As String

    sSQL = "SELECT customer, internal_part_id,internal_part_id_prefix, customer_part_number, product_name"
    sSQL = sSQL & " FROM (SELECT b.name as customer, a.internal_part_id,a.internal_part_id_prefix,a.customer_part_number, a.name as product_name"
    sSQL = sSQL & " FROM sip_production.eng_products a inner join sip_production.mkt_customers b on a.mkt_customer_id = b.id"
    sSQL = sSQL & " WHERE a.internal_part_id_prefix = 'PIA' and a.plant_3 = 1 and a.status_plant_3 = 'active' or a.internal_part_id_prefix = 'PIA'"
    sSQL = sSQL & " and a.plant_4 = 1 and a.status_plant_4 = 'active')  tbl_x "
    sSQL = sSQL & " where customer = '" & cboCustomer.Text & "'"
    sSQL = sSQL & " ORDER BY customer ASC"

    Set Rs_search = New ADODB.Recordset
    If Rs_search.State = adStateOpen Then RS_PRODUCT.Close
    Rs_search.Open sSQL, CN, adOpenDynamic, adLockPessimistic

    With RecordPage
        .Start Rs_search, 500
    End With
    
    FillListview 1, Rs_search

End Sub



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
MAIN.ActivateChild Me

End Sub

Private Sub Form_Load()
On Error GoTo ErrHandler

ButtonList lvList, btnFirst, btnPrev, btnNext, btnLast
AddComboOne "sip_production.mkt_customers", "name", cboCustomer, "status", "active"

If p_plant_mark = "TSSI" Then

    With SQLParser
        .Fields = "customer, internal_part_id,internal_part_id_prefix, customer_part_number, product_name"
        .Tables = "(SELECT b.name as customer, a.internal_part_id,a.internal_part_id_prefix,a.customer_part_number, a.name as product_name " & _
                "FROM sip_production.eng_products a inner join sip_production.mkt_customers b on a.mkt_customer_id = b.id  " & _
                "WHERE substring(a.internal_part_id_prefix,1,2) = 'PI' " & _
                "and a.plant_3 = 1 and a.status_plant_3 = 'active' or  a.plant_4 = 1 and a.status_plant_4 = 'active' )  tbl_x"
        .SortOrder = "customer ASC"
        .SaveStatement
    End With
    
ElseIf p_plant_mark = "Techno" Then

    With SQLParser
        .Fields = "customer, internal_part_id,internal_part_id_prefix, customer_part_number, product_name"
        .Tables = "(SELECT b.name as customer, a.internal_part_id,a.internal_part_id_prefix,a.customer_part_number, a.name as product_name " & _
            "FROM sip_production.eng_products a inner join sip_production.mkt_customers b on a.mkt_customer_id = b.id " & _
            "WHERE a.plant_2 = 1 and a.status_plant_2 = 'active'  AND substring(a.internal_part_id_prefix,1,2) = 'PI')  tbl_x"
        .SortOrder = "customer ASC"
        .SaveStatement
    End With
    
ElseIf p_plant_mark = "Cempaka" Then

    With SQLParser
        .Fields = "customer, internal_part_id,internal_part_id_prefix, customer_part_number, product_name"
        .Tables = "(SELECT b.name as customer, a.internal_part_id,a.internal_part_id_prefix,a.customer_part_number, a.name as product_name " & _
            "FROM sip_production.eng_products a inner join sip_production.mkt_customers b on a.mkt_customer_id = b.id " & _
            "WHERE a.plant_2 = 1 and a.status_plant_2 = 'active'  AND substring(a.internal_part_id_prefix,1,2) = 'PI')  tbl_x"
        .SortOrder = "customer ASC"
        .SaveStatement
    End With
    
End If

Set RS_PRODUCT = New ADODB.Recordset
If RS_PRODUCT.State = adStateOpen Then RS_PRODUCT.Close
RS_PRODUCT.Open SQLParser.SQLStatement, CN, adOpenDynamic, adLockPessimistic
 

With RecordPage
    .Start RS_PRODUCT, 500
End With

FillListview 1, RS_PRODUCT

srcProduct = "NONE"
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
        lvList.Height = Me.ScaleHeight - 1300
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
MAIN.RemoveChild Me.Name
    Set frmDataPDF = Nothing
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

Private Sub FillListview(ByVal whichPage As Long, Rs As Recordset)
On Error Resume Next
Dim i As Integer
RecordPage.CurrentPosition = whichPage
Rs.AbsolutePosition = RecordPage.PageStart
RecordPage.PageEnd
With lvList
    .GridLines = True
    .View = lvwReport

    .ColumnHeaders.Clear
    .ColumnHeaders.Add , , "No."
    .ColumnHeaders.Add , , "customer"
    .ColumnHeaders.Add , , "internal_part_id"
    .ColumnHeaders.Add , , "customer_part_number"
    .ColumnHeaders.Add , , "product_name"
    .ColumnHeaders.Add , , "work_instruction"
    .ColumnHeaders.Add , , "packing_standard"
    .ColumnHeaders.Add , , "critical_check"
    

    .ListItems.Clear
    Do While Not Rs.EOF
    i = i + 1
    Set srcItem = .ListItems.Add(, , i, 1, 1)
        srcItem.SubItems(1) = Format(Rs.Fields("customer"), "yyyy-mm-dd")
        srcItem.SubItems(2) = Rs.Fields("internal_part_id")
        srcItem.SubItems(3) = Rs.Fields("customer_part_number")
        srcItem.SubItems(4) = Rs.Fields("product_name")

        If Dir(App.Path & "\Picture\WI\" & Rs.Fields("internal_part_id") & ".pdf") <> "" Then
            srcItem.SubItems(5) = "Ready"
        Else
            srcItem.SubItems(5) = "-"
        End If

        If Dir(App.Path & "\Picture\PS\" & Rs.Fields("internal_part_id") & ".pdf") <> "" Then
            srcItem.SubItems(6) = "Ready"
        Else
            srcItem.SubItems(6) = "-"
        End If

        If Dir(App.Path & "\Picture\CCP\" & Rs.Fields("internal_part_id") & ".pdf") <> "" Then
            srcItem.SubItems(7) = "Ready"
        Else
            srcItem.SubItems(7) = "-"
        End If
        
        If Rs.AbsolutePosition >= RecordPage.PageEnd Then
            Exit Do
        Else
            Rs.MoveNext
        End If
    Loop
End With
SetNavigation btnFirst, btnPrev, btnNext, btnLast
'Display the page information
lblPageInfo.Caption = "Record " & RecordPage.PageInfo
'Display the selected record
Call RefreshRecSum
Call lvSizeColumns(lvList)
'AltLVBackground lvList, vbWhite, &H8000000F, frmDataPDF
End Sub


Private Sub RefreshRecSum()
    lblRecSum.Caption = "Record: " & srcProduct & " of " & lvList.ListItems.Count
End Sub


Private Sub btnFirst_Click()
    If RecordPage.PAGE_CURRENT <> 1 Then FillListview 1, RS_PRODUCT
End Sub

Private Sub btnLast_Click()
    If RecordPage.PAGE_CURRENT <> RecordPage.PAGE_TOTAL Then FillListview RecordPage.PAGE_TOTAL, RS_PRODUCT
End Sub

Private Sub btnNext_Click()
    If RecordPage.PAGE_CURRENT <> RecordPage.PAGE_TOTAL Then FillListview RecordPage.PAGE_NEXT, RS_PRODUCT
End Sub

Private Sub btnPrev_Click()
    If RecordPage.PAGE_CURRENT <> 1 Then FillListview RecordPage.PAGE_PREVIOUS, RS_PRODUCT
End Sub

Public Sub FilterRecord(ByVal srcCondition As String)
    SQLParser.RestoreStatement
    SQLParser.wCondition = srcCondition
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

FillListview 1, RS_PRODUCT

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





