VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "LvButton.ocx"
Begin VB.Form frmRptLabel 
   Caption         =   "Laporan Label"
   ClientHeight    =   5160
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   9615
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmRptLabel.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5160
   ScaleWidth      =   9615
   Begin VB.PictureBox picFooter 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   405
      Left            =   0
      ScaleHeight     =   405
      ScaleWidth      =   9615
      TabIndex        =   2
      Top             =   4755
      Width           =   9615
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   345
         Left            =   5175
         ScaleHeight     =   345
         ScaleWidth      =   4155
         TabIndex        =   3
         Top             =   45
         Width           =   4150
         Begin VB.CommandButton btnFirst 
            Height          =   315
            Left            =   2760
            Style           =   1  'Graphical
            TabIndex        =   7
            ToolTipText     =   "First 250"
            Top             =   10
            Width           =   315
         End
         Begin VB.CommandButton btnPrev 
            Height          =   315
            Left            =   3075
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Previous 250"
            Top             =   10
            Width           =   315
         End
         Begin VB.CommandButton btnLast 
            Height          =   315
            Left            =   3705
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Last 250"
            Top             =   10
            Width           =   315
         End
         Begin VB.CommandButton btnNext 
            Height          =   315
            Left            =   3390
            Style           =   1  'Graphical
            TabIndex        =   4
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
            TabIndex        =   8
            Top             =   60
            Width           =   2535
         End
      End
      Begin PRD.Liner Liner1 
         Height          =   30
         Left            =   0
         TabIndex        =   9
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
         TabIndex        =   10
         Top             =   120
         Width           =   690
      End
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
      Left            =   2925
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   45
      Width           =   1365
   End
   Begin MSComCtl2.DTPicker DTAwal 
      Height          =   330
      Left            =   5535
      TabIndex        =   0
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
      Height          =   3960
      Left            =   1440
      TabIndex        =   11
      Top             =   495
      Width           =   7830
      _ExtentX        =   13811
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
      TabIndex        =   12
      Top             =   1620
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
      Image           =   "frmRptLabel.frx":617A
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdRefresh 
      Height          =   450
      Left            =   45
      TabIndex        =   13
      Top             =   1035
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
      Image           =   "frmRptLabel.frx":68F4
      cBack           =   16119285
   End
   Begin lvButton.lvButtons_H cmdClose 
      Height          =   450
      Left            =   45
      TabIndex        =   14
      Top             =   2790
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
      Image           =   "frmRptLabel.frx":706E
      cBack           =   16119285
   End
   Begin lvButton.lvButtons_H cmdExport 
      Height          =   450
      Left            =   45
      TabIndex        =   15
      Top             =   2205
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
      Image           =   "frmRptLabel.frx":A2D0
      cBack           =   16119285
   End
   Begin MSComCtl2.DTPicker DTAkhir 
      Height          =   330
      Left            =   7695
      TabIndex        =   16
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
   Begin lvButton.lvButtons_H cmdNew 
      Height          =   405
      Left            =   45
      TabIndex        =   20
      Top             =   495
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
      Image           =   "frmRptLabel.frx":A66A
      cBack           =   16777215
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
      TabIndex        =   19
      Top             =   90
      Width           =   1185
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
      TabIndex        =   18
      Top             =   90
      Width           =   600
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
      TabIndex        =   17
      Top             =   90
      Width           =   510
   End
End
Attribute VB_Name = "frmRptLabel"
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

Private Sub cmdNew_Click()
    CommandPass "New"
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



Private Sub DTAkhir_Click()
    Call LoadData
End Sub

Private Sub DTAkhir_CloseUp()
    Call LoadData
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
For i = 1 To 45
    cboMesin.AddItem i
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
    .Fields = "id,period_shift,date,shift,product_status,qty,label_product,box_number,plant_mark,machine_no,machine_name,customer_name,internal_part_id,product_name, " & _
                "nik,user_name,status "
    .Tables = "(select a.id,concat(qc_label_product_id,'-',box_number) as label_product,a.box_number,a.plant_mark,b.number as machine_no, b.name as machine_name, " & _
                "c.name as customer_name, d.internal_part_id, d.name as product_name,a.product_status, " & _
                "a.date,a.period_shift,a.shift,a.qty,a.created_by,e.nik, e.name as user_name,a.status from prod_result_logs a " & _
                "left join prod_machines b on a.prod_machine_id = b.id " & _
                "left join mkt_customers c on a.mkt_customer_id = c.id " & _
                "left join eng_products d on a.eng_product_id = d.id " & _
                "left join hrd_employees e on a.created_by = e.id) as label_box"
    .wCondition = "period_shift between '" & Format(DTAwal.Value, "yyyy-mm-dd") & "' and '" & Format(DTAkhir.Value, "yyyy-mm-dd") & "' and machine_no = '" & cboMesin.Text & "'"
    .SortOrder = "period_shift,shift,date,label_product DESC"
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

Private Sub lvButtons_H1_Click()

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

    Case "New" 'Refresh
           With frmRptLabelAE
                .State = AddStateMode
                .Show vbModal
           End With
    Case "Update"
    
           If lvList.ListItems.Count < 1 Then
            MsgBox "Tidak ada data yang di pilih!", vbExclamation
            Exit Sub
            End If
            
            If srcRecord = vbNullString Then
                MsgBox "Tidak ada data yang di pilih!", vbExclamation
                Exit Sub
            End If

                If MsgBox("Apakah Anda yakin ingin menghapus data ini : " & lvList.SelectedItem.SubItems(3) & "-" & lvList.SelectedItem.SubItems(4), vbCritical + vbYesNo) = vbYes Then
                    Dim sSQL As String
                    
                    sSQL = "update prod_result_logs a set a.`status` = 'suspend' where a.id = '" & lvList.SelectedItem.SubItems(15) & "'"
                    sSQL_Update sSQL
                    MsgBox "Record yang dipilih berhasil dihapus!", vbInformation
                    RefreshRecords
                Else
                    Exit Sub
                End If
                
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
'RS_PRODUCT.AbsolutePosition = RecordPage.PageStart
RecordPage.PageEnd
              
With lvList
    .GridLines = True
    .View = lvwReport

    .ColumnHeaders.Clear
    .ColumnHeaders.Add , , "No."
    .ColumnHeaders.Add , , "period_shift"
    .ColumnHeaders.Add , , "date"
    .ColumnHeaders.Add , , "shift"
    .ColumnHeaders.Add , , "product_status"
    .ColumnHeaders.Add , , "qty"
    .ColumnHeaders.Add , , "label_product"
    .ColumnHeaders.Add , , "box_number"
    .ColumnHeaders.Add , , "machine_no"
    .ColumnHeaders.Add , , "machine_name"
    .ColumnHeaders.Add , , "customer_name"
    .ColumnHeaders.Add , , "internal_part_id"
    .ColumnHeaders.Add , , "product_name"
    .ColumnHeaders.Add , , "nik"
    .ColumnHeaders.Add , , "user_name"
    .ColumnHeaders.Add , , "status"
    .ColumnHeaders.Add , , "id", 0
    

    .ListItems.Clear
    Do While Not RS_PRODUCT.EOF
    Set srcItem = .ListItems.Add(, , i, 1, 1)
        srcItem.SubItems(1) = Format(RS_PRODUCT.Fields("period_shift"), "yyyy-mm-dd")
        srcItem.SubItems(2) = Format(RS_PRODUCT.Fields("date"), "yyyy-mm-dd hh:mm:ss")
        srcItem.SubItems(3) = RS_PRODUCT.Fields("shift")
        srcItem.SubItems(4) = RS_PRODUCT.Fields("product_status")
        srcItem.SubItems(5) = RS_PRODUCT.Fields("qty")
        srcItem.SubItems(6) = RS_PRODUCT.Fields("label_product")
        srcItem.SubItems(7) = RS_PRODUCT.Fields("box_number")
        srcItem.SubItems(8) = RS_PRODUCT.Fields("machine_no")
        srcItem.SubItems(9) = RS_PRODUCT.Fields("machine_name")
        srcItem.SubItems(10) = RS_PRODUCT.Fields("customer_name")
        srcItem.SubItems(11) = RS_PRODUCT.Fields("internal_part_id")
        srcItem.SubItems(12) = RS_PRODUCT.Fields("product_name")
        srcItem.SubItems(13) = RS_PRODUCT.Fields("nik")
        srcItem.SubItems(14) = RS_PRODUCT.Fields("user_name")
        srcItem.SubItems(15) = RS_PRODUCT.Fields("status")
        srcItem.SubItems(16) = RS_PRODUCT.Fields("id")

        
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

