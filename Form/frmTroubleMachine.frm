VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "LvButton.ocx"
Begin VB.Form frmTroubleMachine 
   Caption         =   "Data Trouble Machine"
   ClientHeight    =   5190
   ClientLeft      =   3345
   ClientTop       =   2070
   ClientWidth     =   9570
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5190
   ScaleWidth      =   9570
   Begin VB.PictureBox picFooter 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   405
      Left            =   0
      ScaleHeight     =   405
      ScaleWidth      =   9570
      TabIndex        =   1
      Top             =   4785
      Width           =   9570
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   345
         Left            =   5175
         ScaleHeight     =   345
         ScaleWidth      =   4155
         TabIndex        =   2
         Top             =   45
         Width           =   4150
         Begin VB.CommandButton btnFirst 
            Height          =   315
            Left            =   2760
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "First 250"
            Top             =   10
            Width           =   315
         End
         Begin VB.CommandButton btnPrev 
            Height          =   315
            Left            =   3075
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Previous 250"
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
         Begin VB.CommandButton btnNext 
            Height          =   315
            Left            =   3390
            Style           =   1  'Graphical
            TabIndex        =   3
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
            TabIndex        =   7
            Top             =   60
            Width           =   2535
         End
      End
      Begin PRD.Liner Liner1 
         Height          =   30
         Left            =   0
         TabIndex        =   8
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
         TabIndex        =   9
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
      ItemData        =   "frmTroubleMachine.frx":0000
      Left            =   2745
      List            =   "frmTroubleMachine.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   45
      Width           =   1365
   End
   Begin MSComctlLib.ListView lvList 
      Height          =   3960
      Left            =   1440
      TabIndex        =   10
      Top             =   540
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
      Left            =   45
      TabIndex        =   11
      Top             =   1800
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
      Image           =   "frmTroubleMachine.frx":0004
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdRefresh 
      Height          =   405
      Left            =   45
      TabIndex        =   12
      Top             =   1305
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
      Image           =   "frmTroubleMachine.frx":077E
      cBack           =   16119285
   End
   Begin lvButton.lvButtons_H cmdClose 
      Height          =   405
      Left            =   45
      TabIndex        =   13
      Top             =   2745
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
      Image           =   "frmTroubleMachine.frx":0EF8
      cBack           =   16119285
   End
   Begin lvButton.lvButtons_H cmdExport 
      Height          =   405
      Left            =   45
      TabIndex        =   14
      Top             =   2250
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
      Image           =   "frmTroubleMachine.frx":415A
      cBack           =   16119285
   End
   Begin lvButton.lvButtons_H cmdMasuk 
      Height          =   675
      Left            =   45
      TabIndex        =   15
      Top             =   540
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   1191
      Caption         =   "Input Trouble"
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
      Image           =   "frmTroubleMachine.frx":44F4
      cBack           =   16777215
   End
   Begin MSComCtl2.DTPicker DTAwal 
      Height          =   330
      Left            =   5355
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
      Format          =   87293955
      CurrentDate     =   43642
   End
   Begin MSComCtl2.DTPicker DTAkhir 
      Height          =   330
      Left            =   7515
      TabIndex        =   17
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
      Format          =   87293955
      CurrentDate     =   43642
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
      TabIndex        =   20
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
      Left            =   4725
      TabIndex        =   19
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
      Left            =   7110
      TabIndex        =   18
      Top             =   90
      Width           =   510
   End
End
Attribute VB_Name = "frmTroubleMachine"
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
With SQLParser
    .Fields = "*"
    .Tables = "( select a.id,a.plant_mark, a.product_id, c.internal_part_id,c.name as product_name," & _
                "a.prod_machine_id, d.number as machine_no, d.name as machine_name, " & _
                "a.period_shift, a.shift, " & _
                "a.trouble,a.analysis,a.`action`,a.result,a.description, " & _
                "a.created_at,a.created_by,e.name as employee_name " & _
                "from prod_trouble_machines a " & _
                "INNER JOIN eng_products c on a.product_id = c.id " & _
                "INNER JOIN prod_machines d on a.prod_machine_id = d.id " & _
                "INNER JOIN sys_accounts e on a.created_by = e.id) as XX"


        If cboMesin.text = "All" Then
            .wCondition = "period_shift between '" & Format(DTAwal.Value, "yyyy-mm-dd") & "' and '" & Format(DTAkhir.Value, "yyyy-mm-dd") & "'  and plant_mark = '" & p_plant_mark & "'"
        Else
            .wCondition = "period_shift between '" & Format(DTAwal.Value, "yyyy-mm-dd") & "' and '" & Format(DTAkhir.Value, "yyyy-mm-dd") & "' and machine_no = '" & cboMesin.text & "'  and plant_mark = '" & p_plant_mark & "'"
        End If
        
    .SortOrder = "period_shift"
    .SaveStatement
End With

Set RS_TROUBLE = New ADODB.Recordset
If RS_TROUBLE.State = adStateOpen Then RS_TROUBLE.Close
RS_TROUBLE.Open SQLParser.SQLStatement, CN, adOpenDynamic, adLockPessimistic


With RecordPage
    .Start RS_TROUBLE, 100
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
'AltLVBackground lvList, vbWhite, &H8000000F, frmTroubleMachine
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


With SQLParser
    .Fields = "*"
    .Tables = "( select a.id,a.plant_mark, a.product_id, c.internal_part_id,c.name as product_name," & _
                "a.prod_machine_id, d.number as machine_no, d.name as machine_name, " & _
                "a.period_shift, a.shift, " & _
                "a.trouble,a.analysis,a.`action`,a.result,a.description, " & _
                "a.created_at,a.created_by,e.name as employee_name " & _
                "from prod_trouble_machines a " & _
                "INNER JOIN eng_products c on a.product_id = c.id " & _
                "INNER JOIN prod_machines d on a.prod_machine_id = d.id " & _
                "INNER JOIN sys_accounts e on a.created_by = e.id) as XX"


        If cboMesin.text = "All" Then
            .wCondition = "period_shift between '" & Format(DTAwal.Value, "yyyy-mm-dd") & "' and '" & Format(DTAkhir.Value, "yyyy-mm-dd") & "'  and plant_mark = '" & p_plant_mark & "'  "
        Else
            .wCondition = "period_shift between '" & Format(DTAwal.Value, "yyyy-mm-dd") & "' and '" & Format(DTAkhir.Value, "yyyy-mm-dd") & "' and machine_no = '" & cboMesin.text & "'  and plant_mark = '" & p_plant_mark & "'"
        End If
        
    .SortOrder = "period_shift"
    .SaveStatement
End With

Set RS_TROUBLE = New ADODB.Recordset
If RS_TROUBLE.State = adStateOpen Then RS_TROUBLE.Close
RS_TROUBLE.Open SQLParser.SQLStatement, CN, adOpenDynamic, adLockPessimistic


With RecordPage
    .Start RS_TROUBLE, 100
End With

FillListview 1

srcMonitoring = "NONE"
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
        lvList.Height = Me.ScaleHeight - 1000
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
    MAIN.RemoveChild Me.Name
    Set frmTroubleMachine = Nothing
    Set RS_TROUBLE = Nothing
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
    srcRecord = lvList.ListItems.Item(srcMonitoring).SubItems(11)
    Call RefreshRecSum
End Sub


Public Sub CommandPass(ByVal srcPerformWhat As String)
On Error GoTo errPerformWhat
Select Case srcPerformWhat
 
    Case "Masuk" 'Refresh
           With frmTroubleMachineAE
                .State = AddStateMode
                .Show
           End With
            

    Case "Update" 'Update
            If srcRecord = vbNullString Then
                MsgBox "Tidak ada data yang di pilih!", vbExclamation
                Exit Sub
            Else
                With frmTroubleMachineAE
                    .State = EditStateMode
                    .PK = srcRecord
                    .Show
                End With
            End If
            
    Case "Refresh" 'Refresh
           RefreshRecords
    
    Case "Search" 'Search
            With frmSearch
                Set .srcForm = Me
                Set .srcColumnHeaders = lvList.ColumnHeaders
                .Show
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
'RS_TROUBLE.AbsolutePosition = RecordPage.PageStart
RecordPage.PageEnd
With lvList
    .GridLines = True
    .View = lvwReport

    .ColumnHeaders.Clear
    .ColumnHeaders.Add , , "No."
    .ColumnHeaders.Add , , "Period_shift"
    .ColumnHeaders.Add , , "Machine_no"
    .ColumnHeaders.Add , , "Shift"
    .ColumnHeaders.Add , , "trouble"
    .ColumnHeaders.Add , , "analysis"
    .ColumnHeaders.Add , , "Action"
    .ColumnHeaders.Add , , "Result"
    .ColumnHeaders.Add , , "description"
    .ColumnHeaders.Add , , "Employee_name"
    .ColumnHeaders.Add , , "product_name"
    .ColumnHeaders.Add , , "id"


    .ListItems.Clear
    Do While Not RS_TROUBLE.EOF
    Set srcItem = .ListItems.Add(, , i, 1, 1)
        
        srcItem.SubItems(1) = RS_TROUBLE.Fields("Period_shift")
        srcItem.SubItems(2) = RS_TROUBLE.Fields("machine_no")
        srcItem.SubItems(3) = RS_TROUBLE.Fields("Shift")
        srcItem.SubItems(4) = RS_TROUBLE.Fields("trouble")
        srcItem.SubItems(5) = RS_TROUBLE.Fields("analysis")
        srcItem.SubItems(6) = RS_TROUBLE.Fields("Action")
        srcItem.SubItems(7) = RS_TROUBLE.Fields("Result")
        srcItem.SubItems(8) = RS_TROUBLE.Fields("description")
        srcItem.SubItems(9) = RS_TROUBLE.Fields("employee_name")
        srcItem.SubItems(10) = RS_TROUBLE.Fields("product_name")
        srcItem.SubItems(11) = RS_TROUBLE.Fields("id")


        If RS_TROUBLE.AbsolutePosition >= RecordPage.PageEnd Then
            Exit Do
        Else
            RS_TROUBLE.MoveNext
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
    
    Set RS_TROUBLE = New ADODB.Recordset
    If RS_TROUBLE.State = adStateOpen Then RS_TROUBLE.Close
    RS_TROUBLE.Open srcSQL, CN, adOpenDynamic, adLockPessimistic

With RecordPage
    .Start RS_TROUBLE, 100
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







