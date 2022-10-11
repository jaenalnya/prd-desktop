VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmUser 
   Caption         =   "Data User"
   ClientHeight    =   5820
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9300
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5820
   ScaleWidth      =   9300
   Begin VB.PictureBox picFooter 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   405
      Left            =   0
      ScaleHeight     =   405
      ScaleWidth      =   9300
      TabIndex        =   0
      Top             =   5415
      Width           =   9300
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
      Begin HRD.Liner Liner1 
         Height          =   30
         Left            =   0
         TabIndex        =   7
         Top             =   0
         Width           =   7170
         _ExtentX        =   12647
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
      Height          =   5175
      Left            =   1440
      TabIndex        =   9
      Top             =   90
      Width           =   7560
      _ExtentX        =   13335
      _ExtentY        =   9128
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
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
      Top             =   2070
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
      Image           =   "frmUser.frx":0000
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdUpdate 
      Height          =   405
      Left            =   45
      TabIndex        =   11
      Top             =   585
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   714
      Caption         =   "&Ubah [F3]"
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
      Image           =   "frmUser.frx":077A
      cBack           =   16119285
   End
   Begin lvButton.lvButtons_H cmdNew 
      Height          =   405
      Left            =   45
      TabIndex        =   12
      Top             =   90
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   714
      Caption         =   "&Baru [F2]"
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
      Image           =   "frmUser.frx":0909
      cBack           =   16119285
   End
   Begin lvButton.lvButtons_H cmdDelete 
      Height          =   405
      Left            =   45
      TabIndex        =   13
      Top             =   1080
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   714
      Caption         =   "&Hapus [F4]"
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
      Image           =   "frmUser.frx":0A63
      cBack           =   16119285
   End
   Begin lvButton.lvButtons_H cmdRefresh 
      Height          =   405
      Left            =   45
      TabIndex        =   14
      Top             =   1575
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
      Image           =   "frmUser.frx":3D00
      cBack           =   16119285
   End
   Begin lvButton.lvButtons_H cmdClose 
      Height          =   405
      Left            =   45
      TabIndex        =   15
      Top             =   3555
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
      Image           =   "frmUser.frx":447A
      cBack           =   16119285
   End
   Begin lvButton.lvButtons_H cmdExport 
      Height          =   405
      Left            =   45
      TabIndex        =   16
      Top             =   3060
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
      Image           =   "frmUser.frx":76DC
      cBack           =   16119285
   End
   Begin lvButton.lvButtons_H CmdPrint 
      Height          =   405
      Left            =   45
      TabIndex        =   17
      Top             =   2565
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   714
      Caption         =   "&Print [F7]"
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
      Image           =   "frmUser.frx":7A76
      cBack           =   16119285
   End
End
Attribute VB_Name = "frmUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim srcItem                        As ListItem
Dim srcRecord                      As String
Dim srcUser                        As Variant
Dim srcSQL                         As String
Dim SQLParser                      As New clsSQLSelectParser

Private Sub cmdNew_Click()
On Error Resume Next
    CommandPass "New"
End Sub

Private Sub CmdPrint_Click()
On Error Resume Next
    CommandPass "Print"
End Sub

Private Sub CmdSearch_Click()
On Error Resume Next
    CommandPass "Search"
End Sub

Private Sub cmdUpdate_Click()
On Error Resume Next
    CommandPass "Update"
End Sub

Private Sub CmdDelete_Click()
On Error Resume Next
    CommandPass "Delete"
End Sub

Private Sub cmdRefresh_Click()
On Error Resume Next
    CommandPass "Refresh"
End Sub

Private Sub cmdExport_Click()
On Error Resume Next
    CommandPass "Export"
End Sub

Private Sub cmdClose_Click()
On Error Resume Next
    CommandPass "Close"
End Sub


Private Sub Form_Activate()
On Error Resume Next

With MAIN
    Me.BackColor = .ACPMenu.BackColor
    Me.Picture = .ACPMenu.LoadBackground
    picFooter.BackColor = .ACPMenu.BackColor
    Picture2.BackColor = .ACPMenu.BackColor
End With

MAIN.ActivateChild Me
lvList.FlatScrollBar = False

End Sub

Private Sub Form_Load()
On Error GoTo ErrHandler
Dim sSQL As String
ButtonList lvList, btnFirst, btnPrev, btnNext, btnLast

With SQLParser
    .Fields = "*"
    .Tables = "tblUsers"
    .SortOrder = "KodeUser ASC"
    .SaveStatement
End With

Set RS_USER = New ADODB.Recordset
If RS_USER.State = adStateOpen Then RS_USER.Close
RS_USER.Open SQLParser.SQLStatement, CN, adOpenDynamic, adLockPessimistic

With RecordPage
    .Start RS_USER, 100
End With

FillListview 1

srcUser = "NONE"
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
        lvList.Height = Me.ScaleHeight - 600
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    MAIN.RemoveChild Me.Name
    Set RS_USER = Nothing
    Set frmUser = Nothing
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
On Error Resume Next
CommandPass "Update"
End Sub

Private Sub LvList_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error Resume Next
srcUser = lvList.SelectedItem.Index
srcRecord = lvList.ListItems.Item(srcUser).Text
Call RefreshRecSum
End Sub


Public Sub CommandPass(ByVal srcPerformWhat As String)
On Error GoTo errPerformWhat
Select Case srcPerformWhat
    Case "New" 'New
            With frmUserAE
                .State = AddStateMode
                .Show vbModal
            End With
    Case "Update" 'Update
            If srcRecord = vbNullString Then
                MsgBox "Tidak ada data yang di pilih!", vbExclamation
                Exit Sub
            Else
                With frmUserAE
                    .State = EditStateMode
                    .PK = srcRecord
                    .Show vbModal
                End With
            End If
            
    Case "Delete" 'Delete
            If lvList.ListItems.Count < 1 Then
                MsgBox "Tidak ada data untuk di hapus!", vbExclamation
                Exit Sub
            End If
            
            If srcRecord = vbNullString Then
                MsgBox "Tidak ada data yang di pilih!", vbExclamation
                Exit Sub
            End If
            If CStr(lvList.SelectedItem.Text) = ACTIVE_USER.KODEUSER Then
                MsgBox "Data tidak bisa di hapus, karena masih  dipergunakan.", vbExclamation
                Exit Sub
            Else
                If MsgBox("Apakah Anda yakin ingin menghapus data ini?", vbCritical + vbYesNo) = vbYes Then
                    sSQL_Delete "DELETE FROM tblUsers WHERE KodeUser='" & srcRecord & "'"
                    MsgBox "Record yang dipilih berhasil dihapus!", vbInformation, Me.Caption
                    RefreshRecords
                Else
                    Exit Sub
                End If
            End If
            
    Case "Refresh" 'Refresh
           RefreshRecords
           
    Case "Export" 'Preview
            With lvList
                If .ListItems.Count = 0 Then
                    MsgBox "There's no records to export!Please check it.", vbExclamation
                    Exit Sub
                End If
            End With
                         
            XLSFILENAME = ""
            
            With MAIN.CDExporter
                .Filter = "Excel Files(*.xls)|*.xls|Excel 2007 Files(*.xlsx)|*.xlsx"
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
    Case "Print"
            Dim sSQL As String
            sSQL = "SELECT tblUsers.* FROM tblUsers ORDER BY KodeUser ASC"
            
            Set RS_PRINT = New ADODB.Recordset
            If RS_PRINT.State = adStateOpen Then RS_PRINT.Close
            RS_PRINT.Open sSQL, CN, adOpenDynamic, adLockPessimistic
            With rptAllUser
                .DTRpt.Recordset = RS_PRINT
                .lblDate.Caption = Now
                .lblCompany.Caption = ACTIVE_COMPANY.Perusahaan
                .lblAlamat.Caption = ACTIVE_COMPANY.Alamat
                
                .txtKodeUser.DataField = "KodeUser"
                .txtNamaAwal.DataField = "NamaAwal"
                .txtNamaAkhir.DataField = "NamaAkhir"
                .txtLoginName.DataField = "Username"
                .txtStatus.DataField = "StatusCD"
                .txtKeterangan.DataField = "Keterangan"
                
                .Show
            End With

    Case "Search" 'Search
            With frmSearch
                Set .srcForm = Me
                Set .srcColumnHeaders = lvList.ColumnHeaders
                .Show vbModal
            End With
            
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
RecordPage.CurrentPosition = whichPage
RS_USER.AbsolutePosition = RecordPage.PageStart
RecordPage.PageEnd
With lvList
    .View = lvwReport
    .GridLines = False
    
    .ColumnHeaders.Clear
    .ColumnHeaders.Add , , "KodeUser"
    .ColumnHeaders.Add , , "Nama Awal"
    .ColumnHeaders.Add , , "Nama Akhir"
    .ColumnHeaders.Add , , "User Name"
    .ColumnHeaders.Add , , "IsAdmin"
    .ColumnHeaders.Add , , "StatusCD"
    .ColumnHeaders.Add , , "Keterangan"
    
    .ListItems.Clear
    Do While Not RS_USER.EOF
    Set srcItem = .ListItems.Add(, , RS_USER.Fields("KodeUser"), 1, 1)
        srcItem.SubItems(1) = RS_USER.Fields("NamaAwal")
        srcItem.SubItems(2) = RS_USER.Fields("NamaAkhir")
        srcItem.SubItems(3) = RS_USER.Fields("UserName")
        srcItem.SubItems(4) = RS_USER.Fields("IsAdmin")
        srcItem.SubItems(5) = RS_USER.Fields("StatusCD")
        srcItem.SubItems(6) = RS_USER.Fields("Keterangan")
        
        If RS_USER.AbsolutePosition >= RecordPage.PageEnd Then
            Exit Do
        Else
            RS_USER.MoveNext
        End If
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
    lblRecSum.Caption = "Record: " & srcUser & " of " & lvList.ListItems.Count
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
On Error Resume Next
    SQLParser.RestoreStatement
    SQLParser.wCondition = srcCondition
    ReloadRecords SQLParser.SQLStatement
End Sub

'Procedure for reloadingrecords
Public Sub ReloadRecords(ByVal srcSQL As String)
On Error Resume Next
    
    Set RS_USER = New ADODB.Recordset
    If RS_USER.State = adStateOpen Then RS_USER.Close
    RS_USER.Open srcSQL, CN, adOpenDynamic, adLockPessimistic


With RecordPage
    .Start RS_USER, 100
End With

FillListview 1

End Sub
Public Sub RefreshRecords()
On Error Resume Next
    SQLParser.RestoreStatement
    ReloadRecords SQLParser.SQLStatement
End Sub

Private Sub picFooter_Resize()
On Error Resume Next
Picture2.Left = picFooter.ScaleWidth - Picture2.ScaleWidth
End Sub

Private Sub lvList_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then PopupMenu MAIN.mnuAction
End Sub
'Pasti dicopy
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
    Select Case KeyCode
    Case vbKeyF2
        CommandPass "New"
    Case vbKeyF3
        CommandPass "Update"
    Case vbKeyF4
        CommandPass "Delete"
    Case vbKeyF5
        CommandPass "Refresh"
    Case vbKeyF6
        CommandPass "Search"
    Case vbKeyF7
        CommandPass "Print"
    Case vbKeyF8
        CommandPass "Export"
    Case vbKeyEscape
        CommandPass "Close"
    End Select
End Sub
