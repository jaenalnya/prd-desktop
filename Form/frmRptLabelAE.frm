VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmRptLabelAE 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Adjustment"
   ClientHeight    =   7785
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7920
   Icon            =   "frmRptLabelAE.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7785
   ScaleWidth      =   7920
   StartUpPosition =   2  'CenterScreen
   Begin PRD.Liner Liner1 
      Height          =   30
      Left            =   90
      TabIndex        =   30
      Top             =   7020
      Width           =   7755
      _ExtentX        =   13679
      _ExtentY        =   53
   End
   Begin VB.TextBox txtBoxbarcode 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2295
      TabIndex        =   28
      Top             =   6480
      Width           =   4830
   End
   Begin VB.Frame Frame1 
      Caption         =   "Input Barcode/ID Label"
      Height          =   2490
      Left            =   270
      TabIndex        =   18
      Top             =   3870
      Width           =   7440
      Begin VB.TextBox txtQty 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5265
         Locked          =   -1  'True
         TabIndex        =   23
         Top             =   1575
         Width           =   1590
      End
      Begin VB.TextBox txtBoxNumber 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2025
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   1575
         Width           =   1590
      End
      Begin VB.TextBox txtBarcode 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2025
         TabIndex        =   19
         Top             =   270
         Width           =   4830
      End
      Begin VB.Label lblBarcode 
         BackStyle       =   0  'Transparent
         Caption         =   "ID"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2025
         TabIndex        =   27
         Top             =   945
         Width           =   4830
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "*Enter untuk check label"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   255
         Index           =   3
         Left            =   2025
         TabIndex        =   26
         Top             =   675
         Width           =   3930
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "* Informasi label harus sama dengan machine, customer dan product"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   255
         Index           =   1
         Left            =   135
         TabIndex        =   25
         Top             =   2115
         Width           =   7215
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Qty"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4455
         TabIndex        =   24
         Top             =   1620
         Width           =   870
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Box Number"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   135
         TabIndex        =   22
         Top             =   1620
         Width           =   1905
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "ID Barcode/Label"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   135
         TabIndex        =   20
         Top             =   315
         Width           =   1905
      End
   End
   Begin VB.ComboBox cboProduct 
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
      ItemData        =   "frmRptLabelAE.frx":15162
      Left            =   2250
      List            =   "frmRptLabelAE.frx":15164
      TabIndex        =   14
      Text            =   "cboProduct"
      Top             =   1980
      Width           =   5460
   End
   Begin VB.ComboBox cboCustomer 
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
      ItemData        =   "frmRptLabelAE.frx":15166
      Left            =   2250
      List            =   "frmRptLabelAE.frx":15170
      TabIndex        =   13
      Text            =   "cboCustomer"
      Top             =   1530
      Width           =   5460
   End
   Begin VB.ComboBox cboMachine 
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
      ItemData        =   "frmRptLabelAE.frx":15184
      Left            =   2250
      List            =   "frmRptLabelAE.frx":15186
      TabIndex        =   12
      Text            =   "cboMachine"
      Top             =   1080
      Width           =   1635
   End
   Begin VB.ComboBox cboshift 
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
      ItemData        =   "frmRptLabelAE.frx":15188
      Left            =   2250
      List            =   "frmRptLabelAE.frx":15195
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   3375
      Width           =   1635
   End
   Begin lvButton.lvButtons_H cmdSave 
      Height          =   420
      Left            =   4500
      TabIndex        =   9
      Top             =   7245
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   741
      Caption         =   "Save [F5]"
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
      Image           =   "frmRptLabelAE.frx":151A2
      cBack           =   -2147483633
   End
   Begin MSComCtl2.DTPicker DTPeriode 
      Height          =   330
      Left            =   2250
      TabIndex        =   8
      Top             =   2925
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   582
      _Version        =   393216
      CustomFormat    =   "dd/MMM/yyyy"
      Format          =   85327875
      CurrentDate     =   43998
   End
   Begin VB.PictureBox picHeader 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   945
      Left            =   0
      ScaleHeight     =   945
      ScaleWidth      =   7920
      TabIndex        =   5
      Top             =   0
      Width           =   7920
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Input Label Barang"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   870
         TabIndex        =   7
         Top             =   150
         Width           =   5355
      End
      Begin VB.Image Image1 
         Height          =   720
         Left            =   0
         Picture         =   "frmRptLabelAE.frx":2C23C
         Top             =   90
         Width           =   720
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Silahkan Masukan Data-Data Secara Benar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   255
         Index           =   2
         Left            =   870
         TabIndex        =   6
         Top             =   480
         Width           =   5100
      End
   End
   Begin lvButton.lvButtons_H cmdCancel 
      Height          =   420
      Left            =   6210
      TabIndex        =   10
      Top             =   7245
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   741
      Caption         =   "Close [Esc]"
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
      Image           =   "frmRptLabelAE.frx":323B6
      cBack           =   -2147483633
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Box Barcode"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   405
      TabIndex        =   29
      Top             =   6525
      Width           =   1905
   End
   Begin VB.Label lblInternalPart 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   2295
      TabIndex        =   17
      Top             =   2475
      Width           =   5370
   End
   Begin VB.Label lblMachine_name 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   4095
      TabIndex        =   16
      Top             =   1125
      Width           =   2805
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "*Tanggal shfit"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   4140
      TabIndex        =   15
      Top             =   2970
      Width           =   2805
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Product"
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
      Left            =   270
      TabIndex        =   4
      Top             =   2025
      Width           =   1770
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Machine"
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
      Left            =   270
      TabIndex        =   3
      Top             =   1125
      Width           =   1905
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Customer name"
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
      Index           =   0
      Left            =   270
      TabIndex        =   2
      Top             =   1575
      Width           =   1770
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Shift"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   270
      TabIndex        =   1
      Top             =   3420
      Width           =   1770
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Periode Shift"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   270
      TabIndex        =   0
      Top             =   2970
      Width           =   1905
   End
End
Attribute VB_Name = "frmRptLabelAE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public State                        As FORM_STATE
Public PK                           As String
Dim sSQL                            As String
Dim Id_prod_machine                 As String
Dim Id_customer                     As String
Dim Id_product                      As String



Private Sub cboMachine_Click()
On Error GoTo ErrHandler
    Dim Rs As New ADODB.Recordset
    Rs.CursorLocation = adUseClient
    
    lblMachine_name.Caption = ""
    Id_prod_machine = ""
    
    Rs.Open "Select * From prod_machines Where sys_plant_id = '" & p_sys_plant & "' and number = '" & cboMachine.Text & "'", CN, adOpenStatic, adLockOptimistic
    If Rs.RecordCount > 0 Then
        Id_prod_machine = Rs!id
        lblMachine_name.Caption = Rs!Name
    End If
    Set Rs = Nothing

Exit Sub
ErrHandler:
    MsgBox "Error Number: " & Err.Number & vbCrLf & "Description: " & Err.Description, vbExclamation
End Sub

Private Sub cboMachine_KeyPress(KeyAscii As Integer)
    KeyAscii = AutoMatchCBBox(cboMachine, KeyAscii)
End Sub

Private Sub cboCustomer_Click()
On Error GoTo ErrHandler
    Dim Rs As New ADODB.Recordset
    Rs.CursorLocation = adUseClient
    Id_customer = ""
    
    Rs.Open "select a.id,a.name as customer_name from mkt_customers a where a.name = '" & cboCustomer.Text & "'", CN, adOpenStatic, adLockOptimistic
    If Rs.RecordCount > 0 Then
        Id_customer = Rs!id
        
        Call ComboProduct
        
    End If
    Set Rs = Nothing
    
Exit Sub
ErrHandler:
    MsgBox "Error Number: " & Err.Number & vbCrLf & "Description: " & Err.Description, vbExclamation
End Sub

Private Sub cboCustomer_KeyPress(KeyAscii As Integer)
    KeyAscii = AutoMatchCBBox(cboCustomer, KeyAscii)
End Sub


Private Sub cboProduct_Click()
On Error GoTo ErrHandler
    Dim Rs As New ADODB.Recordset
    Rs.CursorLocation = adUseClient
    Dim output() As String
    output() = Split(cboProduct.Text, "-")
    
    lblInternalPart.Caption = ""
    Id_product = ""
    
    Rs.Open "select a.id,a.internal_part_id,a.name as product_name from eng_products a where a.internal_part_id = '" & output(0) & "'", CN, adOpenStatic, adLockOptimistic
    If Rs.RecordCount > 0 Then
        Id_product = Rs!id
        lblInternalPart.Caption = Rs!internal_part_id
    End If
    Set Rs = Nothing
Exit Sub
ErrHandler:
    MsgBox "Error Number: " & Err.Number & vbCrLf & "Description: " & Err.Description, vbExclamation
End Sub

Private Sub cboProduct_KeyPress(KeyAscii As Integer)
    KeyAscii = AutoMatchCBBox(cboProduct, KeyAscii)
    cboProduct_Click
End Sub


Private Sub cmdCancel_Click()
  Unload Me
End Sub


Private Sub cmdSave_Click()
On Error GoTo ErrHandler
    If cboMachine.Text = "" Then
        MsgBox "Machine harus di isi, silahkan cek kembali!", vbExclamation
        Exit Sub
    ElseIf cboCustomer.Text = "" Then
        MsgBox "Customer harus di isi, silahkan cek kembali!", vbExclamation
        Exit Sub
    ElseIf cboProduct.Text = "" Then
        MsgBox "Product harus di isi, silahkan cek kembali!", vbExclamation
        Exit Sub
    ElseIf cboshift.Text = "" Then
        MsgBox "Shift harus di isi, silahkan cek kembali!", vbExclamation
        Exit Sub
    ElseIf lblBarcode.Caption = "" Then
        MsgBox "Barcode harus di isi, silahkan cek kembali!", vbExclamation
        Exit Sub
    ElseIf txtQty.Text = "" Then
        MsgBox "Qty harus di isi, silahkan cek kembali!", vbExclamation
        Exit Sub
    End If
    
    If State = AddStateMode Then

        sSQL = "Insert Into prod_result_logs (plant_mark"
        sSQL = sSQL & " ,prod_machine_id"
        sSQL = sSQL & " ,mkt_customer_id"
        sSQL = sSQL & " ,eng_product_id"
        sSQL = sSQL & " ,date"
        sSQL = sSQL & " ,period_shift"
        sSQL = sSQL & " ,shift"
        sSQL = sSQL & " ,product_status"
        sSQL = sSQL & " ,qty"
        sSQL = sSQL & " ,qc_label_product_id"
        sSQL = sSQL & " ,box_number"
        sSQL = sSQL & " ,box_barcode"
        sSQL = sSQL & " ,created_at"
        sSQL = sSQL & " ,created_by)"
        sSQL = sSQL & " values ('" & p_plant_mark & "'"
        sSQL = sSQL & " ,'" & Id_prod_machine & "'"
        sSQL = sSQL & " ,'" & Id_customer & "'"
        sSQL = sSQL & " ,'" & Id_product & "'"
        sSQL = sSQL & " ,'" & Format(DTPeriode, "yyyy-mm-dd hh:mm:ss") & "'"
        sSQL = sSQL & " ,'" & Format(DTPeriode, "yyyy-mm-dd") & "'"
        sSQL = sSQL & " ,'" & cboshift.Text & "'"
        sSQL = sSQL & " ,'ok'"
        sSQL = sSQL & " ,'" & txtQty.Text & "'"
        sSQL = sSQL & " ,'" & lblBarcode.Caption & "'"
        sSQL = sSQL & " ,'" & txtBoxNumber.Text & "'"
        sSQL = sSQL & " ,'" & txtBoxbarcode.Text & "'"
        sSQL = sSQL & " ,'" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "'"
        sSQL = sSQL & " ,'" & ACTIVE_USER.SYSID & "')"
        
        sSQL_Insert sSQL
        
        MsgBox "Data baru berhasil disimpan!", vbInformation
        
        Unload Me

    End If
Exit Sub
ErrHandler:
    MsgBox "Error Number: " & Err.Number & vbCrLf & " Description : " & Err.Description, vbExclamation, Me.Caption

End Sub

Private Sub Form_Activate()
On Error Resume Next
    Me.BackColor = MAIN.ACPMenu.BackColor
    Frame1.BackColor = MAIN.ACPMenu.BackColor
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If KeyAscii = 27 Then
        Unload Me
    ElseIf KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub Form_Load()
On Error GoTo ErrHandler
CenterForm frmRptLabelAE

Call ComboMachine
Call ComboCustomer

DTPeriode.Value = Now

If State = AddStateMode Then
    Me.Caption = "Buat Baru"

    sSQL = "SELECT prod_result_logs.* FROM prod_result_logs limit 1"

    Set RS_PRODUCT = New ADODB.Recordset
    If RS_PRODUCT.State = adStateOpen Then RS_PRODUCT.Close
    RS_PRODUCT.Open sSQL, CN, adOpenDynamic, adLockOptimistic
    
End If

Exit Sub
ErrHandler:
    MsgBox "Error Number: " & Err.Number & vbCrLf & "Description: " & Err.Description, vbExclamation
End Sub

Private Sub Form_Unload(Cancel As Integer)
frmRptLabel.CommandPass "Refresh"
Set frmRptLabelAE = Nothing
Set RS_PRODUCT = Nothing
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
    Select Case KeyCode
    Case vbKeyF5
        cmdSave_Click
    Case vbKeyEscape
        cmdCancel_Click
    End Select
End Sub




Private Sub txtBarcode_KeyPress(KeyAscii As Integer)

On Error GoTo ErrHandler

Dim Rs As New ADODB.Recordset
Dim Rs_search As New ADODB.Recordset
Dim sSQL As String
Dim i As Integer
Dim barcode() As String

   
    If KeyAscii = 13# Then
    
        If cboMachine.Text = "" Then
            MsgBox "Machine harus di isi, silahkan cek kembali!", vbExclamation
            txtBarcode.Text = ""
            Exit Sub
        ElseIf cboCustomer.Text = "" Then
            MsgBox "Customer harus di isi, silahkan cek kembali!", vbExclamation
            txtBarcode.Text = ""
            Exit Sub
        ElseIf cboProduct.Text = "" Then
            MsgBox "Product harus di isi, silahkan cek kembali!", vbExclamation
            txtBarcode.Text = ""
            Exit Sub
        End If
    
        barcode = Split(txtBarcode.Text, "-")
        Rs_search.CursorLocation = adUseClient
        Rs_search.Open "select * from prod_result_logs a where a.qc_label_product_id = '" & barcode(0) & "' and " & _
                        "a.box_number = '" & barcode(1) & "'", CN, adOpenStatic, adLockOptimistic
        If Rs_search.RecordCount > 0 Then
            MsgBox "Data Sudah Pernah di Scan..!", vbExclamation
    
            txtBarcode.Text = ""
            txtBoxNumber.Text = ""
            txtQty.Text = ""
            txtBarcode.SetFocus
            Exit Sub
        Else
            Rs.CursorLocation = adUseClient
            sSQL = "select a.id,a.seq,a.cavity,a.engine_number,a.quantity_box,a.eng_product_id"
            sSQL = sSQL & " from qc_label_products a"
            sSQL = sSQL & " where a.id = '" & barcode(0) & "'"
            
            Rs.Open sSQL, CN, adOpenStatic, adLockOptimistic
            If Rs.RecordCount < 1 Then
                MsgBox "Data Tidak ditemukan", vbExclamation
    
                txtBarcode.Text = ""
                txtBoxNumber.Text = ""
                txtQty.Text = ""
                txtBarcode.SetFocus
                Exit Sub
            Else
            
            
                If Rs.Fields("engine_number") <> cboMachine.Text Then
                    MsgBox "Machine tidak sama dengan label, silahkan cek kembali!", vbExclamation
                    txtBarcode.Text = ""
                    Exit Sub
                ElseIf Rs.Fields("eng_product_id") <> Id_product Then
                    MsgBox "Product tidak sama dengan label, silahkan cek kembali!", vbExclamation
                    txtBarcode.Text = ""
                    Exit Sub
                Else
                    txtBoxNumber.Text = barcode(1)
                    txtQty.Text = Rs.Fields("quantity_box")
                    lblBarcode.Caption = barcode(0)
                    txtBarcode.Text = ""
                End If

            End If
        
            Set Rs = Nothing
            Set Rs_search = Nothing
        End If
        
    End If
    
Exit Sub
ErrHandler:
    MsgBox "Error Number: " & Err.Number & vbCrLf & "Description: " & Err.Description, vbExclamation, Me.Caption
     
   
End Sub


Private Sub txtQty_KeyPress(KeyAscii As Integer)
    KeyAscii = isNumber(KeyAscii)
End Sub

Public Sub ComboProduct()
    Dim Rs As New ADODB.Recordset
    Dim sSQL As String
    Rs.CursorLocation = adUseClient
    
    sSQL = "select a.id,a.internal_part_id,a.name as product_name from eng_products a where a.mkt_customer_id = '" & Id_customer & "'"
    Rs.Open sSQL, CN, adOpenStatic, adLockOptimistic
        
    If Rs.RecordCount > 0 Then
        cboProduct.Clear
        Rs.MoveFirst
        Do While Not Rs.EOF
            cboProduct.AddItem Rs.Fields("internal_part_id") & "-" & Rs.Fields("product_name")
            Rs.MoveNext
        Loop
    End If
    Set Rs = Nothing
End Sub


Public Sub ComboCustomer()
    Dim Rs As New ADODB.Recordset
    Dim sSQL As String
    Rs.CursorLocation = adUseClient
    
    sSQL = "select a.id,a.name as customer_name from mkt_customers a where a.`status` = 'active'"
    Rs.Open sSQL, CN, adOpenStatic, adLockOptimistic
        
    If Rs.RecordCount > 0 Then
        cboCustomer.Clear
        Rs.MoveFirst
        Do While Not Rs.EOF
            cboCustomer.AddItem Rs.Fields("customer_name")
            Rs.MoveNext
        Loop
    End If
    Set Rs = Nothing
End Sub


Public Sub ComboMachine()
    Dim Rs As New ADODB.Recordset
    Dim sSQL As String
    Rs.CursorLocation = adUseClient
    
    sSQL = "select a.id,a.sys_plant_id,a.number as machine_no,a.name as machine_name from prod_machines a where a.sys_plant_id = '" & p_sys_plant & "'"
    Rs.Open sSQL, CN, adOpenStatic, adLockOptimistic
        
    If Rs.RecordCount > 0 Then
        cboMachine.Clear
        Rs.MoveFirst
        Do While Not Rs.EOF
            cboMachine.AddItem Rs.Fields("machine_no")
            Rs.MoveNext
        Loop
    End If
    Set Rs = Nothing
End Sub

