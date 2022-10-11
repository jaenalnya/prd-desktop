VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmLabelBarangAE 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Adjustment"
   ClientHeight    =   6000
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7920
   Icon            =   "frmLabelBarangAE.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6000
   ScaleWidth      =   7920
   StartUpPosition =   2  'CenterScreen
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
      ItemData        =   "frmLabelBarangAE.frx":15162
      Left            =   2250
      List            =   "frmLabelBarangAE.frx":15164
      TabIndex        =   21
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
      ItemData        =   "frmLabelBarangAE.frx":15166
      Left            =   2250
      List            =   "frmLabelBarangAE.frx":15170
      TabIndex        =   20
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
      ItemData        =   "frmLabelBarangAE.frx":15184
      Left            =   2250
      List            =   "frmLabelBarangAE.frx":15186
      TabIndex        =   19
      Text            =   "cboMachine"
      Top             =   1080
      Width           =   1635
   End
   Begin VB.TextBox txtKeterangan 
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
      Left            =   2250
      TabIndex        =   17
      Top             =   4680
      Width           =   5460
   End
   Begin VB.ComboBox cboJam 
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
      ItemData        =   "frmLabelBarangAE.frx":15188
      Left            =   2250
      List            =   "frmLabelBarangAE.frx":151D4
      Style           =   2  'Dropdown List
      TabIndex        =   16
      Top             =   3375
      Width           =   1635
   End
   Begin lvButton.lvButtons_H cmdSave 
      Height          =   420
      Left            =   4545
      TabIndex        =   14
      Top             =   5400
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
      Image           =   "frmLabelBarangAE.frx":15238
      cBack           =   -2147483633
   End
   Begin VB.ComboBox cboAdjust 
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
      ItemData        =   "frmLabelBarangAE.frx":2C2D2
      Left            =   2250
      List            =   "frmLabelBarangAE.frx":2C2DC
      Style           =   2  'Dropdown List
      TabIndex        =   13
      Top             =   3825
      Width           =   1635
   End
   Begin MSComCtl2.DTPicker DTPeriode 
      Height          =   330
      Left            =   2250
      TabIndex        =   12
      Top             =   2925
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   582
      _Version        =   393216
      CustomFormat    =   "dd/MMM/yyyy"
      Format          =   85131267
      CurrentDate     =   43998
   End
   Begin PRD.Liner Liner1 
      Height          =   30
      Left            =   45
      TabIndex        =   11
      Top             =   5265
      Width           =   7755
      _ExtentX        =   13679
      _ExtentY        =   53
   End
   Begin VB.TextBox txtQty 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2250
      TabIndex        =   5
      Text            =   "0"
      Top             =   4275
      Width           =   1590
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
      TabIndex        =   8
      Top             =   0
      Width           =   7920
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Adjustment Shot"
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
         Height          =   255
         Left            =   870
         TabIndex        =   10
         Top             =   150
         Width           =   5355
      End
      Begin VB.Image Image1 
         Height          =   720
         Left            =   0
         Picture         =   "frmLabelBarangAE.frx":2C2F0
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
         TabIndex        =   9
         Top             =   480
         Width           =   5100
      End
   End
   Begin lvButton.lvButtons_H cmdCancel 
      Height          =   420
      Left            =   6210
      TabIndex        =   15
      Top             =   5400
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
      Image           =   "frmLabelBarangAE.frx":3246A
      cBack           =   -2147483633
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
      TabIndex        =   24
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
      TabIndex        =   23
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
      TabIndex        =   22
      Top             =   2970
      Width           =   2805
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
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
      TabIndex        =   18
      Top             =   4725
      Width           =   1770
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Adjutment"
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
      TabIndex        =   7
      Top             =   3870
      Width           =   1905
   End
   Begin VB.Label Label7 
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
      Left            =   270
      TabIndex        =   6
      Top             =   4320
      Width           =   1770
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
      Caption         =   "Jam"
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
Attribute VB_Name = "frmLabelBarangAE"
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


Private Sub Add_counter()
'On Error GoTo ErrHandler

    Dim Rs As New ADODB.Recordset
    Dim sSQL As String
    Dim Counter As Variant
    Dim eng_prod_1 As String, eng_prod_2 As String

    Counter = 0
    Rs.CursorLocation = adUseClient
    
    sSQL = "select a.* from prod_runnings a where "
    sSQL = sSQL & " a.plant_mark = '" & p_plant_mark & "' "
    sSQL = sSQL & " and a.prod_machine_id = '" & Id_prod_machine & "'"
    sSQL = sSQL & " and a.mkt_customer_id = '" & Id_customer & "'"
    sSQL = sSQL & " and a.eng_product_id = '" & Id_product & "'"
    sSQL = sSQL & " and a.period_shift = '" & Format(DTPeriode.Value, "yyyy-mm-dd") & "'"
    sSQL = sSQL & " and a.period_hour = '" & cboJam.Text & "'"
    
    If Rs.State = adStateOpen Then Rs.Close
    Rs.Open sSQL, CN, adOpenDynamic, adLockPessimistic
    
    If Rs.RecordCount < 1 Then
        If cboAdjust.Text = "Tambah" Then
            sSQL = "insert into prod_runnings (plant_mark,prod_machine_id,mkt_customer_id,eng_product_id,"
            sSQL = sSQL & " date,period_shift,period_hour,counter_ok,created_at,created_by) values"
            sSQL = sSQL & " ('" & p_plant_mark & "','" & Id_prod_machine & "','" & Id_customer & "','" & Id_product & "'"
            sSQL = sSQL & " ,'" & Format(DTPeriode.Value, "yyyy-mm-dd hh:mm:ss") & "','" & Format(DTPeriode.Value, "yyyy-mm-dd") & "','" & cboJam.Text & "','" & txtQty.Text & "'"
            sSQL = sSQL & " ,'" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "','" & ACTIVE_USER.SYSID & "')"
            
            sSQL_Insert sSQL
        End If
        
    Else
        If cboAdjust.Text = "Tambah" Then
        
            Counter = Val(Rs.Fields("counter_ok")) + Val(txtQty.Text)
            sSQL = "update prod_runnings set counter_ok = '" & Counter & "',"
                sSQL = sSQL & " date = '" & Format(DTPeriode.Value, "yyyy-mm-dd hh:mm:ss") & "' where "
                sSQL = sSQL & " plant_mark = '" & p_plant_mark & "' "
                sSQL = sSQL & " and prod_machine_id = '" & Id_prod_machine & "'"
                sSQL = sSQL & " and mkt_customer_id = '" & Id_customer & "'"
                sSQL = sSQL & " and eng_product_id = '" & Id_product & "'"
                sSQL = sSQL & " and period_shift = '" & Format(DTPeriode.Value, "yyyy-mm-dd") & "'"
                sSQL = sSQL & " and period_hour = '" & cboJam.Text & "'"
            
            sSQL_Update sSQL
        ElseIf cboAdjust.Text = "Kurang" Then
            Counter = Val(Rs.Fields("counter_ok")) - Val(txtQty.Text)
            sSQL = "update prod_runnings set counter_ok = '" & Counter & "',"
                sSQL = sSQL & " date = '" & Format(DTPeriode.Value, "yyyy-mm-dd hh:mm:ss") & "' where "
                sSQL = sSQL & " plant_mark = '" & p_plant_mark & "' "
                sSQL = sSQL & " and prod_machine_id = '" & Id_prod_machine & "'"
                sSQL = sSQL & " and mkt_customer_id = '" & Id_customer & "'"
                sSQL = sSQL & " and eng_product_id = '" & Id_product & "'"
                sSQL = sSQL & " and period_shift = '" & Format(DTPeriode.Value, "yyyy-mm-dd") & "'"
                sSQL = sSQL & " and period_hour = '" & cboJam.Text & "'"
            
            sSQL_Update sSQL
        End If
        
        Set Rs = Nothing
     End If
    
Exit Sub
ErrHandler:
    MsgBox "Error Number: " & Err.Number & vbCrLf & "Description: " & Err.Description, vbExclamation, Me.Caption
     
     
End Sub



Private Sub cboMachine_Click()
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
End Sub

Private Sub cboMachine_KeyPress(KeyAscii As Integer)
    KeyAscii = AutoMatchCBBox(cboMachine, KeyAscii)
End Sub

Private Sub cboCustomer_Click()
    Dim Rs As New ADODB.Recordset
    Rs.CursorLocation = adUseClient
    Id_customer = ""
    
    Rs.Open "select a.id,a.name as customer_name from mkt_customers a where a.sys_plant_id = '" & p_sys_plant & "' and a.name = '" & cboCustomer.Text & "'", CN, adOpenStatic, adLockOptimistic
    If Rs.RecordCount > 0 Then
        Id_customer = Rs!id
        
        Call ComboProduct
        
    End If
    Set Rs = Nothing
    
    
End Sub

Private Sub cboCustomer_KeyPress(KeyAscii As Integer)
    KeyAscii = AutoMatchCBBox(cboCustomer, KeyAscii)
End Sub


Private Sub cboProduct_Click()
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
End Sub

Private Sub cboProduct_KeyPress(KeyAscii As Integer)
    KeyAscii = AutoMatchCBBox(cboProduct, KeyAscii)
End Sub



Private Sub cmdCancel_Click()
  Unload Me
End Sub


Private Sub cmdSave_Click()
'On Error GoTo ErrHandler
    If cboMachine.Text = "" Then
        MsgBox "Machine harus di isi, silahkan cek kembali!", vbExclamation
        Exit Sub
    ElseIf cboCustomer.Text = "" Then
        MsgBox "Customer harus di isi, silahkan cek kembali!", vbExclamation
        Exit Sub
    ElseIf cboProduct.Text = "" Then
        MsgBox "Product harus di isi, silahkan cek kembali!", vbExclamation
        Exit Sub
    ElseIf cboJam.Text = "" Then
        MsgBox "Jam harus di isi, silahkan cek kembali!", vbExclamation
        Exit Sub
    ElseIf lblInternalPart.Caption = "" Then
        MsgBox "Product harus di Klik, silahkan cek kembali!", vbExclamation
        Exit Sub
    ElseIf cboAdjust.Text = "" Then
        MsgBox "JAdjustment harus di isi, silahkan cek kembali!", vbExclamation
        Exit Sub
    ElseIf txtQty.Text = "" Then
        MsgBox "Qty harus di isi, silahkan cek kembali!", vbExclamation
        Exit Sub
    End If
    
    If State = AddStateMode Then

        RS_ADJUST.AddNew
        RS_ADJUST.Fields("plant_mark") = p_plant_mark
        RS_ADJUST.Fields("prod_machine_id") = Id_prod_machine
        RS_ADJUST.Fields("mkt_customer_id") = Id_customer
        RS_ADJUST.Fields("eng_product_id") = Id_product
        RS_ADJUST.Fields("date") = Format(DTPeriode, "yyyy-mm-dd hh:mm:ss")
        RS_ADJUST.Fields("period_shift") = Format(DTPeriode, "yyyy-mm-dd")
        RS_ADJUST.Fields("period_hour") = cboJam.Text
        RS_ADJUST.Fields("adjust") = cboAdjust.Text
        RS_ADJUST.Fields("qty") = txtQty.Text
        RS_ADJUST.Fields("description") = txtKeterangan.Text
        RS_ADJUST.Fields("created_at") = Format(Now, "yyyy-mm-dd hh:mm:ss")
        RS_ADJUST.Fields("created_by") = ACTIVE_USER.SYSID
        RS_ADJUST.Update
        
        Call Add_counter
        
        MsgBox "Data baru berhasil disimpan!", vbInformation
        
        Unload Me

    End If
Exit Sub
ErrHandler:
    MsgBox "Error Number: " & Err.Number & vbCrLf & " Description : " & Err.Description, vbExclamation, Me.Caption

End Sub

Private Sub Form_Activate()
'On Error Resume Next
    Me.BackColor = MAIN.ACPMenu.BackColor
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
'On Error Resume Next
    If KeyAscii = 27 Then
        Unload Me
    ElseIf KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub Form_Load()
'On Error GoTo ErrHandler
CenterForm frmAdjustmentAE

Call ComboMachine
Call ComboCustomer

DTPeriode.Value = Now

If State = AddStateMode Then
    Me.Caption = "Buat Baru"

    sSQL = "SELECT prod_adjustment_shots.* FROM prod_adjustment_shots limit 1"

    Set RS_ADJUST = New ADODB.Recordset
    If RS_ADJUST.State = adStateOpen Then RS_ADJUST.Close
    RS_ADJUST.Open sSQL, CN, adOpenDynamic, adLockOptimistic
    
End If

Exit Sub
ErrHandler:
    MsgBox "Error Number: " & Err.Number & vbCrLf & "Description: " & Err.Description, vbExclamation
End Sub

Private Sub Form_Unload(Cancel As Integer)
frmAdjustment.CommandPass "Refresh"
Set frmAdjustmentAE = Nothing
Set RS_ADJUST = Nothing
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error Resume Next
    Select Case KeyCode
    Case vbKeyF5
        cmdSave_Click
    Case vbKeyEscape
        cmdCancel_Click
    End Select
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
    
    sSQL = "select a.id,a.name as customer_name from mkt_customers a where a.sys_plant_id = '" & p_sys_plant & "' and a.`status` = 'active'"
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

