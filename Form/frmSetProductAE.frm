VERSION 5.00
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmSetProductAE 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4275
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9555
   Icon            =   "frmSetProductAE.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4275
   ScaleWidth      =   9555
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cboSetting 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   6
      Left            =   5355
      TabIndex        =   21
      Top             =   2790
      Width           =   3795
   End
   Begin VB.ComboBox cboSetting 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   5
      Left            =   5355
      TabIndex        =   20
      Top             =   2385
      Width           =   3795
   End
   Begin PRD.Liner Liner3 
      Height          =   30
      Left            =   45
      TabIndex        =   18
      Top             =   3330
      Width           =   9510
      _ExtentX        =   16775
      _ExtentY        =   53
   End
   Begin VB.ComboBox cboSetting 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   4
      Left            =   7335
      TabIndex        =   17
      Top             =   1305
      Width           =   1815
   End
   Begin VB.ComboBox cboSetting 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   3
      Left            =   1305
      TabIndex        =   4
      Top             =   2790
      Width           =   3795
   End
   Begin VB.ComboBox cboSetting 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   2
      Left            =   1305
      TabIndex        =   3
      Top             =   2340
      Width           =   3795
   End
   Begin PRD.Liner Liner2 
      Height          =   30
      Left            =   45
      TabIndex        =   14
      Top             =   1845
      Width           =   9555
      _ExtentX        =   16854
      _ExtentY        =   53
   End
   Begin VB.ComboBox cboSetting 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   1
      Left            =   1305
      TabIndex        =   2
      Top             =   1305
      Width           =   3795
   End
   Begin VB.ComboBox cboSetting 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   0
      Left            =   1305
      TabIndex        =   1
      Top             =   855
      Width           =   3795
   End
   Begin VB.TextBox txtentry 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Index           =   4
      Left            =   1305
      MaxLength       =   100
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Tag             =   "Name"
      Top             =   3465
      Width           =   3795
   End
   Begin VB.PictureBox picHeader 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   720
      Left            =   0
      ScaleHeight     =   720
      ScaleWidth      =   9555
      TabIndex        =   0
      Top             =   0
      Width           =   9555
      Begin VB.PictureBox Liner1 
         Height          =   30
         Left            =   0
         ScaleHeight     =   30
         ScaleWidth      =   9465
         TabIndex        =   7
         Top             =   945
         Width           =   9465
      End
      Begin VB.Image Image1 
         Height          =   720
         Left            =   45
         Picture         =   "frmSetProductAE.frx":617A
         Top             =   -45
         Width           =   720
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Menu ini digunakan merubah data Product untuk Setiap Mesin"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   240
         Left            =   855
         TabIndex        =   9
         Top             =   360
         Width           =   4695
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "SETTING PRODUCT ON MACHINE"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   855
         TabIndex        =   8
         Top             =   135
         Width           =   3495
      End
   End
   Begin lvButton.lvButtons_H cmdSave 
      Height          =   390
      Left            =   6390
      TabIndex        =   6
      Top             =   3555
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   688
      Caption         =   "&Simpan [F5]"
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
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      Image           =   "frmSetProductAE.frx":C2F4
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdCancel 
      Height          =   390
      Left            =   7830
      TabIndex        =   10
      Top             =   3555
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   688
      Caption         =   "&Batal [ESC]"
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
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      Image           =   "frmSetProductAE.frx":1247E
      cBack           =   -2147483633
   End
   Begin VB.Label lblproduct 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Status"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   5
      Left            =   6570
      TabIndex        =   19
      Top             =   1350
      Width           =   525
   End
   Begin VB.Label lblproduct 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Product 2"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   4
      Left            =   135
      TabIndex        =   16
      Top             =   2835
      Width           =   810
   End
   Begin VB.Label lblproduct 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Product 1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   0
      Left            =   135
      TabIndex        =   15
      Top             =   2385
      Width           =   810
   End
   Begin VB.Label lblproduct 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mesin"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   1
      Left            =   135
      TabIndex        =   13
      Top             =   900
      Width           =   450
   End
   Begin VB.Label lblproduct 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Keterangan"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   3
      Left            =   180
      TabIndex        =   12
      Top             =   3555
      Width           =   945
   End
   Begin VB.Label lblproduct 
      AutoSize        =   -1  'True
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
      Height          =   210
      Index           =   2
      Left            =   135
      TabIndex        =   11
      Top             =   1350
      Width           =   780
   End
End
Attribute VB_Name = "frmSetProductAE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public State                        As FORM_STATE
Public PK                           As String
Dim sSQL                            As String
Dim idmachine                       As Variant
Dim idcustomer                      As Variant
Dim idproduct_1                     As Variant
Dim idproduct_2                     As Variant

Private Sub cboSetting_Change(Index As Integer)

    Select Case Index
        Case 2
            GetPartname 2, 5
        Case 3
            GetPartname 3, 6
        Case 5
            GetInternalID 5, 2
        Case 6
            GetInternalID 6, 3
    End Select
End Sub

Private Sub cboSetting_Click(Index As Integer)

    Select Case Index
        Case 2
            GetPartname 2, 5
        Case 3
            GetPartname 3, 6
        Case 5
            GetInternalID 5, 2
        Case 6
            GetInternalID 6, 3
    End Select
End Sub

Private Sub GetPartname(ByVal X As Integer, ByVal Y As Integer)
    Dim Rs As New Recordset
    Rs.CursorLocation = adUseClient
    Rs.Open "Select * From eng_products Where internal_part_id = '" & cboSetting(X).Text & "'", CN, adOpenStatic, adLockOptimistic
    If Rs.RecordCount > 0 Then
        If Rs!internal_part_id <> "" Then
            cboSetting(Y).Text = Rs!customer_part_name
            If X = 2 Then
                idproduct_1 = Rs!id
            ElseIf X = 3 Then
                idproduct_2 = Rs!id
            End If
        End If
    End If
    Set Rs = Nothing
End Sub

Private Sub GetInternalID(ByVal X As Integer, ByVal Y As Integer)
    Dim Rs As New Recordset
    Rs.CursorLocation = adUseClient

    Rs.Open "Select * From eng_products Where customer_part_name = '" & cboSetting(X).Text & "'", CN, adOpenStatic, adLockOptimistic
    If Rs.RecordCount > 0 Then
        If Rs!customer_part_name <> "" Then
            cboSetting(Y).Text = Rs!internal_part_id
            If X = 5 Then
                idproduct_1 = Rs!id
            ElseIf X = 6 Then
                idproduct_2 = Rs!id
            End If
        End If
    End If
    Set Rs = Nothing
End Sub

Private Sub cboSetting_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyAscii = AutoMatchCBBox(cboSetting(Index), KeyAscii)
End Sub


Private Sub cmdCancel_Click()
  Unload Me
End Sub


Private Sub cmdSave_Click()
'On Error GoTo ErrHandler
    If cboSetting(0).Text = "" Then
        MsgBox "Mesin harus di isi, silahkan cek kembali!", vbExclamation
        Exit Sub
    ElseIf cboSetting(1).Text = "" Then
        MsgBox "Customer harus di isi, silahkan cek kembali!", vbExclamation
        Exit Sub
    ElseIf cboSetting(2).Text = "" Then
        MsgBox "Product harus di isi, silahkan cek kembali!", vbExclamation
        Exit Sub
    ElseIf cboSetting(4).Text = "" Then
        MsgBox "Loading Mesin harus di isi, silahkan cek kembali!", vbExclamation
        Exit Sub
    End If
        
        If cboSetting(3).Text = "" Or cboSetting(6).Text = "" Then
            idproduct_2 = "NULL"
        End If
    
    If State = AddStateMode Then

    ElseIf State = EditStateMode Then
    
        sSQL = "update prod_running_products set eng_product_1 = " & idproduct_1 & ",eng_product_2 = " & idproduct_2 & ", " & _
            "description = '" & txtentry(4).Text & "',machine_status = '" & cboSetting(4).Text & "', " & _
            "updated_at = '" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "',updated_by = " & ACTIVE_USER.KODEUSER & " " & _
            "Where prod_machine_id = '" & idmachine & "'"
        sSQL_Update sSQL
        MsgBox "Data berhasil disimpan!", vbInformation
        Unload Me
    End If
Exit Sub
ErrHandler:
    MsgBox "Error Number: " & Err.Number & vbCrLf & " Description : " & Err.Description, vbExclamation, Me.Caption

End Sub

Private Sub Form_Activate()
'On Error Resume Next
    Dim i As Integer
    cboSetting(1).SetFocus
    Me.BackColor = MAIN.ACPMenu.BackColor

    If MAIN.ACPMenu.Theme = 0 Then
        For i = 1 To 5
        lblproduct(i).ForeColor = &HFFFFFF
        Next i
    Else
        For i = 1 To 5
        lblproduct(i).ForeColor = &H0&
        Next i
    End If
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
CenterForm frmSetProductAE

    AddComboField "prod_machines", "number", "name", cboSetting(0), "sys_plant_id", "3"
    AddComboOne "mkt_customers", "name", cboSetting(1), "sys_plant_id", "3"
    AddComboOne "eng_products", "internal_part_id", cboSetting(2), "status_plant_3", "active"
    AddComboOne "eng_products", "internal_part_id", cboSetting(3), "status_plant_3", "active"
    
    AddComboOne "eng_products", "customer_part_name", cboSetting(5), "status_plant_3", "active"
    AddComboOne "eng_products", "customer_part_name", cboSetting(6), "status_plant_3", "active"
    
If State = AddStateMode Then
    Me.Caption = "Buat Baru"
    sSQL = "SELECT tblBarang.* FROM prod_running_products"

    Set RS_PRODUCT = New ADODB.Recordset
    If RS_PRODUCT.State = adStateOpen Then RS_PRODUCT.Close
    RS_PRODUCT.Open sSQL, CN, adOpenDynamic, adLockOptimistic
    
ElseIf State = EditStateMode Then
    Me.Caption = "Ubah Data"
    
    sSQL = "SELECT a.id,a.plant_mark,a.prod_machine_id, b.machine_no,b.machine_name, b.tonnage, a.mkt_customer_id,"
    sSQL = sSQL & " c.customer_name,a.eng_product_1, d.prod_name_1,d.internal_part_1,d.customer_part_1,"
    sSQL = sSQL & " a.eng_product_2, e.prod_name_2,e.internal_part_2,e.customer_part_2,a.Description , a.machine_status,a.updated_at,a.updated_by"
    sSQL = sSQL & " FROM prod_running_products a Left Join"
    sSQL = sSQL & " (SELECT prod_running_products.id, prod_machines.number as machine_no,prod_machines.name as machine_name, "
    sSQL = sSQL & " prod_machines.tonnage FROM prod_running_products INNER JOIN prod_machines ON "
    sSQL = sSQL & " prod_running_products.prod_machine_id = prod_machines.id) b ON a.id = b.id"
    sSQL = sSQL & " Left Join (SELECT prod_running_products.id, mkt_customers.name as customer_name FROM prod_running_products "
    sSQL = sSQL & " INNER JOIN mkt_customers ON prod_running_products.mkt_customer_id = mkt_customers.id) c ON a.id = c.id"
    sSQL = sSQL & " Left Join (SELECT prod_running_products.id, eng_products.name as prod_name_1,"
    sSQL = sSQL & " eng_products.internal_part_id as internal_part_1,eng_products.customer_part_number as customer_part_1 "
    sSQL = sSQL & " FROM prod_running_products INNER JOIN eng_products ON prod_running_products.eng_product_1 = eng_products.id) d"
    sSQL = sSQL & " ON a.id = d.id Left Join (SELECT prod_running_products.id, eng_products.name as prod_name_2,"
    sSQL = sSQL & " eng_products.internal_part_id as internal_part_2, eng_products.customer_part_number as customer_part_2 "
    sSQL = sSQL & " FROM prod_running_products INNER JOIN eng_products ON prod_running_products.eng_product_2 = eng_products.id) e"
    sSQL = sSQL & " ON a.id = e.id WHERE b.machine_no = " & PK & ""

    Set RS_PRODUCT = New ADODB.Recordset
    If RS_PRODUCT.State = adStateOpen Then RS_PRODUCT.Close
    RS_PRODUCT.Open sSQL, CN, adOpenDynamic, adLockOptimistic
    
    If RS_PRODUCT.RecordCount > 0 Then
        With RS_PRODUCT
            idmachine = .Fields("prod_machine_id")
            idcustomer = .Fields("mkt_customer_id")
            idproduct_1 = IIf(IsNull(.Fields("eng_product_1")), 0, .Fields("eng_product_1"))
            idproduct_2 = IIf(IsNull(.Fields("eng_product_2")), 0, .Fields("eng_product_2"))
            
            cboSetting(0).Text = .Fields("machine_no") & " - " & .Fields("machine_name")
            cboSetting(1).Text = .Fields("customer_name")
            cboSetting(2).Text = IIf(IsNull(.Fields("internal_part_1")), "", .Fields("internal_part_1"))
            cboSetting(3).Text = IIf(IsNull(.Fields("internal_part_2")), "", .Fields("internal_part_2"))
            cboSetting(4).Text = .Fields("machine_status")
            cboSetting(5).Text = IIf(IsNull(.Fields("prod_name_1")), "", .Fields("prod_name_1"))
            cboSetting(6).Text = IIf(IsNull(.Fields("prod_name_2")), "", .Fields("prod_name_2"))

            
        End With
        
        cboSetting(0).Enabled = False
    End If

End If

Exit Sub
ErrHandler:
    MsgBox "Error Number: " & Err.Number & vbCrLf & "Description: " & Err.Description, vbExclamation
End Sub

Private Sub Form_Unload(Cancel As Integer)

frmSetProduct.CommandPass "Refresh"
Set frmSetProductAE = Nothing
Set RS_PRODUCT = Nothing

End Sub

Private Sub txtEntry_GotFocus(Index As Integer)
    HLText txtentry(Index)
End Sub

Private Sub txtEntry_LostFocus(Index As Integer)
    unHLText txtentry(Index)
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


