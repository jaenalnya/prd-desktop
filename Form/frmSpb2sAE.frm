VERSION 5.00
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmSpb2sAE 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SPB 2 Sementara"
   ClientHeight    =   5040
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   8475
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSpb2sAE.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5040
   ScaleWidth      =   8475
   StartUpPosition =   2  'CenterScreen
   Begin PRD.Liner Liner2 
      Height          =   30
      Left            =   45
      TabIndex        =   29
      Top             =   4275
      Width           =   8160
      _ExtentX        =   14393
      _ExtentY        =   53
   End
   Begin VB.TextBox txtEntry 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   12
      Left            =   2205
      Locked          =   -1  'True
      TabIndex        =   26
      Top             =   1035
      Width           =   3255
   End
   Begin lvButton.lvButtons_H cmdOk 
      Height          =   465
      Left            =   3285
      TabIndex        =   25
      Top             =   4410
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   820
      Caption         =   "OK"
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
      Image           =   "frmSpb2sAE.frx":617A
      cBack           =   -2147483633
   End
   Begin VB.TextBox txtEntry 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   11
      Left            =   6165
      Locked          =   -1  'True
      TabIndex        =   23
      Top             =   3735
      Width           =   1950
   End
   Begin VB.TextBox txtEntry 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   10
      Left            =   6165
      Locked          =   -1  'True
      TabIndex        =   21
      Top             =   3285
      Width           =   1950
   End
   Begin VB.TextBox txtEntry 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   9
      Left            =   6165
      Locked          =   -1  'True
      TabIndex        =   19
      Top             =   2835
      Width           =   1950
   End
   Begin VB.TextBox txtEntry 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   8
      Left            =   6165
      Locked          =   -1  'True
      TabIndex        =   17
      Top             =   2385
      Width           =   1950
   End
   Begin VB.TextBox txtEntry 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   7
      Left            =   2205
      Locked          =   -1  'True
      TabIndex        =   15
      Top             =   3735
      Width           =   1950
   End
   Begin VB.TextBox txtEntry 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   6
      Left            =   2205
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   3285
      Width           =   1950
   End
   Begin VB.TextBox txtEntry 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   5
      Left            =   2205
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   2835
      Width           =   1950
   End
   Begin VB.TextBox txtEntry 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   4
      Left            =   2205
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   2385
      Width           =   1950
   End
   Begin VB.TextBox txtEntry 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   3
      Left            =   5715
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   1935
      Width           =   2400
   End
   Begin VB.TextBox txtEntry 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   2
      Left            =   5715
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   1485
      Width           =   2400
   End
   Begin VB.TextBox txtEntry 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   1
      Left            =   2205
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   1935
      Width           =   3255
   End
   Begin VB.TextBox txtEntry 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   0
      Left            =   2205
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   1485
      Width           =   3255
   End
   Begin PRD.Liner Liner1 
      Height          =   30
      Left            =   45
      TabIndex        =   1
      Top             =   900
      Width           =   8115
      _ExtentX        =   20743
      _ExtentY        =   53
   End
   Begin VB.TextBox txtBarcode 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   2250
      TabIndex        =   0
      Top             =   90
      Width           =   3210
   End
   Begin lvButton.lvButtons_H cmdCancel 
      Height          =   420
      Left            =   7020
      TabIndex        =   28
      Top             =   135
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   741
      Caption         =   "Close"
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
      Image           =   "frmSpb2sAE.frx":C304
      cBack           =   -2147483633
   End
   Begin VB.Label lbleng_product_id 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      Height          =   285
      Left            =   5760
      TabIndex        =   31
      Top             =   1080
      Width           =   2355
   End
   Begin VB.Label lblBarcode 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      Height          =   285
      Left            =   2250
      TabIndex        =   30
      Top             =   630
      Width           =   3210
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Customer Tujuan"
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
      Left            =   180
      TabIndex        =   27
      Top             =   1080
      Width           =   1815
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Scan Barcode "
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
      Left            =   180
      TabIndex        =   24
      Top             =   180
      Width           =   1815
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "No Mesin"
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
      Left            =   4815
      TabIndex        =   22
      Top             =   3780
      Width           =   1275
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Cavity"
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
      Left            =   4815
      TabIndex        =   20
      Top             =   3330
      Width           =   1275
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Qty / Box"
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
      Left            =   4815
      TabIndex        =   18
      Top             =   2880
      Width           =   1275
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Jumlah Label"
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
      Left            =   4815
      TabIndex        =   16
      Top             =   2430
      Width           =   1275
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Operator"
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
      Left            =   180
      TabIndex        =   14
      Top             =   3780
      Width           =   1815
   End
   Begin VB.Label Label5 
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
      Height          =   330
      Left            =   180
      TabIndex        =   12
      Top             =   3330
      Width           =   1815
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Model"
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
      Left            =   180
      TabIndex        =   10
      Top             =   2880
      Width           =   1815
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Kode Unik"
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
      Left            =   180
      TabIndex        =   8
      Top             =   2430
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Customer Part Number"
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
      Left            =   180
      TabIndex        =   4
      Top             =   1980
      Width           =   1995
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Internal Part Number"
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
      Left            =   180
      TabIndex        =   2
      Top             =   1530
      Width           =   1815
   End
End
Attribute VB_Name = "frmSpb2sAE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()
Dim sSQL As String
    If txtEntry(12) <> p_customer_name Then
        MsgBox "Data Tidak Sesuai Customer", vbExclamation
        Exit Sub
    ElseIf txtEntry(11).Text <> p_machine_no Then
        MsgBox "Data Tidak Sesuai No Mesin..!", vbExclamation
        Exit Sub
    Else
    
        Dim p_shift As Date
        If Format(Now, "HH") >= 0 And Format(Now, "HH") <= 7 Then
            p_shift = Format(DateAdd("d", -1, Format(Now, "yyyy-mm-dd")), "yyyy-mm-dd")
        Else
            p_shift = Format(Now, "yyyy-mm-dd")
        End If
    
        sSQL = "Insert Into sip_production.prod_spb2s_logs (plant_mark,qc_label_product_id,prod_machine_id,"
        sSQL = sSQL & " mkt_customer_id,eng_product_id,date,period_shift,qty,created_at,created_by)"
        sSQL = sSQL & " values ('" & p_plant_mark & "','" & lblBarcode.Caption & "','" & p_prod_machine_id & "'"
        sSQL = sSQL & " ,'" & p_mkt_customer_id & "','" & lbleng_product_id.Caption & "','" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "'"
        sSQL = sSQL & " ,'" & Format(p_shift, "yyyy-mm-dd") & "','" & txtEntry(9).Text & "'"
        sSQL = sSQL & " ,'" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "','" & ACTIVE_USER.KODEUSER & "')"
        
        sSQL_Insert sSQL
        
        Unload Me
    End If

        
End Sub

Private Sub Form_Activate()
    With MAIN
        Me.BackColor = .ACPMenu.BackColor
        Me.Picture = .ACPMenu.LoadBackground
    End With
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
    Select Case KeyCode
    Case vbKeyEscape
        Unload Me
    End Select
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If KeyAscii = 27 Then
        Unload Me
    ElseIf KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmSpb2sAE = Nothing
    'Set RS_SPB2E = Nothing
End Sub


Private Sub txtBarcode_KeyPress(KeyAscii As Integer)
On Error Resume Next
Dim Rs As New Recordset
Dim Rs_search As New Recordset
Dim sSQL As String
Dim i As Integer
If KeyAscii = 13# Then
    Rs_search.CursorLocation = adUseClient
    Rs_search.Open "select * from sip_production.prod_spb2s_logs a where a.qc_label_product_id = ''", CN, adOpenStatic, adLockOptimistic
    If Rs_search.RecordCount > 0 Then
        MsgBox "Data Sudah Pernah di Scan..!", vbExclamation
        For i = 0 To 12
            txtEntry(i).Text = ""
        Next i
        txtBarcode.Text = ""
        txtBarcode.SetFocus
        Exit Sub
    Else
        Rs.CursorLocation = adUseClient
        sSQL = "select a.id,a.seq,a.cavity,a.engine_number,a.quantity,a.quantity_box,a.eng_product_id,"
        sSQL = sSQL & " b.name as plant_name,c.name as customer_name,d.shift,e.nik,e.name as employee_name,"
        sSQL = sSQL & " f.internal_part_id,f.name as product_name, f.customer_part_number,f.customer_part_name,"
        sSQL = sSQL & " f.unix_code,f.model from sip_234.qc_label_products a"
        sSQL = sSQL & " left join sip_234.sys_plants b on a.sys_plant_id = b.id"
        sSQL = sSQL & " left join sip_234.mkt_customers c on a.mkt_customer_id = c.id"
        sSQL = sSQL & " left join sip_234.hrd_work_shifts d on a.hrd_work_shift_id = d.id"
        sSQL = sSQL & " left join sip_234.hrd_employees e on a.hrd_employee_id = e.id"
        sSQL = sSQL & " left join sip_234.eng_products f on a.eng_product_id = f.id"
        sSQL = sSQL & " where a.id = '" & txtBarcode.Text & "'"
        
        Rs.Open sSQL, CN, adOpenStatic, adLockOptimistic
        If Rs.RecordCount < 1 Then
            MsgBox "Data Tidak ditemukan", vbExclamation
            For i = 0 To 12
                txtEntry(i).Text = ""
            Next i
            txtBarcode.Text = ""
            txtBarcode.SetFocus
            Exit Sub
        Else
            
            txtEntry(0).Text = Rs.Fields("product_name")
            txtEntry(1).Text = Rs.Fields("customer_part_name")
            txtEntry(2).Text = Rs.Fields("internal_part_id")
            txtEntry(3).Text = Rs.Fields("customer_part_number")
            txtEntry(4).Text = IIf(IsNull(Rs.Fields("unix_code")), "", Rs.Fields("unix_code"))
            txtEntry(5).Text = IIf(IsNull(Rs.Fields("model")), "", Rs.Fields("model"))
            txtEntry(6).Text = Rs.Fields("shift")
            txtEntry(7).Text = Rs.Fields("employee_name")
            txtEntry(8).Text = Rs.Fields("quantity")
            txtEntry(9).Text = Rs.Fields("quantity_box")
            txtEntry(10).Text = Rs.Fields("cavity")
            txtEntry(11).Text = Rs.Fields("engine_number")
            txtEntry(12).Text = Rs.Fields("customer_name")
            lblBarcode.Caption = Rs.Fields("id")
            lbleng_product_id.Caption = Rs.Fields("eng_product_id")
            txtBarcode.Text = ""
        End If
    
        Set Rs = Nothing
    Set Rs_search = Nothing
    End If
End If
End Sub
