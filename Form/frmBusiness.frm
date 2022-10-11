VERSION 5.00
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmBusiness 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5925
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6945
   Icon            =   "frmBusiness.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5925
   ScaleWidth      =   6945
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtFooter 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   705
      Left            =   1740
      MaxLength       =   200
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   20
      Top             =   4410
      Width           =   5025
   End
   Begin VB.TextBox txtCatatan 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   930
      Left            =   1740
      MaxLength       =   200
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   18
      Top             =   3420
      Width           =   5025
   End
   Begin VB.PictureBox picHeader 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   810
      Left            =   0
      ScaleHeight     =   810
      ScaleWidth      =   6945
      TabIndex        =   14
      Top             =   0
      Width           =   6945
      Begin VB.Image Image1 
         Height          =   720
         Left            =   30
         Picture         =   "frmBusiness.frx":617A
         Top             =   30
         Width           =   720
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Catatan : Pengisian data harus lengkap"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   240
         Index           =   1
         Left            =   960
         TabIndex        =   16
         Top             =   390
         Width           =   2895
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "INFORMASI PERUSAHAAN"
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
         Left            =   960
         TabIndex        =   15
         Top             =   180
         Width           =   3495
      End
   End
   Begin VB.TextBox txtAlamat 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   930
      Left            =   1740
      MaxLength       =   200
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   1710
      Width           =   5025
   End
   Begin VB.TextBox txtPerusahaan 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   1740
      MaxLength       =   50
      TabIndex        =   1
      Top             =   1305
      Width           =   5025
   End
   Begin VB.TextBox txtIDPerusahaan 
      BackColor       =   &H00FFFFFF&
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
      Height          =   330
      Left            =   1740
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   0
      Top             =   900
      Width           =   2400
   End
   Begin VB.TextBox txtNoTelepon 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   1740
      MaxLength       =   20
      TabIndex        =   3
      Top             =   2670
      Width           =   2505
   End
   Begin VB.TextBox txtNoFax 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   5070
      MaxLength       =   20
      TabIndex        =   4
      Top             =   2670
      Width           =   1680
   End
   Begin VB.TextBox txtEmail 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   1740
      MaxLength       =   30
      TabIndex        =   5
      Top             =   3030
      Width           =   2505
   End
   Begin lvButton.lvButtons_H cmdSave 
      Height          =   345
      Left            =   3840
      TabIndex        =   6
      Top             =   5445
      Width           =   1485
      _ExtentX        =   2619
      _ExtentY        =   609
      Caption         =   "&Simpan"
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
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdCancel 
      Height          =   345
      Left            =   5400
      TabIndex        =   7
      Top             =   5445
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   609
      Caption         =   "&Keluar"
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
      cBack           =   -2147483633
   End
   Begin HRD.Liner Liner2 
      Height          =   30
      Left            =   30
      TabIndex        =   17
      Top             =   5265
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   53
   End
   Begin VB.Label lblBussines 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Catatan Footer"
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
      Index           =   6
      Left            =   150
      TabIndex        =   21
      Top             =   4410
      Width           =   1230
   End
   Begin VB.Label lblBussines 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Catatan"
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
      Left            =   135
      TabIndex        =   19
      Top             =   3375
      Width           =   630
   End
   Begin VB.Label lblBussines 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Alamat"
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
      Left            =   180
      TabIndex        =   13
      Top             =   1665
      Width           =   555
   End
   Begin VB.Label lblBussines 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nama Perusahaan"
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
      Left            =   180
      TabIndex        =   12
      Top             =   1305
      Width           =   1440
   End
   Begin VB.Label lblBussines 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ID Perusahaan"
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
      Left            =   165
      TabIndex        =   11
      Top             =   900
      Width           =   1170
   End
   Begin VB.Label lblBussines 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Business Tel. No."
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
      Left            =   135
      TabIndex        =   10
      Top             =   2670
      Width           =   1395
   End
   Begin VB.Label lblBussines 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fax No."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   7
      Left            =   4395
      TabIndex        =   9
      Top             =   2715
      Width           =   570
   End
   Begin VB.Label lblBussines 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Email Address"
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
      TabIndex        =   8
      Top             =   3030
      Width           =   1110
   End
End
Attribute VB_Name = "frmBusiness"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public State                        As FORM_STATE

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdSave_Click()
Dim obj As Control
            For Each obj In Me
            If TypeOf obj Is TextBox Or TypeOf obj Is ComboBox Then
                If obj.Text = "" Then
                    MsgBox obj.Name & " Data tidak boleh kosong, silahkan lengkapi kembali.", vbExclamation, Me.Caption
                    obj.SetFocus
                    Exit Sub
                End If
            End If
            Next obj
            
            If State = EditStateMode Then
            
                Set RS_COMPANY = New ADODB.Recordset
                sSQL_Update "UPDATE Company_Info SET Perusahaan= '" & txtPerusahaan.Text & "', Alamat= '" & txtAlamat.Text & _
                "',NoTelepon='" & txtNoTelepon.Text & "',NoFax='" & txtNoFax.Text & "', Email='" & txtEmail.Text & "',LastDateModified= '" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "', ModifiedBy= '" & ACTIVE_USER.USERNAME & _
                "',Catatan = '" & txtCatatan.Text & "',Footerprint = '" & txtFooter.Text & "' WHERE IDPerusahaan='" & txtIDPerusahaan.Text & "'"
                
                
                With ACTIVE_COMPANY
                    .IDPerusahaan = txtIDPerusahaan.Text
                    .Perusahaan = txtPerusahaan.Text
                    .Alamat = txtAlamat.Text
                    .NoTelepon = txtNoTelepon.Text
                    .NoFax = txtNoFax.Text
                    .Email = txtEmail.Text
                    .Catatan = txtCatatan.Text
                    .FooterPrint = txtFooter.Text
                End With
                
                MsgBox "Data berhasil disimpan!", vbInformation
                Unload Me
            End If
End Sub

Private Sub Form_Activate()

On Error Resume Next
    Dim i As Integer
    Me.BackColor = MAIN.ACPMenu.BackColor
    
    If MAIN.ACPMenu.Theme = 0 Then
        For i = 0 To 7
        lblBussines(i).ForeColor = &HFFFFFF
        Next i
    Else
        For i = 0 To 7
        lblBussines(i).ForeColor = &H0&
        Next i
    End If
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
Dim sSQL As String
CenterForm frmBusiness


State = EditStateMode

sSQL = "SELECT Company_Info.* " & _
            "FROM Company_Info "

Set RS_COMPANY = New ADODB.Recordset
If RS_COMPANY.State = adStateOpen Then RS_COMPANY.Close
RS_COMPANY.Open sSQL, CN, adOpenDynamic, adLockOptimistic

With RS_COMPANY
    txtIDPerusahaan.Text = .Fields("IDPerusahaan")
    txtPerusahaan.Text = .Fields("Perusahaan")
    txtAlamat.Text = .Fields("Alamat")
    txtNoTelepon.Text = .Fields("NoTelepon")
    txtNoFax.Text = .Fields("NoFax")
    txtEmail.Text = .Fields("Email")
    txtCatatan.Text = .Fields("Catatan")
    txtFooter.Text = .Fields("FooterPrint")
End With

Exit Sub
ErrHandler:
    MsgBox "Error Number: " & Err.Number & vbCrLf & "Description: " & Err.Description, vbExclamation, Me.Caption
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmBusiness = Nothing
Set RS_COMPANY = Nothing
End Sub

Private Sub txtAlamat_GotFocus()
HLText txtAlamat
End Sub

Private Sub txtAlamat_LostFocus()
unHLText txtAlamat

End Sub

Private Sub txtNoTelepon_GotFocus()
HLText txtNoTelepon
End Sub

Private Sub txtNoTelepon_LostFocus()
unHLText txtNoTelepon

End Sub

Private Sub txtIDPerusahaan_GotFocus()
HLText txtIDPerusahaan
End Sub

Private Sub txtIDPerusahaan_LostFocus()
unHLText txtIDPerusahaan

End Sub

Private Sub txtPerusahaan_GotFocus()
HLText txtPerusahaan
End Sub

Private Sub txtPerusahaan_LostFocus()
unHLText txtPerusahaan

End Sub

Private Sub txtEmail_GotFocus()
HLText txtEmail
End Sub

Private Sub txtEmail_LostFocus()
unHLText txtEmail

End Sub

Private Sub txtNoFax_GotFocus()
HLText txtNoFax
End Sub


Private Sub txtNoFax_LostFocus()
unHLText txtNoFax

End Sub


Private Sub txtCatatan_GotFocus()
HLText txtCatatan
End Sub

Private Sub txtCatatan_LostFocus()
unHLText txtCatatan
End Sub

Private Sub txtFooter_GotFocus()
HLText txtFooter
End Sub

Private Sub txtFooter_LostFocus()
unHLText txtFooter
End Sub
