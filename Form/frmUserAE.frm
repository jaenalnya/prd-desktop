VERSION 5.00
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmUserAE 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Buat Baru"
   ClientHeight    =   4035
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8685
   Icon            =   "frmUserAE.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4035
   ScaleWidth      =   8685
   StartUpPosition =   3  'Windows Default
   Begin HRD.Liner Liner2 
      Height          =   30
      Left            =   75
      TabIndex        =   19
      Top             =   3315
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   53
   End
   Begin HRD.Liner Liner1 
      Height          =   30
      Left            =   0
      TabIndex        =   18
      Top             =   780
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   53
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Administrator?"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1395
      TabIndex        =   16
      Top             =   2820
      Width           =   2415
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
      ScaleWidth      =   8685
      TabIndex        =   14
      Top             =   0
      Width           =   8685
      Begin VB.Image Image1 
         Height          =   720
         Left            =   45
         Picture         =   "frmUserAE.frx":038A
         Top             =   45
         Width           =   720
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Menu ini digunakan untuk menambah data User Pengguna"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   240
         Index           =   2
         Left            =   960
         TabIndex        =   17
         Top             =   405
         Width           =   5235
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "USER"
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
         Left            =   945
         TabIndex        =   15
         Top             =   135
         Width           =   2175
      End
   End
   Begin VB.TextBox txtRemarks 
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
      Left            =   5190
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Top             =   1290
      Width           =   3300
   End
   Begin VB.TextBox txtPassword 
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
      IMEMode         =   3  'DISABLE
      Left            =   1395
      Locked          =   -1  'True
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   2460
      Width           =   2400
   End
   Begin VB.TextBox txtUsername 
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
      Left            =   1395
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   2100
      Width           =   2400
   End
   Begin VB.TextBox txtLName 
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
      Left            =   1395
      TabIndex        =   2
      Top             =   1605
      Width           =   2400
   End
   Begin VB.TextBox txtFName 
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
      Left            =   1395
      TabIndex        =   1
      Top             =   1245
      Width           =   2400
   End
   Begin VB.TextBox txtKodeUser 
      BackColor       =   &H00C0FFFF&
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
      Left            =   1395
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   885
      Width           =   2400
   End
   Begin VB.ComboBox cboStatusCD 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "frmUserAE.frx":6504
      Left            =   5190
      List            =   "frmUserAE.frx":650E
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   885
      Width           =   1695
   End
   Begin lvButton.lvButtons_H cmdSave 
      Height          =   390
      Left            =   5895
      TabIndex        =   20
      Top             =   3465
      Width           =   1245
      _ExtentX        =   2196
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
      Image           =   "frmUserAE.frx":6524
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdCancel 
      Height          =   390
      Left            =   7335
      TabIndex        =   21
      Top             =   3465
      Width           =   1215
      _ExtentX        =   2143
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
      Image           =   "frmUserAE.frx":C6AE
      cBack           =   -2147483633
   End
   Begin VB.Label lblUser 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Keterangan"
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
      Index           =   6
      Left            =   4050
      TabIndex        =   13
      Top             =   1290
      Width           =   840
   End
   Begin VB.Label lblUser 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "StatusCD"
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
      Index           =   5
      Left            =   4065
      TabIndex        =   12
      Top             =   885
      Width           =   675
   End
   Begin VB.Label lblUser 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
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
      Index           =   3
      Left            =   120
      TabIndex        =   11
      Top             =   2460
      Width           =   690
   End
   Begin VB.Label lblUser 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Login Name"
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
      Index           =   4
      Left            =   120
      TabIndex        =   10
      Top             =   2100
      Width           =   825
   End
   Begin VB.Label lblUser 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nama akhir"
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
      Index           =   2
      Left            =   105
      TabIndex        =   9
      Top             =   1605
      Width           =   795
   End
   Begin VB.Label lblUser 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nama awal"
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
      Index           =   1
      Left            =   90
      TabIndex        =   8
      Top             =   1245
      Width           =   780
   End
   Begin VB.Label lblUser 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Kode User"
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
      Index           =   0
      Left            =   105
      TabIndex        =   7
      Top             =   885
      Width           =   735
   End
End
Attribute VB_Name = "frmUserAE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public State                        As FORM_STATE
Public PK                           As String
Dim srcSQL                          As String


Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdSave_Click()
On Error GoTo ErrTrapper
    If txtKodeUser.Text = "" Then
        MsgBox "Kode User harus di isi, silahkan cek kembali!", vbExclamation
        Exit Sub
    ElseIf txtFName.Text = "" Then
        MsgBox "Nama Awal harus di isi, silahkan cek kembali!", vbExclamation
        Exit Sub
    ElseIf txtUsername.Text = "" Then
        MsgBox "Username harus di isi, silahkan cek kembali!", vbExclamation
        Exit Sub
    ElseIf txtPassword.Text = "" Then
        MsgBox "Password harus di isi, silahkan cek kembali!", vbExclamation
        Exit Sub
    End If
    

        
    If State = AddStateMode Then
        If isRecordExist("tblUsers", "Username", txtUsername.Text, True) = True Then
            MsgBox "Nama Username sudah ada, silahkan isi yang lain!!", vbExclamation
            Exit Sub
        End If
        RS_USER.AddNew
        RS_USER.Fields("KodeUser") = txtKodeUser.Text
        RS_USER.Fields("NamaAwal") = txtFName.Text
        RS_USER.Fields("NamaAkhir") = txtLName.Text
        RS_USER.Fields("Username") = txtUsername.Text
        RS_USER.Fields("Password") = txtPassword.Text
        RS_USER.Fields("IsAdmin") = changeYNValue(Check1.Value)
        RS_USER.Fields("StatusCD") = cboStatusCD.Text
        RS_USER.Fields("Keterangan") = txtRemarks.Text
        RS_USER.Fields("DateEncoded") = Format(Now, "M/d/yyyy")
        RS_USER.Fields("EncodedBy") = ACTIVE_USER.USERNAME
        RS_USER.Update
        
        MsgBox "User baru telah berhasil di simpan!", vbInformation
        Unload Me
    
    ElseIf State = EditStateMode Then
        RS_USER.Fields("KodeUser") = txtKodeUser.Text
        RS_USER.Fields("NamaAwal") = txtFName.Text
        RS_USER.Fields("NamaAkhir") = txtLName.Text
        RS_USER.Fields("Username") = txtUsername.Text
        RS_USER.Fields("Password") = txtPassword.Text
        RS_USER.Fields("IsAdmin") = changeYNValue(Check1.Value)
        RS_USER.Fields("StatusCD") = cboStatusCD.Text
        RS_USER.Fields("Keterangan") = txtRemarks.Text
        RS_USER.Fields("LastDateModified") = Now
        RS_USER.Fields("ModifiedBy") = ACTIVE_USER.USERNAME
        RS_USER.Update
        
    MsgBox "Data berhasil disimpan!", vbInformation, Me.Caption
    Unload Me
    
    End If
Exit Sub
ErrTrapper:
    MsgBox "Error Number: " & Err.Number & vbCrLf & "Description: " & Err.Description, vbExclamation
End Sub

Private Sub Form_Activate()
On Error Resume Next
Dim i As Integer
    cboStatusCD.ListIndex = 0
    txtKodeUser.SetFocus
    Check1.BackColor = MAIN.ACPMenu.BackColor

    Me.BackColor = MAIN.ACPMenu.BackColor
    If MAIN.ACPMenu.Theme = 0 Then
        For i = 0 To 6
        lblUser(i).ForeColor = &HFFFFFF
        Next i
    Else
        For i = 0 To 6
        lblUser(i).ForeColor = &H0&
        Next i
    End If
End Sub

Private Sub Form_Load()
On Error GoTo ErrTrapper

CenterForm frmUserAE

If State = AddStateMode Then
    Me.Caption = "Buat Baru"
    txtKodeUser.Text = AutoID("tblUsers", "KodeUser", "US-")
    cboStatusCD.Text = "ACTIVE"
    txtUsername.Locked = False
    txtPassword.Locked = False

Else
    Me.Caption = "Ubah Data"
    txtKodeUser.Locked = True
    txtUsername.Locked = True
    txtPassword.Locked = True
    
    srcSQL = "SELECT * FROM tblUsers WHERE (((tblUsers.KodeUser)='" & PK & "'))"

    Set RS_USER = New ADODB.Recordset
    If RS_USER.State = adStateOpen Then RS_USER.Close
    RS_USER.Open srcSQL, CN, adOpenDynamic, adLockOptimistic
    
    DisplayForEditing
    
End If

Exit Sub
ErrTrapper:
    MsgBox "Error Number: " & Err.Number & vbCrLf & "Description: " & Err.Description, vbExclamation
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
Set frmUserAE = Nothing
Set RS_USER = Nothing
frmUser.CommandPass "Refresh"
End Sub

Private Sub DisplayForEditing()
On Error GoTo ErrHandler
    If RS_USER.RecordCount > 0 Then
        With RS_USER
            txtKodeUser.Text = .Fields("KodeUser")
            txtFName.Text = .Fields("NamaAwal")
            txtLName.Text = .Fields("NamaAkhir")
            Check1.Value = changeYNValue(.Fields("IsAdmin"))
            cboStatusCD.Text = .Fields("StatusCD")
            txtUsername.Text = .Fields("Username")
            txtPassword.Text = .Fields("Password")
            txtRemarks.Text = .Fields("Keterangan")
        End With
    End If
    Exit Sub
ErrHandler:
    If Err.Number = 94 Then Resume Next
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

