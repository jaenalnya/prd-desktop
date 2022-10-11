VERSION 5.00
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "LvButton.ocx"
Object = "{9EDDC69F-10E8-4DE7-9648-EC8A45F005C0}#1.0#0"; "B8CONT~2.OCX"
Begin VB.Form frmAbsensiAE 
   BackColor       =   &H80000007&
   ClientHeight    =   4560
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8640
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   4560
   ScaleWidth      =   8640
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2850
      Left            =   135
      ScaleHeight     =   2820
      ScaleWidth      =   5340
      TabIndex        =   8
      Top             =   1080
      Visible         =   0   'False
      Width           =   5370
      Begin lvButton.lvButtons_H cmdNumber 
         Height          =   735
         Index           =   0
         Left            =   90
         TabIndex        =   9
         Top             =   90
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1296
         Caption         =   "1"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin lvButton.lvButtons_H cmdNumber 
         Height          =   735
         Index           =   1
         Left            =   1035
         TabIndex        =   10
         Top             =   90
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1296
         Caption         =   "2"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin lvButton.lvButtons_H cmdNumber 
         Height          =   735
         Index           =   2
         Left            =   1980
         TabIndex        =   11
         Top             =   90
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1296
         Caption         =   "3"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin lvButton.lvButtons_H cmdNumber 
         Height          =   735
         Index           =   3
         Left            =   90
         TabIndex        =   12
         Top             =   990
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1296
         Caption         =   "4"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin lvButton.lvButtons_H cmdNumber 
         Height          =   735
         Index           =   4
         Left            =   1035
         TabIndex        =   13
         Top             =   990
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1296
         Caption         =   "5"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin lvButton.lvButtons_H cmdNumber 
         Height          =   735
         Index           =   5
         Left            =   1980
         TabIndex        =   14
         Top             =   990
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1296
         Caption         =   "6"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin lvButton.lvButtons_H cmdNumber 
         Height          =   735
         Index           =   6
         Left            =   90
         TabIndex        =   15
         Top             =   1890
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1296
         Caption         =   "7"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin lvButton.lvButtons_H cmdNumber 
         Height          =   735
         Index           =   7
         Left            =   1035
         TabIndex        =   16
         Top             =   1890
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1296
         Caption         =   "8"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin lvButton.lvButtons_H cmdNumber 
         Height          =   735
         Index           =   8
         Left            =   1980
         TabIndex        =   17
         Top             =   1890
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1296
         Caption         =   "9"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin lvButton.lvButtons_H cmdNumber 
         Height          =   735
         Index           =   9
         Left            =   2925
         TabIndex        =   18
         Top             =   1890
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1296
         Caption         =   "0"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin lvButton.lvButtons_H cmdNumber 
         Height          =   735
         Index           =   10
         Left            =   3915
         TabIndex        =   19
         Top             =   1890
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   1296
         Caption         =   "Enter"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin lvButton.lvButtons_H cmdNumber 
         Height          =   735
         Index           =   11
         Left            =   3915
         TabIndex        =   20
         Top             =   990
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   1296
         Caption         =   "Hapus"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin lvButton.lvButtons_H cmdNumber 
         Height          =   735
         Index           =   12
         Left            =   2925
         TabIndex        =   21
         Top             =   990
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1296
         Caption         =   "@"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin lvButton.lvButtons_H cmdNumber 
         Height          =   735
         Index           =   13
         Left            =   2925
         TabIndex        =   22
         Top             =   90
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1296
         Caption         =   "#"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin lvButton.lvButtons_H cmdNumber 
         Height          =   735
         Index           =   14
         Left            =   3915
         TabIndex        =   23
         Top             =   90
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   1296
         Caption         =   "< --- Backspace"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
   End
   Begin VB.TextBox txtUsername 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   405
      Left            =   5730
      TabIndex        =   0
      Top             =   2040
      Width           =   2040
   End
   Begin VB.TextBox txtPassword 
      BackColor       =   &H00FFFFFF&
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
      Height          =   405
      IMEMode         =   3  'DISABLE
      Left            =   5730
      PasswordChar    =   "#"
      TabIndex        =   1
      Top             =   2715
      Width           =   2040
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
      ScaleWidth      =   8640
      TabIndex        =   4
      Top             =   0
      Width           =   8640
      Begin PRD.Liner Liner1 
         Height          =   30
         Left            =   0
         TabIndex        =   5
         Top             =   960
         Width           =   10215
         _ExtentX        =   18018
         _ExtentY        =   53
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
         Left            =   960
         TabIndex        =   7
         Top             =   570
         Width           =   5100
      End
      Begin VB.Image Image1 
         Height          =   720
         Left            =   90
         Picture         =   "frmAbsensiAE.frx":0000
         Top             =   90
         Width           =   720
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "A B S E N S I   K A R Y A W A N"
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
         Left            =   960
         TabIndex        =   6
         Top             =   240
         Width           =   5355
      End
   End
   Begin lvButton.lvButtons_H cmdKeyUser 
      Height          =   420
      Left            =   7830
      TabIndex        =   24
      Top             =   2025
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   741
      Caption         =   "..."
      CapAlign        =   2
      BackStyle       =   7
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
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdLogin 
      Height          =   480
      Left            =   5745
      TabIndex        =   2
      Top             =   3330
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   847
      Caption         =   "&MASUK"
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
      Image           =   "frmAbsensiAE.frx":169B2
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdCancel 
      Height          =   480
      Left            =   7305
      TabIndex        =   3
      Top             =   3330
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   847
      Caption         =   "&EXIT"
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
      Image           =   "frmAbsensiAE.frx":1D214
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdKeyPass 
      Height          =   420
      Left            =   7830
      TabIndex        =   25
      Top             =   2700
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   741
      Caption         =   "..."
      CapAlign        =   2
      BackStyle       =   7
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
      cBack           =   -2147483633
   End
   Begin VB.Label lblAbsensi 
      BackStyle       =   0  'Transparent
      Caption         =   "> > >   M A S U K "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   5265
      TabIndex        =   30
      Top             =   1305
      Width           =   2985
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "http://tri-saudara.com"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   120
      TabIndex        =   29
      Top             =   4200
      Width           =   1620
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright © By jaenal. All Rights Reserved 2019"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   4725
      TabIndex        =   28
      Top             =   4230
      Width           =   3510
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&P I N"
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
      Left            =   5055
      TabIndex        =   27
      Top             =   2805
      Width           =   345
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&N I K"
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
      Left            =   5040
      TabIndex        =   26
      Top             =   2130
      Width           =   345
   End
   Begin VB.Image Image2 
      Height          =   3075
      Left            =   45
      Picture         =   "frmAbsensiAE.frx":20476
      Stretch         =   -1  'True
      Top             =   990
      Width           =   4425
   End
   Begin b8Controls4.b83DRect b83DRect1 
      Height          =   3120
      Left            =   0
      Top             =   960
      Width           =   9270
      _ExtentX        =   16351
      _ExtentY        =   5503
      Color1          =   16777215
      Color2          =   16777215
      Color3          =   14737632
      Color4          =   14737632
      BackColor       =   16119285
   End
End
Attribute VB_Name = "frmAbsensiAE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim dwLen                           As Long
Dim strString                       As String
Dim clsDS2                          As New clsDS2
Dim sPass                           As Byte
Dim sSQL                            As String

Private Sub cmdCancel_Click()
    Unload Me
End Sub
'
'Private Sub cmdClear_Click()
'
'    If MsgBox("Apakah akan Logout sebagai User 2 ?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
'
'    Dim iSQL As String
'    iSQL = "INSERT INTO hrd_login_logs (plant_mark,loc_code,hrd_employee_id,emp_code"
'    iSQL = iSQL & " ,tr_date,tr_time,acc_code,created_at,created_by)"
'    iSQL = iSQL & " VALUES ('" & p_plant_mark & "','" & NoMesin & "','" & ACTIVE_USER_2.KODEUSER & "','" & ACTIVE_USER_2.KODEPIN & "'"
'    iSQL = iSQL & " ,'" & Format(Now, "yyyy-mm-dd") & "','" & Format(Now, "hh:mm:ss") & "','2'"
'    iSQL = iSQL & " ,'" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "','" & ACTIVE_USER_2.SYSID & "')"
'
'    sSQL_Insert iSQL
'
'    ACTIVE_USER_2.SYSID = ""
'    ACTIVE_USER_2.KODENIK = ""
'    ACTIVE_USER_2.KODEPIN = ""
'    ACTIVE_USER_2.KODEUSER = ""
'    ACTIVE_USER_2.FULLNAME = ""
'    ACTIVE_USER_2.USERNAME = ""
'    ACTIVE_USER_2.PASSWORD = ""
'
'    With MAIN.StatusBar.Panels
'        .Item(6).Text = ""
'        .Item(7).Text = ""
'    End With
'
'    Unload Me
'End Sub

Private Sub cmdKeyPass_Click()
    sPass = 2
    Picture1.Visible = True
End Sub

Private Sub cmdKeyUser_Click()
    sPass = 1
    Picture1.Visible = True
End Sub

Private Sub cmdLogin_Click()
On Error GoTo ErrHandler

If txtUsername.text = "" Then
    MsgBox "NIK and/or PIN is incorrect.Try Again!", vbExclamation
    txtUsername.SetFocus
    Exit Sub
End If

If txtPassword.text = "" Then
    MsgBox "NIK and/or PIN is incorrect.Try Again!", vbExclamation
    txtPassword.SetFocus
    Exit Sub
End If

sSQL = "SELECT * FROM hrd_employees WHERE NIK = '" & txtUsername.text & "' AND PIN = '" & txtPassword.text & "' AND STATUS ='" & "active" & "'"


Set RS_ABSENSI = New ADODB.Recordset
If RS_ABSENSI.State = adStateOpen Then RS_ABSENSI.Close
RS_ABSENSI.Open sSQL, CN, adOpenStatic, adLockReadOnly

If RS_ABSENSI.BOF Or RS_ABSENSI.EOF = True Then
    MsgBox "NIK and/or PIN Tidak Ditemukan, Silahkan Coba Lagi!", vbExclamation
    Exit Sub

ElseIf RS_ABSENSI.Fields("status") = "suspend" Then
    MsgBox "User account is no longer active.Contact your administrator to re-activate your account!", vbExclamation
    Exit Sub

ElseIf Not RS_ABSENSI.Fields("nik") = txtUsername.text Then
    MsgBox "NIK Salah, Silahkan Coba Lagi!", vbExclamation
    Exit Sub

ElseIf Not RS_ABSENSI.Fields("pin") = txtPassword.text Then
    MsgBox "PIN Salah, Silahkan Coba Lagi!", vbExclamation
    Exit Sub

Else
    Dim iSQL As String
    
    If cmdLogin.Caption = "&MASUK" Then
    
        
        iSQL = "INSERT INTO hrd_absensi_logs (plant_mark,loc_code,hrd_employee_id,emp_code"
        iSQL = iSQL & " ,emp_nik,tr_date,tr_time,acc_code,created_at,created_by)"
        iSQL = iSQL & " VALUES ('" & p_plant_mark & "','" & NoMesin & "','" & RS_ABSENSI.Fields("id") & "','" & RS_ABSENSI.Fields("pin") & "'"
        iSQL = iSQL & " ,'" & RS_ABSENSI.Fields("nik") & "','" & Format(Now, "yyyy-mm-dd") & "','" & Format(Now, "hh:mm:ss") & "','1'"
        iSQL = iSQL & " ,'" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "','" & ACTIVE_USER.SYSID & "')"
        
        sSQL_Insert iSQL
    
    ElseIf cmdLogin.Caption = "&KELUAR" Then
    
        iSQL = "INSERT INTO hrd_absensi_logs (plant_mark,loc_code,hrd_employee_id,emp_code"
        iSQL = iSQL & " ,emp_nik,tr_date,tr_time,acc_code,created_at,created_by)"
        iSQL = iSQL & " VALUES ('" & p_plant_mark & "','" & NoMesin & "','" & RS_ABSENSI.Fields("id") & "','" & RS_ABSENSI.Fields("pin") & "'"
        iSQL = iSQL & " ,'" & RS_ABSENSI.Fields("nik") & "','" & Format(Now, "yyyy-mm-dd") & "','" & Format(Now, "hh:mm:ss") & "','2'"
        iSQL = iSQL & " ,'" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "','" & ACTIVE_USER.SYSID & "')"
        
        sSQL_Insert iSQL
    
    End If
    
    Unload Me

End If

Exit Sub
ErrHandler:
    MsgBox "Error Number: " & Err.Number & vbCrLf & "Description: " & Err.Description, vbExclamation, Me.Caption
    
End Sub



Private Sub cmdNumber_Click(Index As Integer)
If Index = 11 Then
    If sPass = 1 Then
        txtUsername.text = ""
    ElseIf sPass = 2 Then
        txtPassword.text = ""
    End If
ElseIf Index = 10 Then
    sPass = 0
    Picture1.Visible = False
ElseIf Index = 14 Then
    If sPass = 1 Then
        If Len(txtUsername.text) = 0 Then Exit Sub
        txtUsername.text = Mid(txtUsername.text, 1, Len(txtUsername.text) - 1)
    ElseIf sPass = 2 Then
        If Len(txtPassword.text) = 0 Then Exit Sub
        txtPassword.text = Mid(txtPassword.text, 1, Len(txtPassword.text) - 1)
    End If
Else
    If sPass = 1 Then
        txtUsername.text = txtUsername & cmdNumber(Index).Caption
    ElseIf sPass = 2 Then
        txtPassword.text = txtPassword & cmdNumber(Index).Caption
    End If
End If
End Sub

Private Sub Form_Activate()
On Error Resume Next
If END_APP = True Then Unload Me: Exit Sub
   
End Sub


Private Sub Form_Load()
On Error GoTo ErrHandler
CenterForm frmAbsensiAE

If Connected2DB = False Then END_APP = True: Unload Me: Exit Sub

Exit Sub
ErrHandler:
    MsgBox "Error Number: " & Err.Number & vbCrLf & "Description: " & Err.Description, vbExclamation, Me.Caption
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 27 Then
    End
ElseIf KeyAscii = 13 Then

    SendKeys "{TAB}"
End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If UnloadMode = 0 Then END_APP = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmAbsensiAE = Nothing
    Set RS_ABSENSI = Nothing
    frmAbsensi.CommandPass "Refresh"
End Sub


Private Sub txtPassword_GotFocus()
HLText txtPassword
End Sub

Private Sub txtPassword_LostFocus()
unHLText txtPassword
End Sub


Private Sub txtUsername_GotFocus()
HLText txtUsername
End Sub

Private Sub txtUsername_LostFocus()
unHLText txtUsername
End Sub


