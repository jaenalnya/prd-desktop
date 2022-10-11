VERSION 5.00
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{9EDDC69F-10E8-4DE7-9648-EC8A45F005C0}#1.0#0"; "b8Controls4.ocx"
Begin VB.Form frmUnlock 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4590
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   8460
   Icon            =   "frmUnlock.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4590
   ScaleWidth      =   8460
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picHeader 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   945
      Left            =   0
      ScaleHeight     =   945
      ScaleWidth      =   8460
      TabIndex        =   18
      Top             =   0
      Width           =   8460
      Begin PRD.Liner Liner1 
         Height          =   30
         Left            =   0
         TabIndex        =   19
         Top             =   960
         Width           =   10215
         _ExtentX        =   18018
         _ExtentY        =   53
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "P  R O D U C T I O N  S Y S T E M ®™"
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
         TabIndex        =   21
         Top             =   240
         Width           =   5355
      End
      Begin VB.Image Image1 
         Height          =   780
         Left            =   120
         Picture         =   "frmUnlock.frx":617A
         Top             =   120
         Width           =   720
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "PT Tri-Saudara Sentosa Industri"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   255
         Index           =   2
         Left            =   960
         TabIndex        =   20
         Top             =   525
         Width           =   3255
      End
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
      TabIndex        =   17
      Top             =   2445
      Width           =   2040
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
      TabIndex        =   16
      Top             =   1905
      Width           =   2040
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2805
      Left            =   90
      ScaleHeight     =   2775
      ScaleWidth      =   5430
      TabIndex        =   0
      Top             =   1080
      Visible         =   0   'False
      Width           =   5460
      Begin lvButton.lvButtons_H cmdNumber 
         Height          =   735
         Index           =   0
         Left            =   90
         TabIndex        =   1
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
         TabIndex        =   2
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
         TabIndex        =   3
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
         TabIndex        =   4
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
         TabIndex        =   5
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
         TabIndex        =   6
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
         TabIndex        =   7
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
         TabIndex        =   8
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
         TabIndex        =   9
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
         TabIndex        =   10
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
         TabIndex        =   11
         Top             =   1890
         Width           =   1365
         _ExtentX        =   2408
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
         TabIndex        =   12
         Top             =   990
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   1296
         Caption         =   "Clear"
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
         TabIndex        =   13
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
         TabIndex        =   14
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
         TabIndex        =   30
         Top             =   90
         Width           =   1365
         _ExtentX        =   2408
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
   Begin lvButton.lvButtons_H cmdKeyUser 
      Height          =   420
      Left            =   7830
      TabIndex        =   15
      Top             =   1890
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
      TabIndex        =   22
      Top             =   3105
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   847
      Caption         =   "&Login"
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
      Image           =   "frmUnlock.frx":6F67
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdCancel 
      Height          =   480
      Left            =   7170
      TabIndex        =   23
      Top             =   3105
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   847
      Caption         =   "&Cancel"
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
      Image           =   "frmUnlock.frx":D7C9
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdKeyPass 
      Height          =   420
      Left            =   7830
      TabIndex        =   24
      Top             =   2430
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
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ">>> USER LOGIN "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   4650
      TabIndex        =   25
      Top             =   1080
      Width           =   2025
   End
   Begin VB.Image Image3 
      Height          =   600
      Left            =   4530
      Picture         =   "frmUnlock.frx":10A2B
      Top             =   960
      Width           =   3960
   End
   Begin VB.Image Image2 
      Height          =   3075
      Left            =   45
      Picture         =   "frmUnlock.frx":10E7A
      Stretch         =   -1  'True
      Top             =   990
      Width           =   4425
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Login Name"
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
      Left            =   4725
      TabIndex        =   29
      Top             =   1995
      Width           =   825
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Password"
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
      Left            =   4740
      TabIndex        =   28
      Top             =   2535
      Width           =   690
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
      TabIndex        =   27
      Top             =   4230
      Width           =   3510
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
      TabIndex        =   26
      Top             =   4200
      Width           =   1620
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
Attribute VB_Name = "frmUnlock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim dwLen                            As Long
Dim strString                        As String
Dim clsDS2                           As New clsDS2
Dim sPass                            As Byte
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdKeyPass_Click()
    sPass = 2
    Picture1.Visible = True
End Sub

Private Sub cmdKeyUser_Click()
    sPass = 1
    Picture1.Visible = True
End Sub

Private Sub cmdLogin_Click()
Dim sSQL As String

If txtUsername.Text = "" Then
    MsgBox "Username and/or Password is incorrect.Try Again!", vbExclamation
    txtUsername.SetFocus
    Exit Sub
End If

If txtPassword.Text = "" Then
    MsgBox "Username and/or Password is incorrect.Try Again!", vbExclamation
    txtPassword.SetFocus
    Exit Sub
End If

Set RS_UNLOCK = New ADODB.Recordset
If RS_UNLOCK.State = adStateOpen Then RS_UNLOCK.Close

sSQL = "SELECT * FROM sys_accounts WHERE user='" & txtUsername.Text & "' " & _
            "AND password_clear ='" & txtPassword.Text & "' " & _
            "AND app_pms_access ='1' " & _
            "AND status ='" & "active" & "'"
            
RS_UNLOCK.Open sSQL, CN, adOpenStatic, adLockReadOnly

If RS_UNLOCK.BOF Or RS_UNLOCK.EOF = True Then
    MsgBox "User Tidak Mempunyai Hak Akses..!", vbExclamation
    Exit Sub

ElseIf RS_UNLOCK.Fields("status") = "suspend" Then
    MsgBox "User account is no longer active.Contact your administrator to re-activate your account!", vbExclamation
    Exit Sub

ElseIf Not RS_UNLOCK.Fields("user") = txtUsername.Text Then
    MsgBox "Username atau Password Salah, Silahkan Coba Lagi!", vbExclamation
    Exit Sub

ElseIf Not RS_UNLOCK.Fields("password_clear") = txtPassword.Text Then
    MsgBox "Username atau Password Salah, Silahkan Coba Lagi!", vbExclamation
    Exit Sub

    
Else

    ACTIVE_ADMIN.KODEUSER = RS_UNLOCK.Fields("hrd_employee_id")
    ACTIVE_ADMIN.FULLNAME = RS_UNLOCK.Fields("name")
    ACTIVE_ADMIN.USERNAME = RS_UNLOCK.Fields("user")
    ACTIVE_ADMIN.PASSWORD = RS_UNLOCK.Fields("password_clear")
    ACTIVE_ADMIN.ISADMIN = RS_UNLOCK.Fields("app_pms_access")
    
    LOG_APP = True
    Unload Me

End If
End Sub



Private Sub cmdNumber_Click(Index As Integer)
If Index = 11 Then
    If sPass = 1 Then
        txtUsername.Text = ""
    ElseIf sPass = 2 Then
        txtPassword.Text = ""
    End If
ElseIf Index = 10 Then
    sPass = 0
    Picture1.Visible = False
ElseIf Index = 14 Then
    If sPass = 1 Then
        If Len(txtUsername.Text) = 0 Then Exit Sub
        txtUsername.Text = Mid(txtUsername.Text, 1, Len(txtUsername.Text) - 1)
    ElseIf sPass = 2 Then
        If Len(txtPassword.Text) = 0 Then Exit Sub
        txtPassword.Text = Mid(txtPassword.Text, 1, Len(txtPassword.Text) - 1)
    End If
Else
    If sPass = 1 Then
        txtUsername.Text = txtUsername & cmdNumber(Index).Caption
    ElseIf sPass = 2 Then
        txtPassword.Text = txtPassword & cmdNumber(Index).Caption
    End If
End If
End Sub

Private Sub Form_Load()
On Error GoTo ErrHandler
CenterForm frmUnlock

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


Private Sub Form_Unload(Cancel As Integer)
    Set frmUnlock = Nothing
    Set RS_UNLOCK = Nothing
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

