VERSION 5.00
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmChangePassword 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3390
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5985
   Icon            =   "frmChangePassword.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3390
   ScaleWidth      =   5985
   StartUpPosition =   3  'Windows Default
   Begin PRD.Liner Liner2 
      Height          =   30
      Left            =   120
      TabIndex        =   13
      Top             =   2760
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   53
   End
   Begin VB.PictureBox picHeader 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   765
      Left            =   0
      ScaleHeight     =   765
      ScaleWidth      =   5985
      TabIndex        =   10
      Top             =   0
      Width           =   5985
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "CHANGE PASSWORD"
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
         TabIndex        =   12
         Top             =   150
         Width           =   2175
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Note:Please fill all required parameters."
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
         Left            =   960
         TabIndex        =   11
         Top             =   390
         Width           =   2895
      End
      Begin VB.Image Image1 
         Height          =   720
         Left            =   75
         Picture         =   "frmChangePassword.frx":617A
         Top             =   30
         Width           =   720
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
      Height          =   330
      Left            =   2235
      TabIndex        =   0
      Top             =   885
      Width           =   3500
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
      Height          =   330
      IMEMode         =   3  'DISABLE
      Left            =   2235
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1290
      Width           =   3500
   End
   Begin VB.TextBox txtNewPassword 
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
      Height          =   330
      IMEMode         =   3  'DISABLE
      Left            =   2235
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1695
      Width           =   3500
   End
   Begin VB.TextBox txtRetype 
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
      Height          =   330
      IMEMode         =   3  'DISABLE
      Left            =   2235
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   2100
      Width           =   3500
   End
   Begin lvButton.lvButtons_H cmdSave 
      Height          =   390
      Left            =   3240
      TabIndex        =   4
      Top             =   2880
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   688
      Caption         =   "&Update"
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
      Height          =   390
      Left            =   4560
      TabIndex        =   5
      Top             =   2880
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   688
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
      cBack           =   -2147483633
   End
   Begin VB.Label lblForm 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Login Name"
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
      Left            =   225
      TabIndex        =   9
      Top             =   885
      Width           =   960
   End
   Begin VB.Label lblForm 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Old Password"
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
      Left            =   240
      TabIndex        =   8
      Top             =   1290
      Width           =   1095
   End
   Begin VB.Label lblForm 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "New Password"
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
      Left            =   225
      TabIndex        =   7
      Top             =   1695
      Width           =   1200
   End
   Begin VB.Label lblForm 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Re-type Password"
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
      Left            =   210
      TabIndex        =   6
      Top             =   2100
      Width           =   1470
   End
End
Attribute VB_Name = "frmChangePassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public State                        As FORM_STATE
Dim srcSQL                          As String

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdSave_Click()
Dim obj As Control
Dim sRS As Recordset
            For Each obj In Me
            If TypeOf obj Is TextBox Or TypeOf obj Is ComboBox Then
                If obj.Text = "" Then
                    MsgBox obj.Name & " could not be left blank. Please complete the field.", vbExclamation, Me.Caption
                    obj.SetFocus
                    Exit Sub
                End If
            End If
            Next obj
            
            If txtNewPassword.Text <> txtRetype.Text Then
                MsgBox "Passwords did not match.Please check it!", vbExclamation
                Exit Sub
            End If
            
            Set sRS = New ADODB.Recordset
            If sRS.State = adStateOpen Then sRS.Close
            sRS.Open "SELECT * FROM sys_accounts WHERE user='" & ACTIVE_USER.USERNAME & "'", CN, adOpenStatic, adLockReadOnly
            
            If sRS.Fields("password_clear") <> txtPassword.Text Then
                MsgBox "Password did not match.Please check it!", vbExclamation
                Exit Sub
            End If
            
            If State = EditStateMode Then
                Set RS_USER = New ADODB.Recordset
                sSQL_Update "UPDATE sys_accounts SET sys_accounts.password_clear='" & txtNewPassword.Text & "' WHERE sys_accounts.user='" & ACTIVE_USER.USERNAME & "'"
                
                MsgBox "Password has been successfully updated!", vbInformation
                Unload Me
            End If
End Sub

Private Sub Form_Activate()
On Error Resume Next
    Dim i As Integer
    Me.BackColor = MAIN.ACPMenu.BackColor
    
    If MAIN.ACPMenu.Theme = 0 Then
        For i = 0 To 3
        lblForm(i).ForeColor = &HFFFFFF
        Next i
    Else
        For i = 0 To 3
        lblForm(i).ForeColor = &H0&
        Next i
    End If
txtPassword.SetFocus
End Sub

Private Sub Form_Load()
On Error GoTo ErrHandler
CenterForm frmChangePassword

State = EditStateMode

    srcSQL = "SELECT sys_accounts.* " & _
            "FROM sys_accounts " & _
            "WHERE (((sys_accounts.user)='" & ACTIVE_USER.USERNAME & "'))"

    Set RS_USER = New ADODB.Recordset
    If RS_USER.State = adStateOpen Then RS_USER.Close
    RS_USER.Open srcSQL, CN, adOpenDynamic, adLockOptimistic
    
    With RS_USER
        txtUsername.Text = .Fields("user")
    End With

Exit Sub
ErrHandler:
    MsgBox "Error Number: " & Err.Number & vbCrLf & "Description: " & Err.Description, vbExclamation, Me.Caption
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
Set frmChangePassword = Nothing
Set RS_USER = Nothing
End Sub


Private Sub txtUsername_GotFocus()
HLText txtUsername
End Sub


