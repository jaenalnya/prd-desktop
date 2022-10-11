VERSION 5.00
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{8E048CF2-F435-45C9-8A6F-4646F9E1B5F4}#1.0#0"; "XPTab.ocx"
Begin VB.Form frmSettingServer 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Setting Server"
   ClientHeight    =   4290
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5445
   Icon            =   "frmSettingServer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4290
   ScaleWidth      =   5445
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtporton 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1710
      TabIndex        =   18
      Top             =   1755
      Width           =   1005
   End
   Begin prjXTab.XTab XTab1 
      Height          =   2895
      Left            =   45
      TabIndex        =   5
      Top             =   765
      Width           =   5280
      _ExtentX        =   9313
      _ExtentY        =   5106
      TabCount        =   2
      TabCaption(0)   =   "Setting Database"
      TabContCtrlCnt(0)=   10
      Tab(0)ContCtrlCap(1)=   "TxtDBServer"
      Tab(0)ContCtrlCap(2)=   "txtDBName"
      Tab(0)ContCtrlCap(3)=   "txtDBUser"
      Tab(0)ContCtrlCap(4)=   "txtDBPassword"
      Tab(0)ContCtrlCap(5)=   "txtPort"
      Tab(0)ContCtrlCap(6)=   "Label1"
      Tab(0)ContCtrlCap(7)=   "Label2"
      Tab(0)ContCtrlCap(8)=   "Label4"
      Tab(0)ContCtrlCap(9)=   "Label5"
      Tab(0)ContCtrlCap(10)=   "Label6"
      TabCaption(1)   =   "Setting Aplikasi"
      TabContCtrlCnt(1)=   3
      Tab(1)ContCtrlCap(1)=   "txtNoMachine"
      Tab(1)ContCtrlCap(2)=   "Label9"
      Tab(1)ContCtrlCap(3)=   "Label7"
      TabTheme        =   1
      ActiveTabBackStartColor=   16514555
      ActiveTabBackEndColor=   16514555
      InActiveTabBackStartColor=   16777215
      InActiveTabBackEndColor=   15397104
      BeginProperty ActiveTabFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty InActiveTabFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OuterBorderColor=   10198161
      DisabledTabBackColor=   -2147483633
      DisabledTabForeColor=   10526880
      Begin VB.TextBox txtNoMachine 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   -73335
         TabIndex        =   16
         Text            =   "1"
         Top             =   540
         Width           =   1005
      End
      Begin VB.TextBox TxtDBServer 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2250
         TabIndex        =   10
         Text            =   "192.168.150.254"
         Top             =   540
         Width           =   2670
      End
      Begin VB.TextBox txtDBName 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2250
         TabIndex        =   9
         Text            =   "PIN"
         Top             =   990
         Width           =   2670
      End
      Begin VB.TextBox txtDBUser 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2250
         TabIndex        =   8
         Text            =   "jaenal"
         Top             =   1440
         Width           =   2670
      End
      Begin VB.TextBox txtDBPassword 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   2250
         PasswordChar    =   "*"
         TabIndex        =   7
         Text            =   "1234"
         Top             =   1890
         Width           =   2670
      End
      Begin VB.TextBox txtPort 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   2250
         TabIndex        =   6
         Text            =   "1234"
         Top             =   2340
         Width           =   2670
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Port On"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -74685
         TabIndex        =   20
         Top             =   1035
         Width           =   1860
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Nomor Mesin"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -74685
         TabIndex        =   17
         Top             =   585
         Width           =   1860
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Database Server"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   315
         TabIndex        =   15
         Top             =   585
         Width           =   1680
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Database Name"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   315
         TabIndex        =   14
         Top             =   1035
         Width           =   1680
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Database User"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   315
         TabIndex        =   13
         Top             =   1485
         Width           =   1680
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Database Password"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   315
         TabIndex        =   12
         Top             =   1935
         Width           =   1860
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Port"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   315
         TabIndex        =   11
         Top             =   2385
         Width           =   1860
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   675
      Left            =   0
      ScaleHeight     =   675
      ScaleWidth      =   5355
      TabIndex        =   0
      Top             =   0
      Width           =   5355
      Begin VB.Image Image2 
         Height          =   480
         Left            =   45
         Picture         =   "frmSettingServer.frx":0EE2
         Top             =   45
         Width           =   480
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Menu ini digunakan untuk setting database"
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
         Index           =   3
         Left            =   720
         TabIndex        =   2
         Top             =   345
         Width           =   3795
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "SETTING APLIKASI"
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
         Index           =   0
         Left            =   720
         TabIndex        =   1
         Top             =   90
         Width           =   2760
      End
   End
   Begin lvButton.lvButtons_H CmdCancel 
      Height          =   420
      Left            =   3915
      TabIndex        =   3
      Top             =   3735
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   741
      Caption         =   "E&xit"
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
      Image           =   "frmSettingServer.frx":16044
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H CmdSave 
      Height          =   420
      Left            =   2430
      TabIndex        =   4
      Top             =   3735
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   741
      Caption         =   "&Save"
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
      Image           =   "frmSettingServer.frx":2CA06
      cBack           =   -2147483633
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Nomor Mesin"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   360
      TabIndex        =   19
      Top             =   1800
      Width           =   1860
   End
End
Attribute VB_Name = "frmSettingServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit
Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdSave_Click()
On Error Resume Next
    Call WriteINI("SERVER", "DBSERVER", TxtDBServer.Text, App.Path & "\Database.ini")
    Call WriteINI("SERVER", "DBNAME", txtDBName.Text, App.Path & "\Database.ini")
    Call WriteINI("SERVER", "DBUSER", txtDBUser.Text, App.Path & "\Database.ini")
    Call WriteINI("SERVER", "DBPASS", txtDBPassword.Text, App.Path & "\Database.ini")
    Call WriteINI("SERVER", "DBPORT", txtPort.Text, App.Path & "\Database.ini")
    Call WriteINI("SETTING", "MACHINE", txtNoMachine.Text, App.Path & "\Database.ini")
    Call WriteINI("SETTING", "SERIAL", txtporton.Text, App.Path & "\Database.ini")
    MsgBox "Data Sudah di setting", vbInformation
    Call LoadProduct
    Unload Me
End Sub

Private Sub Form_Load()
On Error Resume Next
    TxtDBServer.Text = ReadINI("Server", "DBSERVER", App.Path & "\Database.ini")
    txtDBName.Text = ReadINI("Server", "DBNAME", App.Path & "\Database.ini")
    txtDBUser.Text = ReadINI("Server", "DBUSER", App.Path & "\Database.ini")
    txtDBPassword.Text = ReadINI("Server", "DBPASS", App.Path & "\Database.ini")
    txtPort.Text = ReadINI("Server", "DBPORT", App.Path & "\Database.ini")
    txtNoMachine.Text = ReadINI("SETTING", "MACHINE", App.Path & "\Database.ini")
    txtporton.Text = ReadINI("SETTING", "SERIAL", App.Path & "\Database.ini")
End Sub
