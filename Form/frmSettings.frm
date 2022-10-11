VERSION 5.00
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "LvButton.ocx"
Object = "{8E048CF2-F435-45C9-8A6F-4646F9E1B5F4}#1.0#0"; "XPTab.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmSettings 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Setting Parameter"
   ClientHeight    =   9810
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   16665
   Icon            =   "frmSettings.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9810
   ScaleWidth      =   16665
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cboWaktu 
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
      ItemData        =   "frmSettings.frx":617A
      Left            =   1395
      List            =   "frmSettings.frx":618A
      Style           =   2  'Dropdown List
      TabIndex        =   45
      Top             =   9315
      Width           =   870
   End
   Begin prjXTab.XTab XTab1 
      Height          =   8430
      Left            =   45
      TabIndex        =   6
      Top             =   765
      Width           =   16530
      _ExtentX        =   29157
      _ExtentY        =   14870
      TabCaption(0)   =   "Setting Informasi"
      TabContCtrlCnt(0)=   20
      Tab(0)ContCtrlCap(1)=   "txtInfo9"
      Tab(0)ContCtrlCap(2)=   "txtInfo8"
      Tab(0)ContCtrlCap(3)=   "txtInfo7"
      Tab(0)ContCtrlCap(4)=   "txtInfo6"
      Tab(0)ContCtrlCap(5)=   "txtInfo5"
      Tab(0)ContCtrlCap(6)=   "txtInfo4"
      Tab(0)ContCtrlCap(7)=   "txtInfo3"
      Tab(0)ContCtrlCap(8)=   "txtInfo2"
      Tab(0)ContCtrlCap(9)=   "txtInfo1"
      Tab(0)ContCtrlCap(10)=   "txtInfo0"
      Tab(0)ContCtrlCap(11)=   "lblInfo9"
      Tab(0)ContCtrlCap(12)=   "lblInfo8"
      Tab(0)ContCtrlCap(13)=   "lblInfo7"
      Tab(0)ContCtrlCap(14)=   "lblInfo6"
      Tab(0)ContCtrlCap(15)=   "lblInfo5"
      Tab(0)ContCtrlCap(16)=   "lblInfo4"
      Tab(0)ContCtrlCap(17)=   "lblInfo3"
      Tab(0)ContCtrlCap(18)=   "lblInfo2"
      Tab(0)ContCtrlCap(19)=   "lblInfo1"
      Tab(0)ContCtrlCap(20)=   "lblInfo0"
      TabCaption(1)   =   "Setting Mesin"
      TabContCtrlCnt(1)=   19
      Tab(1)ContCtrlCap(1)=   "cboRelay"
      Tab(1)ContCtrlCap(2)=   "Check7"
      Tab(1)ContCtrlCap(3)=   "cboMesin"
      Tab(1)ContCtrlCap(4)=   "Check1"
      Tab(1)ContCtrlCap(5)=   "txtSensor"
      Tab(1)ContCtrlCap(6)=   "txtWI"
      Tab(1)ContCtrlCap(7)=   "txtPS"
      Tab(1)ContCtrlCap(8)=   "txtCCP"
      Tab(1)ContCtrlCap(9)=   "Check2"
      Tab(1)ContCtrlCap(10)=   "Check3"
      Tab(1)ContCtrlCap(11)=   "Check4"
      Tab(1)ContCtrlCap(12)=   "Check5"
      Tab(1)ContCtrlCap(13)=   "Check6"
      Tab(1)ContCtrlCap(14)=   "Label5"
      Tab(1)ContCtrlCap(15)=   "Label1"
      Tab(1)ContCtrlCap(16)=   "Label2"
      Tab(1)ContCtrlCap(17)=   "Label40"
      Tab(1)ContCtrlCap(18)=   "Label41"
      Tab(1)ContCtrlCap(19)=   "Label42"
      TabCaption(2)   =   "Gambar Informasi"
      TabContCtrlCnt(2)=   4
      Tab(2)ContCtrlCap(1)=   "cmdSaveImage"
      Tab(2)ContCtrlCap(2)=   "dlgOpenFile"
      Tab(2)ContCtrlCap(3)=   "cmdSetImage"
      Tab(2)ContCtrlCap(4)=   "Picture1"
      TabTheme        =   1
      ActiveTabBackStartColor=   16514555
      ActiveTabBackEndColor=   16514555
      InActiveTabBackStartColor=   16777215
      InActiveTabBackEndColor=   15397104
      BeginProperty ActiveTabFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
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
      Begin lvButton.lvButtons_H cmdSaveImage 
         Height          =   420
         Left            =   -73515
         TabIndex        =   51
         Top             =   7920
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   741
         Caption         =   "Simpan Gambar"
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
         cBack           =   -2147483633
      End
      Begin MSComDlg.CommonDialog dlgOpenFile 
         Left            =   -59205
         Top             =   7875
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin lvButton.lvButtons_H cmdSetImage 
         Height          =   420
         Left            =   -74865
         TabIndex        =   50
         Top             =   7920
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   741
         Caption         =   "Cari "
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
         cBack           =   -2147483633
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00FFFFFF&
         Height          =   7305
         Left            =   -74865
         ScaleHeight     =   7245
         ScaleWidth      =   16110
         TabIndex        =   49
         Top             =   540
         Width           =   16170
         Begin VB.Image Image2 
            Height          =   7125
            Left            =   45
            Stretch         =   -1  'True
            Top             =   45
            Width           =   15990
         End
      End
      Begin VB.ComboBox cboRelay 
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
         ItemData        =   "frmSettings.frx":619C
         Left            =   -73065
         List            =   "frmSettings.frx":619E
         Style           =   2  'Dropdown List
         TabIndex        =   47
         Top             =   3465
         Width           =   1230
      End
      Begin VB.CheckBox Check7 
         Caption         =   "Enable Show Information"
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
         Left            =   -71490
         TabIndex        =   46
         Top             =   3465
         Width           =   2580
      End
      Begin VB.TextBox txtInfo 
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
         Height          =   375
         Index           =   9
         Left            =   315
         MaxLength       =   65
         TabIndex        =   42
         Top             =   7740
         Width           =   6270
      End
      Begin VB.TextBox txtInfo 
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
         Height          =   375
         Index           =   8
         Left            =   315
         MaxLength       =   65
         TabIndex        =   40
         Top             =   6975
         Width           =   6270
      End
      Begin VB.TextBox txtInfo 
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
         Height          =   375
         Index           =   7
         Left            =   315
         MaxLength       =   65
         TabIndex        =   38
         Top             =   6210
         Width           =   6270
      End
      Begin VB.TextBox txtInfo 
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
         Height          =   375
         Index           =   6
         Left            =   315
         MaxLength       =   65
         TabIndex        =   36
         Top             =   5445
         Width           =   6270
      End
      Begin VB.TextBox txtInfo 
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
         Height          =   375
         Index           =   5
         Left            =   315
         MaxLength       =   65
         TabIndex        =   34
         Top             =   4725
         Width           =   6270
      End
      Begin VB.TextBox txtInfo 
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
         Height          =   375
         Index           =   4
         Left            =   315
         MaxLength       =   65
         TabIndex        =   32
         Top             =   3960
         Width           =   6270
      End
      Begin VB.TextBox txtInfo 
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
         Height          =   375
         Index           =   3
         Left            =   315
         MaxLength       =   65
         TabIndex        =   30
         Top             =   3195
         Width           =   6270
      End
      Begin VB.TextBox txtInfo 
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
         Height          =   375
         Index           =   2
         Left            =   315
         MaxLength       =   65
         TabIndex        =   28
         Top             =   2430
         Width           =   6270
      End
      Begin VB.TextBox txtInfo 
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
         Height          =   375
         Index           =   1
         Left            =   315
         MaxLength       =   65
         TabIndex        =   26
         Top             =   1710
         Width           =   6270
      End
      Begin VB.TextBox txtInfo 
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
         Height          =   375
         Index           =   0
         Left            =   315
         MaxLength       =   50
         TabIndex        =   24
         Text            =   "INFORMASI PENTING"
         Top             =   990
         Width           =   6270
      End
      Begin VB.ComboBox cboMesin 
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
         Left            =   -73065
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   810
         Width           =   1230
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Auto Close Menu NG"
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
         Left            =   -71490
         TabIndex        =   20
         Top             =   810
         Width           =   2580
      End
      Begin VB.TextBox txtSensor 
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
         Left            =   -73065
         TabIndex        =   15
         Text            =   "95"
         Top             =   1215
         Width           =   1230
      End
      Begin VB.TextBox txtWI 
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
         Left            =   -73065
         TabIndex        =   14
         Text            =   "95"
         Top             =   1665
         Width           =   1230
      End
      Begin VB.TextBox txtPS 
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
         Left            =   -73065
         TabIndex        =   13
         Text            =   "95"
         Top             =   2115
         Width           =   1230
      End
      Begin VB.TextBox txtCCP 
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
         Left            =   -73065
         TabIndex        =   12
         Text            =   "95"
         Top             =   2565
         Width           =   1230
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Auto Idle"
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
         Left            =   -71490
         TabIndex        =   11
         Top             =   1215
         Width           =   2580
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Show Sensor Information"
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
         Left            =   -71490
         TabIndex        =   10
         Top             =   1665
         Width           =   2580
      End
      Begin VB.CheckBox Check4 
         Caption         =   "Show Working Instruction"
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
         Left            =   -71490
         TabIndex        =   9
         Top             =   2115
         Width           =   2580
      End
      Begin VB.CheckBox Check5 
         Caption         =   "Enable Toolbar Menu"
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
         Left            =   -71490
         TabIndex        =   8
         Top             =   2565
         Width           =   2580
      End
      Begin VB.CheckBox Check6 
         Caption         =   "Enable Select Machine"
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
         Left            =   -71490
         TabIndex        =   7
         Top             =   3015
         Width           =   2580
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "ID RELAY"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   -74370
         TabIndex        =   48
         Top             =   3510
         Width           =   1095
      End
      Begin VB.Label lblInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "Baris - 9 (max 65 character)"
         Height          =   240
         Index           =   9
         Left            =   315
         TabIndex        =   41
         Top             =   7470
         Width           =   3300
      End
      Begin VB.Label lblInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "Baris - 8 (max 65 character)"
         Height          =   240
         Index           =   8
         Left            =   315
         TabIndex        =   39
         Top             =   6705
         Width           =   3300
      End
      Begin VB.Label lblInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "Baris - 7 (max 65 character)"
         Height          =   240
         Index           =   7
         Left            =   315
         TabIndex        =   37
         Top             =   5940
         Width           =   3300
      End
      Begin VB.Label lblInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "Baris - 6 (max 65 character)"
         Height          =   240
         Index           =   6
         Left            =   315
         TabIndex        =   35
         Top             =   5175
         Width           =   3300
      End
      Begin VB.Label lblInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "Baris - 5 (max 65 character)"
         Height          =   240
         Index           =   5
         Left            =   315
         TabIndex        =   33
         Top             =   4455
         Width           =   3300
      End
      Begin VB.Label lblInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "Baris - 4 (max 65 character)"
         Height          =   240
         Index           =   4
         Left            =   315
         TabIndex        =   31
         Top             =   3690
         Width           =   3300
      End
      Begin VB.Label lblInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "Baris - 3 (max 65 character)"
         Height          =   240
         Index           =   3
         Left            =   315
         TabIndex        =   29
         Top             =   2925
         Width           =   3300
      End
      Begin VB.Label lblInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "Baris - 2  (max 65 character)"
         Height          =   240
         Index           =   2
         Left            =   315
         TabIndex        =   27
         Top             =   2160
         Width           =   3300
      End
      Begin VB.Label lblInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "Baris - 1  (max 65 character)"
         Height          =   240
         Index           =   1
         Left            =   315
         TabIndex        =   25
         Top             =   1440
         Width           =   3300
      End
      Begin VB.Label lblInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "Baris Judul Informasi (max 50 chracter)"
         Height          =   240
         Index           =   0
         Left            =   315
         TabIndex        =   23
         Top             =   720
         Width           =   3300
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Machine No"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   -74370
         TabIndex        =   22
         Top             =   855
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Sensor ID"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   -74370
         TabIndex        =   19
         Top             =   1260
         Width           =   1095
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Size WI"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   -74370
         TabIndex        =   18
         Top             =   1710
         Width           =   1095
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Size PS"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   -74370
         TabIndex        =   17
         Top             =   2160
         Width           =   1095
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Size CCP"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   -74370
         TabIndex        =   16
         Top             =   2610
         Width           =   1095
      End
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
      ScaleWidth      =   16665
      TabIndex        =   0
      Top             =   0
      Width           =   16665
      Begin VB.PictureBox Liner1 
         Height          =   30
         Left            =   0
         ScaleHeight     =   30
         ScaleWidth      =   9465
         TabIndex        =   1
         Top             =   945
         Width           =   9465
      End
      Begin VB.Image Image1 
         Height          =   720
         Left            =   90
         Picture         =   "frmSettings.frx":61A0
         Top             =   45
         Width           =   720
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Menu ini digunakan untuk merubah menu aplikasi "
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
         Left            =   765
         TabIndex        =   3
         Top             =   360
         Width           =   3840
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "SETTING PARAMETER"
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
         Left            =   765
         TabIndex        =   2
         Top             =   135
         Width           =   3495
      End
   End
   Begin lvButton.lvButtons_H cmdSave 
      Height          =   435
      Left            =   13905
      TabIndex        =   4
      Top             =   9270
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   767
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
      Image           =   "frmSettings.frx":C31A
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdCancel 
      Height          =   435
      Left            =   15345
      TabIndex        =   5
      Top             =   9270
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   767
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
      Image           =   "frmSettings.frx":124A4
      cBack           =   -2147483633
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Minute"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   4
      Left            =   2430
      TabIndex        =   44
      Top             =   9360
      Width           =   690
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Waktu Show"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   3
      Left            =   135
      TabIndex        =   43
      Top             =   9360
      Width           =   1095
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Rs As New Recordset
Dim sSQL As String
    
    
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
On Error GoTo ErrHandler
    Call WriteINI("SETTING", "MACHINE", cboMesin.Text, App.Path & "\Settings.ini")
    Call WriteINI("SETTING", "RELAY", cboRelay.Text, App.Path & "\Settings.ini")
    Call WriteINI("SETTING", "SERIAL", txtSensor.Text, App.Path & "\Settings.ini")
    Call WriteINI("SETTING", "WI", txtWI.Text, App.Path & "\Settings.ini")
    Call WriteINI("SETTING", "PS", txtPS.Text, App.Path & "\Settings.ini")
    Call WriteINI("SETTING", "CCP", txtCCP.Text, App.Path & "\Settings.ini")
    
    Call WriteINI("SETTING", "NGAUTOCLOSE", Check1.Value, App.Path & "\Settings.ini")
    Call WriteINI("SETTING", "IDLEON", Check2.Value, App.Path & "\Settings.ini")
    Call WriteINI("SETTING", "SHOWSENSOR", Check3.Value, App.Path & "\Settings.ini")
    Call WriteINI("SETTING", "SHOWWI", Check4.Value, App.Path & "\Settings.ini")
    Call WriteINI("SETTING", "STOOLBAR", Check5.Value, App.Path & "\Settings.ini")
    Call WriteINI("SETTING", "ENMACHINE", Check6.Value, App.Path & "\Settings.ini")
    
    Call WriteINI("SETTING", "INFO", cboWaktu.Text, App.Path & "\Settings.ini")
    Call WriteINI("SETTING", "SHOWINFO", Check7.Value, App.Path & "\Settings.ini")

    Rs.CursorLocation = adUseClient

    Rs.Open "Select * From prod_informations", CN, adOpenStatic, adLockOptimistic

    If Rs.RecordCount < 1 Then
        sSQL = "insert into prod_informations (info_header,info_1,info_2,info_3,info_4,info_5,"
        sSQL = sSQL & " info_6,info_7,info_8,info_9,created_at,created_by) values"
        sSQL = sSQL & " ('" & txtInfo(0).Text & "','" & txtInfo(1).Text & "','" & txtInfo(2).Text & "','" & txtInfo(3).Text & "',"
        sSQL = sSQL & " '" & txtInfo(4).Text & "','" & txtInfo(5).Text & "','" & txtInfo(6).Text & "','" & txtInfo(7).Text & "',"
        sSQL = sSQL & " '" & txtInfo(8).Text & "','" & txtInfo(9).Text & "','" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "','" & ACTIVE_USER.KODEUSER & "')"
        
        sSQL_Insert sSQL
        
    Else
        sSQL = "update prod_informations set info_header = '" & txtInfo(0).Text & "',"
        sSQL = sSQL & " info_1 = '" & txtInfo(1).Text & "',"
        sSQL = sSQL & " info_2 = '" & txtInfo(2).Text & "',"
        sSQL = sSQL & " info_3 = '" & txtInfo(3).Text & "',"
        sSQL = sSQL & " info_4 = '" & txtInfo(4).Text & "',"
        sSQL = sSQL & " info_5 = '" & txtInfo(5).Text & "',"
        sSQL = sSQL & " info_6 = '" & txtInfo(6).Text & "',"
        sSQL = sSQL & " info_7 = '" & txtInfo(7).Text & "',"
        sSQL = sSQL & " info_8 = '" & txtInfo(8).Text & "',"
        sSQL = sSQL & " info_9 = '" & txtInfo(9).Text & "',"
        sSQL = sSQL & " updated_at = '" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "',"
        sSQL = sSQL & " updated_by = '" & ACTIVE_USER.KODEUSER & "'"
        
        sSQL_Update sSQL
    End If
    
    Set Rs = Nothing

    MsgBox "Data Berhasil diSimpan", vbExclamation
    Unload Me

Exit Sub
ErrHandler:
    MsgBox "Error Number: " & Err.Number & vbCrLf & "Description: " & Err.Description, vbExclamation


End Sub

Private Sub cmdSaveImage_Click()
    If p_plant_mark = "TSSI" Then
        UploadGambar dlgOpenFile.FileName, "\\192.168.131.253\pmsupdate$\informasi.jpg"
    ElseIf p_plant_mark = "Techno" Then
        UploadGambar dlgOpenFile.FileName, "\\192.168.121.251\pmsupdate\informasi.jpg"
    ElseIf p_plant_mark = "Cempaka" Then
        UploadGambar dlgOpenFile.FileName, "\\192.168.151.250\pmsupdate\informasi.jpg"
    End If

End Sub

Private Sub cmdSetImage_Click()
Dim fnum As Integer

    On Error Resume Next
    dlgOpenFile.ShowOpen
    If Err.Number = cdlCancel Then
        ' The user canceled.
        Exit Sub
    ElseIf Err.Number <> 0 Then
        ' Unknown error.
        MsgBox "Error " & Format$(Err.Number) & _
            " selecting file." & vbCrLf & _
            Err.Description
        Exit Sub
    End If
    On Error GoTo 0

    ' Read the file.
    Image2.Picture = LoadPicture(dlgOpenFile.FileName)
    
End Sub

Private Sub Form_Load()
On Error GoTo ErrHandler
    Dim i As Integer
    For i = 1 To 40
        cboMesin.AddItem i
    Next i
    
    cboRelay.AddItem "HURTM"
    cboRelay.AddItem "BITFT"

    cboMesin.Text = ReadINI("SETTING", "MACHINE", App.Path & "\Settings.ini")
    
    txtSensor.Text = ReadINI("SETTING", "SERIAL", App.Path & "\Settings.ini")
    txtWI.Text = ReadINI("SETTING", "WI", App.Path & "\Settings.ini")
    txtPS.Text = ReadINI("SETTING", "PS", App.Path & "\Settings.ini")
    txtCCP.Text = ReadINI("SETTING", "CCP", App.Path & "\Settings.ini")
    
    Check1.Value = ReadINI("SETTING", "NGAUTOCLOSE", App.Path & "\Settings.ini")
    Check2.Value = ReadINI("SETTING", "IDLEON", App.Path & "\Settings.ini")
    Check3.Value = ReadINI("SETTING", "SHOWSENSOR", App.Path & "\Settings.ini")
    Check4.Value = ReadINI("SETTING", "SHOWWI", App.Path & "\Settings.ini")
    Check5.Value = ReadINI("SETTING", "STOOLBAR", App.Path & "\Settings.ini")
    Check6.Value = ReadINI("SETTING", "ENMACHINE", App.Path & "\Settings.ini")
    Check7.Value = ReadINI("SETTING", "SHOWINFO", App.Path & "\Settings.ini")

    Rs.CursorLocation = adUseClient

    Rs.Open "Select * From prod_informations", CN, adOpenStatic, adLockOptimistic

    If Rs.RecordCount > 0 Then
        txtInfo(0).Text = Rs.Fields("info_header")
        txtInfo(1).Text = Rs.Fields("info_1")
        txtInfo(2).Text = Rs.Fields("info_2")
        txtInfo(3).Text = Rs.Fields("info_3")
        txtInfo(4).Text = Rs.Fields("info_4")
        txtInfo(5).Text = Rs.Fields("info_5")
        txtInfo(6).Text = Rs.Fields("info_6")
        txtInfo(7).Text = Rs.Fields("info_7")
        txtInfo(8).Text = Rs.Fields("info_8")
        txtInfo(9).Text = Rs.Fields("info_9")
        
    End If
    
    Set Rs = Nothing
    
    If ReadINI("SETTING", "INFO", App.Path & "\Settings.ini") <> "" Then
        cboWaktu.Text = ReadINI("SETTING", "INFO", App.Path & "\Settings.ini")
    Else
        cboWaktu.ListIndex = 1
    End If
    
    cboRelay.Text = ReadINI("SETTING", "RELAY", App.Path & "\Settings.ini")

    dlgOpenFile.InitDir = App.Path
    dlgOpenFile.Filter = "JPEG image (*.jpg)|*.jpg|All Files (*.*)|*.*"
    dlgOpenFile.FilterIndex = 1
    dlgOpenFile.DialogTitle = "Open File"
    dlgOpenFile.Flags = _
        cdlOFNFileMustExist + _
        cdlOFNHideReadOnly + _
        cdlOFNLongNames + _
        cdlOFNExplorer
    dlgOpenFile.CancelError = True
    
    
 Exit Sub
ErrHandler:
    MsgBox "Error Number: " & Err.Number & vbCrLf & "Description: " & Err.Description, vbExclamation

 
 
End Sub
