VERSION 5.00
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "LvButton.ocx"
Object = "{8E048CF2-F435-45C9-8A6F-4646F9E1B5F4}#1.0#0"; "XPTab.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmInformasi 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Informasi "
   ClientHeight    =   10695
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   18150
   Icon            =   "frmInformasi.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10695
   ScaleWidth      =   18150
   StartUpPosition =   2  'CenterScreen
   Begin lvButton.lvButtons_H cmdExit 
      Height          =   465
      Left            =   8280
      TabIndex        =   13
      Top             =   10035
      Width           =   2265
      _ExtentX        =   3995
      _ExtentY        =   820
      Caption         =   "EXIT"
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
      Image           =   "frmInformasi.frx":617A
      cBack           =   -2147483633
   End
   Begin prjXTab.XTab XTab1 
      Height          =   9150
      Left            =   135
      TabIndex        =   2
      Top             =   810
      Width           =   17925
      _ExtentX        =   31618
      _ExtentY        =   16140
      TabCount        =   2
      TabCaption(0)   =   "TEXT"
      TabContCtrlCnt(0)=   9
      Tab(0)ContCtrlCap(1)=   "lblInfo1"
      Tab(0)ContCtrlCap(2)=   "lblInfo2"
      Tab(0)ContCtrlCap(3)=   "lblInfo3"
      Tab(0)ContCtrlCap(4)=   "lblInfo4"
      Tab(0)ContCtrlCap(5)=   "lblInfo5"
      Tab(0)ContCtrlCap(6)=   "lblInfo6"
      Tab(0)ContCtrlCap(7)=   "lblInfo7"
      Tab(0)ContCtrlCap(8)=   "lblInfo8"
      Tab(0)ContCtrlCap(9)=   "lblInfo9"
      TabCaption(1)   =   "IMAGE"
      TabContCtrlCnt(1)=   1
      Tab(1)ContCtrlCap(1)=   "Picture1"
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
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00FFFFFF&
         Height          =   8610
         Left            =   -74865
         ScaleHeight     =   8550
         ScaleWidth      =   17595
         TabIndex        =   12
         Top             =   450
         Width           =   17655
         Begin VB.Image Image1 
            Height          =   8475
            Left            =   45
            Stretch         =   -1  'True
            Top             =   45
            Width           =   17475
         End
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Konsep 5R"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   21.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Index           =   1
         Left            =   315
         TabIndex        =   11
         Top             =   630
         Width           =   16305
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00FFFFFF&
         Caption         =   "1. Ringkas"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   21.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Index           =   2
         Left            =   315
         TabIndex        =   10
         Top             =   1350
         Width           =   16305
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00FFFFFF&
         Caption         =   "2. Rapi"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   21.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Index           =   3
         Left            =   315
         TabIndex        =   9
         Top             =   2070
         Width           =   16305
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00FFFFFF&
         Caption         =   "3. Resik"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   21.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Index           =   4
         Left            =   315
         TabIndex        =   8
         Top             =   2790
         Width           =   16305
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00FFFFFF&
         Caption         =   "4. Rawat"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   21.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Index           =   5
         Left            =   315
         TabIndex        =   7
         Top             =   3510
         Width           =   16305
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00FFFFFF&
         Caption         =   "5. Rajin"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   21.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Index           =   6
         Left            =   315
         TabIndex        =   6
         Top             =   4230
         Width           =   16305
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   21.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Index           =   7
         Left            =   315
         TabIndex        =   5
         Top             =   4950
         Width           =   16305
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   21.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Index           =   8
         Left            =   315
         TabIndex        =   4
         Top             =   5670
         Width           =   16305
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   21.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Index           =   9
         Left            =   315
         TabIndex        =   3
         Top             =   6390
         Width           =   16305
      End
   End
   Begin MSComDlg.CommonDialog dlgDialog 
      Left            =   0
      Top             =   9675
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   0
      Top             =   10035
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "KLIK SEMBARANG UNTUK KELUAR MENU INFORMASI INI"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   13230
      TabIndex        =   1
      Top             =   10395
      Width           =   4875
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      BackColor       =   &H000040C0&
      Caption         =   "INFORMASI PENTING "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   645
      Index           =   0
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   18105
   End
End
Attribute VB_Name = "frmInformasi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim X As Long
Dim Y As Long


Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub Form_Click()
    Unload Me
End Sub

Private Sub Form_Load()
On Error Resume Next


Dim Rs As New Recordset
Dim sSQL As String

    Rs.CursorLocation = adUseClient

    Rs.Open "Select * From prod_informations", CN, adOpenStatic, adLockOptimistic

    If Rs.RecordCount > 0 Then
        lblInfo(0).Caption = Rs.Fields("info_header")
        lblInfo(1).Caption = Rs.Fields("info_1")
        lblInfo(2).Caption = Rs.Fields("info_2")
        lblInfo(3).Caption = Rs.Fields("info_3")
        lblInfo(4).Caption = Rs.Fields("info_4")
        lblInfo(5).Caption = Rs.Fields("info_5")
        lblInfo(6).Caption = Rs.Fields("info_6")
        lblInfo(7).Caption = Rs.Fields("info_7")
        lblInfo(8).Caption = Rs.Fields("info_8")
        lblInfo(9).Caption = Rs.Fields("info_9")
        
    End If
    
    Set Rs = Nothing
    
    Timer1.Enabled = True
    If p_plant_mark = "TSSI" Then
        UpdateGambar App.Path & "\informasi.jpg", "\\192.168.131.253\pmsupdate$\informasi.jpg"
    ElseIf p_plant_mark = "Techno" Then
        UpdateGambar App.Path & "\informasi.jpg", "\\192.168.121.251\pmsupdate\informasi.jpg"
    ElseIf p_plant_mark = "Cempaka" Then
        UpdateGambar App.Path & "\informasi.jpg", "\\192.168.151.250\pmsupdate\informasi.jpg"
    End If
    
    Image1.Picture = LoadPicture(App.Path & "\informasi.jpg")
         
End Sub

Private Sub lblInfo_Click(Index As Integer)
If Index = 0 Or Index = 1 Or Index = 2 Or Index = 3 Or Index = 4 Or Index = 5 Or _
    Index = 6 Or Index = 7 Or Index = 8 Or Index = 9 Then
    Unload Me
End If
End Sub

Private Sub Timer1_Timer()
    X = X + 1
        If X >= 60 Then
        Y = Y + 1
        X = 0
    End If
    
    If Y >= 1 Then
        Y = 0
        Unload Me
    End If
End Sub
