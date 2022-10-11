VERSION 5.00
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmCall 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Call QC / MTC"
   ClientHeight    =   3045
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   5295
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmCall.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3045
   ScaleWidth      =   5295
   StartUpPosition =   2  'CenterScreen
   Begin PRD.Liner Liner1 
      Height          =   30
      Left            =   90
      TabIndex        =   8
      Top             =   2205
      Width           =   5190
      _ExtentX        =   9155
      _ExtentY        =   53
   End
   Begin lvButton.lvButtons_H cmdExit 
      Height          =   555
      Left            =   3510
      TabIndex        =   7
      Top             =   2340
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   979
      Caption         =   "&EXIT"
      CapAlign        =   2
      BackStyle       =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFore           =   16777215
      cFHover         =   16777215
      cBhover         =   33023
      LockHover       =   1
      cGradient       =   65535
      Mode            =   0
      Value           =   0   'False
      Image           =   "frmCall.frx":617A
      cBack           =   4210752
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   1485
      Picture         =   "frmCall.frx":1CB3C
      Top             =   1125
      Width           =   480
   End
   Begin VB.Label lblTimeStop 
      BackStyle       =   0  'Transparent
      Caption         =   "00:00:00"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   330
      Left            =   2205
      TabIndex        =   6
      Top             =   1215
      Width           =   2715
   End
   Begin VB.Label lblTimeStart 
      BackStyle       =   0  'Transparent
      Caption         =   "00:00:00"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   330
      Left            =   2205
      TabIndex        =   5
      Top             =   765
      Width           =   2895
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "TIME STOP :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   225
      TabIndex        =   4
      Top             =   765
      Width           =   1770
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "NAMA IDLE :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   0
      Left            =   225
      TabIndex        =   3
      Top             =   225
      Width           =   1770
   End
   Begin VB.Label lblIdleName 
      BackStyle       =   0  'Transparent
      Caption         =   "CALL QC"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   2205
      TabIndex        =   2
      Top             =   225
      Width           =   1815
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "LOSS TIME :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   225
      TabIndex        =   1
      Top             =   1620
      Width           =   1770
   End
   Begin VB.Label lblLossTime 
      BackStyle       =   0  'Transparent
      Caption         =   "00:00:00"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   330
      Left            =   2205
      TabIndex        =   0
      Top             =   1620
      Width           =   2715
   End
End
Attribute VB_Name = "frmCall"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdExit_Click()
    Unload Me
End Sub
