VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmProduction 
   Caption         =   "Production"
   ClientHeight    =   10815
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   19080
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmWi.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   10815
   ScaleWidth      =   19080
   Begin VB.Frame Frame1 
      Height          =   5025
      Left            =   2700
      TabIndex        =   23
      Top             =   720
      Width           =   8895
   End
   Begin lvButton.lvButtons_H cmdPS 
      Height          =   660
      Left            =   2700
      TabIndex        =   20
      ToolTipText     =   "Packing Standard"
      Top             =   45
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   1164
      Caption         =   "PS - L"
      CapAlign        =   2
      BackStyle       =   6
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFHover         =   0
      cBhover         =   33023
      LockHover       =   3
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      Image           =   "frmWi.frx":617A
      cBack           =   16777215
   End
   Begin VB.TextBox txtOk_2 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   1350
      TabIndex        =   17
      Text            =   "0"
      Top             =   4725
      Width           =   1230
   End
   Begin VB.TextBox txtNg_2 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   1350
      TabIndex        =   14
      Text            =   "0"
      Top             =   3600
      Width           =   1230
   End
   Begin VB.TextBox txtgross_2 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   1350
      TabIndex        =   12
      Text            =   "0"
      Top             =   2475
      Width           =   1230
   End
   Begin VB.TextBox txtOk_1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   45
      TabIndex        =   7
      Text            =   "0"
      Top             =   4725
      Width           =   1230
   End
   Begin VB.TextBox txtNg_1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   45
      TabIndex        =   6
      Text            =   "0"
      Top             =   3600
      Width           =   1230
   End
   Begin VB.TextBox txtgross_1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   45
      TabIndex        =   5
      Text            =   "0"
      Top             =   2475
      Width           =   1230
   End
   Begin VB.TextBox txtShot_1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   84
      TabIndex        =   3
      Text            =   "0"
      Top             =   1125
      Width           =   2508
   End
   Begin lvButton.lvButtons_H cmdAddIdle 
      Height          =   645
      Left            =   10125
      TabIndex        =   1
      Top             =   45
      Width           =   1590
      _ExtentX        =   2805
      _ExtentY        =   1138
      Caption         =   " IDLE TIME"
      CapAlign        =   2
      BackStyle       =   6
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFHover         =   0
      cBhover         =   33023
      LockHover       =   3
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      Image           =   "frmWi.frx":C304
      cBack           =   16777215
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   4000
      Left            =   270
      Top             =   5670
   End
   Begin lvButton.lvButtons_H cmdClose 
      Height          =   630
      Left            =   17145
      TabIndex        =   0
      Top             =   45
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1111
      Caption         =   "&CLOSE [ESC]"
      CapAlign        =   2
      BackStyle       =   6
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFHover         =   0
      cBhover         =   33023
      LockHover       =   3
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      Image           =   "frmWi.frx":1248E
      cBack           =   16777215
   End
   Begin lvButton.lvButtons_H cmdAddNG 
      Height          =   645
      Left            =   11880
      TabIndex        =   2
      Top             =   45
      Width           =   1590
      _ExtentX        =   2805
      _ExtentY        =   1138
      Caption         =   " NOT GOOD (NG)"
      CapAlign        =   2
      BackStyle       =   6
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFHover         =   0
      cBhover         =   33023
      LockHover       =   3
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      Image           =   "frmWi.frx":156F0
      cBack           =   16777215
   End
   Begin lvButton.lvButtons_H cmdWI 
      Height          =   660
      Left            =   90
      TabIndex        =   21
      ToolTipText     =   "Work Instruction"
      Top             =   45
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   1164
      Caption         =   "WI - L"
      CapAlign        =   2
      BackStyle       =   6
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFHover         =   0
      cBhover         =   33023
      LockHover       =   3
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      Image           =   "frmWi.frx":1B87A
      cBack           =   16777215
   End
   Begin lvButton.lvButtons_H cmdCP 
      Height          =   660
      Left            =   5310
      TabIndex        =   22
      ToolTipText     =   "Critical Point Check"
      Top             =   45
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   1164
      Caption         =   "CPC - L"
      CapAlign        =   2
      BackStyle       =   6
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFHover         =   0
      cBhover         =   33023
      LockHover       =   3
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      Image           =   "frmWi.frx":21A04
      cBack           =   16777215
   End
   Begin lvButton.lvButtons_H cmdPS2 
      Height          =   660
      Left            =   3960
      TabIndex        =   24
      ToolTipText     =   "Packing Standard"
      Top             =   45
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   1164
      Caption         =   "PS - R"
      CapAlign        =   2
      BackStyle       =   6
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFHover         =   0
      cBhover         =   33023
      LockHover       =   3
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      Image           =   "frmWi.frx":27B8E
      cBack           =   16777215
   End
   Begin lvButton.lvButtons_H cmdCP2 
      Height          =   660
      Left            =   6570
      TabIndex        =   25
      ToolTipText     =   "Critical Point Check"
      Top             =   45
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   1164
      Caption         =   "CPC - R"
      CapAlign        =   2
      BackStyle       =   6
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFHover         =   0
      cBhover         =   33023
      LockHover       =   3
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      Image           =   "frmWi.frx":2DD18
      cBack           =   16777215
   End
   Begin lvButton.lvButtons_H cmdWI2 
      Height          =   660
      Left            =   1350
      TabIndex        =   26
      ToolTipText     =   "Work Instruction"
      Top             =   45
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   1164
      Caption         =   "WI - R"
      CapAlign        =   2
      BackStyle       =   6
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFHover         =   0
      cBhover         =   33023
      LockHover       =   3
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      Image           =   "frmWi.frx":33EA2
      cBack           =   16777215
   End
   Begin PRD.InactiveTimer itmrClose 
      Left            =   180
      Top             =   8595
      _ExtentX        =   847
      _ExtentY        =   847
      Enabled         =   0   'False
   End
   Begin lvButton.lvButtons_H cmdHasilProd 
      Height          =   645
      Left            =   13635
      TabIndex        =   27
      Top             =   45
      Width           =   1590
      _ExtentX        =   2805
      _ExtentY        =   1138
      Caption         =   " HASIL PRODUKSI"
      CapAlign        =   2
      BackStyle       =   6
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFHover         =   0
      cBhover         =   33023
      LockHover       =   3
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      Image           =   "frmWi.frx":3A02C
      cBack           =   16777215
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   345
      Left            =   0
      TabIndex        =   28
      Top             =   10470
      Width           =   19080
      _ExtentX        =   33655
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   12
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   529
            MinWidth        =   529
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   2117
            MinWidth        =   2117
            Text            =   "ACTIVE USER:"
            TextSave        =   "ACTIVE USER:"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   19041
            MinWidth        =   4410
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   529
            MinWidth        =   529
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "Today:"
            TextSave        =   "Today:"
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "3/25/2019"
         EndProperty
         BeginProperty Panel8 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "13:50"
         EndProperty
         BeginProperty Panel9 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   3
            Enabled         =   0   'False
            Object.Width           =   882
            MinWidth        =   882
            TextSave        =   "INS"
         EndProperty
         BeginProperty Panel10 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Object.Width           =   882
            MinWidth        =   882
            TextSave        =   "NUM"
         EndProperty
         BeginProperty Panel11 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Enabled         =   0   'False
            Object.Width           =   882
            MinWidth        =   882
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel12 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   4
            Enabled         =   0   'False
            Object.Width           =   882
            MinWidth        =   882
            TextSave        =   "SCRL"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin lvButton.lvButtons_H cmdChangeUser 
      Height          =   645
      Left            =   15390
      TabIndex        =   29
      Top             =   45
      Width           =   1590
      _ExtentX        =   2805
      _ExtentY        =   1138
      Caption         =   "CHANGE USER"
      CapAlign        =   2
      BackStyle       =   6
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFHover         =   0
      cBhover         =   33023
      LockHover       =   3
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      Image           =   "frmWi.frx":401B6
      cBack           =   16777215
   End
   Begin lvButton.lvButtons_H cmdDashboard 
      Height          =   645
      Left            =   8370
      TabIndex        =   30
      Top             =   45
      Width           =   1590
      _ExtentX        =   2805
      _ExtentY        =   1138
      Caption         =   "DASHBOARD"
      CapAlign        =   2
      BackStyle       =   6
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFHover         =   0
      cBhover         =   33023
      LockHover       =   3
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      Image           =   "frmWi.frx":46340
      cBack           =   16777215
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "PROD-R"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   1350
      TabIndex        =   19
      Top             =   4455
      Width           =   1230
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "PROD-L"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   45
      TabIndex        =   18
      Top             =   4455
      Width           =   1230
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      Caption         =   "OK PRODUCT"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   45
      TabIndex        =   10
      Top             =   4170
      Width           =   2535
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      Caption         =   "DATA GROSS"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   45
      TabIndex        =   8
      Top             =   1935
      Width           =   2535
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      Caption         =   "NG PRODUCT"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   45
      TabIndex        =   9
      Top             =   3060
      Width           =   2535
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "PROD-R"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   1350
      TabIndex        =   16
      Top             =   3330
      Width           =   1230
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "PROD-L"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   45
      TabIndex        =   15
      Top             =   3330
      Width           =   1230
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "PROD-R"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   1350
      TabIndex        =   13
      Top             =   2205
      Width           =   1230
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "PROD-L"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   45
      TabIndex        =   11
      Top             =   2205
      Width           =   1230
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      Caption         =   "SHOT MACHINE"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   90
      TabIndex        =   4
      Top             =   810
      Width           =   2505
   End
End
Attribute VB_Name = "frmProduction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_objPDF As AcroPDFLibCtl.AcroPDF 'Declare an object of type AcroPDF
Private m_strFilePath As String 'Declare a string for the PDF Filename and Path
Private Sub cmdAddIdle_Click()
    frmIdleTime.Show 1
End Sub

Private Sub cmdAddNG_Click()
    frmNg.Show 1
End Sub

Private Sub cmdChangeUser_Click()
    If MsgBox("Please log out first before switching to another user.Proceed?", vbQuestion + vbYesNo) = vbNo Then Exit Sub

    MAIN.StatusBar.Panels(3).Text = vbNullString
    MAIN.StatusBar.Panels(4).Text = vbNullString
    
    ACTIVE_USER.USERNAME = vbNullString
    ACTIVE_USER.USERTYPE = vbNullString
    'Unload Me
    frmLogin.Show 1
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdCP_Click()
On Error GoTo ErrHandler
    m_strFilePath = App.Path & "\Picture\CCP\" & p_int_part_1 & ".pdf"
    
    m_objPDF.LoadFile m_strFilePath
    m_objPDF.setShowToolbar False
    m_objPDF.setLayoutMode "SinglePage"
    m_objPDF.setPageMode "none"
    
    'Set the Zoom view according to the value specified. ranges from 0 and onwards
    m_objPDF.setZoom 80
    
    'Move and Resize the object in relation to its container/form
    With m_objPDF
       '.Move 125, 175, Me.Width, Me.Height 'x-position, y-position, width, height
       .Move 125, 175, Frame1.Width - 300, Frame1.Height - 300 'x-position, y-position, width, height
    End With
    
    m_objPDF.Visible = True
    
Exit Sub
ErrHandler:
    MsgBox "Error Number: " & Err.Number & vbCrLf & "Description: " & Err.Description, vbExclamation
End Sub

Private Sub cmdCP2_Click()
On Error GoTo ErrHandler
    m_strFilePath = App.Path & "\Picture\CCP\" & p_int_part_2 & ".pdf"
    
    m_objPDF.LoadFile m_strFilePath
    m_objPDF.setShowToolbar False
    m_objPDF.setLayoutMode "SinglePage"
    m_objPDF.setPageMode "none"
    
    'Set the Zoom view according to the value specified. ranges from 0 and onwards
    m_objPDF.setZoom 80
    
    'Move and Resize the object in relation to its container/form
    With m_objPDF
       '.Move 125, 175, Me.Width, Me.Height 'x-position, y-position, width, height
       .Move 125, 175, Frame1.Width - 300, Frame1.Height - 300 'x-position, y-position, width, height
    End With
    
    m_objPDF.Visible = True


Exit Sub
ErrHandler:
    MsgBox "Error Number: " & Err.Number & vbCrLf & "Description: " & Err.Description, vbExclamation
End Sub

Private Sub cmdDashboard_Click()
    LoadForm frmDashboard
End Sub

Private Sub cmdHasilProd_Click()
    frmInputResult.Show 1
End Sub

Private Sub cmdPS_Click()
On Error GoTo ErrHandler
    m_strFilePath = App.Path & "\Picture\PS\" & p_int_part_1 & ".pdf"
    
    m_objPDF.LoadFile m_strFilePath
    m_objPDF.setShowToolbar False
    m_objPDF.setLayoutMode "SinglePage"
    m_objPDF.setPageMode "none"
    
    'Set the Zoom view according to the value specified. ranges from 0 and onwards
    m_objPDF.setZoom 55
    
    'Move and Resize the object in relation to its container/form
    With m_objPDF
       '.Move 125, 175, Me.Width, Me.Height 'x-position, y-position, width, height
       .Move 125, 175, Frame1.Width - 300, Frame1.Height - 300 'x-position, y-position, width, height
    End With
    
    m_objPDF.Visible = True
    
    itmrClose.InactiveInterval = 1000 * 10
    itmrClose.Enabled = True
    
Exit Sub
ErrHandler:
    MsgBox "Error Number: " & Err.Number & vbCrLf & "Description: " & Err.Description, vbExclamation
End Sub

Private Sub LoadPDF()
    m_objPDF.LoadFile m_strFilePath
    m_objPDF.setShowToolbar False
    m_objPDF.setLayoutMode "SinglePage"
    m_objPDF.setPageMode "none"
    
    'Set the Zoom view according to the value specified. ranges from 0 and onwards
    m_objPDF.setZoom 80
    
    'Move and Resize the object in relation to its container/form
    With m_objPDF
       '.Move 125, 175, Me.Width, Me.Height 'x-position, y-position, width, height
       .Move 125, 175, Frame1.Width - 300, Frame1.Height - 300 'x-position, y-position, width, height
    End With
    
    m_objPDF.Visible = True
End Sub

Private Sub cmdPS2_Click()
On Error GoTo ErrHandler
    m_strFilePath = App.Path & "\Picture\PS\" & p_int_part_2 & ".pdf"
    
    m_objPDF.LoadFile m_strFilePath
    m_objPDF.setShowToolbar False
    m_objPDF.setLayoutMode "SinglePage"
    m_objPDF.setPageMode "none"
    
    'Set the Zoom view according to the value specified. ranges from 0 and onwards
    m_objPDF.setZoom 55
    
    'Move and Resize the object in relation to its container/form
    With m_objPDF
       '.Move 125, 175, Me.Width, Me.Height 'x-position, y-position, width, height
       .Move 125, 175, Frame1.Width - 300, Frame1.Height - 300 'x-position, y-position, width, height
    End With
    
    m_objPDF.Visible = True
    
    itmrClose.InactiveInterval = 1000 * 10
    itmrClose.Enabled = True

    itmrClose.InactiveInterval = 1000 * 10
    itmrClose.Enabled = True
    
Exit Sub
ErrHandler:
    MsgBox "Error Number: " & Err.Number & vbCrLf & "Description: " & Err.Description, vbExclamation
End Sub

Private Sub cmdWI_Click()
On Error GoTo ErrHandler
    m_strFilePath = App.Path & "\Picture\WI\" & p_int_part_1 & ".pdf"

    m_objPDF.LoadFile m_strFilePath
    m_objPDF.setShowToolbar False
    m_objPDF.setLayoutMode "SinglePage"
    m_objPDF.setPageMode "none"
    
    'Set the Zoom view according to the value specified. ranges from 0 and onwards
    m_objPDF.setZoom 85
    
    'Move and Resize the object in relation to its container/form
    With m_objPDF
       '.Move 125, 175, Me.Width, Me.Height 'x-position, y-position, width, height
       .Move 125, 175, Frame1.Width - 300, Frame1.Height - 300 'x-position, y-position, width, height
    End With
    
    m_objPDF.Visible = True
    

    itmrClose.InactiveInterval = 1000 * 10
    itmrClose.Enabled = True
    
Exit Sub
ErrHandler:
    MsgBox "Error Number: " & Err.Number & vbCrLf & "Description: " & Err.Description, vbExclamation
End Sub

Private Sub cmdWI2_Click()
On Error GoTo ErrHandler
    m_strFilePath = App.Path & "\Picture\WI\" & p_int_part_2 & ".pdf"
    
    m_objPDF.LoadFile m_strFilePath
    m_objPDF.setShowToolbar False
    m_objPDF.setLayoutMode "SinglePage"
    m_objPDF.setPageMode "none"
    
    'Set the Zoom view according to the value specified. ranges from 0 and onwards
    m_objPDF.setZoom 85
    
    'Move and Resize the object in relation to its container/form
    With m_objPDF
       '.Move 125, 175, Me.Width, Me.Height 'x-position, y-position, width, height
       .Move 125, 175, Frame1.Width - 300, Frame1.Height - 300 'x-position, y-position, width, height
    End With
    
    m_objPDF.Visible = True

    itmrClose.InactiveInterval = 1000 * 10
    itmrClose.Enabled = True
    
Exit Sub
ErrHandler:
    MsgBox "Error Number: " & Err.Number & vbCrLf & "Description: " & Err.Description, vbExclamation
End Sub

Private Sub Form_Load()

   m_strFilePath = App.Path & "\Picture\CCP\" & p_int_part_1 & ".pdf" 'Change this to the path and filename of your PDF File
   Set m_objPDF = Controls.Add("AcroPDF.PDF.1", "AcroPDF1") 'This will add the PDF Browser control to the form on runtime. The "Test" is the control's name
   Set m_objPDF.Container = Frame1 'Attach the PDF Browser control to a container.

   
    Summary p_eng_product_1, txtgross_1, txtNg_1, txtOk_1
    Summary p_eng_product_2, txtgross_2, txtNg_2, txtOk_2

    With frmProduction.StatusBar.Panels
        .Item(3).Text = ACTIVE_USER.FULLNAME
        .Item(4).Text = ACTIVE_USER.USERNAME
    End With
    
End Sub

Private Sub Summary(sProd As String, tGross As TextBox, tNG As TextBox, tOK As TextBox)
    Dim Rs As New Recordset
    Dim sSQL As String
    Dim p_shift As Date
    If Format(Now, "HH") >= 0 And Format(Now, "HH") <= 7 Then
        p_shift = Format(DateAdd("d", -1, Format(Now, "yyyy-mm-dd")), "yyyy-mm-dd")
    Else
        p_shift = Format(Now, "yyyy-mm-dd")
    End If
    
    Rs.CursorLocation = adUseClient
    
    sSQL = "select data_ok.plant_mark,data_ok.prod_machine_id,data_ok.eng_product_id,data_ok.period_shift,"
    sSQL = sSQL & " data_ok.shot,data_ok.gross,IFNULL(data_ng.shot_ng,0) as shot_ng,(data_ok.gross-IFNULL(data_ng.shot_ng,0)) as total_ok"
    sSQL = sSQL & " from (select a.plant_mark,a.prod_machine_id,a.eng_product_id,a.period_shift,"
    sSQL = sSQL & " sum(a.counter_ok) as shot, sum(a.counter_ok * b.cavity) as gross"
    sSQL = sSQL & " from sip_production.prod_runnings a inner join sip_234.eng_products b on a.eng_product_id = b.id"
    sSQL = sSQL & " group by a.plant_mark,a.prod_machine_id,a.eng_product_id,a.period_shift) data_ok"
    sSQL = sSQL & " left join (select x.plant_mark,x.prod_machine_id,x.eng_product_id,x.period_shift,"
    sSQL = sSQL & " count(x.prod_ng_id) as shot_ng from sip_production.prod_ng_logs x"
    sSQL = sSQL & " group by x.plant_mark,x.prod_machine_id,x.eng_product_id,x.period_shift) data_ng"
    sSQL = sSQL & " on data_ok.plant_mark = data_ng.plant_mark and data_ok.prod_machine_id = data_ng.prod_machine_id"
    sSQL = sSQL & " and data_ok.eng_product_id = data_ng.eng_product_id and data_ok.period_shift = data_ng.period_shift"
    sSQL = sSQL & " where data_ok.plant_mark = '" & p_plant_mark & "' and"
    sSQL = sSQL & " data_ok.prod_machine_id = '" & p_prod_machine_id & "' and"
    sSQL = sSQL & " data_ok.eng_product_id = '" & sProd & "' and"
    sSQL = sSQL & " data_ok.period_shift = '" & Format(p_shift, "yyyy-mm-dd") & "'"

    Rs.Open sSQL, CN, adOpenStatic, adLockOptimistic
    
    If Rs.RecordCount > 0 Then
        txtShot_1 = IIf(IsNull(Rs.Fields("shot")), "0", Rs.Fields("shot"))
        tGross.Text = IIf(IsNull(Rs.Fields("gross")), "0", Rs.Fields("gross"))
        tNG.Text = IIf(IsNull(Rs.Fields("shot_ng")), "0", Rs.Fields("shot_ng"))
        tOK.Text = IIf(IsNull(Rs.Fields("total_ok")), "0", Rs.Fields("total_ok"))
    End If
    Set Rs = Nothing
End Sub

Private Sub Form_Resize()
On Error Resume Next
    If WindowState <> vbMinimized Then
        If Me.Width < 9195 Then Me.Width = 9195
        If Me.Height < 4500 Then Me.Height = 4500

        'Move and Resize the object in relation to its container/form
        Frame1.Width = Me.ScaleWidth - 3000
        Frame1.Height = Me.ScaleHeight - 1200
    End If
    Timer1.Enabled = True
    
End Sub




Private Sub itmrClose_UserInactive()
    cmdCP_Click
    itmrClose.Enabled = False
End Sub



Private Sub Timer1_Timer()
On Error Resume Next
    
    Summary p_eng_product_1, txtgross_1, txtNg_1, txtOk_1
    Summary p_eng_product_2, txtgross_2, txtNg_2, txtOk_2
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
 
    'The Browser control will load a blank
    m_objPDF.LoadFile ""
    
    'Set object to nothing
    Set m_objPDF = Nothing
    'MAIN.RemoveChild Me.Name
    Set frmProduction = Nothing
End Sub

Private Sub Form_Activate()
On Error Resume Next

With MAIN
    Me.BackColor = .ACPMenu.BackColor
    Frame1.BackColor = .ACPMenu.BackColor
    Me.Picture = .ACPMenu.LoadBackground
End With

Call LoadPDF
   
'MAIN.ActivateChild Me
End Sub

