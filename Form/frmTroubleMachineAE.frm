VERSION 5.00
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "LvButton.ocx"
Begin VB.Form frmTroubleMachineAE 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6270
   ClientLeft      =   5295
   ClientTop       =   2025
   ClientWidth     =   9885
   Icon            =   "frmTroubleMachineAE.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6270
   ScaleWidth      =   9885
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picHeader 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   450
      Left            =   0
      ScaleHeight     =   450
      ScaleWidth      =   9885
      TabIndex        =   32
      Top             =   0
      Width           =   9885
      Begin PRD.Liner Liner1 
         Height          =   30
         Left            =   0
         TabIndex        =   33
         Top             =   960
         Width           =   10215
         _ExtentX        =   18018
         _ExtentY        =   53
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "T R O U B L E  M A C H I N E ®™"
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
         Left            =   135
         TabIndex        =   34
         Top             =   90
         Width           =   5355
      End
   End
   Begin VB.Frame frameLogin 
      BackColor       =   &H00FFFFFF&
      Height          =   1995
      Left            =   135
      TabIndex        =   22
      Top             =   1215
      Visible         =   0   'False
      Width           =   4245
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
         Left            =   1335
         PasswordChar    =   "#"
         TabIndex        =   24
         Top             =   810
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
         Left            =   1335
         TabIndex        =   23
         Top             =   270
         Width           =   2040
      End
      Begin lvButton.lvButtons_H cmdKeyUser 
         Height          =   420
         Left            =   3435
         TabIndex        =   25
         Top             =   255
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
         Left            =   1350
         TabIndex        =   26
         Top             =   1350
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
         Image           =   "frmTroubleMachineAE.frx":6612
         cBack           =   -2147483633
      End
      Begin lvButton.lvButtons_H cmdCancel 
         Height          =   480
         Left            =   2775
         TabIndex        =   27
         Top             =   1335
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   847
         Caption         =   "&BATAL"
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
         Image           =   "frmTroubleMachineAE.frx":CE74
         cBack           =   -2147483633
      End
      Begin lvButton.lvButtons_H cmdKeyPass 
         Height          =   420
         Left            =   3435
         TabIndex        =   28
         Top             =   795
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
         Left            =   270
         TabIndex        =   30
         Top             =   360
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
         Left            =   345
         TabIndex        =   29
         Top             =   900
         Width           =   690
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2850
      Left            =   4365
      ScaleHeight     =   2820
      ScaleWidth      =   5340
      TabIndex        =   5
      Top             =   1305
      Visible         =   0   'False
      Width           =   5370
      Begin lvButton.lvButtons_H cmdNumber 
         Height          =   735
         Index           =   13
         Left            =   2925
         TabIndex        =   6
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
         Index           =   0
         Left            =   90
         TabIndex        =   7
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
         TabIndex        =   8
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
         TabIndex        =   9
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
         TabIndex        =   10
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
         TabIndex        =   11
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
         TabIndex        =   12
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
         TabIndex        =   13
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
         TabIndex        =   14
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
         TabIndex        =   15
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
         TabIndex        =   16
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
         TabIndex        =   17
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
         TabIndex        =   18
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
         TabIndex        =   19
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
         Index           =   14
         Left            =   3915
         TabIndex        =   20
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
   Begin lvButton.lvButtons_H CmdStart 
      Height          =   510
      Left            =   135
      TabIndex        =   21
      Top             =   540
      Width           =   1725
      _ExtentX        =   3043
      _ExtentY        =   900
      Caption         =   "INPUT DATA"
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
   Begin lvButton.lvButtons_H cmdExit 
      Height          =   510
      Left            =   8010
      TabIndex        =   35
      Top             =   540
      Width           =   1725
      _ExtentX        =   3043
      _ExtentY        =   900
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
      cBack           =   -2147483633
   End
   Begin VB.Frame frameCheck 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Input Data"
      Enabled         =   0   'False
      Height          =   4965
      Left            =   135
      TabIndex        =   31
      Top             =   1215
      Width           =   9600
      Begin VB.ComboBox cboShift 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "frmTroubleMachineAE.frx":100D6
         Left            =   1710
         List            =   "frmTroubleMachineAE.frx":100E3
         Style           =   2  'Dropdown List
         TabIndex        =   43
         Top             =   315
         Width           =   1275
      End
      Begin PRD.Liner Liner2 
         Height          =   30
         Left            =   90
         TabIndex        =   41
         Top             =   810
         Width           =   9285
         _ExtentX        =   16378
         _ExtentY        =   53
      End
      Begin VB.TextBox txtDescription 
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   1710
         MaxLength       =   250
         MultiLine       =   -1  'True
         TabIndex        =   4
         Top             =   4095
         Width           =   7575
      End
      Begin VB.TextBox txtResult 
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   1710
         MaxLength       =   250
         MultiLine       =   -1  'True
         TabIndex        =   3
         Top             =   3330
         Width           =   7575
      End
      Begin VB.TextBox txtAction 
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   1710
         MaxLength       =   250
         MultiLine       =   -1  'True
         TabIndex        =   2
         Top             =   2565
         Width           =   7575
      End
      Begin VB.TextBox txtAnalysis 
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   1710
         MaxLength       =   250
         MultiLine       =   -1  'True
         TabIndex        =   1
         Top             =   1800
         Width           =   7575
      End
      Begin VB.TextBox txtTrouble 
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   1710
         MaxLength       =   250
         MultiLine       =   -1  'True
         TabIndex        =   0
         Top             =   1035
         Width           =   7575
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Shift"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   135
         TabIndex        =   42
         Top             =   360
         Width           =   1545
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "DESCRIPTION / KETERANGAN"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   180
         TabIndex        =   40
         Top             =   4185
         Width           =   1275
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "RESULT / HASIL"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   180
         TabIndex        =   39
         Top             =   3420
         Width           =   1140
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "ACTION / TINDAKAN"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   180
         TabIndex        =   38
         Top             =   2655
         Width           =   1140
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "ANALYSIS / ANALISA"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   180
         TabIndex        =   37
         Top             =   1890
         Width           =   1140
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "TROUBLE / MASALAH"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   180
         TabIndex        =   36
         Top             =   1080
         Width           =   1140
      End
   End
End
Attribute VB_Name = "frmTroubleMachineAE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const MAX_COMPUTERNAME_LENGTH As Long = 31
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Dim dwLen                           As Long

Dim strString                       As String
Dim clsDS2                          As New clsDS2
Dim sPass                           As Byte
Dim sSQL                            As String

Public State                        As FORM_STATE
Public PK                           As String

Private Sub cmdCancel_Click()
    frameLogin.Visible = False
    Picture1.Visible = False
End Sub

Private Sub cmdExit_Click()
    If CmdStart.Caption = "SAVE DATA" Then
        MsgBox "Silahkan SAVE DATA jika mau Exit..!!", vbExclamation
        Exit Sub
    Else
        Unload Me
    End If
    
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
On Error GoTo ErrHandler

If txtUsername.text = "" Then
    MsgBox "Username and/or Password is incorrect.Try Again!", vbExclamation
    txtUsername.SetFocus
    Exit Sub
End If

If txtPassword.text = "" Then
    MsgBox "Username and/or Password is incorrect.Try Again!", vbExclamation
    txtPassword.SetFocus
    Exit Sub
End If


sSQL = "SELECT a.*,b.nik,b.pin FROM sys_accounts a"
sSQL = sSQL & " LEFT JOIN hrd_employees b ON a.hrd_employee_id = b.id"
sSQL = sSQL & " WHERE a.user = '" & txtUsername.text & "' AND a.password_clear = '" & txtPassword.text & "' AND a.status ='" & "active" & "'"

Set RS_USER = New ADODB.Recordset
If RS_USER.State = adStateOpen Then RS_USER.Close
RS_USER.Open sSQL, CN, adOpenStatic, adLockReadOnly

If RS_USER.BOF Or RS_USER.EOF = True Then
    MsgBox "Username atau Password Salah, Silahkan Coba Lagi!", vbExclamation
    Exit Sub

ElseIf RS_USER.Fields("status") = "suspend" Then
    MsgBox "User account is no longer active.Contact your administrator to re-activate your account!", vbExclamation
    Exit Sub

ElseIf Not RS_USER.Fields("user") = txtUsername.text Then
    MsgBox "Username atau Password Salah, Silahkan Coba Lagi!", vbExclamation
    Exit Sub

ElseIf Not RS_USER.Fields("password_clear") = txtPassword.text Then
    MsgBox "Username atau Password Salah, Silahkan Coba Lagi!", vbExclamation
    Exit Sub

Else

    frameCheck.Enabled = True
    CmdStart.Caption = "SAVE DATA"

    frameLogin.Visible = False
    Picture1.Visible = False

    txtUsername.text = ""
    txtPassword.text = ""



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

Private Sub CmdStart_Click()

    If CmdStart.Caption = "INPUT DATA" Then
        frameLogin.Visible = True
        
    Else

        If cboShift.text = "" Then
            MsgBox "Shift harus di isi, silahkan cek kembali!", vbExclamation
            Exit Sub
        End If
    
        Call UpdateData

        frameCheck.Enabled = False
        CmdStart.Caption = "INPUT DATA"
        frameLogin.Visible = False
        Picture1.Visible = False
    End If


    
End Sub

Private Sub Form_Load()
On Error GoTo ErrHandler
CenterForm frmTroubleMachine

If State = AddStateMode Then
    Me.Caption = "Buat Baru"

    sSQL = "SELECT * FROM sip_production.prod_trouble_machines "

    Set RS_TROUBLE = New ADODB.Recordset
    If RS_TROUBLE.State = adStateOpen Then RS_TROUBLE.Close
    RS_TROUBLE.Open sSQL, CN, adOpenDynamic, adLockOptimistic
    
ElseIf State = EditStateMode Then
    Me.Caption = "Ubah Data"
    
    sSQL = "SELECT * FROM sip_production.prod_trouble_machines WHERE id = '" & PK & "'"

    
    Set RS_TROUBLE = New ADODB.Recordset
    If RS_TROUBLE.State = adStateOpen Then RS_TROUBLE.Close
    RS_TROUBLE.Open sSQL, CN, adOpenDynamic, adLockOptimistic
    
    If RS_TROUBLE.RecordCount > 0 Then
        With RS_TROUBLE
            cboShift.text = .Fields("shift")
            txtTrouble.text = .Fields("trouble")
            txtAnalysis.text = .Fields("analysis")
            txtAction.text = .Fields("action")
            txtResult.text = .Fields("result")
            txtDescription.text = .Fields("description")
        End With
    Else
        MsgBox "Tidak record data di database", vbExclamation
    End If

End If

Exit Sub
ErrHandler:
    MsgBox "Error Number: " & Err.Number & vbCrLf & "Description: " & Err.Description, vbExclamation
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
    frmTroubleMachine.CommandPass "Refresh"
    Set frmTroubleMachine = Nothing
    Set RS_TROUBLE = Nothing
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


Private Sub UpdateData()
  
On Error GoTo ErrHandler

    
    If State = AddStateMode Then

        RS_TROUBLE.AddNew

        RS_TROUBLE.Fields("plant_mark") = p_plant_mark
        RS_TROUBLE.Fields("product_id") = p_eng_product_1
        RS_TROUBLE.Fields("prod_machine_id") = p_prod_machine_id
        RS_TROUBLE.Fields("sys_plant_id") = p_sys_plant_id
        RS_TROUBLE.Fields("period_shift") = Format(p_shift, "yyyy-mm-dd")

        RS_TROUBLE.Fields("shift") = cboShift.text
        RS_TROUBLE.Fields("Trouble") = txtTrouble.text
        RS_TROUBLE.Fields("Analysis") = txtAnalysis.text
        RS_TROUBLE.Fields("Action") = txtAction.text
        RS_TROUBLE.Fields("Result") = txtResult.text
        RS_TROUBLE.Fields("Description") = txtDescription.text
        RS_TROUBLE.Fields("status") = "active"
        RS_TROUBLE.Fields("created_at") = Format(Now, "yyyy-mm-dd hh:mm:ss")
        RS_TROUBLE.Fields("created_by") = RS_USER.Fields("id")

        RS_TROUBLE.Update

        MsgBox "Data baru berhasil disimpan!", vbInformation

    ElseIf State = EditStateMode Then
        RS_TROUBLE.Fields("plant_mark") = p_plant_mark
        RS_TROUBLE.Fields("product_id") = p_eng_product_1
        RS_TROUBLE.Fields("prod_machine_id") = p_prod_machine_id
        RS_TROUBLE.Fields("sys_plant_id") = p_sys_plant_id
        RS_TROUBLE.Fields("period_shift") = Format(p_shift, "yyyy-mm-dd")

        RS_TROUBLE.Fields("shift") = cboShift.text
        RS_TROUBLE.Fields("Trouble") = txtTrouble.text
        RS_TROUBLE.Fields("Analysis") = txtAnalysis.text
        RS_TROUBLE.Fields("Action") = txtAction.text
        RS_TROUBLE.Fields("Result") = txtResult.text
        RS_TROUBLE.Fields("Description") = txtDescription.text
        RS_TROUBLE.Fields("status") = "active"
        RS_TROUBLE.Fields("created_at") = Format(Now, "yyyy-mm-dd hh:mm:ss")
        RS_TROUBLE.Fields("created_by") = RS_USER.Fields("id")

        RS_TROUBLE.Update

        MsgBox "Data berhasil disimpan!", vbInformation
    End If
Exit Sub
ErrHandler:
    MsgBox "Error Number: " & Err.Number & vbCrLf & " Description : " & Err.Description, vbExclamation, Me.Caption
    
    
End Sub

