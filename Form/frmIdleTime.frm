VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "LvButton.ocx"
Object = "{8E048CF2-F435-45C9-8A6F-4646F9E1B5F4}#1.0#0"; "XPTab.ocx"
Begin VB.Form frmIdleTime 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "INPUT DATA IDLE MESIN"
   ClientHeight    =   9090
   ClientLeft      =   1245
   ClientTop       =   1500
   ClientWidth     =   17640
   ControlBox      =   0   'False
   Icon            =   "frmIdleTime.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9090
   ScaleWidth      =   17640
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtDescription 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   6120
      TabIndex        =   78
      Top             =   8325
      Width           =   11355
   End
   Begin VB.Timer Timer3 
      Interval        =   10
      Left            =   14580
      Top             =   315
   End
   Begin prjXTab.XTab XTab1 
      Height          =   8160
      Left            =   90
      TabIndex        =   49
      Top             =   765
      Width           =   5955
      _ExtentX        =   10504
      _ExtentY        =   14393
      TabCount        =   4
      TabCaption(0)   =   "MACHINE"
      TabContCtrlCnt(0)=   9
      Tab(0)ContCtrlCap(1)=   "frameProd"
      Tab(0)ContCtrlCap(2)=   "cmdIdle26"
      Tab(0)ContCtrlCap(3)=   "cmdIdle22"
      Tab(0)ContCtrlCap(4)=   "cmdIdle21"
      Tab(0)ContCtrlCap(5)=   "cmdIdle9"
      Tab(0)ContCtrlCap(6)=   "cmdIdle8"
      Tab(0)ContCtrlCap(7)=   "cmdIdle5"
      Tab(0)ContCtrlCap(8)=   "cmdIdle4"
      Tab(0)ContCtrlCap(9)=   "cmdIdle1"
      TabCaption(1)   =   "METHODE"
      TabContCtrlCnt(1)=   8
      Tab(1)ContCtrlCap(1)=   "cmdIdle25"
      Tab(1)ContCtrlCap(2)=   "cmdIdle24"
      Tab(1)ContCtrlCap(3)=   "cmdIdle23"
      Tab(1)ContCtrlCap(4)=   "cmdIdle20"
      Tab(1)ContCtrlCap(5)=   "cmdIdle19"
      Tab(1)ContCtrlCap(6)=   "cmdIdle18"
      Tab(1)ContCtrlCap(7)=   "cmdIdle17"
      Tab(1)ContCtrlCap(8)=   "cmdIdle10"
      TabCaption(2)   =   "MAN"
      TabContCtrlCnt(2)=   4
      Tab(2)ContCtrlCap(1)=   "cmdIdle11"
      Tab(2)ContCtrlCap(2)=   "cmdIdle7"
      Tab(2)ContCtrlCap(3)=   "cmdIdle6"
      Tab(2)ContCtrlCap(4)=   "cmdIdle2"
      TabCaption(3)   =   "MATERIAL"
      TabContCtrlCnt(3)=   6
      Tab(3)ContCtrlCap(1)=   "cmdIdle16"
      Tab(3)ContCtrlCap(2)=   "cmdIdle15"
      Tab(3)ContCtrlCap(3)=   "cmdIdle14"
      Tab(3)ContCtrlCap(4)=   "cmdIdle13"
      Tab(3)ContCtrlCap(5)=   "cmdIdle12"
      Tab(3)ContCtrlCap(6)=   "cmdIdle3"
      ActiveTabHeight =   30
      InActiveTabHeight=   23
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
      XRadius         =   20
      PictureSize     =   1
      Begin lvButton.lvButtons_H cmdIdle 
         Height          =   555
         Index           =   16
         Left            =   -73785
         TabIndex        =   77
         Top             =   5805
         Width           =   3120
         _ExtentX        =   5503
         _ExtentY        =   979
         Caption         =   "[P] KEKURANGAN PACKING"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFore           =   0
         cFHover         =   0
         cBhover         =   33023
         LockHover       =   1
         cGradient       =   16512
         Mode            =   0
         Value           =   0   'False
         cBack           =   16777215
      End
      Begin lvButton.lvButtons_H cmdIdle 
         Height          =   555
         Index           =   15
         Left            =   -73785
         TabIndex        =   76
         Top             =   4815
         Width           =   3120
         _ExtentX        =   5503
         _ExtentY        =   979
         Caption         =   "[O] KEKURANGAN MATERIAL"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFore           =   0
         cFHover         =   0
         cBhover         =   33023
         LockHover       =   1
         cGradient       =   16512
         Mode            =   0
         Value           =   0   'False
         cBack           =   16777215
      End
      Begin lvButton.lvButtons_H cmdIdle 
         Height          =   555
         Index           =   14
         Left            =   -73785
         TabIndex        =   75
         Top             =   3825
         Width           =   3120
         _ExtentX        =   5503
         _ExtentY        =   979
         Caption         =   "[N] PROBLEM SUPPLY MTRL"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFore           =   0
         cFHover         =   0
         cBhover         =   33023
         LockHover       =   1
         cGradient       =   16512
         Mode            =   0
         Value           =   0   'False
         cBack           =   16777215
      End
      Begin lvButton.lvButtons_H cmdIdle 
         Height          =   555
         Index           =   13
         Left            =   -73785
         TabIndex        =   74
         Top             =   2880
         Width           =   3120
         _ExtentX        =   5503
         _ExtentY        =   979
         Caption         =   "[M] PROBLEM MTRL PROSES"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFore           =   0
         cFHover         =   0
         cBhover         =   33023
         LockHover       =   1
         cGradient       =   16512
         Mode            =   0
         Value           =   0   'False
         cBack           =   16777215
      End
      Begin lvButton.lvButtons_H cmdIdle 
         Height          =   555
         Index           =   12
         Left            =   -73785
         TabIndex        =   73
         Top             =   1935
         Width           =   3120
         _ExtentX        =   5503
         _ExtentY        =   979
         Caption         =   "[L] PROBLEM MATERIAL"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFore           =   0
         cFHover         =   0
         cBhover         =   33023
         LockHover       =   1
         cGradient       =   16512
         Mode            =   0
         Value           =   0   'False
         cBack           =   16777215
      End
      Begin lvButton.lvButtons_H cmdIdle 
         Height          =   555
         Index           =   3
         Left            =   -73785
         TabIndex        =   72
         Top             =   990
         Width           =   3120
         _ExtentX        =   5503
         _ExtentY        =   979
         Caption         =   "[C] PERSIAPAN MTRL AWAL"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFore           =   0
         cFHover         =   0
         cBhover         =   33023
         LockHover       =   1
         cGradient       =   16512
         Mode            =   0
         Value           =   0   'False
         cBack           =   16777215
      End
      Begin lvButton.lvButtons_H cmdIdle 
         Height          =   555
         Index           =   11
         Left            =   -73695
         TabIndex        =   71
         Top             =   3780
         Width           =   3120
         _ExtentX        =   5503
         _ExtentY        =   979
         Caption         =   "[K] MAN POWER"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFore           =   0
         cFHover         =   0
         cBhover         =   33023
         LockHover       =   1
         cGradient       =   16512
         Mode            =   0
         Value           =   0   'False
         cBack           =   16777215
      End
      Begin lvButton.lvButtons_H cmdIdle 
         Height          =   555
         Index           =   7
         Left            =   -73695
         TabIndex        =   70
         Top             =   2835
         Width           =   3120
         _ExtentX        =   5503
         _ExtentY        =   979
         Caption         =   "[G] STICKING KARENA MAN"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFore           =   0
         cFHover         =   0
         cBhover         =   33023
         LockHover       =   1
         cGradient       =   16512
         Mode            =   0
         Value           =   0   'False
         cBack           =   16777215
      End
      Begin lvButton.lvButtons_H cmdIdle 
         Height          =   555
         Index           =   6
         Left            =   -73695
         TabIndex        =   69
         Top             =   1890
         Width           =   3120
         _ExtentX        =   5503
         _ExtentY        =   979
         Caption         =   "[F] MOLD PROBLEM MAN"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFore           =   0
         cFHover         =   0
         cBhover         =   33023
         LockHover       =   1
         cGradient       =   16512
         Mode            =   0
         Value           =   0   'False
         cBack           =   16777215
      End
      Begin lvButton.lvButtons_H cmdIdle 
         Height          =   555
         Index           =   2
         Left            =   -73695
         TabIndex        =   68
         Top             =   945
         Width           =   3120
         _ExtentX        =   5503
         _ExtentY        =   979
         Caption         =   "[B] SETUP AWAL"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFore           =   0
         cFHover         =   0
         cBhover         =   33023
         LockHover       =   1
         cGradient       =   16512
         Mode            =   0
         Value           =   0   'False
         cBack           =   16777215
      End
      Begin lvButton.lvButtons_H cmdIdle 
         Height          =   555
         Index           =   25
         Left            =   -73740
         TabIndex        =   67
         Top             =   7020
         Width           =   3120
         _ExtentX        =   5503
         _ExtentY        =   979
         Caption         =   "[Y] TRIAL KHUSUS"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFore           =   0
         cFHover         =   0
         cBhover         =   33023
         LockHover       =   1
         cGradient       =   16512
         Mode            =   0
         Value           =   0   'False
         cBack           =   16777215
      End
      Begin lvButton.lvButtons_H cmdIdle 
         Height          =   555
         Index           =   24
         Left            =   -73740
         TabIndex        =   66
         Top             =   6120
         Width           =   3120
         _ExtentX        =   5503
         _ExtentY        =   979
         Caption         =   "[X] TRIAL MOLDING"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFore           =   0
         cFHover         =   0
         cBhover         =   33023
         LockHover       =   1
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   16777215
      End
      Begin lvButton.lvButtons_H cmdIdle 
         Height          =   555
         Index           =   23
         Left            =   -73740
         TabIndex        =   65
         Top             =   5220
         Width           =   3120
         _ExtentX        =   5503
         _ExtentY        =   979
         Caption         =   "[W] PEMADAMAN INTERNAL"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFore           =   0
         cFHover         =   0
         cBhover         =   33023
         LockHover       =   1
         cGradient       =   16512
         Mode            =   0
         Value           =   0   'False
         cBack           =   16777215
      End
      Begin lvButton.lvButtons_H cmdIdle 
         Height          =   555
         Index           =   20
         Left            =   -73740
         TabIndex        =   64
         Top             =   4365
         Width           =   3120
         _ExtentX        =   5503
         _ExtentY        =   979
         Caption         =   "[T] PEMADAMAN LISTRIK PLN"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFore           =   0
         cFHover         =   0
         cBhover         =   33023
         LockHover       =   1
         cGradient       =   16512
         Mode            =   0
         Value           =   0   'False
         cBack           =   16777215
      End
      Begin lvButton.lvButtons_H cmdIdle 
         Height          =   555
         Index           =   19
         Left            =   -73740
         TabIndex        =   63
         Top             =   3510
         Width           =   3120
         _ExtentX        =   5503
         _ExtentY        =   979
         Caption         =   "[S] PROSES IMPROVEMENT"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFore           =   0
         cFHover         =   0
         cBhover         =   33023
         LockHover       =   1
         cGradient       =   16512
         Mode            =   0
         Value           =   0   'False
         cBack           =   16777215
      End
      Begin lvButton.lvButtons_H cmdIdle 
         Height          =   555
         Index           =   18
         Left            =   -73740
         TabIndex        =   62
         Top             =   2655
         Width           =   3120
         _ExtentX        =   5503
         _ExtentY        =   979
         Caption         =   "[R] TRIAL"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFore           =   0
         cFHover         =   0
         cBhover         =   33023
         LockHover       =   1
         cGradient       =   16512
         Mode            =   0
         Value           =   0   'False
         cBack           =   16777215
      End
      Begin lvButton.lvButtons_H cmdIdle 
         Height          =   555
         Index           =   17
         Left            =   -73740
         TabIndex        =   61
         Top             =   1755
         Width           =   3120
         _ExtentX        =   5503
         _ExtentY        =   979
         Caption         =   "[Q] NO ORDER"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFore           =   0
         cFHover         =   0
         cBhover         =   33023
         LockHover       =   1
         cGradient       =   16512
         Mode            =   0
         Value           =   0   'False
         cBack           =   16777215
      End
      Begin lvButton.lvButtons_H cmdIdle 
         Height          =   555
         Index           =   10
         Left            =   -73740
         TabIndex        =   60
         Top             =   900
         Width           =   3120
         _ExtentX        =   5503
         _ExtentY        =   979
         Caption         =   "[J] KETIDAK STABILAN PROSES"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFore           =   0
         cFHover         =   0
         cBhover         =   33023
         LockHover       =   1
         cGradient       =   16512
         Mode            =   0
         Value           =   0   'False
         cBack           =   16777215
      End
      Begin VB.Frame frameProd 
         BackColor       =   &H00FFFFFF&
         Height          =   735
         Left            =   270
         TabIndex        =   58
         Top             =   7200
         Width           =   5235
         Begin VB.TextBox txtLain 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Left            =   45
            TabIndex        =   59
            Top             =   135
            Width           =   5100
         End
      End
      Begin lvButton.lvButtons_H cmdIdle 
         Height          =   555
         Index           =   26
         Left            =   1260
         TabIndex        =   57
         Top             =   6210
         Width           =   3165
         _ExtentX        =   5583
         _ExtentY        =   979
         Caption         =   "[Z] LAIN-LAIN"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFore           =   0
         cFHover         =   0
         cBhover         =   33023
         LockHover       =   1
         cGradient       =   16512
         Mode            =   0
         Value           =   0   'False
         cBack           =   16777215
      End
      Begin lvButton.lvButtons_H cmdIdle 
         Height          =   555
         Index           =   22
         Left            =   1305
         TabIndex        =   56
         Top             =   5445
         Width           =   3120
         _ExtentX        =   5503
         _ExtentY        =   979
         Caption         =   "[V] MANTENANCE MESIN"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFore           =   0
         cFHover         =   0
         cBhover         =   33023
         LockHover       =   1
         cGradient       =   16512
         Mode            =   0
         Value           =   0   'False
         cBack           =   16777215
      End
      Begin lvButton.lvButtons_H cmdIdle 
         Height          =   555
         Index           =   21
         Left            =   1305
         TabIndex        =   55
         Top             =   4680
         Width           =   3120
         _ExtentX        =   5503
         _ExtentY        =   979
         Caption         =   "[U] MAINTENANCE MOLD"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFore           =   0
         cFHover         =   0
         cBhover         =   33023
         LockHover       =   1
         cGradient       =   16512
         Mode            =   0
         Value           =   0   'False
         cBack           =   16777215
      End
      Begin lvButton.lvButtons_H cmdIdle 
         Height          =   555
         Index           =   9
         Left            =   1305
         TabIndex        =   54
         Top             =   3915
         Width           =   3120
         _ExtentX        =   5503
         _ExtentY        =   979
         Caption         =   "[I] MESIN PROBLEM MAN"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFore           =   0
         cFHover         =   0
         cBhover         =   33023
         LockHover       =   1
         cGradient       =   16512
         Mode            =   0
         Value           =   0   'False
         cBack           =   16777215
      End
      Begin lvButton.lvButtons_H cmdIdle 
         Height          =   555
         Index           =   8
         Left            =   1305
         TabIndex        =   53
         Top             =   3150
         Width           =   3120
         _ExtentX        =   5503
         _ExtentY        =   979
         Caption         =   "[H] MESIN PROBLEM NORMAL"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFore           =   0
         cFHover         =   0
         cBhover         =   33023
         LockHover       =   1
         cGradient       =   16512
         Mode            =   0
         Value           =   0   'False
         cBack           =   16777215
      End
      Begin lvButton.lvButtons_H cmdIdle 
         Height          =   555
         Index           =   5
         Left            =   1305
         TabIndex        =   52
         Top             =   2385
         Width           =   3120
         _ExtentX        =   5503
         _ExtentY        =   979
         Caption         =   "[E] MOLD PROBLEM PROSES"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFore           =   0
         cFHover         =   0
         cBhover         =   33023
         LockHover       =   1
         cGradient       =   16512
         Mode            =   0
         Value           =   0   'False
         cBack           =   16777215
      End
      Begin lvButton.lvButtons_H cmdIdle 
         Height          =   555
         Index           =   4
         Left            =   1305
         TabIndex        =   51
         Top             =   1620
         Width           =   3120
         _ExtentX        =   5503
         _ExtentY        =   979
         Caption         =   "[D] MOLD PROBLEM NORMAL"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFore           =   0
         cFHover         =   0
         cBhover         =   33023
         LockHover       =   1
         cGradient       =   16512
         Mode            =   0
         Value           =   0   'False
         cBack           =   16777215
      End
      Begin lvButton.lvButtons_H cmdIdle 
         Height          =   555
         Index           =   1
         Left            =   1305
         TabIndex        =   50
         Top             =   900
         Width           =   3120
         _ExtentX        =   5503
         _ExtentY        =   979
         Caption         =   "[A] SETUP MOLD"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFore           =   0
         cFHover         =   0
         cBhover         =   33023
         LockHover       =   1
         cGradient       =   16512
         Mode            =   0
         Value           =   0   'False
         cBack           =   16777215
      End
   End
   Begin VB.Timer TimerPort 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   9765
      Top             =   45
   End
   Begin VB.Frame Frame3 
      Caption         =   "Keyboard"
      Height          =   3300
      Left            =   7425
      TabIndex        =   15
      Top             =   4275
      Visible         =   0   'False
      Width           =   8430
      Begin lvButton.lvButtons_H cmdKey 
         Height          =   600
         Index           =   0
         Left            =   585
         TabIndex        =   16
         Top             =   1035
         Width           =   690
         _ExtentX        =   1217
         _ExtentY        =   1058
         Caption         =   "A"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   11.25
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
      Begin lvButton.lvButtons_H cmdKey 
         Height          =   600
         Index           =   1
         Left            =   4230
         TabIndex        =   17
         Top             =   1755
         Width           =   690
         _ExtentX        =   1217
         _ExtentY        =   1058
         Caption         =   "B"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   11.25
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
      Begin lvButton.lvButtons_H cmdKey 
         Height          =   600
         Index           =   2
         Left            =   2610
         TabIndex        =   18
         Top             =   1755
         Width           =   690
         _ExtentX        =   1217
         _ExtentY        =   1058
         Caption         =   "C"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   11.25
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
      Begin lvButton.lvButtons_H cmdKey 
         Height          =   600
         Index           =   3
         Left            =   2205
         TabIndex        =   19
         Top             =   1035
         Width           =   690
         _ExtentX        =   1217
         _ExtentY        =   1058
         Caption         =   "D"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   11.25
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
      Begin lvButton.lvButtons_H cmdKey 
         Height          =   600
         Index           =   4
         Left            =   1800
         TabIndex        =   20
         Top             =   315
         Width           =   690
         _ExtentX        =   1217
         _ExtentY        =   1058
         Caption         =   "E"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   11.25
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
      Begin lvButton.lvButtons_H cmdKey 
         Height          =   600
         Index           =   5
         Left            =   3015
         TabIndex        =   21
         Top             =   1035
         Width           =   690
         _ExtentX        =   1217
         _ExtentY        =   1058
         Caption         =   "F"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   11.25
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
      Begin lvButton.lvButtons_H cmdKey 
         Height          =   600
         Index           =   6
         Left            =   3825
         TabIndex        =   22
         Top             =   1035
         Width           =   690
         _ExtentX        =   1217
         _ExtentY        =   1058
         Caption         =   "G"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   11.25
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
      Begin lvButton.lvButtons_H cmdKey 
         Height          =   600
         Index           =   7
         Left            =   4635
         TabIndex        =   23
         Top             =   1035
         Width           =   690
         _ExtentX        =   1217
         _ExtentY        =   1058
         Caption         =   "H"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   11.25
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
      Begin lvButton.lvButtons_H cmdKey 
         Height          =   600
         Index           =   8
         Left            =   5850
         TabIndex        =   24
         Top             =   315
         Width           =   690
         _ExtentX        =   1217
         _ExtentY        =   1058
         Caption         =   "I"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   11.25
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
      Begin lvButton.lvButtons_H cmdKey 
         Height          =   600
         Index           =   9
         Left            =   5445
         TabIndex        =   25
         Top             =   1035
         Width           =   690
         _ExtentX        =   1217
         _ExtentY        =   1058
         Caption         =   "J"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   11.25
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
      Begin lvButton.lvButtons_H cmdKey 
         Height          =   600
         Index           =   10
         Left            =   6255
         TabIndex        =   26
         Top             =   1035
         Width           =   690
         _ExtentX        =   1217
         _ExtentY        =   1058
         Caption         =   "K"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   11.25
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
      Begin lvButton.lvButtons_H cmdKey 
         Height          =   600
         Index           =   11
         Left            =   7065
         TabIndex        =   27
         Top             =   1035
         Width           =   690
         _ExtentX        =   1217
         _ExtentY        =   1058
         Caption         =   "L"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   11.25
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
      Begin lvButton.lvButtons_H cmdKey 
         Height          =   600
         Index           =   12
         Left            =   5850
         TabIndex        =   28
         Top             =   1755
         Width           =   690
         _ExtentX        =   1217
         _ExtentY        =   1058
         Caption         =   "M"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   11.25
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
      Begin lvButton.lvButtons_H cmdKey 
         Height          =   600
         Index           =   13
         Left            =   5040
         TabIndex        =   29
         Top             =   1755
         Width           =   690
         _ExtentX        =   1217
         _ExtentY        =   1058
         Caption         =   "N"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   11.25
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
      Begin lvButton.lvButtons_H cmdKey 
         Height          =   600
         Index           =   14
         Left            =   6660
         TabIndex        =   30
         Top             =   315
         Width           =   690
         _ExtentX        =   1217
         _ExtentY        =   1058
         Caption         =   "O"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   11.25
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
      Begin lvButton.lvButtons_H cmdKey 
         Height          =   600
         Index           =   15
         Left            =   7470
         TabIndex        =   31
         Top             =   315
         Width           =   690
         _ExtentX        =   1217
         _ExtentY        =   1058
         Caption         =   "P"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   11.25
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
      Begin lvButton.lvButtons_H cmdKey 
         Height          =   600
         Index           =   16
         Left            =   180
         TabIndex        =   32
         Top             =   315
         Width           =   690
         _ExtentX        =   1217
         _ExtentY        =   1058
         Caption         =   "Q"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   11.25
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
      Begin lvButton.lvButtons_H cmdKey 
         Height          =   600
         Index           =   17
         Left            =   2610
         TabIndex        =   33
         Top             =   315
         Width           =   690
         _ExtentX        =   1217
         _ExtentY        =   1058
         Caption         =   "R"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   11.25
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
      Begin lvButton.lvButtons_H cmdKey 
         Height          =   600
         Index           =   18
         Left            =   1395
         TabIndex        =   34
         Top             =   1035
         Width           =   690
         _ExtentX        =   1217
         _ExtentY        =   1058
         Caption         =   "S"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   11.25
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
      Begin lvButton.lvButtons_H cmdKey 
         Height          =   600
         Index           =   19
         Left            =   3420
         TabIndex        =   35
         Top             =   315
         Width           =   690
         _ExtentX        =   1217
         _ExtentY        =   1058
         Caption         =   "T"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   11.25
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
      Begin lvButton.lvButtons_H cmdKey 
         Height          =   600
         Index           =   20
         Left            =   5040
         TabIndex        =   36
         Top             =   315
         Width           =   690
         _ExtentX        =   1217
         _ExtentY        =   1058
         Caption         =   "U"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   11.25
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
      Begin lvButton.lvButtons_H cmdKey 
         Height          =   600
         Index           =   21
         Left            =   3420
         TabIndex        =   37
         Top             =   1755
         Width           =   690
         _ExtentX        =   1217
         _ExtentY        =   1058
         Caption         =   "V"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   11.25
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
      Begin lvButton.lvButtons_H cmdKey 
         Height          =   600
         Index           =   22
         Left            =   990
         TabIndex        =   38
         Top             =   315
         Width           =   690
         _ExtentX        =   1217
         _ExtentY        =   1058
         Caption         =   "W"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   11.25
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
      Begin lvButton.lvButtons_H cmdKey 
         Height          =   600
         Index           =   23
         Left            =   1800
         TabIndex        =   39
         Top             =   1755
         Width           =   690
         _ExtentX        =   1217
         _ExtentY        =   1058
         Caption         =   "X"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   11.25
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
      Begin lvButton.lvButtons_H cmdKey 
         Height          =   600
         Index           =   24
         Left            =   4230
         TabIndex        =   40
         Top             =   315
         Width           =   690
         _ExtentX        =   1217
         _ExtentY        =   1058
         Caption         =   "Y"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   11.25
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
      Begin lvButton.lvButtons_H cmdKey 
         Height          =   600
         Index           =   25
         Left            =   990
         TabIndex        =   41
         Top             =   1755
         Width           =   690
         _ExtentX        =   1217
         _ExtentY        =   1058
         Caption         =   "Z"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   11.25
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
      Begin lvButton.lvButtons_H cmdKey 
         Height          =   600
         Index           =   26
         Left            =   3015
         TabIndex        =   42
         Top             =   2475
         Width           =   3525
         _ExtentX        =   6218
         _ExtentY        =   1058
         Caption         =   "Space"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   11.25
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
      Begin lvButton.lvButtons_H cmdKey 
         Height          =   600
         Index           =   27
         Left            =   6750
         TabIndex        =   43
         Top             =   2475
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   1058
         Caption         =   "Enter"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   11.25
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
      Begin lvButton.lvButtons_H cmdKey 
         Height          =   600
         Index           =   28
         Left            =   180
         TabIndex        =   44
         Top             =   2475
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   1058
         Caption         =   "Clear"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   11.25
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
      Begin lvButton.lvButtons_H cmdKey 
         Height          =   600
         Index           =   29
         Left            =   6750
         TabIndex        =   47
         Top             =   1755
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   1058
         Caption         =   "<----           Back Space"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
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
      Begin lvButton.lvButtons_H cmdKey 
         Height          =   600
         Index           =   30
         Left            =   2160
         TabIndex        =   48
         Top             =   2475
         Width           =   690
         _ExtentX        =   1217
         _ExtentY        =   1058
         Caption         =   ","
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   11.25
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
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   8730
      Top             =   45
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   9225
      Top             =   45
   End
   Begin MSComctlLib.ImageList i16x16 
      Left            =   8100
      Top             =   45
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   15
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIdleTime.frx":617A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIdleTime.frx":6B8C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIdleTime.frx":759E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIdleTime.frx":7938
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIdleTime.frx":7CD2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIdleTime.frx":806C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIdleTime.frx":8406
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIdleTime.frx":8E18
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIdleTime.frx":982A
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIdleTime.frx":A23C
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIdleTime.frx":AC4E
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIdleTime.frx":B660
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIdleTime.frx":C072
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIdleTime.frx":CA84
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIdleTime.frx":D020
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView LvList 
      Height          =   4560
      Left            =   6120
      TabIndex        =   0
      Top             =   3240
      Width           =   11355
      _ExtentX        =   20029
      _ExtentY        =   8043
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "i16x16"
      SmallIcons      =   "i16x16"
      ColHdrIcons     =   "i16x16"
      ForeColor       =   -2147483640
      BackColor       =   16777215
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin lvButton.lvButtons_H cmdSwitchUser 
      Height          =   555
      Left            =   15750
      TabIndex        =   9
      Top             =   2340
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   979
      Caption         =   "GANTI USER"
      CapAlign        =   2
      BackStyle       =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
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
      Image           =   "frmIdleTime.frx":D5BC
      cBack           =   4210752
   End
   Begin lvButton.lvButtons_H cmdAddNG 
      Height          =   555
      Left            =   15750
      TabIndex        =   10
      Top             =   1620
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   979
      Caption         =   "INPUT NG"
      CapAlign        =   2
      BackStyle       =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
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
      Image           =   "frmIdleTime.frx":13746
      cBack           =   4210752
   End
   Begin lvButton.lvButtons_H cmdExit 
      Height          =   555
      Left            =   15750
      TabIndex        =   11
      Top             =   180
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
      Image           =   "frmIdleTime.frx":198D0
      cBack           =   4210752
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Description :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6165
      TabIndex        =   79
      Top             =   7965
      Width           =   1545
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      Height          =   735
      Left            =   12825
      Top             =   2160
      Width           =   735
   End
   Begin VB.Label lbloff 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "off"
      Height          =   240
      Left            =   13770
      TabIndex        =   46
      Top             =   2565
      Width           =   510
   End
   Begin VB.Label lblOn 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "on"
      Height          =   240
      Left            =   13770
      TabIndex        =   45
      Top             =   2295
      Width           =   510
   End
   Begin VB.Image Image2 
      Height          =   720
      Left            =   90
      Picture         =   "frmIdleTime.frx":30292
      Top             =   -45
      Width           =   720
   End
   Begin VB.Label lblMachineName 
      BackStyle       =   0  'Transparent
      Caption         =   "--"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   3420
      TabIndex        =   14
      Top             =   135
      Width           =   3930
   End
   Begin VB.Label lblnNoMachine 
      BackStyle       =   0  'Transparent
      Caption         =   "--"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   2565
      TabIndex        =   13
      Top             =   135
      Width           =   600
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Machine :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   1080
      TabIndex        =   12
      Top             =   135
      Width           =   1320
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
      Left            =   8100
      TabIndex        =   8
      Top             =   2475
      Width           =   2715
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
      Left            =   6120
      TabIndex        =   7
      Top             =   2475
      Width           =   1770
   End
   Begin VB.Label lblKodeIdle 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   8100
      TabIndex        =   6
      Top             =   1080
      Width           =   375
   End
   Begin VB.Label lblIdleName 
      BackStyle       =   0  'Transparent
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
      Left            =   8505
      TabIndex        =   5
      Top             =   1080
      Width           =   3795
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
      Left            =   6120
      TabIndex        =   4
      Top             =   1080
      Width           =   1770
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
      Left            =   6120
      TabIndex        =   3
      Top             =   1620
      Width           =   1770
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
      Left            =   8100
      TabIndex        =   2
      Top             =   1620
      Width           =   2715
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
      Left            =   8100
      TabIndex        =   1
      Top             =   2070
      Width           =   2715
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   7380
      Picture         =   "frmIdleTime.frx":3640C
      Top             =   1980
      Width           =   480
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   735
      Left            =   13680
      Shape           =   3  'Circle
      Top             =   2160
      Width           =   735
   End
End
Attribute VB_Name = "frmIdleTime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim srcItem                         As ListItem
Dim srcRecord                       As String

Dim PortOn                          As String
Dim IdleOn                       As Boolean
 
Private Declare Function ShellExecute Lib "shell32.dll" _
Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Dim iShot As Integer
Dim eng_prod_1 As String, eng_prod_2 As String

Dim sSQL As String
Dim sQL As String

Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Dim x1 As Boolean
Dim xShot As Integer


Private Sub lbloff_DblClick()
    lbloff.Caption = "95"
End Sub


'Private Sub Timer3_Timer()
'    If GetAsyncKeyState(119) < 0 Then
'        If x1 = False Then
'            Shape2.BackColor = vbRed
'            If xShot = 5 Then
'                Call AutoUpdate
'                xShot = 0
'
'            Else
'                xShot = xShot + 1
'            End If
'
'            Debug.Print "Tekan"
'            x1 = 1
'        End If
'    Else
'        If x1 = True Then
'            Shape2.BackColor = vbWhite
'            Debug.Print "Angkat"
'            x1 = 0
'        End If
'    End If
'End Sub

Private Sub TimerPort_Timer()
    'ShowSensor = ReadINI("SETTING", "SHOWSENSOR", App.Path & "\Settings.ini")
    If ShowSensor = True Then
        lbloff.Caption = PortIn(PortAddress)
    End If
End Sub

'Private Sub lbloff_Change()
'
'    If lbloff.Caption = PortOn Then
'        Shape1.FillColor = vbGreen
'        If iShot = 5 Then
'            Call AutoUpdate
'            iShot = 0
'
'        Else
'            iShot = iShot + 1
'        End If
'    Else
'        Shape1.FillColor = vbRed
'    End If
'
'End Sub

Private Sub AutoUpdate()
On Error GoTo ErrHandler

    If bIdle = True Then
    
        'insert data
        sQL = "insert into sip_production.prod_machine_idles"
        sQL = sQL & " (plant_mark,prod_machine_id,mkt_customer_id,eng_product_1,eng_product_2,prod_idletime_id,"
        sQL = sQL & " period_shift,start_idle,proses,created_at,created_by,end_idle,idle_time,hrd_employee_id,description) values"
        sQL = sQL & " ('" & p_plant_mark & "','" & p_prod_machine_id & "','" & p_mkt_customer_id & "'," & eng_prod_1 & ""
        sQL = sQL & " ," & eng_prod_2 & ",'26','" & Format(p_shift, "yyyy-mm-dd") & "','" & lblTimeStart.Caption & "'"
        sQL = sQL & " ,'N','" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "','" & ACTIVE_USER.KODEUSER & "'"
        sQL = sQL & " ,'" & lblTimeStop.Caption & "','" & lblLossTime.Caption & "','" & ACTIVE_USER.KODEUSER & "','Auto Close System')"
        
        sSQL_Insert sQL

    Else
        sQL = "update sip_production.prod_machine_idles set end_idle = '" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "'"
            sQL = sQL & " ,idle_time = '" & lblLossTime.Caption & "',proses = 'N',hrd_employee_id = '" & ACTIVE_USER.KODEUSER & "',description = 'Auto Close System' where "
            sQL = sQL & " plant_mark = '" & p_plant_mark & "' "
            sQL = sQL & " and prod_machine_id = '" & p_prod_machine_id & "'"
            sQL = sQL & " and mkt_customer_id = '" & p_mkt_customer_id & "'"
            sQL = sQL & " and proses = 'Y' and prod_idletime_id = '" & lblKodeIdle.Caption & "'"
        sSQL_Update sQL
        
    End If

    Timer1.Enabled = False
    bIdle = False
    MAIN.Label4.Caption = ""
    MAIN.Label1.Caption = Format(Now, "hh:mm:ss")
    
    Unload Me

Exit Sub
ErrHandler:
    MsgBox "Error Number: " & Err.Number & vbCrLf & "Description: " & Err.Description, vbExclamation, Me.Caption
    
    
End Sub
Private Sub cmdAddNG_Click()
    frmNg.Show 1
End Sub

Private Sub cmdExit_Click()
On Error GoTo ErrHandler

    If cmdExit.Caption = "&EXIT" Then
        bIdle = False
        MAIN.Label4.Caption = ""
        MAIN.Label1.Caption = Format(Now, "hh:mm:ss")
        Unload Me
    Else
        
        If lblKodeIdle.Caption = "" Then
            MsgBox "Idle name belum di pilih..!", vbExclamation
            Exit Sub
        End If
                
        frmUnlock.Show 1
        
        If LOG_APP = True Then
         
'            If bIdle = True Then
'
'                'insert data
'                sQL = "insert into sip_production.prod_machine_idles"
'                sQL = sQL & " (plant_mark,prod_machine_id,mkt_customer_id,eng_product_1,eng_product_2,prod_idletime_id,"
'                sQL = sQL & " period_shift,start_idle,proses,created_at,created_by,end_idle,idle_time,hrd_employee_id) values"
'                sQL = sQL & " ('" & p_plant_mark & "','" & p_prod_machine_id & "','" & p_mkt_customer_id & "'," & eng_prod_1 & ""
'                sQL = sQL & " ," & eng_prod_2 & ",'" & lblKodeIdle.Caption & "','" & Format(p_shift, "yyyy-mm-dd") & "','" & lblTimeStart.Caption & "'"
'                sQL = sQL & " ,'N','" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "','" & ACTIVE_USER.KODEUSER & "'"
'                sQL = sQL & " ,'" & lblTimeStop.Caption & "','" & lblLossTime.Caption & "','" & ACTIVE_ADMIN.KODEUSER & "')"
'
'                sSQL_Insert sQL
'
'            Else
                sQL = "update sip_production.prod_machine_idles set end_idle = '" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "'"
                    sQL = sQL & " ,idle_time = '" & lblLossTime.Caption & "',proses = 'N',hrd_employee_id = '" & ACTIVE_ADMIN.KODEUSER & "' "
                    sQL = sQL & " ,description = '" & txtDescription.text & "' Where "
                    sQL = sQL & " plant_mark = '" & p_plant_mark & "' "
                    sQL = sQL & " and prod_machine_id = '" & p_prod_machine_id & "'"
                    sQL = sQL & " and mkt_customer_id = '" & p_mkt_customer_id & "'"
                    sQL = sQL & " and proses = 'Y' and prod_idletime_id = '" & lblKodeIdle.Caption & "'"
                sSQL_Update sQL
                
'            End If
'
            Call FillListview
            
            XTab1.Enabled = True
            cmdExit.Caption = "&EXIT"
            Timer1.Enabled = False
            
        End If

    End If

Exit Sub
ErrHandler:
    MsgBox "Error Number: " & Err.Number & vbCrLf & "Description: " & Err.Description, vbExclamation, Me.Caption
    
End Sub


Private Sub cmdIdle_Click(Index As Integer)
On Error GoTo ErrHandler

        If Index = 26 Then
            If Frame3.Visible = True Then Frame3.Visible = False Else Frame3.Visible = True
        Else
            If bIdle <> True Then
                Timer1.Enabled = True
                lblTimeStart.Caption = Format(Now, "yyyy-mm-dd hh:mm:ss")
            End If
                
            XTab1.Enabled = False
            cmdExit.Caption = "RUNNING"
            LOG_APP = False
            lblKodeIdle.Caption = Index
            lblIdleName.Caption = cmdIdle(Index).Caption

            'insert data
            sQL = "insert into sip_production.prod_machine_idles"
            sQL = sQL & " (plant_mark,prod_machine_id,mkt_customer_id,eng_product_1,eng_product_2,prod_idletime_id,"
            sQL = sQL & " period_shift,start_idle,proses,created_at,created_by,description) values"
            sQL = sQL & " ('" & p_plant_mark & "','" & p_prod_machine_id & "','" & p_mkt_customer_id & "'," & eng_prod_1 & ""
            sQL = sQL & " ," & eng_prod_2 & ",'" & Index & "','" & Format(p_shift, "yyyy-mm-dd") & "','" & lblTimeStart.Caption & "'"
            sQL = sQL & " ,'Y','" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "','" & ACTIVE_USER.KODEUSER & "','" & txtLain.text & "')"
            
            sSQL_Insert sQL
            
            Call FillListview
            
            Timer2.Enabled = True
        End If

Exit Sub
ErrHandler:
    MsgBox "Error Number: " & Err.Number & vbCrLf & "Description: " & Err.Description, vbExclamation, Me.Caption
 
End Sub



Private Sub cmdKey_Click(Index As Integer)
On Error GoTo ErrHandler

        If Index = 27 Then
            If txtLain.text = "" Then
                MsgBox "Keterangan Belum di Isi..!", vbExclamation
                Exit Sub
            Else

                If bIdle <> True Then
                    Timer1.Enabled = True
                    lblTimeStart.Caption = Format(Now, "yyyy-mm-dd hh:mm:ss")
                End If
                
                XTab1.Enabled = False
                cmdExit.Caption = "RUNNING"
                LOG_APP = False

                lblKodeIdle.Caption = 26
                lblIdleName.Caption = txtLain.text
        
                'insert data
                sQL = "insert into sip_production.prod_machine_idles"
                sQL = sQL & " (plant_mark,prod_machine_id,mkt_customer_id,eng_product_1,eng_product_2,prod_idletime_id,"
                sQL = sQL & " period_shift,start_idle,proses,created_at,created_by,description) values"
                sQL = sQL & " ('" & p_plant_mark & "','" & p_prod_machine_id & "','" & p_mkt_customer_id & "'," & eng_prod_1 & ""
                sQL = sQL & " ," & eng_prod_2 & ",'" & lblKodeIdle.Caption & "','" & Format(p_shift, "yyyy-mm-dd") & "','" & lblTimeStart.Caption & "'"
                sQL = sQL & " ,'Y','" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "','" & ACTIVE_USER.KODEUSER & "','" & txtLain.text & "')"
                
                sSQL_Insert sQL
                
                Call FillListview
                
                Timer2.Enabled = True
                Frame3.Visible = False
                
            End If
        ElseIf Index = 28 Then
            txtLain.text = ""
        ElseIf Index = 29 Then
            If Len(txtLain.text) = 0 Then Exit Sub
            txtLain.text = Mid(txtLain.text, 1, Len(txtLain.text) - 1)
        ElseIf Index = 26 Then
            txtLain.text = txtLain.text & " "
        Else
            txtLain.text = txtLain.text & cmdKey(Index).Caption
        End If

    
Exit Sub
ErrHandler:
    MsgBox "Error Number: " & Err.Number & vbCrLf & "Description: " & Err.Description, vbExclamation, Me.Caption
    
End Sub

Private Sub cmdSwitchUser_Click()
    If MsgBox("Silahkan Logout sebelum mengganti dengan User yang lain, apakah mau diganti ?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
    
    Dim iSQL As String
    iSQL = "INSERT INTO hrd_login_logs (plant_mark,loc_code,hrd_employee_id,emp_code"
    iSQL = iSQL & " ,tr_date,tr_time,acc_code,created_at,created_by)"
    iSQL = iSQL & " VALUES ('" & p_plant_mark & "','" & NoMesin & "','" & ACTIVE_USER.KODEUSER & "','" & ACTIVE_USER.KODEPIN & "'"
    iSQL = iSQL & " ,'" & Format(Now, "yyyy-mm-dd") & "','" & Format(Now, "hh:mm:ss") & "','2'"
    iSQL = iSQL & " ,'" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "','" & ACTIVE_USER.SYSID & "')"
    
    sSQL_Insert iSQL
    frmLogin.Show vbModal
End Sub



Private Sub Form_Activate()
With MAIN
    Me.BackColor = .ACPMenu.BackColor
    Frame3.BackColor = .ACPMenu.BackColor
    Me.Picture = .ACPMenu.LoadBackground
End With

    'AltLVBackground lvList, vbWhite, &HFFFFC0, frmIdleTime
End Sub

Private Sub Form_Load()
On Error GoTo ErrHandler

    lblnNoMachine.Caption = p_machine_no
    lblMachineName.Caption = p_machine_name

    If p_eng_product_1 = "" Then
        eng_prod_1 = "NULL"
    Else
        eng_prod_1 = p_eng_product_1
    End If
    If p_eng_product_2 = "" Then
        eng_prod_2 = "NULL"
    Else
        eng_prod_2 = p_eng_product_2
    End If
            
    FillListview
    
    If bIdle = True Then
        cmdExit.Caption = "RUNNING"
        Timer1.Enabled = True
        lblTimeStart.Caption = Format(Now, "yyyy-mm-dd hh:mm:ss")
    End If
    
    LOG_APP = False
    
    bformIdle = True

    PortOn = ReadINI("SETTING", "SERIAL", App.Path & "\Settings.ini")
    IdleOn = ReadINI("SETTING", "IDLEON", App.Path & "\Settings.ini")
    lblOn.Caption = PortOn
    
    PortAddress = &H379
    
    TimerPort.Enabled = True
    
    iShot = 0
    
    formIdle = True
    
Exit Sub
ErrHandler:
    MsgBox "Error Number: " & Err.Number & vbCrLf & "Description: " & Err.Description, vbExclamation, Me.Caption
    
End Sub



Private Sub Form_Unload(Cancel As Integer)
    formIdle = False
    bformIdle = False
End Sub

Private Sub Timer1_Timer()
    lblTimeStop.Caption = Format(Now, "yyyy-mm-dd hh:mm:ss")
    lblLossTime.Caption = Format(CDate(CDate(lblTimeStop.Caption) - CDate(lblTimeStart.Caption)), "hh:mm:ss")
    
    If Format(Now, "HH:MM") = "08:00" Then
        Call AutoUpdate
    End If
    
End Sub

Private Sub FillListview()
On Error GoTo ErrHandler
    Dim Rs As New Recordset

    Rs.CursorLocation = adUseClient
    sSQL = "select @no:=@no+1 AS nomor, a.start_idle,a.end_idle,a.idle_time,a.description, b.number,b.name as machine_name, c.name as idle_name,d.name as karyawan"
    sSQL = sSQL & " from sip_production.prod_machine_idles a"
    sSQL = sSQL & " JOIN (SELECT @no:=0) as no"
    sSQL = sSQL & " INNER JOIN sip_production.prod_machines b ON a.prod_machine_id = b.id"
    sSQL = sSQL & " INNER JOIN sip_production.prod_idletimes c ON a.prod_idletime_id = c.id"
    sSQL = sSQL & " LEFT JOIN sip_production.hrd_employees d ON a.hrd_employee_id = d.id "
    sSQL = sSQL & " WHERE b.number = '" & lblnNoMachine.Caption & "'"
    sSQL = sSQL & " AND date(a.period_shift) = '" & Format(p_shift, "yyyy-mm-dd") & "' ORDER BY nomor DESC"

    Rs.Open sSQL, CN, adOpenStatic, adLockOptimistic
    
    With lvList
        .GridLines = True
        .View = lvwReport
        
        .ColumnHeaders.Clear
        .ColumnHeaders.Add , , "NO."
        .ColumnHeaders.Add , , "START TIME"
        .ColumnHeaders.Add , , "END TIME"
        .ColumnHeaders.Add , , "IDLE TIME"
        .ColumnHeaders.Add , , "IDLE NAME"
        .ColumnHeaders.Add , , "KARYAWAN"
        .ColumnHeaders.Add , , "DESCRIPTION"
        .ListItems.Clear
        Do While Not Rs.EOF
        Set srcItem = .ListItems.Add(, , Rs.Fields("nomor"), 1, 1)
            srcItem.SubItems(1) = Format(Rs.Fields("start_idle"), "yyyy-mm-dd hh:mm:ss")
            srcItem.SubItems(2) = IIf(IsNull(Format(Rs.Fields("end_idle"), "yyyy-mm-dd hh:mm:ss")), "", Format(Rs.Fields("end_idle"), "yyyy-mm-dd hh:mm:ss"))
            srcItem.SubItems(3) = IIf(IsNull(Format(Rs.Fields("idle_time"), "hh:mm:ss")), "", Format(Rs.Fields("idle_time"), "hh:mm:ss"))
            srcItem.SubItems(4) = IIf(IsNull(Rs.Fields("idle_name")), "", Rs.Fields("idle_name"))
            srcItem.SubItems(5) = IIf(IsNull(Rs.Fields("karyawan")), "", Rs.Fields("karyawan"))
            srcItem.SubItems(6) = IIf(IsNull(Rs.Fields("description")), "", Rs.Fields("description"))
    
            Rs.MoveNext
        Loop
    End With
    Call lvSizeColumns(lvList)

Exit Sub
ErrHandler:
    MsgBox "Error Number: " & Err.Number & vbCrLf & "Description: " & Err.Description, vbExclamation, Me.Caption
    
End Sub

Private Sub Timer2_Timer()

    If cmdExit.Caption <> "&EXIT" Then

        sQL = "update sip_production.prod_machine_idles set end_idle = '" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "'"
            sQL = sQL & " ,idle_time = '" & lblLossTime.Caption & "' where "
            sQL = sQL & " plant_mark = '" & p_plant_mark & "' "
            sQL = sQL & " and prod_machine_id = '" & p_prod_machine_id & "'"
            sQL = sQL & " and mkt_customer_id = '" & p_mkt_customer_id & "'"
            sQL = sQL & " and proses = 'Y' and prod_idletime_id = '" & lblKodeIdle.Caption & "'"
        sSQL_Update sQL
    
        Call FillListview
    End If
End Sub

