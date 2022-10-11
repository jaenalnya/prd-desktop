VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form frmInfo 
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   10080
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   17655
   Icon            =   "frmInfo.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   10080
   ScaleWidth      =   17655
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Frame frameDesign 
      BackColor       =   &H00FFFFFF&
      Height          =   2175
      Index           =   0
      Left            =   0
      TabIndex        =   999
      Top             =   630
      Width           =   2850
      Begin VB.TextBox txtInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   0
         Left            =   630
         MultiLine       =   -1  'True
         TabIndex        =   1011
         Top             =   135
         Width           =   2175
      End
      Begin VB.TextBox txtidle 
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
         Index           =   0
         Left            =   630
         TabIndex        =   1010
         Top             =   1710
         Width           =   735
      End
      Begin VB.TextBox txtShot 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   0
         Left            =   630
         TabIndex        =   1009
         Top             =   810
         Width           =   555
      End
      Begin VB.TextBox txtShot 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   1
         Left            =   1170
         TabIndex        =   1008
         Top             =   810
         Width           =   555
      End
      Begin VB.TextBox txtShot 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   2
         Left            =   1710
         TabIndex        =   1007
         Top             =   810
         Width           =   555
      End
      Begin VB.TextBox txtShot 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   3
         Left            =   2250
         TabIndex        =   1006
         Top             =   810
         Width           =   555
      End
      Begin VB.TextBox txtidle 
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
         Index           =   1
         Left            =   1350
         TabIndex        =   1005
         Top             =   1710
         Width           =   735
      End
      Begin VB.TextBox txtidle 
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
         Index           =   2
         Left            =   2070
         TabIndex        =   1004
         Top             =   1710
         Width           =   735
      End
      Begin VB.TextBox txtng 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   0
         Left            =   630
         TabIndex        =   1003
         Top             =   1080
         Width           =   555
      End
      Begin VB.TextBox txtng 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   1
         Left            =   1170
         TabIndex        =   1002
         Top             =   1080
         Width           =   555
      End
      Begin VB.TextBox txtng 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   2
         Left            =   1710
         TabIndex        =   1001
         Top             =   1080
         Width           =   555
      End
      Begin VB.TextBox txtng 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   3
         Left            =   2250
         TabIndex        =   1000
         Top             =   1080
         Width           =   555
      End
      Begin VB.Label lblMesin 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "01"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   465
         Index           =   1
         Left            =   90
         TabIndex        =   1022
         Top             =   135
         Width           =   510
      End
      Begin VB.Label lbljam2R 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "10 : R"
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
         Height          =   240
         Index           =   0
         Left            =   2250
         TabIndex        =   1021
         Top             =   585
         Width           =   555
      End
      Begin VB.Label lbljam2L 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "10 : L"
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
         Height          =   240
         Index           =   0
         Left            =   1710
         TabIndex        =   1020
         Top             =   585
         Width           =   555
      End
      Begin VB.Label lbljam1R 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "9 : R"
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
         Height          =   240
         Index           =   0
         Left            =   1170
         TabIndex        =   1019
         Top             =   585
         Width           =   555
      End
      Begin VB.Label lbljam1L 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "9 : L"
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
         Height          =   240
         Index           =   0
         Left            =   630
         TabIndex        =   1018
         Top             =   585
         Width           =   555
      End
      Begin VB.Label lblIdle 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "        IDLE"
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
         Height          =   600
         Index           =   1
         Left            =   90
         TabIndex        =   1017
         Top             =   1485
         Width           =   510
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "NG"
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
         Height          =   330
         Index           =   1
         Left            =   90
         TabIndex        =   1016
         Top             =   1125
         Width           =   510
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "GROSS"
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
         Height          =   285
         Index           =   2
         Left            =   90
         TabIndex        =   1015
         Top             =   810
         Width           =   510
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SHIFT 3"
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
         Height          =   240
         Index           =   5
         Left            =   2070
         TabIndex        =   1014
         Top             =   1485
         Width           =   735
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SHIFT 2"
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
         Height          =   240
         Index           =   4
         Left            =   1350
         TabIndex        =   1013
         Top             =   1485
         Width           =   735
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SHIFT 1"
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
         Height          =   240
         Index           =   3
         Left            =   630
         TabIndex        =   1012
         Top             =   1485
         Width           =   735
      End
   End
   Begin VB.Frame frameDesign 
      BackColor       =   &H00FFFFFF&
      Height          =   2175
      Index           =   1
      Left            =   0
      TabIndex        =   975
      Top             =   2925
      Width           =   2850
      Begin VB.TextBox txtng 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   4
         Left            =   2250
         TabIndex        =   987
         Top             =   1080
         Width           =   555
      End
      Begin VB.TextBox txtng 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   5
         Left            =   1710
         TabIndex        =   986
         Top             =   1080
         Width           =   555
      End
      Begin VB.TextBox txtng 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   6
         Left            =   1170
         TabIndex        =   985
         Top             =   1080
         Width           =   555
      End
      Begin VB.TextBox txtng 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   7
         Left            =   630
         TabIndex        =   984
         Top             =   1080
         Width           =   555
      End
      Begin VB.TextBox txtidle 
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
         Index           =   3
         Left            =   2070
         TabIndex        =   983
         Top             =   1710
         Width           =   735
      End
      Begin VB.TextBox txtidle 
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
         Index           =   4
         Left            =   1350
         TabIndex        =   982
         Top             =   1710
         Width           =   735
      End
      Begin VB.TextBox txtShot 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   4
         Left            =   2250
         TabIndex        =   981
         Top             =   810
         Width           =   555
      End
      Begin VB.TextBox txtShot 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   5
         Left            =   1710
         TabIndex        =   980
         Top             =   810
         Width           =   555
      End
      Begin VB.TextBox txtShot 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   6
         Left            =   1170
         TabIndex        =   979
         Top             =   810
         Width           =   555
      End
      Begin VB.TextBox txtShot 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   7
         Left            =   630
         TabIndex        =   978
         Top             =   810
         Width           =   555
      End
      Begin VB.TextBox txtidle 
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
         Index           =   5
         Left            =   630
         TabIndex        =   977
         Top             =   1710
         Width           =   735
      End
      Begin VB.TextBox txtInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   1
         Left            =   630
         MultiLine       =   -1  'True
         TabIndex        =   976
         Top             =   135
         Width           =   2175
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SHIFT 1"
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
         Height          =   240
         Index           =   6
         Left            =   630
         TabIndex        =   998
         Top             =   1485
         Width           =   735
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SHIFT 2"
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
         Height          =   240
         Index           =   7
         Left            =   1350
         TabIndex        =   997
         Top             =   1485
         Width           =   735
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SHIFT 3"
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
         Height          =   240
         Index           =   8
         Left            =   2070
         TabIndex        =   996
         Top             =   1485
         Width           =   735
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "GROSS"
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
         Height          =   285
         Index           =   9
         Left            =   90
         TabIndex        =   995
         Top             =   810
         Width           =   510
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "NG"
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
         Height          =   330
         Index           =   10
         Left            =   90
         TabIndex        =   994
         Top             =   1125
         Width           =   510
      End
      Begin VB.Label lblIdle 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "        IDLE"
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
         Height          =   600
         Index           =   2
         Left            =   90
         TabIndex        =   993
         Top             =   1485
         Width           =   510
      End
      Begin VB.Label lbljam1L 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "9 : L"
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
         Height          =   240
         Index           =   1
         Left            =   630
         TabIndex        =   992
         Top             =   585
         Width           =   555
      End
      Begin VB.Label lbljam1R 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "9 : R"
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
         Height          =   240
         Index           =   1
         Left            =   1170
         TabIndex        =   991
         Top             =   585
         Width           =   555
      End
      Begin VB.Label lbljam2L 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "10 : L"
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
         Height          =   240
         Index           =   1
         Left            =   1710
         TabIndex        =   990
         Top             =   585
         Width           =   555
      End
      Begin VB.Label lbljam2R 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "10 : R"
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
         Height          =   240
         Index           =   1
         Left            =   2250
         TabIndex        =   989
         Top             =   585
         Width           =   555
      End
      Begin VB.Label lblMesin 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "02"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   465
         Index           =   2
         Left            =   90
         TabIndex        =   988
         Top             =   135
         Width           =   510
      End
   End
   Begin VB.Frame frameDesign 
      BackColor       =   &H00FFFFFF&
      Height          =   2175
      Index           =   2
      Left            =   0
      TabIndex        =   951
      Top             =   5220
      Width           =   2850
      Begin VB.TextBox txtInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   2
         Left            =   630
         MultiLine       =   -1  'True
         TabIndex        =   963
         Top             =   135
         Width           =   2175
      End
      Begin VB.TextBox txtidle 
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
         Index           =   6
         Left            =   630
         TabIndex        =   962
         Top             =   1710
         Width           =   735
      End
      Begin VB.TextBox txtShot 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   8
         Left            =   630
         TabIndex        =   961
         Top             =   810
         Width           =   555
      End
      Begin VB.TextBox txtShot 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   9
         Left            =   1170
         TabIndex        =   960
         Top             =   810
         Width           =   555
      End
      Begin VB.TextBox txtShot 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   10
         Left            =   1710
         TabIndex        =   959
         Top             =   810
         Width           =   555
      End
      Begin VB.TextBox txtShot 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   11
         Left            =   2250
         TabIndex        =   958
         Top             =   810
         Width           =   555
      End
      Begin VB.TextBox txtidle 
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
         Index           =   7
         Left            =   1350
         TabIndex        =   957
         Top             =   1710
         Width           =   735
      End
      Begin VB.TextBox txtidle 
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
         Index           =   8
         Left            =   2070
         TabIndex        =   956
         Top             =   1710
         Width           =   735
      End
      Begin VB.TextBox txtng 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   8
         Left            =   630
         TabIndex        =   955
         Top             =   1080
         Width           =   555
      End
      Begin VB.TextBox txtng 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   9
         Left            =   1170
         TabIndex        =   954
         Top             =   1080
         Width           =   555
      End
      Begin VB.TextBox txtng 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   10
         Left            =   1710
         TabIndex        =   953
         Top             =   1080
         Width           =   555
      End
      Begin VB.TextBox txtng 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   11
         Left            =   2250
         TabIndex        =   952
         Top             =   1080
         Width           =   555
      End
      Begin VB.Label lblMesin 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "03"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   465
         Index           =   3
         Left            =   90
         TabIndex        =   974
         Top             =   135
         Width           =   510
      End
      Begin VB.Label lbljam2R 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "10 : R"
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
         Height          =   240
         Index           =   2
         Left            =   2250
         TabIndex        =   973
         Top             =   585
         Width           =   555
      End
      Begin VB.Label lbljam2L 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "10 : L"
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
         Height          =   240
         Index           =   2
         Left            =   1710
         TabIndex        =   972
         Top             =   585
         Width           =   555
      End
      Begin VB.Label lbljam1R 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "9 : R"
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
         Height          =   240
         Index           =   2
         Left            =   1170
         TabIndex        =   971
         Top             =   585
         Width           =   555
      End
      Begin VB.Label lbljam1L 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "9 : L"
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
         Height          =   240
         Index           =   2
         Left            =   630
         TabIndex        =   970
         Top             =   585
         Width           =   555
      End
      Begin VB.Label lblIdle 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "        IDLE"
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
         Height          =   600
         Index           =   3
         Left            =   90
         TabIndex        =   969
         Top             =   1485
         Width           =   510
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "NG"
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
         Height          =   330
         Index           =   13
         Left            =   90
         TabIndex        =   968
         Top             =   1125
         Width           =   510
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "GROSS"
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
         Height          =   285
         Index           =   14
         Left            =   90
         TabIndex        =   967
         Top             =   810
         Width           =   510
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SHIFT 3"
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
         Height          =   240
         Index           =   15
         Left            =   2070
         TabIndex        =   966
         Top             =   1485
         Width           =   735
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SHIFT 2"
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
         Height          =   240
         Index           =   16
         Left            =   1350
         TabIndex        =   965
         Top             =   1485
         Width           =   735
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SHIFT 1"
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
         Height          =   240
         Index           =   17
         Left            =   630
         TabIndex        =   964
         Top             =   1485
         Width           =   735
      End
   End
   Begin VB.Frame frameDesign 
      BackColor       =   &H00FFFFFF&
      Height          =   2175
      Index           =   3
      Left            =   0
      TabIndex        =   927
      Top             =   7515
      Width           =   2850
      Begin VB.TextBox txtInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   3
         Left            =   630
         MultiLine       =   -1  'True
         TabIndex        =   939
         Top             =   135
         Width           =   2175
      End
      Begin VB.TextBox txtidle 
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
         Index           =   9
         Left            =   630
         TabIndex        =   938
         Top             =   1710
         Width           =   735
      End
      Begin VB.TextBox txtShot 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   12
         Left            =   630
         TabIndex        =   937
         Top             =   810
         Width           =   555
      End
      Begin VB.TextBox txtShot 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   13
         Left            =   1170
         TabIndex        =   936
         Top             =   810
         Width           =   555
      End
      Begin VB.TextBox txtShot 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   14
         Left            =   1710
         TabIndex        =   935
         Top             =   810
         Width           =   555
      End
      Begin VB.TextBox txtShot 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   15
         Left            =   2250
         TabIndex        =   934
         Top             =   810
         Width           =   555
      End
      Begin VB.TextBox txtidle 
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
         Index           =   10
         Left            =   1350
         TabIndex        =   933
         Top             =   1710
         Width           =   735
      End
      Begin VB.TextBox txtidle 
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
         Index           =   11
         Left            =   2070
         TabIndex        =   932
         Top             =   1710
         Width           =   735
      End
      Begin VB.TextBox txtng 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   12
         Left            =   630
         TabIndex        =   931
         Top             =   1080
         Width           =   555
      End
      Begin VB.TextBox txtng 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   13
         Left            =   1170
         TabIndex        =   930
         Top             =   1080
         Width           =   555
      End
      Begin VB.TextBox txtng 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   14
         Left            =   1710
         TabIndex        =   929
         Top             =   1080
         Width           =   555
      End
      Begin VB.TextBox txtng 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   15
         Left            =   2250
         TabIndex        =   928
         Top             =   1080
         Width           =   555
      End
      Begin VB.Label lblMesin 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "04"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   465
         Index           =   4
         Left            =   90
         TabIndex        =   950
         Top             =   135
         Width           =   510
      End
      Begin VB.Label lbljam2R 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "10 : R"
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
         Height          =   240
         Index           =   3
         Left            =   2250
         TabIndex        =   949
         Top             =   585
         Width           =   555
      End
      Begin VB.Label lbljam2L 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "10 : L"
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
         Height          =   240
         Index           =   3
         Left            =   1710
         TabIndex        =   948
         Top             =   585
         Width           =   555
      End
      Begin VB.Label lbljam1R 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "9 : R"
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
         Height          =   240
         Index           =   3
         Left            =   1170
         TabIndex        =   947
         Top             =   585
         Width           =   555
      End
      Begin VB.Label lbljam1L 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "9 : L"
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
         Height          =   240
         Index           =   3
         Left            =   630
         TabIndex        =   946
         Top             =   585
         Width           =   555
      End
      Begin VB.Label lblIdle 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "        IDLE"
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
         Height          =   600
         Index           =   4
         Left            =   90
         TabIndex        =   945
         Top             =   1485
         Width           =   510
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "NG"
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
         Height          =   330
         Index           =   19
         Left            =   90
         TabIndex        =   944
         Top             =   1125
         Width           =   510
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "GROSS"
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
         Height          =   285
         Index           =   20
         Left            =   90
         TabIndex        =   943
         Top             =   810
         Width           =   510
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SHIFT 3"
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
         Height          =   240
         Index           =   21
         Left            =   2070
         TabIndex        =   942
         Top             =   1485
         Width           =   735
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SHIFT 2"
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
         Height          =   240
         Index           =   22
         Left            =   1350
         TabIndex        =   941
         Top             =   1485
         Width           =   735
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SHIFT 1"
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
         Height          =   240
         Index           =   23
         Left            =   630
         TabIndex        =   940
         Top             =   1485
         Width           =   735
      End
   End
   Begin VB.Frame frameDesign 
      BackColor       =   &H00FFFFFF&
      Height          =   2175
      Index           =   4
      Left            =   0
      TabIndex        =   903
      Top             =   9810
      Width           =   2850
      Begin VB.TextBox txtInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   4
         Left            =   630
         MultiLine       =   -1  'True
         TabIndex        =   915
         Top             =   135
         Width           =   2175
      End
      Begin VB.TextBox txtidle 
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
         Index           =   12
         Left            =   630
         TabIndex        =   914
         Top             =   1710
         Width           =   735
      End
      Begin VB.TextBox txtShot 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   16
         Left            =   630
         TabIndex        =   913
         Top             =   810
         Width           =   555
      End
      Begin VB.TextBox txtShot 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   17
         Left            =   1170
         TabIndex        =   912
         Top             =   810
         Width           =   555
      End
      Begin VB.TextBox txtShot 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   18
         Left            =   1710
         TabIndex        =   911
         Top             =   810
         Width           =   555
      End
      Begin VB.TextBox txtShot 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   19
         Left            =   2250
         TabIndex        =   910
         Top             =   810
         Width           =   555
      End
      Begin VB.TextBox txtidle 
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
         Index           =   13
         Left            =   1350
         TabIndex        =   909
         Top             =   1710
         Width           =   735
      End
      Begin VB.TextBox txtidle 
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
         Index           =   14
         Left            =   2070
         TabIndex        =   908
         Top             =   1710
         Width           =   735
      End
      Begin VB.TextBox txtng 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   16
         Left            =   630
         TabIndex        =   907
         Top             =   1080
         Width           =   555
      End
      Begin VB.TextBox txtng 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   17
         Left            =   1170
         TabIndex        =   906
         Top             =   1080
         Width           =   555
      End
      Begin VB.TextBox txtng 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   18
         Left            =   1710
         TabIndex        =   905
         Top             =   1080
         Width           =   555
      End
      Begin VB.TextBox txtng 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   19
         Left            =   2250
         TabIndex        =   904
         Top             =   1080
         Width           =   555
      End
      Begin VB.Label lblMesin 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "05"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   465
         Index           =   5
         Left            =   90
         TabIndex        =   926
         Top             =   135
         Width           =   510
      End
      Begin VB.Label lbljam2R 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "10 : R"
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
         Height          =   240
         Index           =   4
         Left            =   2250
         TabIndex        =   925
         Top             =   585
         Width           =   555
      End
      Begin VB.Label lbljam2L 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "10 : L"
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
         Height          =   240
         Index           =   4
         Left            =   1710
         TabIndex        =   924
         Top             =   585
         Width           =   555
      End
      Begin VB.Label lbljam1R 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "9 : R"
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
         Height          =   240
         Index           =   4
         Left            =   1170
         TabIndex        =   923
         Top             =   585
         Width           =   555
      End
      Begin VB.Label lbljam1L 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "9 : L"
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
         Height          =   240
         Index           =   4
         Left            =   630
         TabIndex        =   922
         Top             =   585
         Width           =   555
      End
      Begin VB.Label lblIdle 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "        IDLE"
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
         Height          =   600
         Index           =   5
         Left            =   90
         TabIndex        =   921
         Top             =   1485
         Width           =   510
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "NG"
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
         Height          =   330
         Index           =   25
         Left            =   90
         TabIndex        =   920
         Top             =   1125
         Width           =   510
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "GROSS"
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
         Height          =   285
         Index           =   26
         Left            =   90
         TabIndex        =   919
         Top             =   810
         Width           =   510
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SHIFT 3"
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
         Height          =   240
         Index           =   27
         Left            =   2070
         TabIndex        =   918
         Top             =   1485
         Width           =   735
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SHIFT 2"
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
         Height          =   240
         Index           =   28
         Left            =   1350
         TabIndex        =   917
         Top             =   1485
         Width           =   735
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SHIFT 1"
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
         Height          =   240
         Index           =   29
         Left            =   630
         TabIndex        =   916
         Top             =   1485
         Width           =   735
      End
   End
   Begin VB.Frame frameDesign 
      BackColor       =   &H00FFFFFF&
      Height          =   2175
      Index           =   5
      Left            =   2970
      TabIndex        =   879
      Top             =   630
      Width           =   2850
      Begin VB.TextBox txtng 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   20
         Left            =   2250
         TabIndex        =   891
         Top             =   1080
         Width           =   555
      End
      Begin VB.TextBox txtng 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   21
         Left            =   1710
         TabIndex        =   890
         Top             =   1080
         Width           =   555
      End
      Begin VB.TextBox txtng 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   22
         Left            =   1170
         TabIndex        =   889
         Top             =   1080
         Width           =   555
      End
      Begin VB.TextBox txtng 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   23
         Left            =   630
         TabIndex        =   888
         Top             =   1080
         Width           =   555
      End
      Begin VB.TextBox txtidle 
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
         Index           =   15
         Left            =   2070
         TabIndex        =   887
         Top             =   1710
         Width           =   735
      End
      Begin VB.TextBox txtidle 
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
         Index           =   16
         Left            =   1350
         TabIndex        =   886
         Top             =   1710
         Width           =   735
      End
      Begin VB.TextBox txtShot 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   20
         Left            =   2250
         TabIndex        =   885
         Top             =   810
         Width           =   555
      End
      Begin VB.TextBox txtShot 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   21
         Left            =   1710
         TabIndex        =   884
         Top             =   810
         Width           =   555
      End
      Begin VB.TextBox txtShot 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   22
         Left            =   1170
         TabIndex        =   883
         Top             =   810
         Width           =   555
      End
      Begin VB.TextBox txtShot 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   23
         Left            =   630
         TabIndex        =   882
         Top             =   810
         Width           =   555
      End
      Begin VB.TextBox txtidle 
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
         Index           =   17
         Left            =   630
         TabIndex        =   881
         Top             =   1710
         Width           =   735
      End
      Begin VB.TextBox txtInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   5
         Left            =   630
         MultiLine       =   -1  'True
         TabIndex        =   880
         Top             =   135
         Width           =   2175
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SHIFT 1"
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
         Height          =   240
         Index           =   30
         Left            =   630
         TabIndex        =   902
         Top             =   1485
         Width           =   735
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SHIFT 2"
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
         Height          =   240
         Index           =   31
         Left            =   1350
         TabIndex        =   901
         Top             =   1485
         Width           =   735
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SHIFT 3"
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
         Height          =   240
         Index           =   32
         Left            =   2070
         TabIndex        =   900
         Top             =   1485
         Width           =   735
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "GROSS"
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
         Height          =   285
         Index           =   33
         Left            =   90
         TabIndex        =   899
         Top             =   810
         Width           =   510
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "NG"
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
         Height          =   330
         Index           =   34
         Left            =   90
         TabIndex        =   898
         Top             =   1125
         Width           =   510
      End
      Begin VB.Label lblIdle 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "        IDLE"
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
         Height          =   600
         Index           =   6
         Left            =   90
         TabIndex        =   897
         Top             =   1485
         Width           =   510
      End
      Begin VB.Label lbljam1L 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "9 : L"
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
         Height          =   240
         Index           =   5
         Left            =   630
         TabIndex        =   896
         Top             =   585
         Width           =   555
      End
      Begin VB.Label lbljam1R 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "9 : R"
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
         Height          =   240
         Index           =   5
         Left            =   1170
         TabIndex        =   895
         Top             =   585
         Width           =   555
      End
      Begin VB.Label lbljam2L 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "10 : L"
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
         Height          =   240
         Index           =   5
         Left            =   1710
         TabIndex        =   894
         Top             =   585
         Width           =   555
      End
      Begin VB.Label lbljam2R 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "10 : R"
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
         Height          =   240
         Index           =   5
         Left            =   2250
         TabIndex        =   893
         Top             =   585
         Width           =   555
      End
      Begin VB.Label lblMesin 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "06"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   465
         Index           =   6
         Left            =   90
         TabIndex        =   892
         Top             =   135
         Width           =   510
      End
   End
   Begin VB.Frame frameDesign 
      BackColor       =   &H00FFFFFF&
      Height          =   2175
      Index           =   6
      Left            =   2970
      TabIndex        =   855
      Top             =   2925
      Width           =   2850
      Begin VB.TextBox txtInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   6
         Left            =   630
         MultiLine       =   -1  'True
         TabIndex        =   867
         Top             =   135
         Width           =   2175
      End
      Begin VB.TextBox txtidle 
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
         Index           =   18
         Left            =   630
         TabIndex        =   866
         Top             =   1710
         Width           =   735
      End
      Begin VB.TextBox txtShot 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   24
         Left            =   630
         TabIndex        =   865
         Top             =   810
         Width           =   555
      End
      Begin VB.TextBox txtShot 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   25
         Left            =   1170
         TabIndex        =   864
         Top             =   810
         Width           =   555
      End
      Begin VB.TextBox txtShot 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   26
         Left            =   1710
         TabIndex        =   863
         Top             =   810
         Width           =   555
      End
      Begin VB.TextBox txtShot 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   27
         Left            =   2250
         TabIndex        =   862
         Top             =   810
         Width           =   555
      End
      Begin VB.TextBox txtidle 
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
         Index           =   19
         Left            =   1350
         TabIndex        =   861
         Top             =   1710
         Width           =   735
      End
      Begin VB.TextBox txtidle 
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
         Index           =   20
         Left            =   2070
         TabIndex        =   860
         Top             =   1710
         Width           =   735
      End
      Begin VB.TextBox txtng 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   24
         Left            =   630
         TabIndex        =   859
         Top             =   1080
         Width           =   555
      End
      Begin VB.TextBox txtng 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   25
         Left            =   1170
         TabIndex        =   858
         Top             =   1080
         Width           =   555
      End
      Begin VB.TextBox txtng 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   26
         Left            =   1710
         TabIndex        =   857
         Top             =   1080
         Width           =   555
      End
      Begin VB.TextBox txtng 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   27
         Left            =   2250
         TabIndex        =   856
         Top             =   1080
         Width           =   555
      End
      Begin VB.Label lblMesin 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "07"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   465
         Index           =   7
         Left            =   90
         TabIndex        =   878
         Top             =   135
         Width           =   510
      End
      Begin VB.Label lbljam2R 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "10 : R"
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
         Height          =   240
         Index           =   6
         Left            =   2250
         TabIndex        =   877
         Top             =   585
         Width           =   555
      End
      Begin VB.Label lbljam2L 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "10 : L"
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
         Height          =   240
         Index           =   6
         Left            =   1710
         TabIndex        =   876
         Top             =   585
         Width           =   555
      End
      Begin VB.Label lbljam1R 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "9 : R"
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
         Height          =   240
         Index           =   6
         Left            =   1170
         TabIndex        =   875
         Top             =   585
         Width           =   555
      End
      Begin VB.Label lbljam1L 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "9 : L"
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
         Height          =   240
         Index           =   6
         Left            =   630
         TabIndex        =   874
         Top             =   585
         Width           =   555
      End
      Begin VB.Label lblIdle 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "        IDLE"
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
         Height          =   600
         Index           =   7
         Left            =   90
         TabIndex        =   873
         Top             =   1485
         Width           =   510
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "NG"
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
         Height          =   330
         Index           =   37
         Left            =   90
         TabIndex        =   872
         Top             =   1125
         Width           =   510
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "GROSS"
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
         Height          =   285
         Index           =   38
         Left            =   90
         TabIndex        =   871
         Top             =   810
         Width           =   510
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SHIFT 3"
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
         Height          =   240
         Index           =   39
         Left            =   2070
         TabIndex        =   870
         Top             =   1485
         Width           =   735
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SHIFT 2"
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
         Height          =   240
         Index           =   40
         Left            =   1350
         TabIndex        =   869
         Top             =   1485
         Width           =   735
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SHIFT 1"
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
         Height          =   240
         Index           =   41
         Left            =   630
         TabIndex        =   868
         Top             =   1485
         Width           =   735
      End
   End
   Begin VB.Frame frameDesign 
      BackColor       =   &H00FFFFFF&
      Height          =   2175
      Index           =   7
      Left            =   2970
      TabIndex        =   831
      Top             =   5220
      Width           =   2850
      Begin VB.TextBox txtng 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   28
         Left            =   2250
         TabIndex        =   843
         Top             =   1080
         Width           =   555
      End
      Begin VB.TextBox txtng 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   29
         Left            =   1710
         TabIndex        =   842
         Top             =   1080
         Width           =   555
      End
      Begin VB.TextBox txtng 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   30
         Left            =   1170
         TabIndex        =   841
         Top             =   1080
         Width           =   555
      End
      Begin VB.TextBox txtng 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   31
         Left            =   630
         TabIndex        =   840
         Top             =   1080
         Width           =   555
      End
      Begin VB.TextBox txtidle 
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
         Index           =   21
         Left            =   2070
         TabIndex        =   839
         Top             =   1710
         Width           =   735
      End
      Begin VB.TextBox txtidle 
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
         Index           =   22
         Left            =   1350
         TabIndex        =   838
         Top             =   1710
         Width           =   735
      End
      Begin VB.TextBox txtShot 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   28
         Left            =   2250
         TabIndex        =   837
         Top             =   810
         Width           =   555
      End
      Begin VB.TextBox txtShot 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   29
         Left            =   1710
         TabIndex        =   836
         Top             =   810
         Width           =   555
      End
      Begin VB.TextBox txtShot 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   30
         Left            =   1170
         TabIndex        =   835
         Top             =   810
         Width           =   555
      End
      Begin VB.TextBox txtShot 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   31
         Left            =   630
         TabIndex        =   834
         Top             =   810
         Width           =   555
      End
      Begin VB.TextBox txtidle 
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
         Index           =   23
         Left            =   630
         TabIndex        =   833
         Top             =   1710
         Width           =   735
      End
      Begin VB.TextBox txtInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   7
         Left            =   630
         MultiLine       =   -1  'True
         TabIndex        =   832
         Top             =   135
         Width           =   2175
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SHIFT 1"
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
         Height          =   240
         Index           =   42
         Left            =   630
         TabIndex        =   854
         Top             =   1485
         Width           =   735
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SHIFT 2"
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
         Height          =   240
         Index           =   43
         Left            =   1350
         TabIndex        =   853
         Top             =   1485
         Width           =   735
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SHIFT 3"
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
         Height          =   240
         Index           =   44
         Left            =   2070
         TabIndex        =   852
         Top             =   1485
         Width           =   735
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "GROSS"
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
         Height          =   285
         Index           =   45
         Left            =   90
         TabIndex        =   851
         Top             =   810
         Width           =   510
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "NG"
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
         Height          =   330
         Index           =   46
         Left            =   90
         TabIndex        =   850
         Top             =   1125
         Width           =   510
      End
      Begin VB.Label lblIdle 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "        IDLE"
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
         Height          =   600
         Index           =   8
         Left            =   90
         TabIndex        =   849
         Top             =   1485
         Width           =   510
      End
      Begin VB.Label lbljam1L 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "9 : L"
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
         Height          =   240
         Index           =   7
         Left            =   630
         TabIndex        =   848
         Top             =   585
         Width           =   555
      End
      Begin VB.Label lbljam1R 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "9 : R"
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
         Height          =   240
         Index           =   7
         Left            =   1170
         TabIndex        =   847
         Top             =   585
         Width           =   555
      End
      Begin VB.Label lbljam2L 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "10 : L"
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
         Height          =   240
         Index           =   7
         Left            =   1710
         TabIndex        =   846
         Top             =   585
         Width           =   555
      End
      Begin VB.Label lbljam2R 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "10 : R"
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
         Height          =   240
         Index           =   7
         Left            =   2250
         TabIndex        =   845
         Top             =   585
         Width           =   555
      End
      Begin VB.Label lblMesin 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "08"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   465
         Index           =   8
         Left            =   90
         TabIndex        =   844
         Top             =   135
         Width           =   510
      End
   End
   Begin VB.Frame frameDesign 
      BackColor       =   &H00FFFFFF&
      Height          =   2175
      Index           =   8
      Left            =   2970
      TabIndex        =   807
      Top             =   7515
      Width           =   2850
      Begin VB.TextBox txtng 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   32
         Left            =   2250
         TabIndex        =   819
         Top             =   1080
         Width           =   555
      End
      Begin VB.TextBox txtng 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   33
         Left            =   1710
         TabIndex        =   818
         Top             =   1080
         Width           =   555
      End
      Begin VB.TextBox txtng 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   34
         Left            =   1170
         TabIndex        =   817
         Top             =   1080
         Width           =   555
      End
      Begin VB.TextBox txtng 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   35
         Left            =   630
         TabIndex        =   816
         Top             =   1080
         Width           =   555
      End
      Begin VB.TextBox txtidle 
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
         Index           =   24
         Left            =   2070
         TabIndex        =   815
         Top             =   1710
         Width           =   735
      End
      Begin VB.TextBox txtidle 
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
         Index           =   25
         Left            =   1350
         TabIndex        =   814
         Top             =   1710
         Width           =   735
      End
      Begin VB.TextBox txtShot 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   32
         Left            =   2250
         TabIndex        =   813
         Top             =   810
         Width           =   555
      End
      Begin VB.TextBox txtShot 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   33
         Left            =   1710
         TabIndex        =   812
         Top             =   810
         Width           =   555
      End
      Begin VB.TextBox txtShot 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   34
         Left            =   1170
         TabIndex        =   811
         Top             =   810
         Width           =   555
      End
      Begin VB.TextBox txtShot 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   35
         Left            =   630
         TabIndex        =   810
         Top             =   810
         Width           =   555
      End
      Begin VB.TextBox txtidle 
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
         Index           =   26
         Left            =   630
         TabIndex        =   809
         Top             =   1710
         Width           =   735
      End
      Begin VB.TextBox txtInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   8
         Left            =   630
         MultiLine       =   -1  'True
         TabIndex        =   808
         Top             =   135
         Width           =   2175
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SHIFT 1"
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
         Height          =   240
         Index           =   48
         Left            =   630
         TabIndex        =   830
         Top             =   1485
         Width           =   735
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SHIFT 2"
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
         Height          =   240
         Index           =   49
         Left            =   1350
         TabIndex        =   829
         Top             =   1485
         Width           =   735
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SHIFT 3"
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
         Height          =   240
         Index           =   50
         Left            =   2070
         TabIndex        =   828
         Top             =   1485
         Width           =   735
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "GROSS"
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
         Height          =   285
         Index           =   51
         Left            =   90
         TabIndex        =   827
         Top             =   810
         Width           =   510
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "NG"
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
         Height          =   330
         Index           =   52
         Left            =   90
         TabIndex        =   826
         Top             =   1125
         Width           =   510
      End
      Begin VB.Label lblIdle 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "        IDLE"
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
         Height          =   600
         Index           =   9
         Left            =   90
         TabIndex        =   825
         Top             =   1485
         Width           =   510
      End
      Begin VB.Label lbljam1L 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "9 : L"
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
         Height          =   240
         Index           =   8
         Left            =   630
         TabIndex        =   824
         Top             =   585
         Width           =   555
      End
      Begin VB.Label lbljam1R 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "9 : R"
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
         Height          =   240
         Index           =   8
         Left            =   1170
         TabIndex        =   823
         Top             =   585
         Width           =   555
      End
      Begin VB.Label lbljam2L 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "10 : L"
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
         Height          =   240
         Index           =   8
         Left            =   1710
         TabIndex        =   822
         Top             =   585
         Width           =   555
      End
      Begin VB.Label lbljam2R 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "10 : R"
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
         Height          =   240
         Index           =   8
         Left            =   2250
         TabIndex        =   821
         Top             =   585
         Width           =   555
      End
      Begin VB.Label lblMesin 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "09"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   465
         Index           =   9
         Left            =   90
         TabIndex        =   820
         Top             =   135
         Width           =   510
      End
   End
   Begin VB.Frame frameDesign 
      BackColor       =   &H00FFFFFF&
      Height          =   2175
      Index           =   9
      Left            =   2970
      TabIndex        =   783
      Top             =   9810
      Width           =   2850
      Begin VB.TextBox txtInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   9
         Left            =   630
         MultiLine       =   -1  'True
         TabIndex        =   795
         Top             =   135
         Width           =   2175
      End
      Begin VB.TextBox txtidle 
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
         Index           =   27
         Left            =   630
         TabIndex        =   794
         Top             =   1710
         Width           =   735
      End
      Begin VB.TextBox txtShot 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   36
         Left            =   630
         TabIndex        =   793
         Top             =   810
         Width           =   555
      End
      Begin VB.TextBox txtShot 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   37
         Left            =   1170
         TabIndex        =   792
         Top             =   810
         Width           =   555
      End
      Begin VB.TextBox txtShot 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   38
         Left            =   1710
         TabIndex        =   791
         Top             =   810
         Width           =   555
      End
      Begin VB.TextBox txtShot 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   39
         Left            =   2250
         TabIndex        =   790
         Top             =   810
         Width           =   555
      End
      Begin VB.TextBox txtidle 
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
         Index           =   28
         Left            =   1350
         TabIndex        =   789
         Top             =   1710
         Width           =   735
      End
      Begin VB.TextBox txtidle 
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
         Index           =   29
         Left            =   2070
         TabIndex        =   788
         Top             =   1710
         Width           =   735
      End
      Begin VB.TextBox txtng 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   36
         Left            =   630
         TabIndex        =   787
         Top             =   1080
         Width           =   555
      End
      Begin VB.TextBox txtng 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   37
         Left            =   1170
         TabIndex        =   786
         Top             =   1080
         Width           =   555
      End
      Begin VB.TextBox txtng 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   38
         Left            =   1710
         TabIndex        =   785
         Top             =   1080
         Width           =   555
      End
      Begin VB.TextBox txtng 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   39
         Left            =   2250
         TabIndex        =   784
         Top             =   1080
         Width           =   555
      End
      Begin VB.Label lblMesin 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "10"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   465
         Index           =   10
         Left            =   90
         TabIndex        =   806
         Top             =   135
         Width           =   510
      End
      Begin VB.Label lbljam2R 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "10 : R"
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
         Height          =   240
         Index           =   9
         Left            =   2250
         TabIndex        =   805
         Top             =   585
         Width           =   555
      End
      Begin VB.Label lbljam2L 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "10 : L"
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
         Height          =   240
         Index           =   9
         Left            =   1710
         TabIndex        =   804
         Top             =   585
         Width           =   555
      End
      Begin VB.Label lbljam1R 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "9 : R"
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
         Height          =   240
         Index           =   9
         Left            =   1170
         TabIndex        =   803
         Top             =   585
         Width           =   555
      End
      Begin VB.Label lbljam1L 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "9 : L"
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
         Height          =   240
         Index           =   9
         Left            =   630
         TabIndex        =   802
         Top             =   585
         Width           =   555
      End
      Begin VB.Label lblIdle 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "        IDLE"
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
         Height          =   600
         Index           =   10
         Left            =   90
         TabIndex        =   801
         Top             =   1485
         Width           =   510
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "NG"
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
         Height          =   330
         Index           =   55
         Left            =   90
         TabIndex        =   800
         Top             =   1125
         Width           =   510
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "GROSS"
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
         Height          =   285
         Index           =   56
         Left            =   90
         TabIndex        =   799
         Top             =   810
         Width           =   510
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SHIFT 3"
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
         Height          =   240
         Index           =   57
         Left            =   2070
         TabIndex        =   798
         Top             =   1485
         Width           =   735
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SHIFT 2"
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
         Height          =   240
         Index           =   58
         Left            =   1350
         TabIndex        =   797
         Top             =   1485
         Width           =   735
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SHIFT 1"
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
         Height          =   240
         Index           =   59
         Left            =   630
         TabIndex        =   796
         Top             =   1485
         Width           =   735
      End
   End
   Begin VB.Frame frameDesign 
      BackColor       =   &H00FFFFFF&
      Height          =   2175
      Index           =   10
      Left            =   5940
      TabIndex        =   759
      Top             =   630
      Width           =   2850
      Begin VB.TextBox txtInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   10
         Left            =   630
         MultiLine       =   -1  'True
         TabIndex        =   771
         Top             =   135
         Width           =   2175
      End
      Begin VB.TextBox txtidle 
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
         Index           =   30
         Left            =   630
         TabIndex        =   770
         Top             =   1710
         Width           =   735
      End
      Begin VB.TextBox txtShot 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   40
         Left            =   630
         TabIndex        =   769
         Top             =   810
         Width           =   555
      End
      Begin VB.TextBox txtShot 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   41
         Left            =   1170
         TabIndex        =   768
         Top             =   810
         Width           =   555
      End
      Begin VB.TextBox txtShot 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   42
         Left            =   1710
         TabIndex        =   767
         Top             =   810
         Width           =   555
      End
      Begin VB.TextBox txtShot 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   43
         Left            =   2250
         TabIndex        =   766
         Top             =   810
         Width           =   555
      End
      Begin VB.TextBox txtidle 
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
         Index           =   31
         Left            =   1350
         TabIndex        =   765
         Top             =   1710
         Width           =   735
      End
      Begin VB.TextBox txtidle 
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
         Index           =   32
         Left            =   2070
         TabIndex        =   764
         Top             =   1710
         Width           =   735
      End
      Begin VB.TextBox txtng 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   40
         Left            =   630
         TabIndex        =   763
         Top             =   1080
         Width           =   555
      End
      Begin VB.TextBox txtng 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   41
         Left            =   1170
         TabIndex        =   762
         Top             =   1080
         Width           =   555
      End
      Begin VB.TextBox txtng 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   42
         Left            =   1710
         TabIndex        =   761
         Top             =   1080
         Width           =   555
      End
      Begin VB.TextBox txtng 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   43
         Left            =   2250
         TabIndex        =   760
         Top             =   1080
         Width           =   555
      End
      Begin VB.Label lblMesin 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "11"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   465
         Index           =   11
         Left            =   90
         TabIndex        =   782
         Top             =   135
         Width           =   510
      End
      Begin VB.Label lbljam2R 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "10 : R"
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
         Height          =   240
         Index           =   10
         Left            =   2250
         TabIndex        =   781
         Top             =   585
         Width           =   555
      End
      Begin VB.Label lbljam2L 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "10 : L"
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
         Height          =   240
         Index           =   10
         Left            =   1710
         TabIndex        =   780
         Top             =   585
         Width           =   555
      End
      Begin VB.Label lbljam1R 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "9 : R"
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
         Height          =   240
         Index           =   10
         Left            =   1170
         TabIndex        =   779
         Top             =   585
         Width           =   555
      End
      Begin VB.Label lbljam1L 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "9 : L"
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
         Height          =   240
         Index           =   10
         Left            =   630
         TabIndex        =   778
         Top             =   585
         Width           =   555
      End
      Begin VB.Label lblIdle 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "        IDLE"
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
         Height          =   600
         Index           =   11
         Left            =   90
         TabIndex        =   777
         Top             =   1485
         Width           =   510
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "NG"
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
         Height          =   330
         Index           =   61
         Left            =   90
         TabIndex        =   776
         Top             =   1125
         Width           =   510
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "GROSS"
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
         Height          =   285
         Index           =   62
         Left            =   90
         TabIndex        =   775
         Top             =   810
         Width           =   510
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SHIFT 3"
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
         Height          =   240
         Index           =   63
         Left            =   2070
         TabIndex        =   774
         Top             =   1485
         Width           =   735
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SHIFT 2"
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
         Height          =   240
         Index           =   64
         Left            =   1350
         TabIndex        =   773
         Top             =   1485
         Width           =   735
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SHIFT 1"
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
         Height          =   240
         Index           =   65
         Left            =   630
         TabIndex        =   772
         Top             =   1485
         Width           =   735
      End
   End
   Begin VB.Frame frameDesign 
      BackColor       =   &H00FFFFFF&
      Height          =   2175
      Index           =   11
      Left            =   5940
      TabIndex        =   735
      Top             =   2925
      Width           =   2850
      Begin VB.TextBox txtng 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   44
         Left            =   2250
         TabIndex        =   747
         Top             =   1080
         Width           =   555
      End
      Begin VB.TextBox txtng 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   45
         Left            =   1710
         TabIndex        =   746
         Top             =   1080
         Width           =   555
      End
      Begin VB.TextBox txtng 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   46
         Left            =   1170
         TabIndex        =   745
         Top             =   1080
         Width           =   555
      End
      Begin VB.TextBox txtng 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   47
         Left            =   630
         TabIndex        =   744
         Top             =   1080
         Width           =   555
      End
      Begin VB.TextBox txtidle 
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
         Index           =   33
         Left            =   2070
         TabIndex        =   743
         Top             =   1710
         Width           =   735
      End
      Begin VB.TextBox txtidle 
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
         Index           =   34
         Left            =   1350
         TabIndex        =   742
         Top             =   1710
         Width           =   735
      End
      Begin VB.TextBox txtShot 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   44
         Left            =   2250
         TabIndex        =   741
         Top             =   810
         Width           =   555
      End
      Begin VB.TextBox txtShot 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   45
         Left            =   1710
         TabIndex        =   740
         Top             =   810
         Width           =   555
      End
      Begin VB.TextBox txtShot 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   46
         Left            =   1170
         TabIndex        =   739
         Top             =   810
         Width           =   555
      End
      Begin VB.TextBox txtShot 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   47
         Left            =   630
         TabIndex        =   738
         Top             =   810
         Width           =   555
      End
      Begin VB.TextBox txtidle 
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
         Index           =   35
         Left            =   630
         TabIndex        =   737
         Top             =   1710
         Width           =   735
      End
      Begin VB.TextBox txtInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   11
         Left            =   630
         MultiLine       =   -1  'True
         TabIndex        =   736
         Top             =   135
         Width           =   2175
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SHIFT 1"
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
         Height          =   240
         Index           =   66
         Left            =   630
         TabIndex        =   758
         Top             =   1485
         Width           =   735
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SHIFT 2"
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
         Height          =   240
         Index           =   67
         Left            =   1350
         TabIndex        =   757
         Top             =   1485
         Width           =   735
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SHIFT 3"
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
         Height          =   240
         Index           =   68
         Left            =   2070
         TabIndex        =   756
         Top             =   1485
         Width           =   735
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "GROSS"
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
         Height          =   285
         Index           =   69
         Left            =   90
         TabIndex        =   755
         Top             =   810
         Width           =   510
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "NG"
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
         Height          =   330
         Index           =   70
         Left            =   90
         TabIndex        =   754
         Top             =   1125
         Width           =   510
      End
      Begin VB.Label lblIdle 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "        IDLE"
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
         Height          =   600
         Index           =   12
         Left            =   90
         TabIndex        =   753
         Top             =   1485
         Width           =   510
      End
      Begin VB.Label lbljam1L 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "9 : L"
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
         Height          =   240
         Index           =   11
         Left            =   630
         TabIndex        =   752
         Top             =   585
         Width           =   555
      End
      Begin VB.Label lbljam1R 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "9 : R"
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
         Height          =   240
         Index           =   11
         Left            =   1170
         TabIndex        =   751
         Top             =   585
         Width           =   555
      End
      Begin VB.Label lbljam2L 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "10 : L"
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
         Height          =   240
         Index           =   11
         Left            =   1710
         TabIndex        =   750
         Top             =   585
         Width           =   555
      End
      Begin VB.Label lbljam2R 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "10 : R"
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
         Height          =   240
         Index           =   11
         Left            =   2250
         TabIndex        =   749
         Top             =   585
         Width           =   555
      End
      Begin VB.Label lblMesin 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "12"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   465
         Index           =   12
         Left            =   90
         TabIndex        =   748
         Top             =   135
         Width           =   510
      End
   End
   Begin VB.Frame frameDesign 
      BackColor       =   &H00FFFFFF&
      Height          =   2175
      Index           =   12
      Left            =   5940
      TabIndex        =   711
      Top             =   5220
      Width           =   2850
      Begin VB.TextBox txtInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   12
         Left            =   630
         MultiLine       =   -1  'True
         TabIndex        =   723
         Top             =   135
         Width           =   2175
      End
      Begin VB.TextBox txtidle 
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
         Index           =   36
         Left            =   630
         TabIndex        =   722
         Top             =   1710
         Width           =   735
      End
      Begin VB.TextBox txtShot 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   48
         Left            =   630
         TabIndex        =   721
         Top             =   810
         Width           =   555
      End
      Begin VB.TextBox txtShot 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   49
         Left            =   1170
         TabIndex        =   720
         Top             =   810
         Width           =   555
      End
      Begin VB.TextBox txtShot 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   50
         Left            =   1710
         TabIndex        =   719
         Top             =   810
         Width           =   555
      End
      Begin VB.TextBox txtShot 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   51
         Left            =   2250
         TabIndex        =   718
         Top             =   810
         Width           =   555
      End
      Begin VB.TextBox txtidle 
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
         Index           =   37
         Left            =   1350
         TabIndex        =   717
         Top             =   1710
         Width           =   735
      End
      Begin VB.TextBox txtidle 
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
         Index           =   38
         Left            =   2070
         TabIndex        =   716
         Top             =   1710
         Width           =   735
      End
      Begin VB.TextBox txtng 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   48
         Left            =   630
         TabIndex        =   715
         Top             =   1080
         Width           =   555
      End
      Begin VB.TextBox txtng 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   49
         Left            =   1170
         TabIndex        =   714
         Top             =   1080
         Width           =   555
      End
      Begin VB.TextBox txtng 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   50
         Left            =   1710
         TabIndex        =   713
         Top             =   1080
         Width           =   555
      End
      Begin VB.TextBox txtng 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   51
         Left            =   2250
         TabIndex        =   712
         Top             =   1080
         Width           =   555
      End
      Begin VB.Label lblMesin 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "13"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   465
         Index           =   13
         Left            =   90
         TabIndex        =   734
         Top             =   135
         Width           =   510
      End
      Begin VB.Label lbljam2R 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "10 : R"
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
         Height          =   240
         Index           =   12
         Left            =   2250
         TabIndex        =   733
         Top             =   585
         Width           =   555
      End
      Begin VB.Label lbljam2L 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "10 : L"
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
         Height          =   240
         Index           =   12
         Left            =   1710
         TabIndex        =   732
         Top             =   585
         Width           =   555
      End
      Begin VB.Label lbljam1R 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "9 : R"
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
         Height          =   240
         Index           =   12
         Left            =   1170
         TabIndex        =   731
         Top             =   585
         Width           =   555
      End
      Begin VB.Label lbljam1L 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "9 : L"
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
         Height          =   240
         Index           =   12
         Left            =   630
         TabIndex        =   730
         Top             =   585
         Width           =   555
      End
      Begin VB.Label lblIdle 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "        IDLE"
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
         Height          =   600
         Index           =   13
         Left            =   90
         TabIndex        =   729
         Top             =   1440
         Width           =   510
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "NG"
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
         Height          =   330
         Index           =   73
         Left            =   90
         TabIndex        =   728
         Top             =   1080
         Width           =   510
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "GROSS"
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
         Height          =   285
         Index           =   74
         Left            =   90
         TabIndex        =   727
         Top             =   765
         Width           =   510
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SHIFT 3"
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
         Height          =   240
         Index           =   75
         Left            =   2070
         TabIndex        =   726
         Top             =   1485
         Width           =   735
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SHIFT 2"
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
         Height          =   240
         Index           =   76
         Left            =   1350
         TabIndex        =   725
         Top             =   1485
         Width           =   735
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SHIFT 1"
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
         Height          =   240
         Index           =   77
         Left            =   630
         TabIndex        =   724
         Top             =   1485
         Width           =   735
      End
   End
   Begin VB.Frame frameDesign 
      BackColor       =   &H00FFFFFF&
      Height          =   2175
      Index           =   13
      Left            =   5940
      TabIndex        =   687
      Top             =   7515
      Width           =   2850
      Begin VB.TextBox txtInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   13
         Left            =   630
         MultiLine       =   -1  'True
         TabIndex        =   699
         Top             =   135
         Width           =   2175
      End
      Begin VB.TextBox txtidle 
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
         Index           =   39
         Left            =   630
         TabIndex        =   698
         Top             =   1710
         Width           =   735
      End
      Begin VB.TextBox txtShot 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   52
         Left            =   630
         TabIndex        =   697
         Top             =   810
         Width           =   555
      End
      Begin VB.TextBox txtShot 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   53
         Left            =   1170
         TabIndex        =   696
         Top             =   810
         Width           =   555
      End
      Begin VB.TextBox txtShot 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   54
         Left            =   1710
         TabIndex        =   695
         Top             =   810
         Width           =   555
      End
      Begin VB.TextBox txtShot 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   55
         Left            =   2250
         TabIndex        =   694
         Top             =   810
         Width           =   555
      End
      Begin VB.TextBox txtidle 
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
         Index           =   40
         Left            =   1350
         TabIndex        =   693
         Top             =   1710
         Width           =   735
      End
      Begin VB.TextBox txtidle 
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
         Index           =   41
         Left            =   2070
         TabIndex        =   692
         Top             =   1710
         Width           =   735
      End
      Begin VB.TextBox txtng 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   52
         Left            =   630
         TabIndex        =   691
         Top             =   1080
         Width           =   555
      End
      Begin VB.TextBox txtng 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   53
         Left            =   1170
         TabIndex        =   690
         Top             =   1080
         Width           =   555
      End
      Begin VB.TextBox txtng 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   54
         Left            =   1710
         TabIndex        =   689
         Top             =   1080
         Width           =   555
      End
      Begin VB.TextBox txtng 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   55
         Left            =   2250
         TabIndex        =   688
         Top             =   1080
         Width           =   555
      End
      Begin VB.Label lblMesin 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "14"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   465
         Index           =   14
         Left            =   90
         TabIndex        =   710
         Top             =   135
         Width           =   510
      End
      Begin VB.Label lbljam2R 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "10 : R"
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
         Height          =   240
         Index           =   13
         Left            =   2250
         TabIndex        =   709
         Top             =   585
         Width           =   555
      End
      Begin VB.Label lbljam2L 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "10 : L"
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
         Height          =   240
         Index           =   13
         Left            =   1710
         TabIndex        =   708
         Top             =   585
         Width           =   555
      End
      Begin VB.Label lbljam1R 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "9 : R"
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
         Height          =   240
         Index           =   13
         Left            =   1170
         TabIndex        =   707
         Top             =   585
         Width           =   555
      End
      Begin VB.Label lbljam1L 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "9 : L"
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
         Height          =   240
         Index           =   13
         Left            =   630
         TabIndex        =   706
         Top             =   585
         Width           =   555
      End
      Begin VB.Label lblIdle 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "        IDLE"
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
         Height          =   600
         Index           =   14
         Left            =   90
         TabIndex        =   705
         Top             =   1485
         Width           =   510
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "NG"
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
         Height          =   330
         Index           =   79
         Left            =   90
         TabIndex        =   704
         Top             =   1125
         Width           =   510
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "GROSS"
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
         Height          =   285
         Index           =   80
         Left            =   90
         TabIndex        =   703
         Top             =   810
         Width           =   510
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SHIFT 3"
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
         Height          =   240
         Index           =   81
         Left            =   2070
         TabIndex        =   702
         Top             =   1485
         Width           =   735
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SHIFT 2"
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
         Height          =   240
         Index           =   82
         Left            =   1350
         TabIndex        =   701
         Top             =   1485
         Width           =   735
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SHIFT 1"
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
         Height          =   240
         Index           =   83
         Left            =   630
         TabIndex        =   700
         Top             =   1485
         Width           =   735
      End
   End
   Begin VB.Frame frameDesign 
      BackColor       =   &H00FFFFFF&
      Height          =   2175
      Index           =   14
      Left            =   5940
      TabIndex        =   663
      Top             =   9810
      Width           =   2850
      Begin VB.TextBox txtng 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   56
         Left            =   2250
         TabIndex        =   675
         Top             =   1080
         Width           =   555
      End
      Begin VB.TextBox txtng 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   57
         Left            =   1710
         TabIndex        =   674
         Top             =   1080
         Width           =   555
      End
      Begin VB.TextBox txtng 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   58
         Left            =   1170
         TabIndex        =   673
         Top             =   1080
         Width           =   555
      End
      Begin VB.TextBox txtng 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   59
         Left            =   630
         TabIndex        =   672
         Top             =   1080
         Width           =   555
      End
      Begin VB.TextBox txtidle 
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
         Index           =   42
         Left            =   2070
         TabIndex        =   671
         Top             =   1710
         Width           =   735
      End
      Begin VB.TextBox txtidle 
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
         Index           =   43
         Left            =   1350
         TabIndex        =   670
         Top             =   1710
         Width           =   735
      End
      Begin VB.TextBox txtShot 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   56
         Left            =   2250
         TabIndex        =   669
         Top             =   810
         Width           =   555
      End
      Begin VB.TextBox txtShot 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   57
         Left            =   1710
         TabIndex        =   668
         Top             =   810
         Width           =   555
      End
      Begin VB.TextBox txtShot 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   58
         Left            =   1170
         TabIndex        =   667
         Top             =   810
         Width           =   555
      End
      Begin VB.TextBox txtShot 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   59
         Left            =   630
         TabIndex        =   666
         Top             =   810
         Width           =   555
      End
      Begin VB.TextBox txtidle 
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
         Index           =   44
         Left            =   630
         TabIndex        =   665
         Top             =   1710
         Width           =   735
      End
      Begin VB.TextBox txtInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   14
         Left            =   630
         MultiLine       =   -1  'True
         TabIndex        =   664
         Top             =   135
         Width           =   2175
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SHIFT 1"
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
         Height          =   240
         Index           =   84
         Left            =   630
         TabIndex        =   686
         Top             =   1485
         Width           =   735
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SHIFT 2"
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
         Height          =   240
         Index           =   85
         Left            =   1350
         TabIndex        =   685
         Top             =   1485
         Width           =   735
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SHIFT 3"
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
         Height          =   240
         Index           =   86
         Left            =   2070
         TabIndex        =   684
         Top             =   1485
         Width           =   735
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "GROSS"
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
         Height          =   285
         Index           =   87
         Left            =   90
         TabIndex        =   683
         Top             =   810
         Width           =   510
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "NG"
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
         Height          =   330
         Index           =   88
         Left            =   90
         TabIndex        =   682
         Top             =   1125
         Width           =   510
      End
      Begin VB.Label lblIdle 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "        IDLE"
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
         Height          =   600
         Index           =   15
         Left            =   90
         TabIndex        =   681
         Top             =   1485
         Width           =   510
      End
      Begin VB.Label lbljam1L 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "9 : L"
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
         Height          =   240
         Index           =   14
         Left            =   630
         TabIndex        =   680
         Top             =   585
         Width           =   555
      End
      Begin VB.Label lbljam1R 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "9 : R"
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
         Height          =   240
         Index           =   14
         Left            =   1170
         TabIndex        =   679
         Top             =   585
         Width           =   555
      End
      Begin VB.Label lbljam2L 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "10 : L"
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
         Height          =   240
         Index           =   14
         Left            =   1710
         TabIndex        =   678
         Top             =   585
         Width           =   555
      End
      Begin VB.Label lbljam2R 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "10 : R"
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
         Height          =   240
         Index           =   14
         Left            =   2250
         TabIndex        =   677
         Top             =   585
         Width           =   555
      End
      Begin VB.Label lblMesin 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "15"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   465
         Index           =   15
         Left            =   90
         TabIndex        =   676
         Top             =   135
         Width           =   510
      End
   End
   Begin VB.Frame frameDesign 
      BackColor       =   &H00FFFFFF&
      Height          =   2175
      Index           =   15
      Left            =   8910
      TabIndex        =   639
      Top             =   630
      Width           =   2850
      Begin VB.TextBox txtInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   15
         Left            =   630
         MultiLine       =   -1  'True
         TabIndex        =   651
         Top             =   135
         Width           =   2175
      End
      Begin VB.TextBox txtidle 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Index           =   45
         Left            =   630
         TabIndex        =   650
         Top             =   1710
         Width           =   735
      End
      Begin VB.TextBox txtShot 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   60
         Left            =   630
         TabIndex        =   649
         Top             =   810
         Width           =   555
      End
      Begin VB.TextBox txtShot 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   61
         Left            =   1170
         TabIndex        =   648
         Top             =   810
         Width           =   555
      End
      Begin VB.TextBox txtShot 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   62
         Left            =   1710
         TabIndex        =   647
         Top             =   810
         Width           =   555
      End
      Begin VB.TextBox txtShot 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   63
         Left            =   2250
         TabIndex        =   646
         Top             =   810
         Width           =   555
      End
      Begin VB.TextBox txtidle 
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
         Index           =   46
         Left            =   1350
         TabIndex        =   645
         Top             =   1710
         Width           =   735
      End
      Begin VB.TextBox txtidle 
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
         Index           =   47
         Left            =   2070
         TabIndex        =   644
         Top             =   1710
         Width           =   735
      End
      Begin VB.TextBox txtng 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   60
         Left            =   630
         TabIndex        =   643
         Top             =   1080
         Width           =   555
      End
      Begin VB.TextBox txtng 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   61
         Left            =   1170
         TabIndex        =   642
         Top             =   1080
         Width           =   555
      End
      Begin VB.TextBox txtng 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   62
         Left            =   1710
         TabIndex        =   641
         Top             =   1080
         Width           =   555
      End
      Begin VB.TextBox txtng 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   63
         Left            =   2250
         TabIndex        =   640
         Top             =   1080
         Width           =   555
      End
      Begin VB.Label lblMesin 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "16"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   465
         Index           =   16
         Left            =   90
         TabIndex        =   662
         Top             =   135
         Width           =   510
      End
      Begin VB.Label lbljam2R 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "10 : R"
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
         Height          =   240
         Index           =   15
         Left            =   2250
         TabIndex        =   661
         Top             =   585
         Width           =   555
      End
      Begin VB.Label lbljam2L 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "10 : L"
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
         Height          =   240
         Index           =   15
         Left            =   1710
         TabIndex        =   660
         Top             =   585
         Width           =   555
      End
      Begin VB.Label lbljam1R 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "9 : R"
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
         Height          =   240
         Index           =   15
         Left            =   1170
         TabIndex        =   659
         Top             =   585
         Width           =   555
      End
      Begin VB.Label lbljam1L 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "9 : L"
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
         Height          =   240
         Index           =   15
         Left            =   630
         TabIndex        =   658
         Top             =   585
         Width           =   555
      End
      Begin VB.Label lblIdle 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "        IDLE"
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
         Height          =   600
         Index           =   16
         Left            =   90
         TabIndex        =   657
         Top             =   1485
         Width           =   510
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "NG"
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
         Height          =   330
         Index           =   91
         Left            =   90
         TabIndex        =   656
         Top             =   1125
         Width           =   510
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "GROSS"
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
         Height          =   285
         Index           =   92
         Left            =   90
         TabIndex        =   655
         Top             =   810
         Width           =   510
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SHIFT 3"
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
         Height          =   240
         Index           =   93
         Left            =   2070
         TabIndex        =   654
         Top             =   1485
         Width           =   735
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SHIFT 2"
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
         Height          =   240
         Index           =   94
         Left            =   1350
         TabIndex        =   653
         Top             =   1485
         Width           =   735
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SHIFT 1"
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
         Height          =   240
         Index           =   95
         Left            =   630
         TabIndex        =   652
         Top             =   1485
         Width           =   735
      End
   End
   Begin VB.Frame frameDesign 
      BackColor       =   &H00FFFFFF&
      Height          =   2175
      Index           =   16
      Left            =   8910
      TabIndex        =   615
      Top             =   2925
      Width           =   2850
      Begin VB.TextBox txtng 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   64
         Left            =   2250
         TabIndex        =   627
         Top             =   1080
         Width           =   555
      End
      Begin VB.TextBox txtng 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   65
         Left            =   1710
         TabIndex        =   626
         Top             =   1080
         Width           =   555
      End
      Begin VB.TextBox txtng 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   66
         Left            =   1170
         TabIndex        =   625
         Top             =   1080
         Width           =   555
      End
      Begin VB.TextBox txtng 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   67
         Left            =   630
         TabIndex        =   624
         Top             =   1080
         Width           =   555
      End
      Begin VB.TextBox txtidle 
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
         Index           =   48
         Left            =   2070
         TabIndex        =   623
         Top             =   1710
         Width           =   735
      End
      Begin VB.TextBox txtidle 
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
         Index           =   49
         Left            =   1350
         TabIndex        =   622
         Top             =   1710
         Width           =   735
      End
      Begin VB.TextBox txtShot 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   64
         Left            =   2250
         TabIndex        =   621
         Top             =   810
         Width           =   555
      End
      Begin VB.TextBox txtShot 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   65
         Left            =   1710
         TabIndex        =   620
         Top             =   810
         Width           =   555
      End
      Begin VB.TextBox txtShot 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   66
         Left            =   1170
         TabIndex        =   619
         Top             =   810
         Width           =   555
      End
      Begin VB.TextBox txtShot 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   67
         Left            =   630
         TabIndex        =   618
         Top             =   810
         Width           =   555
      End
      Begin VB.TextBox txtidle 
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
         Index           =   50
         Left            =   630
         TabIndex        =   617
         Top             =   1710
         Width           =   735
      End
      Begin VB.TextBox txtInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   16
         Left            =   630
         MultiLine       =   -1  'True
         TabIndex        =   616
         Top             =   135
         Width           =   2175
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SHIFT 1"
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
         Height          =   240
         Index           =   96
         Left            =   630
         TabIndex        =   638
         Top             =   1485
         Width           =   735
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SHIFT 2"
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
         Height          =   240
         Index           =   97
         Left            =   1350
         TabIndex        =   637
         Top             =   1485
         Width           =   735
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SHIFT 3"
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
         Height          =   240
         Index           =   98
         Left            =   2070
         TabIndex        =   636
         Top             =   1485
         Width           =   735
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "GROSS"
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
         Height          =   285
         Index           =   99
         Left            =   90
         TabIndex        =   635
         Top             =   810
         Width           =   510
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "NG"
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
         Height          =   330
         Index           =   100
         Left            =   90
         TabIndex        =   634
         Top             =   1125
         Width           =   510
      End
      Begin VB.Label lblIdle 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "        IDLE"
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
         Height          =   600
         Index           =   17
         Left            =   90
         TabIndex        =   633
         Top             =   1485
         Width           =   510
      End
      Begin VB.Label lbljam1L 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "9 : L"
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
         Height          =   240
         Index           =   16
         Left            =   630
         TabIndex        =   632
         Top             =   585
         Width           =   555
      End
      Begin VB.Label lbljam1R 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "9 : R"
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
         Height          =   240
         Index           =   16
         Left            =   1170
         TabIndex        =   631
         Top             =   585
         Width           =   555
      End
      Begin VB.Label lbljam2L 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "10 : L"
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
         Height          =   240
         Index           =   16
         Left            =   1710
         TabIndex        =   630
         Top             =   585
         Width           =   555
      End
      Begin VB.Label lbljam2R 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "10 : R"
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
         Height          =   240
         Index           =   16
         Left            =   2250
         TabIndex        =   629
         Top             =   585
         Width           =   555
      End
      Begin VB.Label lblMesin 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "17"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   465
         Index           =   17
         Left            =   90
         TabIndex        =   628
         Top             =   135
         Width           =   510
      End
   End
   Begin VB.Frame frameDesign 
      BackColor       =   &H00FFFFFF&
      Height          =   2175
      Index           =   17
      Left            =   8910
      TabIndex        =   591
      Top             =   5220
      Width           =   2850
      Begin VB.TextBox txtInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   17
         Left            =   630
         MultiLine       =   -1  'True
         TabIndex        =   603
         Top             =   135
         Width           =   2175
      End
      Begin VB.TextBox txtidle 
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
         Index           =   51
         Left            =   630
         TabIndex        =   602
         Top             =   1710
         Width           =   735
      End
      Begin VB.TextBox txtShot 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   68
         Left            =   630
         TabIndex        =   601
         Top             =   810
         Width           =   555
      End
      Begin VB.TextBox txtShot 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   69
         Left            =   1170
         TabIndex        =   600
         Top             =   810
         Width           =   555
      End
      Begin VB.TextBox txtShot 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   70
         Left            =   1710
         TabIndex        =   599
         Top             =   810
         Width           =   555
      End
      Begin VB.TextBox txtShot 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   71
         Left            =   2250
         TabIndex        =   598
         Top             =   810
         Width           =   555
      End
      Begin VB.TextBox txtidle 
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
         Index           =   52
         Left            =   1350
         TabIndex        =   597
         Top             =   1710
         Width           =   735
      End
      Begin VB.TextBox txtidle 
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
         Index           =   53
         Left            =   2070
         TabIndex        =   596
         Top             =   1710
         Width           =   735
      End
      Begin VB.TextBox txtng 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   68
         Left            =   630
         TabIndex        =   595
         Top             =   1080
         Width           =   555
      End
      Begin VB.TextBox txtng 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   69
         Left            =   1170
         TabIndex        =   594
         Top             =   1080
         Width           =   555
      End
      Begin VB.TextBox txtng 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   70
         Left            =   1710
         TabIndex        =   593
         Top             =   1080
         Width           =   555
      End
      Begin VB.TextBox txtng 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   71
         Left            =   2250
         TabIndex        =   592
         Top             =   1080
         Width           =   555
      End
      Begin VB.Label lblMesin 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "18"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   465
         Index           =   18
         Left            =   90
         TabIndex        =   614
         Top             =   135
         Width           =   510
      End
      Begin VB.Label lbljam2R 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "10 : R"
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
         Height          =   240
         Index           =   17
         Left            =   2250
         TabIndex        =   613
         Top             =   585
         Width           =   555
      End
      Begin VB.Label lbljam2L 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "10 : L"
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
         Height          =   240
         Index           =   17
         Left            =   1710
         TabIndex        =   612
         Top             =   585
         Width           =   555
      End
      Begin VB.Label lbljam1R 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "9 : R"
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
         Height          =   240
         Index           =   17
         Left            =   1170
         TabIndex        =   611
         Top             =   585
         Width           =   555
      End
      Begin VB.Label lbljam1L 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "9 : L"
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
         Height          =   240
         Index           =   17
         Left            =   630
         TabIndex        =   610
         Top             =   585
         Width           =   555
      End
      Begin VB.Label lblIdle 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "        IDLE"
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
         Height          =   600
         Index           =   18
         Left            =   90
         TabIndex        =   609
         Top             =   1485
         Width           =   510
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "NG"
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
         Height          =   330
         Index           =   103
         Left            =   90
         TabIndex        =   608
         Top             =   1125
         Width           =   510
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "GROSS"
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
         Height          =   285
         Index           =   104
         Left            =   90
         TabIndex        =   607
         Top             =   810
         Width           =   510
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SHIFT 3"
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
         Height          =   240
         Index           =   105
         Left            =   2070
         TabIndex        =   606
         Top             =   1485
         Width           =   735
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SHIFT 2"
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
         Height          =   240
         Index           =   106
         Left            =   1350
         TabIndex        =   605
         Top             =   1485
         Width           =   735
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SHIFT 1"
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
         Height          =   240
         Index           =   107
         Left            =   630
         TabIndex        =   604
         Top             =   1485
         Width           =   735
      End
   End
   Begin VB.Frame frameDesign 
      BackColor       =   &H00FFFFFF&
      Height          =   2175
      Index           =   18
      Left            =   8910
      TabIndex        =   567
      Top             =   7515
      Width           =   2850
      Begin VB.TextBox txtInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   18
         Left            =   630
         MultiLine       =   -1  'True
         TabIndex        =   579
         Top             =   135
         Width           =   2175
      End
      Begin VB.TextBox txtidle 
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
         Index           =   54
         Left            =   630
         TabIndex        =   578
         Top             =   1710
         Width           =   735
      End
      Begin VB.TextBox txtShot 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   72
         Left            =   630
         TabIndex        =   577
         Top             =   810
         Width           =   555
      End
      Begin VB.TextBox txtShot 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   73
         Left            =   1170
         TabIndex        =   576
         Top             =   810
         Width           =   555
      End
      Begin VB.TextBox txtShot 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   74
         Left            =   1710
         TabIndex        =   575
         Top             =   810
         Width           =   555
      End
      Begin VB.TextBox txtShot 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   75
         Left            =   2250
         TabIndex        =   574
         Top             =   810
         Width           =   555
      End
      Begin VB.TextBox txtidle 
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
         Index           =   55
         Left            =   1350
         TabIndex        =   573
         Top             =   1710
         Width           =   735
      End
      Begin VB.TextBox txtidle 
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
         Index           =   56
         Left            =   2070
         TabIndex        =   572
         Top             =   1710
         Width           =   735
      End
      Begin VB.TextBox txtng 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   72
         Left            =   630
         TabIndex        =   571
         Top             =   1080
         Width           =   555
      End
      Begin VB.TextBox txtng 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   73
         Left            =   1170
         TabIndex        =   570
         Top             =   1080
         Width           =   555
      End
      Begin VB.TextBox txtng 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   74
         Left            =   1710
         TabIndex        =   569
         Top             =   1080
         Width           =   555
      End
      Begin VB.TextBox txtng 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   75
         Left            =   2250
         TabIndex        =   568
         Top             =   1080
         Width           =   555
      End
      Begin VB.Label lblMesin 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "19"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   465
         Index           =   19
         Left            =   90
         TabIndex        =   590
         Top             =   135
         Width           =   510
      End
      Begin VB.Label lbljam2R 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "10 : R"
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
         Height          =   240
         Index           =   18
         Left            =   2250
         TabIndex        =   589
         Top             =   585
         Width           =   555
      End
      Begin VB.Label lbljam2L 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "10 : L"
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
         Height          =   240
         Index           =   18
         Left            =   1710
         TabIndex        =   588
         Top             =   585
         Width           =   555
      End
      Begin VB.Label lbljam1R 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "9 : R"
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
         Height          =   240
         Index           =   18
         Left            =   1170
         TabIndex        =   587
         Top             =   585
         Width           =   555
      End
      Begin VB.Label lbljam1L 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "9 : L"
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
         Height          =   240
         Index           =   18
         Left            =   630
         TabIndex        =   586
         Top             =   585
         Width           =   555
      End
      Begin VB.Label lblIdle 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "           IDLE"
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
         Height          =   600
         Index           =   19
         Left            =   90
         TabIndex        =   585
         Top             =   1485
         Width           =   510
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "NG"
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
         Height          =   330
         Index           =   109
         Left            =   90
         TabIndex        =   584
         Top             =   1125
         Width           =   510
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "GROSS"
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
         Height          =   285
         Index           =   110
         Left            =   90
         TabIndex        =   583
         Top             =   810
         Width           =   510
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SHIFT 3"
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
         Height          =   240
         Index           =   111
         Left            =   2070
         TabIndex        =   582
         Top             =   1485
         Width           =   735
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SHIFT 2"
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
         Height          =   240
         Index           =   112
         Left            =   1350
         TabIndex        =   581
         Top             =   1485
         Width           =   735
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SHIFT 1"
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
         Height          =   240
         Index           =   113
         Left            =   630
         TabIndex        =   580
         Top             =   1485
         Width           =   735
      End
   End
   Begin VB.Frame frameDesign 
      BackColor       =   &H00FFFFFF&
      Height          =   2175
      Index           =   19
      Left            =   8910
      TabIndex        =   543
      Top             =   9810
      Width           =   2850
      Begin VB.TextBox txtng 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   76
         Left            =   2250
         TabIndex        =   555
         Top             =   1080
         Width           =   555
      End
      Begin VB.TextBox txtng 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   77
         Left            =   1710
         TabIndex        =   554
         Top             =   1080
         Width           =   555
      End
      Begin VB.TextBox txtng 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   78
         Left            =   1170
         TabIndex        =   553
         Top             =   1080
         Width           =   555
      End
      Begin VB.TextBox txtng 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   79
         Left            =   630
         TabIndex        =   552
         Top             =   1080
         Width           =   555
      End
      Begin VB.TextBox txtidle 
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
         Index           =   57
         Left            =   2070
         TabIndex        =   551
         Top             =   1710
         Width           =   735
      End
      Begin VB.TextBox txtidle 
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
         Index           =   58
         Left            =   1350
         TabIndex        =   550
         Top             =   1710
         Width           =   735
      End
      Begin VB.TextBox txtShot 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   76
         Left            =   2250
         TabIndex        =   549
         Top             =   810
         Width           =   555
      End
      Begin VB.TextBox txtShot 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   77
         Left            =   1710
         TabIndex        =   548
         Top             =   810
         Width           =   555
      End
      Begin VB.TextBox txtShot 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   78
         Left            =   1170
         TabIndex        =   547
         Top             =   810
         Width           =   555
      End
      Begin VB.TextBox txtShot 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   79
         Left            =   630
         TabIndex        =   546
         Top             =   810
         Width           =   555
      End
      Begin VB.TextBox txtidle 
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
         Index           =   59
         Left            =   630
         TabIndex        =   545
         Top             =   1710
         Width           =   735
      End
      Begin VB.TextBox txtInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   19
         Left            =   630
         MultiLine       =   -1  'True
         TabIndex        =   544
         Top             =   135
         Width           =   2175
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SHIFT 1"
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
         Height          =   240
         Index           =   114
         Left            =   630
         TabIndex        =   566
         Top             =   1485
         Width           =   735
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SHIFT 2"
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
         Height          =   240
         Index           =   115
         Left            =   1350
         TabIndex        =   565
         Top             =   1485
         Width           =   735
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SHIFT 3"
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
         Height          =   240
         Index           =   116
         Left            =   2070
         TabIndex        =   564
         Top             =   1485
         Width           =   735
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "GROSS"
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
         Height          =   285
         Index           =   117
         Left            =   90
         TabIndex        =   563
         Top             =   810
         Width           =   510
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "NG"
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
         Height          =   330
         Index           =   118
         Left            =   90
         TabIndex        =   562
         Top             =   1125
         Width           =   510
      End
      Begin VB.Label lblIdle 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "           IDLE"
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
         Height          =   600
         Index           =   20
         Left            =   90
         TabIndex        =   561
         Top             =   1485
         Width           =   510
      End
      Begin VB.Label lbljam1L 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "9 : L"
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
         Height          =   240
         Index           =   19
         Left            =   630
         TabIndex        =   560
         Top             =   585
         Width           =   555
      End
      Begin VB.Label lbljam1R 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "9 : R"
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
         Height          =   240
         Index           =   19
         Left            =   1170
         TabIndex        =   559
         Top             =   585
         Width           =   555
      End
      Begin VB.Label lbljam2L 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "10 : L"
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
         Height          =   240
         Index           =   19
         Left            =   1710
         TabIndex        =   558
         Top             =   585
         Width           =   555
      End
      Begin VB.Label lbljam2R 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "10 : R"
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
         Height          =   240
         Index           =   19
         Left            =   2250
         TabIndex        =   557
         Top             =   585
         Width           =   555
      End
      Begin VB.Label lblMesin 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "20"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   465
         Index           =   20
         Left            =   90
         TabIndex        =   556
         Top             =   135
         Width           =   510
      End
   End
   Begin VB.Frame frameDesign 
      BackColor       =   &H00FFFFFF&
      Height          =   2175
      Index           =   20
      Left            =   11880
      TabIndex        =   519
      Top             =   630
      Width           =   2850
      Begin VB.TextBox txtInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   20
         Left            =   630
         MultiLine       =   -1  'True
         TabIndex        =   531
         Top             =   135
         Width           =   2175
      End
      Begin VB.TextBox txtidle 
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
         Index           =   60
         Left            =   630
         TabIndex        =   530
         Top             =   1710
         Width           =   735
      End
      Begin VB.TextBox txtShot 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   80
         Left            =   630
         TabIndex        =   529
         Top             =   810
         Width           =   555
      End
      Begin VB.TextBox txtShot 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   81
         Left            =   1170
         TabIndex        =   528
         Top             =   810
         Width           =   555
      End
      Begin VB.TextBox txtShot 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   82
         Left            =   1710
         TabIndex        =   527
         Top             =   810
         Width           =   555
      End
      Begin VB.TextBox txtShot 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   83
         Left            =   2250
         TabIndex        =   526
         Top             =   810
         Width           =   555
      End
      Begin VB.TextBox txtidle 
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
         Index           =   61
         Left            =   1350
         TabIndex        =   525
         Top             =   1710
         Width           =   735
      End
      Begin VB.TextBox txtidle 
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
         Index           =   62
         Left            =   2070
         TabIndex        =   524
         Top             =   1710
         Width           =   735
      End
      Begin VB.TextBox txtng 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   80
         Left            =   630
         TabIndex        =   523
         Top             =   1080
         Width           =   555
      End
      Begin VB.TextBox txtng 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   81
         Left            =   1170
         TabIndex        =   522
         Top             =   1080
         Width           =   555
      End
      Begin VB.TextBox txtng 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   82
         Left            =   1710
         TabIndex        =   521
         Top             =   1080
         Width           =   555
      End
      Begin VB.TextBox txtng 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   83
         Left            =   2250
         TabIndex        =   520
         Top             =   1080
         Width           =   555
      End
      Begin VB.Label lblMesin 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "21"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   465
         Index           =   21
         Left            =   90
         TabIndex        =   542
         Top             =   135
         Width           =   510
      End
      Begin VB.Label lbljam2R 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "10 : R"
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
         Height          =   240
         Index           =   20
         Left            =   2250
         TabIndex        =   541
         Top             =   585
         Width           =   555
      End
      Begin VB.Label lbljam2L 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "10 : L"
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
         Height          =   240
         Index           =   20
         Left            =   1710
         TabIndex        =   540
         Top             =   585
         Width           =   555
      End
      Begin VB.Label lbljam1R 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "9 : R"
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
         Height          =   240
         Index           =   20
         Left            =   1170
         TabIndex        =   539
         Top             =   585
         Width           =   555
      End
      Begin VB.Label lbljam1L 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "9 : L"
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
         Height          =   240
         Index           =   20
         Left            =   630
         TabIndex        =   538
         Top             =   585
         Width           =   555
      End
      Begin VB.Label lblIdle 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "           IDLE"
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
         Height          =   600
         Index           =   21
         Left            =   90
         TabIndex        =   537
         Top             =   1485
         Width           =   510
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "NG"
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
         Height          =   330
         Index           =   121
         Left            =   90
         TabIndex        =   536
         Top             =   1125
         Width           =   510
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "GROSS"
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
         Height          =   285
         Index           =   122
         Left            =   90
         TabIndex        =   535
         Top             =   810
         Width           =   510
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SHIFT 3"
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
         Height          =   240
         Index           =   123
         Left            =   2070
         TabIndex        =   534
         Top             =   1485
         Width           =   735
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SHIFT 2"
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
         Height          =   240
         Index           =   124
         Left            =   1350
         TabIndex        =   533
         Top             =   1485
         Width           =   735
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SHIFT 1"
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
         Height          =   240
         Index           =   125
         Left            =   630
         TabIndex        =   532
         Top             =   1485
         Width           =   735
      End
   End
   Begin VB.Frame frameDesign 
      BackColor       =   &H00FFFFFF&
      Height          =   2175
      Index           =   21
      Left            =   11880
      TabIndex        =   495
      Top             =   2925
      Width           =   2850
      Begin VB.TextBox txtng 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   84
         Left            =   2250
         TabIndex        =   507
         Top             =   1080
         Width           =   555
      End
      Begin VB.TextBox txtng 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   85
         Left            =   1710
         TabIndex        =   506
         Top             =   1080
         Width           =   555
      End
      Begin VB.TextBox txtng 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   86
         Left            =   1170
         TabIndex        =   505
         Top             =   1080
         Width           =   555
      End
      Begin VB.TextBox txtng 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   87
         Left            =   630
         TabIndex        =   504
         Top             =   1080
         Width           =   555
      End
      Begin VB.TextBox txtidle 
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
         Index           =   63
         Left            =   2070
         TabIndex        =   503
         Top             =   1710
         Width           =   735
      End
      Begin VB.TextBox txtidle 
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
         Index           =   64
         Left            =   1350
         TabIndex        =   502
         Top             =   1710
         Width           =   735
      End
      Begin VB.TextBox txtShot 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   84
         Left            =   2250
         TabIndex        =   501
         Top             =   810
         Width           =   555
      End
      Begin VB.TextBox txtShot 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   85
         Left            =   1710
         TabIndex        =   500
         Top             =   810
         Width           =   555
      End
      Begin VB.TextBox txtShot 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   86
         Left            =   1170
         TabIndex        =   499
         Top             =   810
         Width           =   555
      End
      Begin VB.TextBox txtShot 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   87
         Left            =   630
         TabIndex        =   498
         Top             =   810
         Width           =   555
      End
      Begin VB.TextBox txtidle 
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
         Index           =   65
         Left            =   630
         TabIndex        =   497
         Top             =   1710
         Width           =   735
      End
      Begin VB.TextBox txtInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   21
         Left            =   630
         MultiLine       =   -1  'True
         TabIndex        =   496
         Top             =   135
         Width           =   2175
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SHIFT 1"
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
         Height          =   240
         Index           =   126
         Left            =   630
         TabIndex        =   518
         Top             =   1485
         Width           =   735
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SHIFT 2"
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
         Height          =   240
         Index           =   127
         Left            =   1350
         TabIndex        =   517
         Top             =   1485
         Width           =   735
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SHIFT 3"
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
         Height          =   240
         Index           =   128
         Left            =   2070
         TabIndex        =   516
         Top             =   1485
         Width           =   735
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "GROSS"
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
         Height          =   285
         Index           =   129
         Left            =   90
         TabIndex        =   515
         Top             =   810
         Width           =   510
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "NG"
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
         Height          =   330
         Index           =   130
         Left            =   90
         TabIndex        =   514
         Top             =   1125
         Width           =   510
      End
      Begin VB.Label lblIdle 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "           IDLE"
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
         Height          =   600
         Index           =   22
         Left            =   90
         TabIndex        =   513
         Top             =   1485
         Width           =   510
      End
      Begin VB.Label lbljam1L 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "9 : L"
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
         Height          =   240
         Index           =   21
         Left            =   630
         TabIndex        =   512
         Top             =   585
         Width           =   555
      End
      Begin VB.Label lbljam1R 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "9 : R"
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
         Height          =   240
         Index           =   21
         Left            =   1170
         TabIndex        =   511
         Top             =   585
         Width           =   555
      End
      Begin VB.Label lbljam2L 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "10 : L"
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
         Height          =   240
         Index           =   21
         Left            =   1710
         TabIndex        =   510
         Top             =   585
         Width           =   555
      End
      Begin VB.Label lbljam2R 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "10 : R"
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
         Height          =   240
         Index           =   21
         Left            =   2250
         TabIndex        =   509
         Top             =   585
         Width           =   555
      End
      Begin VB.Label lblMesin 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "22"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   465
         Index           =   22
         Left            =   90
         TabIndex        =   508
         Top             =   135
         Width           =   510
      End
   End
   Begin VB.Frame frameDesign 
      BackColor       =   &H00FFFFFF&
      Height          =   2175
      Index           =   22
      Left            =   11880
      TabIndex        =   471
      Top             =   5220
      Width           =   2850
      Begin VB.TextBox txtInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   22
         Left            =   630
         MultiLine       =   -1  'True
         TabIndex        =   483
         Top             =   135
         Width           =   2175
      End
      Begin VB.TextBox txtidle 
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
         Index           =   66
         Left            =   630
         TabIndex        =   482
         Top             =   1710
         Width           =   735
      End
      Begin VB.TextBox txtShot 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   88
         Left            =   630
         TabIndex        =   481
         Top             =   810
         Width           =   555
      End
      Begin VB.TextBox txtShot 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   89
         Left            =   1170
         TabIndex        =   480
         Top             =   810
         Width           =   555
      End
      Begin VB.TextBox txtShot 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   90
         Left            =   1710
         TabIndex        =   479
         Top             =   810
         Width           =   555
      End
      Begin VB.TextBox txtShot 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   91
         Left            =   2250
         TabIndex        =   478
         Top             =   810
         Width           =   555
      End
      Begin VB.TextBox txtidle 
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
         Index           =   67
         Left            =   1350
         TabIndex        =   477
         Top             =   1710
         Width           =   735
      End
      Begin VB.TextBox txtidle 
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
         Index           =   68
         Left            =   2070
         TabIndex        =   476
         Top             =   1710
         Width           =   735
      End
      Begin VB.TextBox txtng 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   88
         Left            =   630
         TabIndex        =   475
         Top             =   1080
         Width           =   555
      End
      Begin VB.TextBox txtng 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   89
         Left            =   1170
         TabIndex        =   474
         Top             =   1080
         Width           =   555
      End
      Begin VB.TextBox txtng 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   90
         Left            =   1710
         TabIndex        =   473
         Top             =   1080
         Width           =   555
      End
      Begin VB.TextBox txtng 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   91
         Left            =   2250
         TabIndex        =   472
         Top             =   1080
         Width           =   555
      End
      Begin VB.Label lblMesin 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "23"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   465
         Index           =   23
         Left            =   90
         TabIndex        =   494
         Top             =   135
         Width           =   510
      End
      Begin VB.Label lbljam2R 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "10 : R"
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
         Height          =   240
         Index           =   22
         Left            =   2250
         TabIndex        =   493
         Top             =   585
         Width           =   555
      End
      Begin VB.Label lbljam2L 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "10 : L"
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
         Height          =   240
         Index           =   22
         Left            =   1710
         TabIndex        =   492
         Top             =   585
         Width           =   555
      End
      Begin VB.Label lbljam1R 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "9 : R"
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
         Height          =   240
         Index           =   22
         Left            =   1170
         TabIndex        =   491
         Top             =   585
         Width           =   555
      End
      Begin VB.Label lbljam1L 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "9 : L"
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
         Height          =   240
         Index           =   22
         Left            =   630
         TabIndex        =   490
         Top             =   585
         Width           =   555
      End
      Begin VB.Label lblIdle 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "           IDLE"
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
         Height          =   600
         Index           =   23
         Left            =   90
         TabIndex        =   489
         Top             =   1485
         Width           =   510
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "NG"
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
         Height          =   330
         Index           =   133
         Left            =   90
         TabIndex        =   488
         Top             =   1125
         Width           =   510
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "GROSS"
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
         Height          =   285
         Index           =   134
         Left            =   90
         TabIndex        =   487
         Top             =   810
         Width           =   510
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SHIFT 3"
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
         Height          =   240
         Index           =   135
         Left            =   2070
         TabIndex        =   486
         Top             =   1485
         Width           =   735
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SHIFT 2"
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
         Height          =   240
         Index           =   136
         Left            =   1350
         TabIndex        =   485
         Top             =   1485
         Width           =   735
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SHIFT 1"
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
         Height          =   240
         Index           =   137
         Left            =   630
         TabIndex        =   484
         Top             =   1485
         Width           =   735
      End
   End
   Begin VB.Frame frameDesign 
      BackColor       =   &H00FFFFFF&
      Height          =   2175
      Index           =   23
      Left            =   11880
      TabIndex        =   447
      Top             =   7515
      Width           =   2850
      Begin VB.TextBox txtInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   23
         Left            =   630
         MultiLine       =   -1  'True
         TabIndex        =   459
         Top             =   135
         Width           =   2175
      End
      Begin VB.TextBox txtidle 
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
         Index           =   69
         Left            =   630
         TabIndex        =   458
         Top             =   1710
         Width           =   735
      End
      Begin VB.TextBox txtShot 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   92
         Left            =   630
         TabIndex        =   457
         Top             =   810
         Width           =   555
      End
      Begin VB.TextBox txtShot 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   93
         Left            =   1170
         TabIndex        =   456
         Top             =   810
         Width           =   555
      End
      Begin VB.TextBox txtShot 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   94
         Left            =   1710
         TabIndex        =   455
         Top             =   810
         Width           =   555
      End
      Begin VB.TextBox txtShot 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   95
         Left            =   2250
         TabIndex        =   454
         Top             =   810
         Width           =   555
      End
      Begin VB.TextBox txtidle 
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
         Index           =   70
         Left            =   1350
         TabIndex        =   453
         Top             =   1710
         Width           =   735
      End
      Begin VB.TextBox txtidle 
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
         Index           =   71
         Left            =   2070
         TabIndex        =   452
         Top             =   1710
         Width           =   735
      End
      Begin VB.TextBox txtng 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   92
         Left            =   630
         TabIndex        =   451
         Top             =   1080
         Width           =   555
      End
      Begin VB.TextBox txtng 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   93
         Left            =   1170
         TabIndex        =   450
         Top             =   1080
         Width           =   555
      End
      Begin VB.TextBox txtng 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   94
         Left            =   1710
         TabIndex        =   449
         Top             =   1080
         Width           =   555
      End
      Begin VB.TextBox txtng 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   95
         Left            =   2250
         TabIndex        =   448
         Top             =   1080
         Width           =   555
      End
      Begin VB.Label lblMesin 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "24"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   465
         Index           =   24
         Left            =   90
         TabIndex        =   470
         Top             =   135
         Width           =   510
      End
      Begin VB.Label lbljam2R 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "10 : R"
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
         Height          =   240
         Index           =   23
         Left            =   2250
         TabIndex        =   469
         Top             =   585
         Width           =   555
      End
      Begin VB.Label lbljam2L 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "10 : L"
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
         Height          =   240
         Index           =   23
         Left            =   1710
         TabIndex        =   468
         Top             =   585
         Width           =   555
      End
      Begin VB.Label lbljam1R 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "9 : R"
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
         Height          =   240
         Index           =   23
         Left            =   1170
         TabIndex        =   467
         Top             =   585
         Width           =   555
      End
      Begin VB.Label lbljam1L 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "9 : L"
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
         Height          =   240
         Index           =   23
         Left            =   630
         TabIndex        =   466
         Top             =   585
         Width           =   555
      End
      Begin VB.Label lblIdle 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "           IDLE"
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
         Height          =   600
         Index           =   24
         Left            =   90
         TabIndex        =   465
         Top             =   1485
         Width           =   510
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "NG"
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
         Height          =   330
         Index           =   139
         Left            =   90
         TabIndex        =   464
         Top             =   1125
         Width           =   510
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "GROSS"
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
         Height          =   285
         Index           =   140
         Left            =   90
         TabIndex        =   463
         Top             =   810
         Width           =   510
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SHIFT 3"
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
         Height          =   240
         Index           =   141
         Left            =   2070
         TabIndex        =   462
         Top             =   1485
         Width           =   735
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SHIFT 2"
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
         Height          =   240
         Index           =   142
         Left            =   1350
         TabIndex        =   461
         Top             =   1485
         Width           =   735
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SHIFT 1"
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
         Height          =   240
         Index           =   143
         Left            =   630
         TabIndex        =   460
         Top             =   1485
         Width           =   735
      End
   End
   Begin VB.Frame frameDesign 
      BackColor       =   &H00FFFFFF&
      Height          =   2175
      Index           =   24
      Left            =   11880
      TabIndex        =   423
      Top             =   9810
      Width           =   2850
      Begin VB.TextBox txtng 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   96
         Left            =   2250
         TabIndex        =   435
         Top             =   1080
         Width           =   555
      End
      Begin VB.TextBox txtng 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   97
         Left            =   1710
         TabIndex        =   434
         Top             =   1080
         Width           =   555
      End
      Begin VB.TextBox txtng 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   98
         Left            =   1170
         TabIndex        =   433
         Top             =   1080
         Width           =   555
      End
      Begin VB.TextBox txtng 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   99
         Left            =   630
         TabIndex        =   432
         Top             =   1080
         Width           =   555
      End
      Begin VB.TextBox txtidle 
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
         Index           =   72
         Left            =   2070
         TabIndex        =   431
         Top             =   1710
         Width           =   735
      End
      Begin VB.TextBox txtidle 
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
         Index           =   73
         Left            =   1350
         TabIndex        =   430
         Top             =   1710
         Width           =   735
      End
      Begin VB.TextBox txtShot 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   96
         Left            =   2250
         TabIndex        =   429
         Top             =   810
         Width           =   555
      End
      Begin VB.TextBox txtShot 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   97
         Left            =   1710
         TabIndex        =   428
         Top             =   810
         Width           =   555
      End
      Begin VB.TextBox txtShot 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   98
         Left            =   1170
         TabIndex        =   427
         Top             =   810
         Width           =   555
      End
      Begin VB.TextBox txtShot 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   99
         Left            =   630
         TabIndex        =   426
         Top             =   810
         Width           =   555
      End
      Begin VB.TextBox txtidle 
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
         Index           =   74
         Left            =   630
         TabIndex        =   425
         Top             =   1710
         Width           =   735
      End
      Begin VB.TextBox txtInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   24
         Left            =   630
         MultiLine       =   -1  'True
         TabIndex        =   424
         Top             =   135
         Width           =   2175
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SHIFT 1"
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
         Height          =   240
         Index           =   144
         Left            =   630
         TabIndex        =   446
         Top             =   1485
         Width           =   735
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SHIFT 2"
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
         Height          =   240
         Index           =   145
         Left            =   1350
         TabIndex        =   445
         Top             =   1485
         Width           =   735
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SHIFT 3"
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
         Height          =   240
         Index           =   146
         Left            =   2070
         TabIndex        =   444
         Top             =   1485
         Width           =   735
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "GROSS"
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
         Height          =   285
         Index           =   147
         Left            =   90
         TabIndex        =   443
         Top             =   810
         Width           =   510
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "NG"
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
         Height          =   330
         Index           =   148
         Left            =   90
         TabIndex        =   442
         Top             =   1125
         Width           =   510
      End
      Begin VB.Label lblIdle 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "           IDLE"
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
         Height          =   600
         Index           =   25
         Left            =   90
         TabIndex        =   441
         Top             =   1485
         Width           =   510
      End
      Begin VB.Label lbljam1L 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "9 : L"
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
         Height          =   240
         Index           =   24
         Left            =   630
         TabIndex        =   440
         Top             =   585
         Width           =   555
      End
      Begin VB.Label lbljam1R 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "9 : R"
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
         Height          =   240
         Index           =   24
         Left            =   1170
         TabIndex        =   439
         Top             =   585
         Width           =   555
      End
      Begin VB.Label lbljam2L 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "10 : L"
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
         Height          =   240
         Index           =   24
         Left            =   1710
         TabIndex        =   438
         Top             =   585
         Width           =   555
      End
      Begin VB.Label lbljam2R 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "10 : R"
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
         Height          =   240
         Index           =   24
         Left            =   2250
         TabIndex        =   437
         Top             =   585
         Width           =   555
      End
      Begin VB.Label lblMesin 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "25"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   465
         Index           =   25
         Left            =   90
         TabIndex        =   436
         Top             =   135
         Width           =   510
      End
   End
   Begin VB.Frame frameDesign 
      BackColor       =   &H00FFFFFF&
      Height          =   2175
      Index           =   25
      Left            =   14850
      TabIndex        =   399
      Top             =   630
      Width           =   2850
      Begin VB.TextBox txtInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   25
         Left            =   630
         MultiLine       =   -1  'True
         TabIndex        =   411
         Top             =   135
         Width           =   2175
      End
      Begin VB.TextBox txtidle 
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
         Index           =   75
         Left            =   630
         TabIndex        =   410
         Top             =   1710
         Width           =   735
      End
      Begin VB.TextBox txtShot 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   100
         Left            =   630
         TabIndex        =   409
         Top             =   810
         Width           =   555
      End
      Begin VB.TextBox txtShot 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   101
         Left            =   1170
         TabIndex        =   408
         Top             =   810
         Width           =   555
      End
      Begin VB.TextBox txtShot 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   102
         Left            =   1710
         TabIndex        =   407
         Top             =   810
         Width           =   555
      End
      Begin VB.TextBox txtShot 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   103
         Left            =   2250
         TabIndex        =   406
         Top             =   810
         Width           =   555
      End
      Begin VB.TextBox txtidle 
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
         Index           =   76
         Left            =   1350
         TabIndex        =   405
         Top             =   1710
         Width           =   735
      End
      Begin VB.TextBox txtidle 
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
         Index           =   77
         Left            =   2070
         TabIndex        =   404
         Top             =   1710
         Width           =   735
      End
      Begin VB.TextBox txtng 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   100
         Left            =   630
         TabIndex        =   403
         Top             =   1080
         Width           =   555
      End
      Begin VB.TextBox txtng 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   101
         Left            =   1170
         TabIndex        =   402
         Top             =   1080
         Width           =   555
      End
      Begin VB.TextBox txtng 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   102
         Left            =   1710
         TabIndex        =   401
         Top             =   1080
         Width           =   555
      End
      Begin VB.TextBox txtng 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   103
         Left            =   2250
         TabIndex        =   400
         Top             =   1080
         Width           =   555
      End
      Begin VB.Label lblMesin 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "26"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   465
         Index           =   26
         Left            =   90
         TabIndex        =   422
         Top             =   135
         Width           =   510
      End
      Begin VB.Label lbljam2R 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "10 : R"
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
         Height          =   240
         Index           =   25
         Left            =   2250
         TabIndex        =   421
         Top             =   585
         Width           =   555
      End
      Begin VB.Label lbljam2L 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "10 : L"
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
         Height          =   240
         Index           =   25
         Left            =   1710
         TabIndex        =   420
         Top             =   585
         Width           =   555
      End
      Begin VB.Label lbljam1R 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "9 : R"
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
         Height          =   240
         Index           =   25
         Left            =   1170
         TabIndex        =   419
         Top             =   585
         Width           =   555
      End
      Begin VB.Label lbljam1L 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "9 : L"
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
         Height          =   240
         Index           =   25
         Left            =   630
         TabIndex        =   418
         Top             =   585
         Width           =   555
      End
      Begin VB.Label lblIdle 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "           IDLE"
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
         Height          =   600
         Index           =   26
         Left            =   90
         TabIndex        =   417
         Top             =   1485
         Width           =   510
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "NG"
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
         Height          =   330
         Index           =   151
         Left            =   90
         TabIndex        =   416
         Top             =   1125
         Width           =   510
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "GROSS"
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
         Height          =   285
         Index           =   152
         Left            =   90
         TabIndex        =   415
         Top             =   810
         Width           =   510
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SHIFT 3"
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
         Height          =   240
         Index           =   153
         Left            =   2070
         TabIndex        =   414
         Top             =   1485
         Width           =   735
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SHIFT 2"
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
         Height          =   240
         Index           =   154
         Left            =   1350
         TabIndex        =   413
         Top             =   1485
         Width           =   735
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SHIFT 1"
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
         Height          =   240
         Index           =   155
         Left            =   630
         TabIndex        =   412
         Top             =   1485
         Width           =   735
      End
   End
   Begin VB.Frame frameDesign 
      BackColor       =   &H00FFFFFF&
      Height          =   2175
      Index           =   26
      Left            =   14850
      TabIndex        =   375
      Top             =   2925
      Width           =   2850
      Begin VB.TextBox txtng 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   104
         Left            =   2250
         TabIndex        =   387
         Top             =   1080
         Width           =   555
      End
      Begin VB.TextBox txtng 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   105
         Left            =   1710
         TabIndex        =   386
         Top             =   1080
         Width           =   555
      End
      Begin VB.TextBox txtng 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   106
         Left            =   1170
         TabIndex        =   385
         Top             =   1080
         Width           =   555
      End
      Begin VB.TextBox txtng 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   107
         Left            =   630
         TabIndex        =   384
         Top             =   1080
         Width           =   555
      End
      Begin VB.TextBox txtidle 
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
         Index           =   78
         Left            =   2070
         TabIndex        =   383
         Top             =   1710
         Width           =   735
      End
      Begin VB.TextBox txtidle 
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
         Index           =   79
         Left            =   1350
         TabIndex        =   382
         Top             =   1710
         Width           =   735
      End
      Begin VB.TextBox txtShot 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   104
         Left            =   2250
         TabIndex        =   381
         Top             =   810
         Width           =   555
      End
      Begin VB.TextBox txtShot 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   105
         Left            =   1710
         TabIndex        =   380
         Top             =   810
         Width           =   555
      End
      Begin VB.TextBox txtShot 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   106
         Left            =   1170
         TabIndex        =   379
         Top             =   810
         Width           =   555
      End
      Begin VB.TextBox txtShot 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   107
         Left            =   630
         TabIndex        =   378
         Top             =   810
         Width           =   555
      End
      Begin VB.TextBox txtidle 
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
         Index           =   80
         Left            =   630
         TabIndex        =   377
         Top             =   1710
         Width           =   735
      End
      Begin VB.TextBox txtInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   26
         Left            =   630
         MultiLine       =   -1  'True
         TabIndex        =   376
         Top             =   135
         Width           =   2175
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SHIFT 1"
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
         Height          =   240
         Index           =   156
         Left            =   630
         TabIndex        =   398
         Top             =   1485
         Width           =   735
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SHIFT 2"
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
         Height          =   240
         Index           =   157
         Left            =   1350
         TabIndex        =   397
         Top             =   1485
         Width           =   735
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SHIFT 3"
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
         Height          =   240
         Index           =   158
         Left            =   2070
         TabIndex        =   396
         Top             =   1485
         Width           =   735
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "GROSS"
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
         Height          =   285
         Index           =   159
         Left            =   90
         TabIndex        =   395
         Top             =   810
         Width           =   510
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "NG"
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
         Height          =   330
         Index           =   160
         Left            =   90
         TabIndex        =   394
         Top             =   1125
         Width           =   510
      End
      Begin VB.Label lblIdle 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "           IDLE"
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
         Height          =   600
         Index           =   27
         Left            =   90
         TabIndex        =   393
         Top             =   1485
         Width           =   510
      End
      Begin VB.Label lbljam1L 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "9 : L"
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
         Height          =   240
         Index           =   26
         Left            =   630
         TabIndex        =   392
         Top             =   585
         Width           =   555
      End
      Begin VB.Label lbljam1R 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "9 : R"
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
         Height          =   240
         Index           =   26
         Left            =   1170
         TabIndex        =   391
         Top             =   585
         Width           =   555
      End
      Begin VB.Label lbljam2L 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "10 : L"
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
         Height          =   240
         Index           =   26
         Left            =   1710
         TabIndex        =   390
         Top             =   585
         Width           =   555
      End
      Begin VB.Label lbljam2R 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "10 : R"
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
         Height          =   240
         Index           =   26
         Left            =   2250
         TabIndex        =   389
         Top             =   585
         Width           =   555
      End
      Begin VB.Label lblMesin 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "27"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   465
         Index           =   27
         Left            =   90
         TabIndex        =   388
         Top             =   135
         Width           =   510
      End
   End
   Begin VB.Frame frameDesign 
      BackColor       =   &H00FFFFFF&
      Height          =   2175
      Index           =   27
      Left            =   14850
      TabIndex        =   351
      Top             =   5220
      Width           =   2850
      Begin VB.TextBox txtInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   27
         Left            =   630
         MultiLine       =   -1  'True
         TabIndex        =   363
         Top             =   135
         Width           =   2175
      End
      Begin VB.TextBox txtidle 
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
         Index           =   81
         Left            =   630
         TabIndex        =   362
         Top             =   1710
         Width           =   735
      End
      Begin VB.TextBox txtShot 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   108
         Left            =   630
         TabIndex        =   361
         Top             =   810
         Width           =   555
      End
      Begin VB.TextBox txtShot 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   109
         Left            =   1170
         TabIndex        =   360
         Top             =   810
         Width           =   555
      End
      Begin VB.TextBox txtShot 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   110
         Left            =   1710
         TabIndex        =   359
         Top             =   810
         Width           =   555
      End
      Begin VB.TextBox txtShot 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   111
         Left            =   2250
         TabIndex        =   358
         Top             =   810
         Width           =   555
      End
      Begin VB.TextBox txtidle 
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
         Index           =   82
         Left            =   1350
         TabIndex        =   357
         Top             =   1710
         Width           =   735
      End
      Begin VB.TextBox txtidle 
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
         Index           =   83
         Left            =   2070
         TabIndex        =   356
         Top             =   1710
         Width           =   735
      End
      Begin VB.TextBox txtng 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   108
         Left            =   630
         TabIndex        =   355
         Top             =   1080
         Width           =   555
      End
      Begin VB.TextBox txtng 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   109
         Left            =   1170
         TabIndex        =   354
         Top             =   1080
         Width           =   555
      End
      Begin VB.TextBox txtng 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   110
         Left            =   1710
         TabIndex        =   353
         Top             =   1080
         Width           =   555
      End
      Begin VB.TextBox txtng 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   111
         Left            =   2250
         TabIndex        =   352
         Top             =   1080
         Width           =   555
      End
      Begin VB.Label lblMesin 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "28"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   465
         Index           =   28
         Left            =   90
         TabIndex        =   374
         Top             =   135
         Width           =   510
      End
      Begin VB.Label lbljam2R 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "10 : R"
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
         Height          =   240
         Index           =   27
         Left            =   2250
         TabIndex        =   373
         Top             =   585
         Width           =   555
      End
      Begin VB.Label lbljam2L 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "10 : L"
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
         Height          =   240
         Index           =   27
         Left            =   1710
         TabIndex        =   372
         Top             =   585
         Width           =   555
      End
      Begin VB.Label lbljam1R 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "9 : R"
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
         Height          =   240
         Index           =   27
         Left            =   1170
         TabIndex        =   371
         Top             =   585
         Width           =   555
      End
      Begin VB.Label lbljam1L 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "9 : L"
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
         Height          =   240
         Index           =   27
         Left            =   630
         TabIndex        =   370
         Top             =   585
         Width           =   555
      End
      Begin VB.Label lblIdle 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "           IDLE"
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
         Height          =   600
         Index           =   28
         Left            =   90
         TabIndex        =   369
         Top             =   1485
         Width           =   510
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "NG"
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
         Height          =   330
         Index           =   163
         Left            =   90
         TabIndex        =   368
         Top             =   1125
         Width           =   510
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "GROSS"
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
         Height          =   285
         Index           =   164
         Left            =   90
         TabIndex        =   367
         Top             =   810
         Width           =   510
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SHIFT 3"
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
         Height          =   240
         Index           =   165
         Left            =   2070
         TabIndex        =   366
         Top             =   1485
         Width           =   735
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SHIFT 2"
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
         Height          =   240
         Index           =   166
         Left            =   1350
         TabIndex        =   365
         Top             =   1485
         Width           =   735
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SHIFT 1"
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
         Height          =   240
         Index           =   167
         Left            =   630
         TabIndex        =   364
         Top             =   1485
         Width           =   735
      End
   End
   Begin VB.Frame frameDesign 
      BackColor       =   &H00FFFFFF&
      Height          =   2175
      Index           =   28
      Left            =   14850
      TabIndex        =   327
      Top             =   7515
      Width           =   2850
      Begin VB.TextBox txtInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   28
         Left            =   630
         MultiLine       =   -1  'True
         TabIndex        =   339
         Top             =   135
         Width           =   2175
      End
      Begin VB.TextBox txtidle 
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
         Index           =   84
         Left            =   630
         TabIndex        =   338
         Top             =   1710
         Width           =   735
      End
      Begin VB.TextBox txtShot 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   112
         Left            =   630
         TabIndex        =   337
         Top             =   810
         Width           =   555
      End
      Begin VB.TextBox txtShot 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   113
         Left            =   1170
         TabIndex        =   336
         Top             =   810
         Width           =   555
      End
      Begin VB.TextBox txtShot 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   114
         Left            =   1710
         TabIndex        =   335
         Top             =   810
         Width           =   555
      End
      Begin VB.TextBox txtShot 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   115
         Left            =   2250
         TabIndex        =   334
         Top             =   810
         Width           =   555
      End
      Begin VB.TextBox txtidle 
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
         Index           =   85
         Left            =   1350
         TabIndex        =   333
         Top             =   1710
         Width           =   735
      End
      Begin VB.TextBox txtidle 
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
         Index           =   86
         Left            =   2070
         TabIndex        =   332
         Top             =   1710
         Width           =   735
      End
      Begin VB.TextBox txtng 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   112
         Left            =   630
         TabIndex        =   331
         Top             =   1080
         Width           =   555
      End
      Begin VB.TextBox txtng 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   113
         Left            =   1170
         TabIndex        =   330
         Top             =   1080
         Width           =   555
      End
      Begin VB.TextBox txtng 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   114
         Left            =   1710
         TabIndex        =   329
         Top             =   1080
         Width           =   555
      End
      Begin VB.TextBox txtng 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   115
         Left            =   2250
         TabIndex        =   328
         Top             =   1080
         Width           =   555
      End
      Begin VB.Label lblMesin 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "29"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   465
         Index           =   29
         Left            =   90
         TabIndex        =   350
         Top             =   135
         Width           =   510
      End
      Begin VB.Label lbljam2R 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "10 : R"
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
         Height          =   240
         Index           =   28
         Left            =   2250
         TabIndex        =   349
         Top             =   585
         Width           =   555
      End
      Begin VB.Label lbljam2L 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "10 : L"
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
         Height          =   240
         Index           =   28
         Left            =   1710
         TabIndex        =   348
         Top             =   585
         Width           =   555
      End
      Begin VB.Label lbljam1R 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "9 : R"
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
         Height          =   240
         Index           =   28
         Left            =   1170
         TabIndex        =   347
         Top             =   585
         Width           =   555
      End
      Begin VB.Label lbljam1L 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "9 : L"
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
         Height          =   240
         Index           =   28
         Left            =   630
         TabIndex        =   346
         Top             =   585
         Width           =   555
      End
      Begin VB.Label lblIdle 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "           IDLE"
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
         Height          =   600
         Index           =   29
         Left            =   90
         TabIndex        =   345
         Top             =   1485
         Width           =   510
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "NG"
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
         Height          =   330
         Index           =   169
         Left            =   90
         TabIndex        =   344
         Top             =   1125
         Width           =   510
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "GROSS"
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
         Height          =   285
         Index           =   170
         Left            =   90
         TabIndex        =   343
         Top             =   810
         Width           =   510
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SHIFT 3"
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
         Height          =   240
         Index           =   171
         Left            =   2070
         TabIndex        =   342
         Top             =   1485
         Width           =   735
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SHIFT 2"
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
         Height          =   240
         Index           =   172
         Left            =   1350
         TabIndex        =   341
         Top             =   1485
         Width           =   735
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SHIFT 1"
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
         Height          =   240
         Index           =   173
         Left            =   630
         TabIndex        =   340
         Top             =   1485
         Width           =   735
      End
   End
   Begin VB.Frame frameDesign 
      BackColor       =   &H00FFFFFF&
      Height          =   2175
      Index           =   29
      Left            =   14850
      TabIndex        =   303
      Top             =   9810
      Width           =   2850
      Begin VB.TextBox txtng 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   116
         Left            =   2250
         TabIndex        =   315
         Top             =   1080
         Width           =   555
      End
      Begin VB.TextBox txtng 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   117
         Left            =   1710
         TabIndex        =   314
         Top             =   1080
         Width           =   555
      End
      Begin VB.TextBox txtng 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   118
         Left            =   1170
         TabIndex        =   313
         Top             =   1080
         Width           =   555
      End
      Begin VB.TextBox txtng 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   119
         Left            =   630
         TabIndex        =   312
         Top             =   1080
         Width           =   555
      End
      Begin VB.TextBox txtidle 
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
         Index           =   87
         Left            =   2070
         TabIndex        =   311
         Top             =   1710
         Width           =   735
      End
      Begin VB.TextBox txtidle 
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
         Index           =   88
         Left            =   1350
         TabIndex        =   310
         Top             =   1710
         Width           =   735
      End
      Begin VB.TextBox txtShot 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   116
         Left            =   2250
         TabIndex        =   309
         Top             =   810
         Width           =   555
      End
      Begin VB.TextBox txtShot 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   117
         Left            =   1710
         TabIndex        =   308
         Top             =   810
         Width           =   555
      End
      Begin VB.TextBox txtShot 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   118
         Left            =   1170
         TabIndex        =   307
         Top             =   810
         Width           =   555
      End
      Begin VB.TextBox txtShot 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   119
         Left            =   630
         TabIndex        =   306
         Top             =   810
         Width           =   555
      End
      Begin VB.TextBox txtidle 
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
         Index           =   89
         Left            =   630
         TabIndex        =   305
         Top             =   1710
         Width           =   735
      End
      Begin VB.TextBox txtInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   29
         Left            =   630
         MultiLine       =   -1  'True
         TabIndex        =   304
         Top             =   135
         Width           =   2175
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SHIFT 1"
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
         Height          =   240
         Index           =   174
         Left            =   630
         TabIndex        =   326
         Top             =   1485
         Width           =   735
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SHIFT 2"
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
         Height          =   240
         Index           =   175
         Left            =   1350
         TabIndex        =   325
         Top             =   1485
         Width           =   735
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SHIFT 3"
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
         Height          =   240
         Index           =   176
         Left            =   2070
         TabIndex        =   324
         Top             =   1485
         Width           =   735
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "GROSS"
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
         Height          =   285
         Index           =   177
         Left            =   90
         TabIndex        =   323
         Top             =   810
         Width           =   510
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "NG"
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
         Height          =   330
         Index           =   178
         Left            =   90
         TabIndex        =   322
         Top             =   1125
         Width           =   510
      End
      Begin VB.Label lblIdle 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "           IDLE"
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
         Height          =   600
         Index           =   30
         Left            =   90
         TabIndex        =   321
         Top             =   1485
         Width           =   510
      End
      Begin VB.Label lbljam1L 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "9 : L"
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
         Height          =   240
         Index           =   29
         Left            =   630
         TabIndex        =   320
         Top             =   585
         Width           =   555
      End
      Begin VB.Label lbljam1R 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "9 : R"
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
         Height          =   240
         Index           =   29
         Left            =   1170
         TabIndex        =   319
         Top             =   585
         Width           =   555
      End
      Begin VB.Label lbljam2L 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "10 : L"
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
         Height          =   240
         Index           =   29
         Left            =   1710
         TabIndex        =   318
         Top             =   585
         Width           =   555
      End
      Begin VB.Label lbljam2R 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "10 : R"
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
         Height          =   240
         Index           =   29
         Left            =   2250
         TabIndex        =   317
         Top             =   585
         Width           =   555
      End
      Begin VB.Label lblMesin 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "30"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   465
         Index           =   30
         Left            =   90
         TabIndex        =   316
         Top             =   135
         Width           =   510
      End
   End
   Begin VB.Frame frameDesign 
      BackColor       =   &H00FFFFFF&
      Height          =   2175
      Index           =   30
      Left            =   17820
      TabIndex        =   279
      Top             =   630
      Width           =   2850
      Begin VB.TextBox txtInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   30
         Left            =   630
         MultiLine       =   -1  'True
         TabIndex        =   291
         Top             =   135
         Width           =   2175
      End
      Begin VB.TextBox txtidle 
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
         Index           =   90
         Left            =   630
         TabIndex        =   290
         Top             =   1710
         Width           =   735
      End
      Begin VB.TextBox txtShot 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   120
         Left            =   630
         TabIndex        =   289
         Top             =   810
         Width           =   555
      End
      Begin VB.TextBox txtShot 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   121
         Left            =   1170
         TabIndex        =   288
         Top             =   810
         Width           =   555
      End
      Begin VB.TextBox txtShot 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   122
         Left            =   1710
         TabIndex        =   287
         Top             =   810
         Width           =   555
      End
      Begin VB.TextBox txtShot 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   123
         Left            =   2250
         TabIndex        =   286
         Top             =   810
         Width           =   555
      End
      Begin VB.TextBox txtidle 
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
         Index           =   91
         Left            =   1350
         TabIndex        =   285
         Top             =   1710
         Width           =   735
      End
      Begin VB.TextBox txtidle 
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
         Index           =   92
         Left            =   2070
         TabIndex        =   284
         Top             =   1710
         Width           =   735
      End
      Begin VB.TextBox txtng 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   120
         Left            =   630
         TabIndex        =   283
         Top             =   1080
         Width           =   555
      End
      Begin VB.TextBox txtng 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   121
         Left            =   1170
         TabIndex        =   282
         Top             =   1080
         Width           =   555
      End
      Begin VB.TextBox txtng 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   122
         Left            =   1710
         TabIndex        =   281
         Top             =   1080
         Width           =   555
      End
      Begin VB.TextBox txtng 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   123
         Left            =   2250
         TabIndex        =   280
         Top             =   1080
         Width           =   555
      End
      Begin VB.Label lblMesin 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "31"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   465
         Index           =   31
         Left            =   90
         TabIndex        =   302
         Top             =   135
         Width           =   510
      End
      Begin VB.Label lbljam2R 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "10 : R"
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
         Height          =   240
         Index           =   30
         Left            =   2250
         TabIndex        =   301
         Top             =   585
         Width           =   555
      End
      Begin VB.Label lbljam2L 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "10 : L"
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
         Height          =   240
         Index           =   30
         Left            =   1710
         TabIndex        =   300
         Top             =   585
         Width           =   555
      End
      Begin VB.Label lbljam1R 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "9 : R"
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
         Height          =   240
         Index           =   30
         Left            =   1170
         TabIndex        =   299
         Top             =   585
         Width           =   555
      End
      Begin VB.Label lbljam1L 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "9 : L"
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
         Height          =   240
         Index           =   30
         Left            =   630
         TabIndex        =   298
         Top             =   585
         Width           =   555
      End
      Begin VB.Label lblIdle 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "           IDLE"
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
         Height          =   600
         Index           =   31
         Left            =   90
         TabIndex        =   297
         Top             =   1485
         Width           =   510
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "NG"
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
         Height          =   330
         Index           =   181
         Left            =   90
         TabIndex        =   296
         Top             =   1125
         Width           =   510
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "GROSS"
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
         Height          =   285
         Index           =   182
         Left            =   90
         TabIndex        =   295
         Top             =   810
         Width           =   510
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SHIFT 3"
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
         Height          =   240
         Index           =   183
         Left            =   2070
         TabIndex        =   294
         Top             =   1485
         Width           =   735
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SHIFT 2"
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
         Height          =   240
         Index           =   184
         Left            =   1350
         TabIndex        =   293
         Top             =   1485
         Width           =   735
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SHIFT 1"
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
         Height          =   240
         Index           =   185
         Left            =   630
         TabIndex        =   292
         Top             =   1485
         Width           =   735
      End
   End
   Begin VB.Frame frameDesign 
      BackColor       =   &H00FFFFFF&
      Height          =   2175
      Index           =   31
      Left            =   17820
      TabIndex        =   255
      Top             =   2925
      Width           =   2850
      Begin VB.TextBox txtng 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   124
         Left            =   2250
         TabIndex        =   267
         Top             =   1080
         Width           =   555
      End
      Begin VB.TextBox txtng 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   125
         Left            =   1710
         TabIndex        =   266
         Top             =   1080
         Width           =   555
      End
      Begin VB.TextBox txtng 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   126
         Left            =   1170
         TabIndex        =   265
         Top             =   1080
         Width           =   555
      End
      Begin VB.TextBox txtng 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   127
         Left            =   630
         TabIndex        =   264
         Top             =   1080
         Width           =   555
      End
      Begin VB.TextBox txtidle 
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
         Index           =   93
         Left            =   2070
         TabIndex        =   263
         Top             =   1710
         Width           =   735
      End
      Begin VB.TextBox txtidle 
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
         Index           =   94
         Left            =   1350
         TabIndex        =   262
         Top             =   1710
         Width           =   735
      End
      Begin VB.TextBox txtShot 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   124
         Left            =   2250
         TabIndex        =   261
         Top             =   810
         Width           =   555
      End
      Begin VB.TextBox txtShot 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   125
         Left            =   1710
         TabIndex        =   260
         Top             =   810
         Width           =   555
      End
      Begin VB.TextBox txtShot 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   126
         Left            =   1170
         TabIndex        =   259
         Top             =   810
         Width           =   555
      End
      Begin VB.TextBox txtShot 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   127
         Left            =   630
         TabIndex        =   258
         Top             =   810
         Width           =   555
      End
      Begin VB.TextBox txtidle 
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
         Index           =   95
         Left            =   630
         TabIndex        =   257
         Top             =   1710
         Width           =   735
      End
      Begin VB.TextBox txtInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   31
         Left            =   630
         MultiLine       =   -1  'True
         TabIndex        =   256
         Top             =   135
         Width           =   2175
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SHIFT 1"
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
         Height          =   240
         Index           =   186
         Left            =   630
         TabIndex        =   278
         Top             =   1485
         Width           =   735
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SHIFT 2"
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
         Height          =   240
         Index           =   187
         Left            =   1350
         TabIndex        =   277
         Top             =   1485
         Width           =   735
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SHIFT 3"
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
         Height          =   240
         Index           =   188
         Left            =   2070
         TabIndex        =   276
         Top             =   1485
         Width           =   735
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "GROSS"
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
         Height          =   285
         Index           =   189
         Left            =   90
         TabIndex        =   275
         Top             =   810
         Width           =   510
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "NG"
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
         Height          =   330
         Index           =   190
         Left            =   90
         TabIndex        =   274
         Top             =   1125
         Width           =   510
      End
      Begin VB.Label lblIdle 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "           IDLE"
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
         Height          =   600
         Index           =   32
         Left            =   90
         TabIndex        =   273
         Top             =   1485
         Width           =   510
      End
      Begin VB.Label lbljam1L 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "9 : L"
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
         Height          =   240
         Index           =   31
         Left            =   630
         TabIndex        =   272
         Top             =   585
         Width           =   555
      End
      Begin VB.Label lbljam1R 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "9 : R"
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
         Height          =   240
         Index           =   31
         Left            =   1170
         TabIndex        =   271
         Top             =   585
         Width           =   555
      End
      Begin VB.Label lbljam2L 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "10 : L"
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
         Height          =   240
         Index           =   31
         Left            =   1710
         TabIndex        =   270
         Top             =   585
         Width           =   555
      End
      Begin VB.Label lbljam2R 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "10 : R"
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
         Height          =   240
         Index           =   31
         Left            =   2250
         TabIndex        =   269
         Top             =   585
         Width           =   555
      End
      Begin VB.Label lblMesin 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "32"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   465
         Index           =   32
         Left            =   90
         TabIndex        =   268
         Top             =   135
         Width           =   510
      End
   End
   Begin VB.Frame frameDesign 
      BackColor       =   &H00FFFFFF&
      Height          =   2175
      Index           =   32
      Left            =   17820
      TabIndex        =   231
      Top             =   5220
      Width           =   2850
      Begin VB.TextBox txtInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   32
         Left            =   630
         MultiLine       =   -1  'True
         TabIndex        =   243
         Top             =   135
         Width           =   2175
      End
      Begin VB.TextBox txtidle 
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
         Index           =   96
         Left            =   630
         TabIndex        =   242
         Top             =   1710
         Width           =   735
      End
      Begin VB.TextBox txtShot 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   128
         Left            =   630
         TabIndex        =   241
         Top             =   810
         Width           =   555
      End
      Begin VB.TextBox txtShot 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   129
         Left            =   1170
         TabIndex        =   240
         Top             =   810
         Width           =   555
      End
      Begin VB.TextBox txtShot 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   130
         Left            =   1710
         TabIndex        =   239
         Top             =   810
         Width           =   555
      End
      Begin VB.TextBox txtShot 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   131
         Left            =   2250
         TabIndex        =   238
         Top             =   810
         Width           =   555
      End
      Begin VB.TextBox txtidle 
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
         Index           =   97
         Left            =   1350
         TabIndex        =   237
         Top             =   1710
         Width           =   735
      End
      Begin VB.TextBox txtidle 
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
         Index           =   98
         Left            =   2070
         TabIndex        =   236
         Top             =   1710
         Width           =   735
      End
      Begin VB.TextBox txtng 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   128
         Left            =   630
         TabIndex        =   235
         Top             =   1080
         Width           =   555
      End
      Begin VB.TextBox txtng 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   129
         Left            =   1170
         TabIndex        =   234
         Top             =   1080
         Width           =   555
      End
      Begin VB.TextBox txtng 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   130
         Left            =   1710
         TabIndex        =   233
         Top             =   1080
         Width           =   555
      End
      Begin VB.TextBox txtng 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   131
         Left            =   2250
         TabIndex        =   232
         Top             =   1080
         Width           =   555
      End
      Begin VB.Label lblMesin 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "33"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   465
         Index           =   33
         Left            =   90
         TabIndex        =   254
         Top             =   135
         Width           =   510
      End
      Begin VB.Label lbljam2R 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "10 : R"
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
         Height          =   240
         Index           =   32
         Left            =   2250
         TabIndex        =   253
         Top             =   585
         Width           =   555
      End
      Begin VB.Label lbljam2L 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "10 : L"
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
         Height          =   240
         Index           =   32
         Left            =   1710
         TabIndex        =   252
         Top             =   585
         Width           =   555
      End
      Begin VB.Label lbljam1R 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "9 : R"
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
         Height          =   240
         Index           =   32
         Left            =   1170
         TabIndex        =   251
         Top             =   585
         Width           =   555
      End
      Begin VB.Label lbljam1L 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "9 : L"
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
         Height          =   240
         Index           =   32
         Left            =   630
         TabIndex        =   250
         Top             =   585
         Width           =   555
      End
      Begin VB.Label lblIdle 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "           IDLE"
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
         Height          =   600
         Index           =   33
         Left            =   90
         TabIndex        =   249
         Top             =   1485
         Width           =   510
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "NG"
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
         Height          =   330
         Index           =   193
         Left            =   90
         TabIndex        =   248
         Top             =   1125
         Width           =   510
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "GROSS"
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
         Height          =   285
         Index           =   194
         Left            =   90
         TabIndex        =   247
         Top             =   810
         Width           =   510
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SHIFT 3"
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
         Height          =   240
         Index           =   195
         Left            =   2070
         TabIndex        =   246
         Top             =   1485
         Width           =   735
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SHIFT 2"
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
         Height          =   240
         Index           =   196
         Left            =   1350
         TabIndex        =   245
         Top             =   1485
         Width           =   735
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SHIFT 1"
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
         Height          =   240
         Index           =   197
         Left            =   630
         TabIndex        =   244
         Top             =   1485
         Width           =   735
      End
   End
   Begin VB.Frame frameDesign 
      BackColor       =   &H00FFFFFF&
      Height          =   2175
      Index           =   33
      Left            =   17820
      TabIndex        =   207
      Top             =   7515
      Width           =   2850
      Begin VB.TextBox txtInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   33
         Left            =   630
         MultiLine       =   -1  'True
         TabIndex        =   219
         Top             =   135
         Width           =   2175
      End
      Begin VB.TextBox txtidle 
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
         Index           =   99
         Left            =   630
         TabIndex        =   218
         Top             =   1710
         Width           =   735
      End
      Begin VB.TextBox txtShot 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   132
         Left            =   630
         TabIndex        =   217
         Top             =   810
         Width           =   555
      End
      Begin VB.TextBox txtShot 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   133
         Left            =   1170
         TabIndex        =   216
         Top             =   810
         Width           =   555
      End
      Begin VB.TextBox txtShot 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   134
         Left            =   1710
         TabIndex        =   215
         Top             =   810
         Width           =   555
      End
      Begin VB.TextBox txtShot 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   135
         Left            =   2250
         TabIndex        =   214
         Top             =   810
         Width           =   555
      End
      Begin VB.TextBox txtidle 
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
         Index           =   100
         Left            =   1350
         TabIndex        =   213
         Top             =   1710
         Width           =   735
      End
      Begin VB.TextBox txtidle 
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
         Index           =   101
         Left            =   2070
         TabIndex        =   212
         Top             =   1710
         Width           =   735
      End
      Begin VB.TextBox txtng 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   132
         Left            =   630
         TabIndex        =   211
         Top             =   1080
         Width           =   555
      End
      Begin VB.TextBox txtng 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   133
         Left            =   1170
         TabIndex        =   210
         Top             =   1080
         Width           =   555
      End
      Begin VB.TextBox txtng 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   134
         Left            =   1710
         TabIndex        =   209
         Top             =   1080
         Width           =   555
      End
      Begin VB.TextBox txtng 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   135
         Left            =   2250
         TabIndex        =   208
         Top             =   1080
         Width           =   555
      End
      Begin VB.Label lblMesin 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "34"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   465
         Index           =   34
         Left            =   90
         TabIndex        =   230
         Top             =   135
         Width           =   510
      End
      Begin VB.Label lbljam2R 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "10 : R"
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
         Height          =   240
         Index           =   33
         Left            =   2250
         TabIndex        =   229
         Top             =   585
         Width           =   555
      End
      Begin VB.Label lbljam2L 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "10 : L"
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
         Height          =   240
         Index           =   33
         Left            =   1710
         TabIndex        =   228
         Top             =   585
         Width           =   555
      End
      Begin VB.Label lbljam1R 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "9 : R"
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
         Height          =   240
         Index           =   33
         Left            =   1170
         TabIndex        =   227
         Top             =   585
         Width           =   555
      End
      Begin VB.Label lbljam1L 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "9 : L"
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
         Height          =   240
         Index           =   33
         Left            =   630
         TabIndex        =   226
         Top             =   585
         Width           =   555
      End
      Begin VB.Label lblIdle 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "           IDLE"
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
         Height          =   600
         Index           =   34
         Left            =   90
         TabIndex        =   225
         Top             =   1485
         Width           =   510
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "NG"
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
         Height          =   330
         Index           =   199
         Left            =   90
         TabIndex        =   224
         Top             =   1125
         Width           =   510
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "GROSS"
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
         Height          =   285
         Index           =   200
         Left            =   90
         TabIndex        =   223
         Top             =   810
         Width           =   510
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SHIFT 3"
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
         Height          =   240
         Index           =   201
         Left            =   2070
         TabIndex        =   222
         Top             =   1485
         Width           =   735
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SHIFT 2"
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
         Height          =   240
         Index           =   202
         Left            =   1350
         TabIndex        =   221
         Top             =   1485
         Width           =   735
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SHIFT 1"
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
         Height          =   240
         Index           =   203
         Left            =   630
         TabIndex        =   220
         Top             =   1485
         Width           =   735
      End
   End
   Begin VB.Frame frameDesign 
      BackColor       =   &H00FFFFFF&
      Height          =   2175
      Index           =   34
      Left            =   17820
      TabIndex        =   183
      Top             =   9810
      Width           =   2850
      Begin VB.TextBox txtng 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   136
         Left            =   2250
         TabIndex        =   195
         Top             =   1080
         Width           =   555
      End
      Begin VB.TextBox txtng 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   137
         Left            =   1710
         TabIndex        =   194
         Top             =   1080
         Width           =   555
      End
      Begin VB.TextBox txtng 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   138
         Left            =   1170
         TabIndex        =   193
         Top             =   1080
         Width           =   555
      End
      Begin VB.TextBox txtng 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   139
         Left            =   630
         TabIndex        =   192
         Top             =   1080
         Width           =   555
      End
      Begin VB.TextBox txtidle 
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
         Index           =   102
         Left            =   2070
         TabIndex        =   191
         Top             =   1710
         Width           =   735
      End
      Begin VB.TextBox txtidle 
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
         Index           =   103
         Left            =   1350
         TabIndex        =   190
         Top             =   1710
         Width           =   735
      End
      Begin VB.TextBox txtShot 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   136
         Left            =   2250
         TabIndex        =   189
         Top             =   810
         Width           =   555
      End
      Begin VB.TextBox txtShot 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   137
         Left            =   1710
         TabIndex        =   188
         Top             =   810
         Width           =   555
      End
      Begin VB.TextBox txtShot 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   138
         Left            =   1170
         TabIndex        =   187
         Top             =   810
         Width           =   555
      End
      Begin VB.TextBox txtShot 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   139
         Left            =   630
         TabIndex        =   186
         Top             =   810
         Width           =   555
      End
      Begin VB.TextBox txtidle 
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
         Index           =   104
         Left            =   630
         TabIndex        =   185
         Top             =   1710
         Width           =   735
      End
      Begin VB.TextBox txtInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   34
         Left            =   630
         MultiLine       =   -1  'True
         TabIndex        =   184
         Top             =   135
         Width           =   2175
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SHIFT 1"
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
         Height          =   240
         Index           =   204
         Left            =   630
         TabIndex        =   206
         Top             =   1485
         Width           =   735
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SHIFT 2"
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
         Height          =   240
         Index           =   205
         Left            =   1350
         TabIndex        =   205
         Top             =   1485
         Width           =   735
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SHIFT 3"
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
         Height          =   240
         Index           =   206
         Left            =   2070
         TabIndex        =   204
         Top             =   1485
         Width           =   735
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "GROSS"
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
         Height          =   285
         Index           =   207
         Left            =   90
         TabIndex        =   203
         Top             =   810
         Width           =   510
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "NG"
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
         Height          =   330
         Index           =   208
         Left            =   90
         TabIndex        =   202
         Top             =   1125
         Width           =   510
      End
      Begin VB.Label lblIdle 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "           IDLE"
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
         Height          =   600
         Index           =   35
         Left            =   90
         TabIndex        =   201
         Top             =   1485
         Width           =   510
      End
      Begin VB.Label lbljam1L 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "9 : L"
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
         Height          =   240
         Index           =   34
         Left            =   630
         TabIndex        =   200
         Top             =   585
         Width           =   555
      End
      Begin VB.Label lbljam1R 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "9 : R"
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
         Height          =   240
         Index           =   34
         Left            =   1170
         TabIndex        =   199
         Top             =   585
         Width           =   555
      End
      Begin VB.Label lbljam2L 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "10 : L"
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
         Height          =   240
         Index           =   34
         Left            =   1710
         TabIndex        =   198
         Top             =   585
         Width           =   555
      End
      Begin VB.Label lbljam2R 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "10 : R"
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
         Height          =   240
         Index           =   34
         Left            =   2250
         TabIndex        =   197
         Top             =   585
         Width           =   555
      End
      Begin VB.Label lblMesin 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "35"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   465
         Index           =   35
         Left            =   90
         TabIndex        =   196
         Top             =   135
         Width           =   510
      End
   End
   Begin VB.Frame frameDesign 
      BackColor       =   &H00FFFFFF&
      Height          =   2175
      Index           =   35
      Left            =   20790
      TabIndex        =   159
      Top             =   630
      Width           =   2850
      Begin VB.TextBox txtInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   35
         Left            =   630
         MultiLine       =   -1  'True
         TabIndex        =   171
         Top             =   135
         Width           =   2175
      End
      Begin VB.TextBox txtidle 
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
         Index           =   105
         Left            =   630
         TabIndex        =   170
         Top             =   1710
         Width           =   735
      End
      Begin VB.TextBox txtShot 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   140
         Left            =   630
         TabIndex        =   169
         Top             =   810
         Width           =   555
      End
      Begin VB.TextBox txtShot 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   141
         Left            =   1170
         TabIndex        =   168
         Top             =   810
         Width           =   555
      End
      Begin VB.TextBox txtShot 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   142
         Left            =   1710
         TabIndex        =   167
         Top             =   810
         Width           =   555
      End
      Begin VB.TextBox txtShot 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   143
         Left            =   2250
         TabIndex        =   166
         Top             =   810
         Width           =   555
      End
      Begin VB.TextBox txtidle 
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
         Index           =   106
         Left            =   1350
         TabIndex        =   165
         Top             =   1710
         Width           =   735
      End
      Begin VB.TextBox txtidle 
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
         Index           =   107
         Left            =   2070
         TabIndex        =   164
         Top             =   1710
         Width           =   735
      End
      Begin VB.TextBox txtng 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   140
         Left            =   630
         TabIndex        =   163
         Top             =   1080
         Width           =   555
      End
      Begin VB.TextBox txtng 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   141
         Left            =   1170
         TabIndex        =   162
         Top             =   1080
         Width           =   555
      End
      Begin VB.TextBox txtng 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   142
         Left            =   1710
         TabIndex        =   161
         Top             =   1080
         Width           =   555
      End
      Begin VB.TextBox txtng 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   143
         Left            =   2250
         TabIndex        =   160
         Top             =   1080
         Width           =   555
      End
      Begin VB.Label lblMesin 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "36"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   465
         Index           =   36
         Left            =   90
         TabIndex        =   182
         Top             =   135
         Width           =   510
      End
      Begin VB.Label lbljam2R 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "10 : R"
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
         Height          =   240
         Index           =   35
         Left            =   2250
         TabIndex        =   181
         Top             =   585
         Width           =   555
      End
      Begin VB.Label lbljam2L 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "10 : L"
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
         Height          =   240
         Index           =   35
         Left            =   1710
         TabIndex        =   180
         Top             =   585
         Width           =   555
      End
      Begin VB.Label lbljam1R 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "9 : R"
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
         Height          =   240
         Index           =   35
         Left            =   1170
         TabIndex        =   179
         Top             =   585
         Width           =   555
      End
      Begin VB.Label lbljam1L 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "9 : L"
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
         Height          =   240
         Index           =   35
         Left            =   630
         TabIndex        =   178
         Top             =   585
         Width           =   555
      End
      Begin VB.Label lblIdle 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "           IDLE"
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
         Height          =   600
         Index           =   36
         Left            =   90
         TabIndex        =   177
         Top             =   1485
         Width           =   510
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "NG"
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
         Height          =   330
         Index           =   211
         Left            =   90
         TabIndex        =   176
         Top             =   1125
         Width           =   510
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "GROSS"
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
         Height          =   285
         Index           =   212
         Left            =   90
         TabIndex        =   175
         Top             =   810
         Width           =   510
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SHIFT 3"
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
         Height          =   240
         Index           =   213
         Left            =   2070
         TabIndex        =   174
         Top             =   1485
         Width           =   735
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SHIFT 2"
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
         Height          =   240
         Index           =   214
         Left            =   1350
         TabIndex        =   173
         Top             =   1485
         Width           =   735
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SHIFT 1"
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
         Height          =   240
         Index           =   215
         Left            =   630
         TabIndex        =   172
         Top             =   1485
         Width           =   735
      End
   End
   Begin VB.Frame frameDesign 
      BackColor       =   &H00FFFFFF&
      Height          =   2175
      Index           =   36
      Left            =   20790
      TabIndex        =   135
      Top             =   2925
      Width           =   2850
      Begin VB.TextBox txtng 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   144
         Left            =   2250
         TabIndex        =   147
         Top             =   1080
         Width           =   555
      End
      Begin VB.TextBox txtng 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   145
         Left            =   1710
         TabIndex        =   146
         Top             =   1080
         Width           =   555
      End
      Begin VB.TextBox txtng 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   146
         Left            =   1170
         TabIndex        =   145
         Top             =   1080
         Width           =   555
      End
      Begin VB.TextBox txtng 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   147
         Left            =   630
         TabIndex        =   144
         Top             =   1080
         Width           =   555
      End
      Begin VB.TextBox txtidle 
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
         Index           =   108
         Left            =   2070
         TabIndex        =   143
         Top             =   1710
         Width           =   735
      End
      Begin VB.TextBox txtidle 
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
         Index           =   109
         Left            =   1350
         TabIndex        =   142
         Top             =   1710
         Width           =   735
      End
      Begin VB.TextBox txtShot 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   144
         Left            =   2250
         TabIndex        =   141
         Top             =   810
         Width           =   555
      End
      Begin VB.TextBox txtShot 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   145
         Left            =   1710
         TabIndex        =   140
         Top             =   810
         Width           =   555
      End
      Begin VB.TextBox txtShot 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   146
         Left            =   1170
         TabIndex        =   139
         Top             =   810
         Width           =   555
      End
      Begin VB.TextBox txtShot 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   147
         Left            =   630
         TabIndex        =   138
         Top             =   810
         Width           =   555
      End
      Begin VB.TextBox txtidle 
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
         Index           =   110
         Left            =   630
         TabIndex        =   137
         Top             =   1710
         Width           =   735
      End
      Begin VB.TextBox txtInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   36
         Left            =   630
         MultiLine       =   -1  'True
         TabIndex        =   136
         Top             =   135
         Width           =   2175
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SHIFT 1"
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
         Height          =   240
         Index           =   216
         Left            =   630
         TabIndex        =   158
         Top             =   1485
         Width           =   735
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SHIFT 2"
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
         Height          =   240
         Index           =   217
         Left            =   1350
         TabIndex        =   157
         Top             =   1485
         Width           =   735
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SHIFT 3"
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
         Height          =   240
         Index           =   218
         Left            =   2070
         TabIndex        =   156
         Top             =   1485
         Width           =   735
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "GROSS"
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
         Height          =   285
         Index           =   219
         Left            =   90
         TabIndex        =   155
         Top             =   810
         Width           =   510
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "NG"
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
         Height          =   330
         Index           =   220
         Left            =   90
         TabIndex        =   154
         Top             =   1125
         Width           =   510
      End
      Begin VB.Label lblIdle 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "           IDLE"
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
         Height          =   600
         Index           =   37
         Left            =   90
         TabIndex        =   153
         Top             =   1485
         Width           =   510
      End
      Begin VB.Label lbljam1L 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "9 : L"
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
         Height          =   240
         Index           =   36
         Left            =   630
         TabIndex        =   152
         Top             =   585
         Width           =   555
      End
      Begin VB.Label lbljam1R 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "9 : R"
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
         Height          =   240
         Index           =   36
         Left            =   1170
         TabIndex        =   151
         Top             =   585
         Width           =   555
      End
      Begin VB.Label lbljam2L 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "10 : L"
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
         Height          =   240
         Index           =   36
         Left            =   1710
         TabIndex        =   150
         Top             =   585
         Width           =   555
      End
      Begin VB.Label lbljam2R 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "10 : R"
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
         Height          =   240
         Index           =   36
         Left            =   2250
         TabIndex        =   149
         Top             =   585
         Width           =   555
      End
      Begin VB.Label lblMesin 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "37"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   465
         Index           =   37
         Left            =   90
         TabIndex        =   148
         Top             =   135
         Width           =   510
      End
   End
   Begin VB.Frame frameDesign 
      BackColor       =   &H00FFFFFF&
      Height          =   2175
      Index           =   37
      Left            =   20790
      TabIndex        =   111
      Top             =   5220
      Width           =   2850
      Begin VB.TextBox txtInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   37
         Left            =   630
         MultiLine       =   -1  'True
         TabIndex        =   123
         Top             =   135
         Width           =   2175
      End
      Begin VB.TextBox txtidle 
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
         Index           =   111
         Left            =   630
         TabIndex        =   122
         Top             =   1710
         Width           =   735
      End
      Begin VB.TextBox txtShot 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   148
         Left            =   630
         TabIndex        =   121
         Top             =   810
         Width           =   555
      End
      Begin VB.TextBox txtShot 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   149
         Left            =   1170
         TabIndex        =   120
         Top             =   810
         Width           =   555
      End
      Begin VB.TextBox txtShot 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   150
         Left            =   1710
         TabIndex        =   119
         Top             =   810
         Width           =   555
      End
      Begin VB.TextBox txtShot 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   151
         Left            =   2250
         TabIndex        =   118
         Top             =   810
         Width           =   555
      End
      Begin VB.TextBox txtidle 
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
         Index           =   112
         Left            =   1350
         TabIndex        =   117
         Top             =   1710
         Width           =   735
      End
      Begin VB.TextBox txtidle 
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
         Index           =   113
         Left            =   2070
         TabIndex        =   116
         Top             =   1710
         Width           =   735
      End
      Begin VB.TextBox txtng 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   148
         Left            =   630
         TabIndex        =   115
         Top             =   1080
         Width           =   555
      End
      Begin VB.TextBox txtng 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   149
         Left            =   1170
         TabIndex        =   114
         Top             =   1080
         Width           =   555
      End
      Begin VB.TextBox txtng 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   150
         Left            =   1710
         TabIndex        =   113
         Top             =   1080
         Width           =   555
      End
      Begin VB.TextBox txtng 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   151
         Left            =   2250
         TabIndex        =   112
         Top             =   1080
         Width           =   555
      End
      Begin VB.Label lblMesin 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "38"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   465
         Index           =   38
         Left            =   90
         TabIndex        =   134
         Top             =   135
         Width           =   510
      End
      Begin VB.Label lbljam2R 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "10 : R"
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
         Height          =   240
         Index           =   37
         Left            =   2250
         TabIndex        =   133
         Top             =   585
         Width           =   555
      End
      Begin VB.Label lbljam2L 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "10 : L"
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
         Height          =   240
         Index           =   37
         Left            =   1710
         TabIndex        =   132
         Top             =   585
         Width           =   555
      End
      Begin VB.Label lbljam1R 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "9 : R"
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
         Height          =   240
         Index           =   37
         Left            =   1170
         TabIndex        =   131
         Top             =   585
         Width           =   555
      End
      Begin VB.Label lbljam1L 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "9 : L"
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
         Height          =   240
         Index           =   37
         Left            =   630
         TabIndex        =   130
         Top             =   585
         Width           =   555
      End
      Begin VB.Label lblIdle 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "           IDLE"
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
         Height          =   600
         Index           =   38
         Left            =   90
         TabIndex        =   129
         Top             =   1485
         Width           =   510
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "NG"
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
         Height          =   330
         Index           =   223
         Left            =   90
         TabIndex        =   128
         Top             =   1125
         Width           =   510
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "GROSS"
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
         Height          =   285
         Index           =   224
         Left            =   90
         TabIndex        =   127
         Top             =   810
         Width           =   510
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SHIFT 3"
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
         Height          =   240
         Index           =   225
         Left            =   2070
         TabIndex        =   126
         Top             =   1485
         Width           =   735
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SHIFT 2"
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
         Height          =   240
         Index           =   226
         Left            =   1350
         TabIndex        =   125
         Top             =   1485
         Width           =   735
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SHIFT 1"
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
         Height          =   240
         Index           =   227
         Left            =   630
         TabIndex        =   124
         Top             =   1485
         Width           =   735
      End
   End
   Begin VB.Frame frameDesign 
      BackColor       =   &H00FFFFFF&
      Height          =   2175
      Index           =   38
      Left            =   20790
      TabIndex        =   87
      Top             =   7515
      Width           =   2850
      Begin VB.TextBox txtInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   38
         Left            =   630
         MultiLine       =   -1  'True
         TabIndex        =   99
         Top             =   135
         Width           =   2175
      End
      Begin VB.TextBox txtidle 
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
         Index           =   114
         Left            =   630
         TabIndex        =   98
         Top             =   1710
         Width           =   735
      End
      Begin VB.TextBox txtShot 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   152
         Left            =   630
         TabIndex        =   97
         Top             =   810
         Width           =   555
      End
      Begin VB.TextBox txtShot 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   153
         Left            =   1170
         TabIndex        =   96
         Top             =   810
         Width           =   555
      End
      Begin VB.TextBox txtShot 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   154
         Left            =   1710
         TabIndex        =   95
         Top             =   810
         Width           =   555
      End
      Begin VB.TextBox txtShot 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   155
         Left            =   2250
         TabIndex        =   94
         Top             =   810
         Width           =   555
      End
      Begin VB.TextBox txtidle 
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
         Index           =   115
         Left            =   1350
         TabIndex        =   93
         Top             =   1710
         Width           =   735
      End
      Begin VB.TextBox txtidle 
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
         Index           =   116
         Left            =   2070
         TabIndex        =   92
         Top             =   1710
         Width           =   735
      End
      Begin VB.TextBox txtng 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   152
         Left            =   630
         TabIndex        =   91
         Top             =   1080
         Width           =   555
      End
      Begin VB.TextBox txtng 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   153
         Left            =   1170
         TabIndex        =   90
         Top             =   1080
         Width           =   555
      End
      Begin VB.TextBox txtng 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   154
         Left            =   1710
         TabIndex        =   89
         Top             =   1080
         Width           =   555
      End
      Begin VB.TextBox txtng 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   155
         Left            =   2250
         TabIndex        =   88
         Top             =   1080
         Width           =   555
      End
      Begin VB.Label lblMesin 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "39"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   465
         Index           =   39
         Left            =   90
         TabIndex        =   110
         Top             =   135
         Width           =   510
      End
      Begin VB.Label lbljam2R 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "10 : R"
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
         Height          =   240
         Index           =   38
         Left            =   2250
         TabIndex        =   109
         Top             =   585
         Width           =   555
      End
      Begin VB.Label lbljam2L 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "10 : L"
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
         Height          =   240
         Index           =   38
         Left            =   1710
         TabIndex        =   108
         Top             =   585
         Width           =   555
      End
      Begin VB.Label lbljam1R 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "9 : R"
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
         Height          =   240
         Index           =   38
         Left            =   1170
         TabIndex        =   107
         Top             =   585
         Width           =   555
      End
      Begin VB.Label lbljam1L 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "9 : L"
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
         Height          =   240
         Index           =   38
         Left            =   630
         TabIndex        =   106
         Top             =   585
         Width           =   555
      End
      Begin VB.Label lblIdle 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "           IDLE"
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
         Height          =   600
         Index           =   39
         Left            =   90
         TabIndex        =   105
         Top             =   1485
         Width           =   510
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "NG"
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
         Height          =   330
         Index           =   229
         Left            =   90
         TabIndex        =   104
         Top             =   1125
         Width           =   510
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "GROSS"
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
         Height          =   285
         Index           =   230
         Left            =   90
         TabIndex        =   103
         Top             =   810
         Width           =   510
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SHIFT 3"
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
         Height          =   240
         Index           =   231
         Left            =   2070
         TabIndex        =   102
         Top             =   1485
         Width           =   735
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SHIFT 2"
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
         Height          =   240
         Index           =   232
         Left            =   1350
         TabIndex        =   101
         Top             =   1485
         Width           =   735
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SHIFT 1"
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
         Height          =   240
         Index           =   233
         Left            =   630
         TabIndex        =   100
         Top             =   1485
         Width           =   735
      End
   End
   Begin VB.Frame frameDesign 
      BackColor       =   &H00FFFFFF&
      Height          =   2175
      Index           =   39
      Left            =   20790
      TabIndex        =   63
      Top             =   9810
      Width           =   2850
      Begin VB.TextBox txtng 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   156
         Left            =   2250
         TabIndex        =   75
         Top             =   1080
         Width           =   555
      End
      Begin VB.TextBox txtng 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   157
         Left            =   1710
         TabIndex        =   74
         Top             =   1080
         Width           =   555
      End
      Begin VB.TextBox txtng 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   158
         Left            =   1170
         TabIndex        =   73
         Top             =   1080
         Width           =   555
      End
      Begin VB.TextBox txtng 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   159
         Left            =   630
         TabIndex        =   72
         Top             =   1080
         Width           =   555
      End
      Begin VB.TextBox txtidle 
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
         Index           =   117
         Left            =   2070
         TabIndex        =   71
         Top             =   1710
         Width           =   735
      End
      Begin VB.TextBox txtidle 
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
         Index           =   118
         Left            =   1350
         TabIndex        =   70
         Top             =   1710
         Width           =   735
      End
      Begin VB.TextBox txtShot 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   156
         Left            =   2250
         TabIndex        =   69
         Top             =   810
         Width           =   555
      End
      Begin VB.TextBox txtShot 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   157
         Left            =   1710
         TabIndex        =   68
         Top             =   810
         Width           =   555
      End
      Begin VB.TextBox txtShot 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   158
         Left            =   1170
         TabIndex        =   67
         Top             =   810
         Width           =   555
      End
      Begin VB.TextBox txtShot 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   159
         Left            =   630
         TabIndex        =   66
         Top             =   810
         Width           =   555
      End
      Begin VB.TextBox txtidle 
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
         Index           =   119
         Left            =   630
         TabIndex        =   65
         Top             =   1710
         Width           =   735
      End
      Begin VB.TextBox txtInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   39
         Left            =   630
         MultiLine       =   -1  'True
         TabIndex        =   64
         Top             =   135
         Width           =   2175
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SHIFT 1"
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
         Height          =   240
         Index           =   234
         Left            =   630
         TabIndex        =   86
         Top             =   1485
         Width           =   735
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SHIFT 2"
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
         Height          =   240
         Index           =   235
         Left            =   1350
         TabIndex        =   85
         Top             =   1485
         Width           =   735
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SHIFT 3"
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
         Height          =   240
         Index           =   236
         Left            =   2070
         TabIndex        =   84
         Top             =   1485
         Width           =   735
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "GROSS"
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
         Height          =   285
         Index           =   237
         Left            =   90
         TabIndex        =   83
         Top             =   810
         Width           =   510
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "NG"
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
         Height          =   330
         Index           =   238
         Left            =   90
         TabIndex        =   82
         Top             =   1125
         Width           =   510
      End
      Begin VB.Label lblIdle 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "           IDLE"
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
         Height          =   600
         Index           =   40
         Left            =   90
         TabIndex        =   81
         Top             =   1485
         Width           =   510
      End
      Begin VB.Label lbljam1L 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "9 : L"
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
         Height          =   240
         Index           =   39
         Left            =   630
         TabIndex        =   80
         Top             =   585
         Width           =   555
      End
      Begin VB.Label lbljam1R 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "9 : R"
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
         Height          =   240
         Index           =   39
         Left            =   1170
         TabIndex        =   79
         Top             =   585
         Width           =   555
      End
      Begin VB.Label lbljam2L 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "10 : L"
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
         Height          =   240
         Index           =   39
         Left            =   1710
         TabIndex        =   78
         Top             =   585
         Width           =   555
      End
      Begin VB.Label lbljam2R 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "10 : R"
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
         Height          =   240
         Index           =   39
         Left            =   2250
         TabIndex        =   77
         Top             =   585
         Width           =   555
      End
      Begin VB.Label lblMesin 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "40"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   465
         Index           =   40
         Left            =   90
         TabIndex        =   76
         Top             =   135
         Width           =   510
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   60000
      Left            =   135
      Top             =   0
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Informasi Warna"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   0
      TabIndex        =   52
      Top             =   12060
      Width           =   26610
      Begin VB.Shape Shape1 
         BackColor       =   &H00008000&
         BackStyle       =   1  'Opaque
         Height          =   240
         Left            =   270
         Top             =   270
         Width           =   240
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "LOAD"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   630
         TabIndex        =   62
         Top             =   315
         Width           =   510
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00004080&
         BackStyle       =   1  'Opaque
         Height          =   240
         Left            =   1350
         Top             =   270
         Width           =   240
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "NO LOAD"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1710
         TabIndex        =   61
         Top             =   315
         Width           =   870
      End
      Begin VB.Shape Shape3 
         BackColor       =   &H00008080&
         BackStyle       =   1  'Opaque
         Height          =   240
         Left            =   2700
         Top             =   270
         Width           =   240
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "MAINTENANCE"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3060
         TabIndex        =   60
         Top             =   315
         Width           =   1455
      End
      Begin VB.Shape Shape4 
         BackColor       =   &H00808000&
         BackStyle       =   1  'Opaque
         Height          =   240
         Left            =   4500
         Top             =   270
         Width           =   240
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFFFFF&
         Caption         =   "BROKEN"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   4860
         TabIndex        =   59
         Top             =   315
         Width           =   870
      End
      Begin VB.Shape Shape5 
         BackColor       =   &H00C000C0&
         BackStyle       =   1  'Opaque
         Height          =   240
         Left            =   5715
         Top             =   270
         Width           =   240
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFFFFF&
         Caption         =   "TRIAL"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   6075
         TabIndex        =   58
         Top             =   315
         Width           =   780
      End
      Begin VB.Shape Shape6 
         BackColor       =   &H000000FF&
         BackStyle       =   1  'Opaque
         Height          =   240
         Left            =   6885
         Top             =   270
         Width           =   240
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FFFFFF&
         Caption         =   "IDLE MACHINE"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   7245
         TabIndex        =   57
         Top             =   315
         Width           =   1455
      End
      Begin VB.Label Label8 
         BackColor       =   &H00FFFFFF&
         Caption         =   "GROSS > 98 %"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   11925
         TabIndex        =   56
         Top             =   315
         Width           =   1455
      End
      Begin VB.Shape Shape7 
         BackColor       =   &H0000FF00&
         BackStyle       =   1  'Opaque
         Height          =   240
         Left            =   11565
         Top             =   270
         Width           =   240
      End
      Begin VB.Label Label9 
         BackColor       =   &H00FFFFFF&
         Caption         =   "GROSS < 98 %"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   14850
         TabIndex        =   55
         Top             =   315
         Width           =   1455
      End
      Begin VB.Shape Shape8 
         BackColor       =   &H000000FF&
         BackStyle       =   1  'Opaque
         Height          =   240
         Left            =   14490
         Top             =   270
         Width           =   240
      End
      Begin VB.Shape Shape9 
         BackColor       =   &H0000FF00&
         BackStyle       =   1  'Opaque
         Height          =   240
         Left            =   17370
         Top             =   225
         Width           =   240
      End
      Begin VB.Shape Shape10 
         BackColor       =   &H000000FF&
         BackStyle       =   1  'Opaque
         Height          =   240
         Left            =   20205
         Top             =   225
         Width           =   240
      End
      Begin VB.Label Label10 
         BackColor       =   &H00FFFFFF&
         Caption         =   "NG < 5 %"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   17775
         TabIndex        =   54
         Top             =   270
         Width           =   1455
      End
      Begin VB.Label Label11 
         BackColor       =   &H00FFFFFF&
         Caption         =   "NG > 5 %"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   20610
         TabIndex        =   53
         Top             =   270
         Width           =   1455
      End
   End
   Begin VB.Frame frameDesign 
      BackColor       =   &H00FFFFFF&
      Height          =   2175
      Index           =   40
      Left            =   23760
      TabIndex        =   28
      Top             =   630
      Width           =   2850
      Begin VB.TextBox txtInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   40
         Left            =   630
         MultiLine       =   -1  'True
         TabIndex        =   40
         Top             =   135
         Width           =   2175
      End
      Begin VB.TextBox txtidle 
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
         Index           =   120
         Left            =   630
         TabIndex        =   39
         Top             =   1710
         Width           =   735
      End
      Begin VB.TextBox txtShot 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   160
         Left            =   630
         TabIndex        =   38
         Top             =   810
         Width           =   555
      End
      Begin VB.TextBox txtShot 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   161
         Left            =   1170
         TabIndex        =   37
         Top             =   810
         Width           =   555
      End
      Begin VB.TextBox txtShot 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   162
         Left            =   1710
         TabIndex        =   36
         Top             =   810
         Width           =   555
      End
      Begin VB.TextBox txtShot 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   163
         Left            =   2250
         TabIndex        =   35
         Top             =   810
         Width           =   555
      End
      Begin VB.TextBox txtidle 
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
         Index           =   121
         Left            =   1350
         TabIndex        =   34
         Top             =   1710
         Width           =   735
      End
      Begin VB.TextBox txtidle 
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
         Index           =   122
         Left            =   2070
         TabIndex        =   33
         Top             =   1710
         Width           =   735
      End
      Begin VB.TextBox txtng 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   160
         Left            =   630
         TabIndex        =   32
         Top             =   1080
         Width           =   555
      End
      Begin VB.TextBox txtng 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   161
         Left            =   1170
         TabIndex        =   31
         Top             =   1080
         Width           =   555
      End
      Begin VB.TextBox txtng 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   162
         Left            =   1710
         TabIndex        =   30
         Top             =   1080
         Width           =   555
      End
      Begin VB.TextBox txtng 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   163
         Left            =   2250
         TabIndex        =   29
         Top             =   1080
         Width           =   555
      End
      Begin VB.Label lblMesin 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "41"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   465
         Index           =   41
         Left            =   90
         TabIndex        =   51
         Top             =   135
         Width           =   510
      End
      Begin VB.Label lbljam2R 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "10 : R"
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
         Height          =   240
         Index           =   40
         Left            =   2250
         TabIndex        =   50
         Top             =   585
         Width           =   555
      End
      Begin VB.Label lbljam2L 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "10 : L"
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
         Height          =   240
         Index           =   40
         Left            =   1710
         TabIndex        =   49
         Top             =   585
         Width           =   555
      End
      Begin VB.Label lbljam1R 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "9 : R"
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
         Height          =   240
         Index           =   40
         Left            =   1170
         TabIndex        =   48
         Top             =   585
         Width           =   555
      End
      Begin VB.Label lbljam1L 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "9 : L"
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
         Height          =   240
         Index           =   40
         Left            =   630
         TabIndex        =   47
         Top             =   585
         Width           =   555
      End
      Begin VB.Label lblIdle 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "           IDLE"
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
         Height          =   600
         Index           =   41
         Left            =   90
         TabIndex        =   46
         Top             =   1485
         Width           =   510
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "NG"
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
         Height          =   330
         Index           =   0
         Left            =   90
         TabIndex        =   45
         Top             =   1125
         Width           =   510
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "GROSS"
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
         Height          =   285
         Index           =   11
         Left            =   90
         TabIndex        =   44
         Top             =   810
         Width           =   510
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SHIFT 3"
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
         Height          =   240
         Index           =   12
         Left            =   2070
         TabIndex        =   43
         Top             =   1485
         Width           =   735
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SHIFT 2"
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
         Height          =   240
         Index           =   18
         Left            =   1350
         TabIndex        =   42
         Top             =   1485
         Width           =   735
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SHIFT 1"
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
         Height          =   240
         Index           =   24
         Left            =   630
         TabIndex        =   41
         Top             =   1485
         Width           =   735
      End
   End
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   675
      Top             =   45
   End
   Begin VB.Frame frameDesign 
      BackColor       =   &H00FFFFFF&
      Height          =   2175
      Index           =   42
      Left            =   23760
      TabIndex        =   26
      Top             =   5220
      Width           =   2850
   End
   Begin VB.Frame frameDesign 
      BackColor       =   &H00FFFFFF&
      Height          =   2175
      Index           =   43
      Left            =   23760
      TabIndex        =   25
      Top             =   7515
      Width           =   2850
   End
   Begin VB.Frame frameDesign 
      BackColor       =   &H00FFFFFF&
      Height          =   2175
      Index           =   44
      Left            =   23760
      TabIndex        =   24
      Top             =   9810
      Width           =   2850
   End
   Begin VB.Frame frameDesign 
      BackColor       =   &H00FFFFFF&
      Height          =   2175
      Index           =   41
      Left            =   23760
      TabIndex        =   0
      Top             =   2925
      Width           =   2850
      Begin VB.TextBox txtng 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   164
         Left            =   630
         TabIndex        =   12
         Top             =   1080
         Width           =   555
      End
      Begin VB.TextBox txtng 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   165
         Left            =   1170
         TabIndex        =   11
         Top             =   1080
         Width           =   555
      End
      Begin VB.TextBox txtng 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   166
         Left            =   1710
         TabIndex        =   10
         Top             =   1080
         Width           =   555
      End
      Begin VB.TextBox txtng 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   167
         Left            =   2250
         TabIndex        =   9
         Top             =   1080
         Width           =   555
      End
      Begin VB.TextBox txtidle 
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
         Index           =   123
         Left            =   630
         TabIndex        =   8
         Top             =   1710
         Width           =   735
      End
      Begin VB.TextBox txtidle 
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
         Index           =   124
         Left            =   1350
         TabIndex        =   7
         Top             =   1710
         Width           =   735
      End
      Begin VB.TextBox txtShot 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   164
         Left            =   630
         TabIndex        =   6
         Top             =   810
         Width           =   555
      End
      Begin VB.TextBox txtShot 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   165
         Left            =   1170
         TabIndex        =   5
         Top             =   810
         Width           =   555
      End
      Begin VB.TextBox txtShot 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   166
         Left            =   1710
         TabIndex        =   4
         Top             =   810
         Width           =   555
      End
      Begin VB.TextBox txtShot 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   167
         Left            =   2250
         TabIndex        =   3
         Top             =   810
         Width           =   555
      End
      Begin VB.TextBox txtidle 
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
         Index           =   125
         Left            =   2070
         TabIndex        =   2
         Top             =   1710
         Width           =   735
      End
      Begin VB.TextBox txtInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   41
         Left            =   630
         MultiLine       =   -1  'True
         TabIndex        =   1
         Top             =   135
         Width           =   2175
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SHIFT 1"
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
         Height          =   240
         Index           =   35
         Left            =   630
         TabIndex        =   23
         Top             =   1485
         Width           =   735
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SHIFT 2"
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
         Height          =   240
         Index           =   36
         Left            =   1350
         TabIndex        =   22
         Top             =   1485
         Width           =   735
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SHIFT 3"
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
         Height          =   240
         Index           =   47
         Left            =   2070
         TabIndex        =   21
         Top             =   1485
         Width           =   735
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "GROSS"
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
         Height          =   285
         Index           =   53
         Left            =   90
         TabIndex        =   20
         Top             =   810
         Width           =   510
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "NG"
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
         Height          =   330
         Index           =   54
         Left            =   90
         TabIndex        =   19
         Top             =   1125
         Width           =   510
      End
      Begin VB.Label lblIdle 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "           IDLE"
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
         Height          =   600
         Index           =   42
         Left            =   90
         TabIndex        =   18
         Top             =   1485
         Width           =   510
      End
      Begin VB.Label lbljam1L 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "9 : L"
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
         Height          =   240
         Index           =   41
         Left            =   630
         TabIndex        =   17
         Top             =   585
         Width           =   555
      End
      Begin VB.Label lbljam1R 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "9 : R"
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
         Height          =   240
         Index           =   41
         Left            =   1170
         TabIndex        =   16
         Top             =   585
         Width           =   555
      End
      Begin VB.Label lbljam2L 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "10 : L"
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
         Height          =   240
         Index           =   41
         Left            =   1710
         TabIndex        =   15
         Top             =   585
         Width           =   555
      End
      Begin VB.Label lbljam2R 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "10 : R"
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
         Height          =   240
         Index           =   41
         Left            =   2250
         TabIndex        =   14
         Top             =   585
         Width           =   555
      End
      Begin VB.Label lblMesin 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "42"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   465
         Index           =   42
         Left            =   90
         TabIndex        =   13
         Top             =   135
         Width           =   510
      End
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   285
      Left            =   135
      TabIndex        =   27
      Top             =   135
      Width           =   3345
      _ExtentX        =   5900
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "PRODUCTION MONITORING SISTEM"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   0
      TabIndex        =   1023
      Top             =   45
      Width           =   26610
   End
End
Attribute VB_Name = "frmInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Form_Load()

    Call SetMachine
    
    Call CekTime
    
    Call CekData
    
    Call LoadIdle
    
    Call CheckIdle
    
    Call LockText

End Sub

Private Sub LockText()
Dim i As Integer
For i = 0 To 39
    txtInfo(i).Locked = True
Next i
For i = 0 To 156
    txtShot(i).Locked = True
    txtng(i).Locked = True
Next i
For i = 0 To 117
    txtidle(i).Locked = True
Next i
End Sub

Private Sub SetMachine()
    Dim Rs As New ADODB.Recordset
    Dim sSQL As String
    Dim i As Integer
    Dim X As Integer

    Rs.CursorLocation = adUseClient
    sSQL = "SELECT a.plant_mark,a.prod_machine_id,b.number as machine_no,b.name as machine_name,"
    sSQL = sSQL & " a.eng_product_1,c.internal_part_id,c.NAME AS product_name, a.machine_status"
    sSQL = sSQL & " FROM prod_running_products a"
    sSQL = sSQL & " LEFT JOIN sip_production.prod_machines b on a.prod_machine_id = b.id"
    sSQL = sSQL & " LEFT JOIN sip_production.eng_products c ON a.eng_product_1 = c.id"
    sSQL = sSQL & " WHERE a.plant_mark = '" & p_plant_mark & "' and a.`status` = 'active'  Order by b.number ASC"
    
    Rs.Open sSQL, CN, adOpenStatic, adLockOptimistic

    If Rs.RecordCount > 0 Then
        Rs.MoveFirst
        Do While Not Rs.EOF
            
            
            X = Rs.Fields("machine_no")
            i = X - 1
                txtInfo(i).text = Rs.Fields("product_name")
                If Rs.Fields("machine_status") = "loaded" Then
                    lblMesin(X).BackColor = &H8000&
                ElseIf Rs.Fields("machine_status") = "no_load" Then
                    lblMesin(X).BackColor = &H4080&
                ElseIf Rs.Fields("machine_status") = "maintenance" Then
                    lblMesin(X).BackColor = &H8080&
                ElseIf Rs.Fields("machine_status") = "broken" Then
                    lblMesin(X).BackColor = &H808000
                ElseIf Rs.Fields("machine_status") = "trial" Then
                    lblMesin(X).BackColor = &HC000C0
                End If
            
            Rs.MoveNext
        Loop
    End If
    Set Rs = Nothing
End Sub

Private Sub CheckIdle()
    Dim Rs As New ADODB.Recordset
    Dim sSQL As String
    Dim i As Integer
    Dim X As Integer
    For X = 1 To 40
        lblIdle(X).BackColor = &H404040
    Next X

    Rs.CursorLocation = adUseClient
    sSQL = "SELECT Distinct a.plant_mark,a.prod_machine_id,b.NUMBER AS machine_no,a.period_shift,a.proses FROM sip_production.prod_machine_idles a"
    sSQL = sSQL & " LEFT JOIN sip_production.prod_machines b ON a.prod_machine_id = b.id"
    sSQL = sSQL & " WHERE a.period_shift = '" & Format(p_shift, "yyyy-mm-dd") & "' AND a.proses = 'Y'  Order by a.prod_machine_id ASC "

    Rs.Open sSQL, CN, adOpenStatic, adLockOptimistic

    If Rs.RecordCount > 0 Then
        Rs.MoveFirst
        Do While Not Rs.EOF
            i = Rs.Fields("machine_no")
                lblIdle(i).BackColor = &HFF&

            Rs.MoveNext
        Loop
    End If
    Set Rs = Nothing
End Sub


Private Sub lblMesin_Click(Index As Integer)
On Error GoTo ErrHandler

Dim qSQL As String

qSQL = "select a.plant_mark, a.prod_machine_id, b.number, b.name as machine_name, b.tonnage, a.mkt_customer_id, c.name as customer_name,"
qSQL = qSQL & " a.eng_product_id, d.internal_part_id, d.product_name, d.customer_part_number, d.customer_part_name, d.prod_yield,"
qSQL = qSQL & " d.material_name , d.color_name, d.cavity, d.weight_gr, d.weight_runner_gr, d.cycle_time_ia, a.period_shift,e.nik, e.name as employee,f.nik as nik_2, f.name as employee_2,"
qSQL = qSQL & " Round(((3600 / d.cycle_time_ia) * d.cavity),0) as target_shot"
qSQL = qSQL & " from sip_production.prod_runnings a"
qSQL = qSQL & " left join sip_production.prod_machines b on a.prod_machine_id = b.id"
qSQL = qSQL & " left join sip_production.mkt_customers c on a.mkt_customer_id = c.id"
qSQL = qSQL & " left join (select x.id, x.internal_part_id, x.name as product_name, x.customer_part_number, x.customer_part_name,"
qSQL = qSQL & " y.name as material_name, z.name as color_name,x.cavity, x.weight_gr, x.weight_runner_gr, x.cycle_time_ia, x.prod_yield"
qSQL = qSQL & " from sip_production.eng_products x left join sip_production.eng_materials y on x.eng_material_id = y.id"
qSQL = qSQL & " left join sip_production.eng_colors z on x.eng_color_id = z.id) d on a.eng_product_id = d.id"
qSQL = qSQL & " left join sip_production.hrd_employees e on a.operator_1 = e.id"
qSQL = qSQL & " left join sip_production.hrd_employees f on a.operator_2 = f.id"
qSQL = qSQL & " where a.plant_mark = '" & p_plant_mark & "'"
qSQL = qSQL & " and b.number = '" & Index & "'"
qSQL = qSQL & " and a.period_shift = '" & Format(p_shift, "yyyy-mm-dd") & "'"
qSQL = qSQL & " group by a.period_shift, a.plant_mark, a.prod_machine_id, a.mkt_customer_id, a.eng_product_id, a.period_shift, a.created_by"

Set RS_PRINT = New ADODB.Recordset
If RS_PRINT.State = adStateOpen Then RS_PRINT.Close
RS_PRINT.Open qSQL, CN, adOpenDynamic, adLockPessimistic

    With RptProduksi
        .DTRpt.Recordset = RS_PRINT
        If sPrint = 0 Then
            .Show 1
        End If
    
    End With
    
Exit Sub
ErrHandler:
    MsgBox "Error Number: " & Err.Number & vbCrLf & "Description: " & Err.Description, vbExclamation, Me.Caption

End Sub

Private Sub Timer1_Timer()
Static Xtimer As Integer
    If Minute(Now) = 0 Then
        CekTime
    End If
    
    Call CekData
    
    Call LoadIdle

'    Xtimer = Xtimer + 1
'    If Xtimer > 2 Then
'       Xtimer = 0
'        Call SetMachine
'        Call CheckIdle
'    End If
'
    Call SetMachine
    Call CheckIdle

ProgressBar1.Value = 0

End Sub

Private Sub CekTime()
Dim Jam As String
Dim i As Integer
Jam = Format(Now, "HH")
For i = 0 To 39
    Select Case Jam
        Case "08"
            datajam i, "08", "09"
        Case "09"
            datajam i, "08", "09"
        Case "10"
            datajam i, "09", "10"
        Case "11"
            datajam i, "10", "11"
        Case "12"
            datajam i, "11", "12"
        Case "13"
            datajam i, "12", "13"
        Case "14"
            datajam i, "13", "14"
        Case "15"
            datajam i, "14", "15"
        Case "16"
            datajam i, "15", "16"
        Case "17"
            datajam i, "16", "17"
        Case "18"
            datajam i, "17", "18"
        Case "19"
            datajam i, "18", "19"
        Case "20"
            datajam i, "19", "20"
        Case "21"
            datajam i, "20", "21"
        Case "22"
            datajam i, "21", "22"
        Case "23"
            datajam i, "22", "23"
        Case "00"
            datajam i, "23", "00"
        Case "01"
            datajam i, "00", "01"
        Case "02"
            datajam i, "01", "08"
        Case "03"
            datajam i, "02", "08"
        Case "04"
            datajam i, "03", "08"
        Case "05"
            datajam i, "04", "08"
        Case "06"
            datajam i, "05", "08"
        Case "07"
            datajam i, "06", "07"
    End Select

Next i
End Sub

Private Sub datajam(ix As Integer, jam1 As String, jam2 As String)
    lbljam1L(ix).Caption = jam1 & " : L"
    lbljam1R(ix).Caption = jam1 & " : R"
    lbljam2L(ix).Caption = jam2 & " : L"
    lbljam2R(ix).Caption = jam2 & " : R"
End Sub


Private Sub LoadData(sHour As String)
On Error GoTo ErrHandler

Dim Rs As New Recordset
Dim sSQL As String

Rs.CursorLocation = adUseClient

sSQL = "select a.plant_mark AS plant_mark,a.prod_machine_id AS prod_machine_id,b.number AS machine_no,"
sSQL = sSQL & " b.name AS machine_name,a.mkt_customer_id AS mkt_customer_id,"
sSQL = sSQL & " c.name AS customer_name,a.eng_product_1 AS eng_product_1,d.internal_part_id AS int_part_1,d.name AS prod_name_1,"
sSQL = sSQL & " d.cycle_time_ia AS cycle_time_ia_1,d.cavity AS cavity_1,d.weight_gr AS weight_gr_1,d.weight_runner_gr AS weight_runner_gr_1,"
sSQL = sSQL & " a.eng_product_2 AS eng_product_2,e.internal_part_id AS int_part_2,e.name AS prod_name_2,e.cavity AS cavity_2,a.description AS description,"
sSQL = sSQL & " a.status AS STATUS,a.machine_status AS machine_status,"
sSQL = sSQL & " round(3600/d.cycle_time_ia) as target_ct,"
sSQL = sSQL & " ifnull(f.counter_ok * d.cavity,'') as shot_1,"
sSQL = sSQL & " ifnull(g.counter_ok * e.cavity,'') as shot_2, "
sSQL = sSQL & " ifnull(h.ng_1,'') as ng_1,"
sSQL = sSQL & " ifnull(i.ng_2,'') as ng_2,"
sSQL = sSQL & " round((ifnull(f.counter_ok,0) /round(3600/d.cycle_time_ia)) * 100) as act_shot_1,"
sSQL = sSQL & " round((ifnull(g.counter_ok,0) /round(3600/d.cycle_time_ia)) * 100) as act_shot_2,"
sSQL = sSQL & " round((ifnull(h.ng_1,0) / ifnull(f.counter_ok * d.cavity, 0))*100) as act_ng_1,"
sSQL = sSQL & " round((ifnull(i.ng_2, 0) / ifnull(g.counter_ok * e.cavity, 0)) * 100) As act_ng_2"
sSQL = sSQL & " from ((((prod_running_products a"
sSQL = sSQL & " left join (select * from prod_runnings x where x.period_shift = '" & Format(p_shift, "yyyy-mm-dd") & "' and x.period_hour = '" & sHour & "') f"
                sSQL = sSQL & " on a.plant_mark = f.plant_mark and a.prod_machine_id = f.prod_machine_id and a.eng_product_1 = f.eng_product_id"
                
sSQL = sSQL & " left join (select * from prod_runnings x where x.period_shift = '" & Format(p_shift, "yyyy-mm-dd") & "' and x.period_hour = '" & sHour & "') g"
                sSQL = sSQL & " on a.plant_mark = g.plant_mark and a.prod_machine_id = g.prod_machine_id and a.eng_product_2 = g.eng_product_id"

sSQL = sSQL & " left join (select x.plant_mark,x.prod_machine_id,x.mkt_customer_id,x.eng_product_id,x.period_shift,x.period_hour,sum(x.counter_ng) as ng_1"
                sSQL = sSQL & " from prod_data_ngs x where x.period_shift = '" & Format(p_shift, "yyyy-mm-dd") & "' and x.period_hour = '" & sHour & "'"
                sSQL = sSQL & " group by x.plant_mark,x.prod_machine_id,x.mkt_customer_id,x.eng_product_id,x.period_shift,x.period_hour) h"
                sSQL = sSQL & " on a.plant_mark = h.plant_mark and a.prod_machine_id = h.prod_machine_id and a.eng_product_1 = h.eng_product_id"
                
sSQL = sSQL & " left join (select x.plant_mark,x.prod_machine_id,x.mkt_customer_id,x.eng_product_id,x.period_shift,x.period_hour,sum(x.counter_ng) as ng_2"
                sSQL = sSQL & " from prod_data_ngs x where x.period_shift = '" & Format(p_shift, "yyyy-mm-dd") & "' and x.period_hour = '" & sHour & "'"
                sSQL = sSQL & " group by x.plant_mark,x.prod_machine_id,x.mkt_customer_id,x.eng_product_id,x.period_shift,x.period_hour) i"
                sSQL = sSQL & " on a.plant_mark = i.plant_mark and a.prod_machine_id = i.prod_machine_id and a.eng_product_2 = i.eng_product_id"
sSQL = sSQL & " left join prod_machines b on(a.prod_machine_id = b.id))"
sSQL = sSQL & " left join mkt_customers c on(a.mkt_customer_id = c.id))"
sSQL = sSQL & " left join eng_products d on(a.eng_product_1 = d.id))"
sSQL = sSQL & " left join eng_products e on(a.eng_product_2 = e.id))"
sSQL = sSQL & " where a.status = 'active' and a.plant_mark = '" & p_plant_mark & "' order by b.number"
        
    Rs.Open sSQL, CN, adOpenStatic, adLockOptimistic

    If Rs.RecordCount > 0 Then
        Rs.MoveFirst
        Do While Not Rs.EOF
            Select Case Rs.Fields("machine_no")
                Case 1
                    txtShot(0).text = Rs.Fields("shot_1")
                    txtShot(1).text = Rs.Fields("shot_2")
                    txtng(0).text = Rs.Fields("ng_1")
                    txtng(1).text = Rs.Fields("ng_2")
                    
                    CheckColor_1 Rs, txtShot(0), txtShot(1), txtng(0), txtng(1)
                    
                Case 2
                    txtShot(7).text = Rs.Fields("shot_1")
                    txtShot(6).text = Rs.Fields("shot_2")
                    txtng(7).text = Rs.Fields("ng_1")
                    txtng(6).text = Rs.Fields("ng_2")
                    
                    CheckColor_1 Rs, txtShot(7), txtShot(6), txtng(7), txtng(6)
                    
                 Case 3
                    txtShot(8).text = Rs.Fields("shot_1")
                    txtShot(9).text = Rs.Fields("shot_2")
                    txtng(8).text = Rs.Fields("ng_1")
                    txtng(9).text = Rs.Fields("ng_2")
                    
                    CheckColor_1 Rs, txtShot(8), txtShot(9), txtng(8), txtng(9)
                    
                 Case 4
                    txtShot(12).text = Rs.Fields("shot_1")
                    txtShot(13).text = Rs.Fields("shot_2")
                    txtng(12).text = Rs.Fields("ng_1")
                    txtng(13).text = Rs.Fields("ng_2")
                    
                    CheckColor_1 Rs, txtShot(12), txtShot(13), txtng(12), txtng(13)
                                
               Case 5
                    txtShot(16).text = Rs.Fields("shot_1")
                    txtShot(17).text = Rs.Fields("shot_2")
                    txtng(16).text = Rs.Fields("ng_1")
                    txtng(17).text = Rs.Fields("ng_2")
                    
                    CheckColor_1 Rs, txtShot(16), txtShot(17), txtng(16), txtng(17)
                    
                Case 6
                    txtShot(23).text = Rs.Fields("shot_1")
                    txtShot(22).text = Rs.Fields("shot_2")
                    txtng(23).text = Rs.Fields("ng_1")
                    txtng(22).text = Rs.Fields("ng_2")
                    
                    CheckColor_1 Rs, txtShot(23), txtShot(22), txtng(23), txtng(22)
                    
                 Case 7
                    txtShot(24).text = Rs.Fields("shot_1")
                    txtShot(25).text = Rs.Fields("shot_2")
                    txtng(24).text = Rs.Fields("ng_1")
                    txtng(25).text = Rs.Fields("ng_2")
                    
                    CheckColor_1 Rs, txtShot(24), txtShot(25), txtng(24), txtng(25)
                    
                 Case 8
                    txtShot(31).text = Rs.Fields("shot_1")
                    txtShot(30).text = Rs.Fields("shot_2")
                    txtng(31).text = Rs.Fields("ng_1")
                    txtng(30).text = Rs.Fields("ng_2")
                    
                    CheckColor_1 Rs, txtShot(31), txtShot(30), txtng(31), txtng(30)
                    
                 Case 9
                    txtShot(35).text = Rs.Fields("shot_1")
                    txtShot(34).text = Rs.Fields("shot_2")
                    txtng(35).text = Rs.Fields("ng_1")
                    txtng(34).text = Rs.Fields("ng_2")
                    
                    CheckColor_1 Rs, txtShot(35), txtShot(34), txtng(35), txtng(34)
                    
                Case 10
                    txtShot(36).text = Rs.Fields("shot_1")
                    txtShot(37).text = Rs.Fields("shot_2")
                    txtng(36).text = Rs.Fields("ng_1")
                    txtng(37).text = Rs.Fields("ng_2")
                    
                    CheckColor_1 Rs, txtShot(36), txtShot(37), txtng(36), txtng(37)
                    
                Case 11
                    txtShot(40).text = Rs.Fields("shot_1")
                    txtShot(41).text = Rs.Fields("shot_2")
                    txtng(40).text = Rs.Fields("ng_1")
                    txtng(41).text = Rs.Fields("ng_2")
                    
                    CheckColor_1 Rs, txtShot(40), txtShot(41), txtng(40), txtng(41)
                    
                    
                Case 12
                    txtShot(47).text = Rs.Fields("shot_1")
                    txtShot(46).text = Rs.Fields("shot_2")
                    txtng(47).text = Rs.Fields("ng_1")
                    txtng(46).text = Rs.Fields("ng_2")
                    
                    CheckColor_1 Rs, txtShot(47), txtShot(46), txtng(46), txtng(47)
                    
                    
                Case 13
                    txtShot(48).text = Rs.Fields("shot_1")
                    txtShot(49).text = Rs.Fields("shot_2")
                    txtng(48).text = Rs.Fields("ng_1")
                    txtng(49).text = Rs.Fields("ng_2")
                    
                    CheckColor_1 Rs, txtShot(48), txtShot(49), txtng(48), txtng(49)
                    
                Case 14
                    txtShot(52).text = Rs.Fields("shot_1")
                    txtShot(53).text = Rs.Fields("shot_2")
                    txtng(52).text = Rs.Fields("ng_1")
                    txtng(53).text = Rs.Fields("ng_2")
                    
                    CheckColor_1 Rs, txtShot(52), txtShot(53), txtng(52), txtng(53)
                    
                Case 15
                    txtShot(59).text = Rs.Fields("shot_1")
                    txtShot(58).text = Rs.Fields("shot_2")
                    txtng(59).text = Rs.Fields("ng_1")
                    txtng(58).text = Rs.Fields("ng_2")
                    
                    CheckColor_1 Rs, txtShot(59), txtShot(58), txtng(59), txtng(58)
                    
                Case 16
                    txtShot(60).text = Rs.Fields("shot_1")
                    txtShot(61).text = Rs.Fields("shot_2")
                    txtng(60).text = Rs.Fields("ng_1")
                    txtng(61).text = Rs.Fields("ng_2")
                    
                    CheckColor_1 Rs, txtShot(60), txtShot(61), txtng(60), txtng(61)
                    
                Case 17
                    txtShot(67).text = Rs.Fields("shot_1")
                    txtShot(66).text = Rs.Fields("shot_2")
                    txtng(67).text = Rs.Fields("ng_1")
                    txtng(66).text = Rs.Fields("ng_2")
                    
                    CheckColor_1 Rs, txtShot(67), txtShot(66), txtng(67), txtng(66)
                    
                Case 18
                    txtShot(68).text = Rs.Fields("shot_1")
                    txtShot(69).text = Rs.Fields("shot_2")
                    txtng(68).text = Rs.Fields("ng_1")
                    txtng(69).text = Rs.Fields("ng_2")
                    
                    CheckColor_1 Rs, txtShot(68), txtShot(69), txtng(68), txtng(69)
                    
                Case 19
                    txtShot(72).text = Rs.Fields("shot_1")
                    txtShot(73).text = Rs.Fields("shot_2")
                    txtng(72).text = Rs.Fields("ng_1")
                    txtng(73).text = Rs.Fields("ng_2")
                    
                    CheckColor_1 Rs, txtShot(72), txtShot(73), txtng(72), txtng(73)
                    
                Case 20
                    txtShot(79).text = Rs.Fields("shot_1")
                    txtShot(78).text = Rs.Fields("shot_2")
                    txtng(79).text = Rs.Fields("ng_1")
                    txtng(78).text = Rs.Fields("ng_2")
                    
                    CheckColor_1 Rs, txtShot(79), txtShot(78), txtng(79), txtng(78)
                    
                Case 21
                    txtShot(80).text = Rs.Fields("shot_1")
                    txtShot(81).text = Rs.Fields("shot_2")
                    txtng(80).text = Rs.Fields("ng_1")
                    txtng(81).text = Rs.Fields("ng_2")
                    
                    CheckColor_1 Rs, txtShot(80), txtShot(81), txtng(80), txtng(81)
                    
                Case 22
                    txtShot(87).text = Rs.Fields("shot_1")
                    txtShot(86).text = Rs.Fields("shot_2")
                    txtng(87).text = Rs.Fields("ng_1")
                    txtng(86).text = Rs.Fields("ng_2")
                    
                    CheckColor_1 Rs, txtShot(87), txtShot(86), txtng(87), txtng(86)
                    
                Case 23
                    txtShot(88).text = Rs.Fields("shot_1")
                    txtShot(89).text = Rs.Fields("shot_2")
                    txtng(88).text = Rs.Fields("ng_1")
                    txtng(89).text = Rs.Fields("ng_2")
                    
                    CheckColor_1 Rs, txtShot(88), txtShot(89), txtng(88), txtng(89)
                    
                Case 24
                    txtShot(92).text = Rs.Fields("shot_1")
                    txtShot(93).text = Rs.Fields("shot_2")
                    txtng(92).text = Rs.Fields("ng_1")
                    txtng(93).text = Rs.Fields("ng_2")
                    
                    CheckColor_1 Rs, txtShot(92), txtShot(93), txtng(92), txtng(93)
                    
                Case 25
                    txtShot(99).text = Rs.Fields("shot_1")
                    txtShot(98).text = Rs.Fields("shot_2")
                    txtng(99).text = Rs.Fields("ng_1")
                    txtng(98).text = Rs.Fields("ng_2")
                    
                    CheckColor_1 Rs, txtShot(99), txtShot(98), txtng(99), txtng(98)
                    
                Case 26
                    txtShot(100).text = Rs.Fields("shot_1")
                    txtShot(101).text = Rs.Fields("shot_2")
                    txtng(100).text = Rs.Fields("ng_1")
                    txtng(101).text = Rs.Fields("ng_2")
                    
                    CheckColor_1 Rs, txtShot(100), txtShot(101), txtng(100), txtng(101)
                    
                Case 27
                    txtShot(107).text = Rs.Fields("shot_1")
                    txtShot(106).text = Rs.Fields("shot_2")
                    txtng(107).text = Rs.Fields("ng_1")
                    txtng(106).text = Rs.Fields("ng_2")
                    
                    CheckColor_1 Rs, txtShot(107), txtShot(106), txtng(107), txtng(106)
                    
                Case 28
                    txtShot(108).text = Rs.Fields("shot_1")
                    txtShot(109).text = Rs.Fields("shot_2")
                    txtng(108).text = Rs.Fields("ng_1")
                    txtng(109).text = Rs.Fields("ng_2")
                    
                    CheckColor_1 Rs, txtShot(108), txtShot(109), txtng(108), txtng(109)
                    
                Case 29
                    txtShot(112).text = Rs.Fields("shot_1")
                    txtShot(113).text = Rs.Fields("shot_2")
                    txtng(112).text = Rs.Fields("ng_1")
                    txtng(113).text = Rs.Fields("ng_2")
                    
                    CheckColor_1 Rs, txtShot(112), txtShot(113), txtng(112), txtng(113)
                    
                Case 30
                    txtShot(116).text = Rs.Fields("shot_1")
                    txtShot(117).text = Rs.Fields("shot_2")
                    txtng(116).text = Rs.Fields("ng_1")
                    txtng(117).text = Rs.Fields("ng_2")
                    
                    CheckColor_1 Rs, txtShot(116), txtShot(117), txtng(116), txtng(117)
                    
                Case 31
                    txtShot(120).text = Rs.Fields("shot_1")
                    txtShot(121).text = Rs.Fields("shot_2")
                    txtng(120).text = Rs.Fields("ng_1")
                    txtng(121).text = Rs.Fields("ng_2")
                    
                    CheckColor_1 Rs, txtShot(120), txtShot(121), txtng(120), txtng(121)
                    
                Case 32
                    txtShot(127).text = Rs.Fields("shot_1")
                    txtShot(126).text = Rs.Fields("shot_2")
                    txtng(127).text = Rs.Fields("ng_1")
                    txtng(126).text = Rs.Fields("ng_2")
                    
                    CheckColor_1 Rs, txtShot(127), txtShot(126), txtng(127), txtng(126)
                    
                Case 33
                    txtShot(128).text = Rs.Fields("shot_1")
                    txtShot(129).text = Rs.Fields("shot_2")
                    txtng(128).text = Rs.Fields("ng_1")
                    txtng(129).text = Rs.Fields("ng_2")
                    
                    CheckColor_1 Rs, txtShot(128), txtShot(129), txtng(128), txtng(129)

                Case 34
                    txtShot(132).text = Rs.Fields("shot_1")
                    txtShot(133).text = Rs.Fields("shot_2")
                    txtng(132).text = Rs.Fields("ng_1")
                    txtng(133).text = Rs.Fields("ng_2")
                    
                    CheckColor_1 Rs, txtShot(132), txtShot(133), txtng(132), txtng(133)
                    
                Case 35
                    txtShot(139).text = Rs.Fields("shot_1")
                    txtShot(138).text = Rs.Fields("shot_2")
                    txtng(139).text = Rs.Fields("ng_1")
                    txtng(138).text = Rs.Fields("ng_2")
                    
                    CheckColor_1 Rs, txtShot(139), txtShot(138), txtng(139), txtng(138)
                    
                Case 36
                    txtShot(140).text = Rs.Fields("shot_1")
                    txtShot(141).text = Rs.Fields("shot_2")
                    txtng(140).text = Rs.Fields("ng_1")
                    txtng(141).text = Rs.Fields("ng_2")
                    
                    CheckColor_1 Rs, txtShot(140), txtShot(141), txtng(140), txtng(141)
                    
                Case 37
                    txtShot(147).text = Rs.Fields("shot_1")
                    txtShot(146).text = Rs.Fields("shot_2")
                    txtng(147).text = Rs.Fields("ng_1")
                    txtng(146).text = Rs.Fields("ng_2")
                    
                    CheckColor_1 Rs, txtShot(147), txtShot(146), txtng(147), txtng(146)
                    
                Case 38
                    txtShot(148).text = Rs.Fields("shot_1")
                    txtShot(149).text = Rs.Fields("shot_2")
                    txtng(148).text = Rs.Fields("ng_1")
                    txtng(149).text = Rs.Fields("ng_2")
                    
                    CheckColor_1 Rs, txtShot(148), txtShot(149), txtng(148), txtng(149)
                    
                Case 39
                    txtShot(152).text = Rs.Fields("shot_1")
                    txtShot(153).text = Rs.Fields("shot_2")
                    txtng(152).text = Rs.Fields("ng_1")
                    txtng(153).text = Rs.Fields("ng_2")
                    
                    CheckColor_1 Rs, txtShot(152), txtShot(153), txtng(152), txtng(153)
                    
                Case 40
                    txtShot(159).text = Rs.Fields("shot_1")
                    txtShot(158).text = Rs.Fields("shot_2")
                    txtng(159).text = Rs.Fields("ng_1")
                    txtng(158).text = Rs.Fields("ng_2")
                    
                    CheckColor_1 Rs, txtShot(159), txtShot(158), txtng(159), txtng(158)
                    
                Case 41
                    txtShot(160).text = Rs.Fields("shot_1")
                    txtShot(161).text = Rs.Fields("shot_2")
                    txtng(160).text = Rs.Fields("ng_1")
                    txtng(161).text = Rs.Fields("ng_2")
                    
                    CheckColor_1 Rs, txtShot(160), txtShot(161), txtng(160), txtng(161)

                Case 42
                    txtShot(164).text = Rs.Fields("shot_1")
                    txtShot(165).text = Rs.Fields("shot_2")
                    txtng(164).text = Rs.Fields("ng_1")
                    txtng(165).text = Rs.Fields("ng_2")
                    
                    CheckColor_1 Rs, txtShot(164), txtShot(165), txtng(164), txtng(165)
                    

            End Select

            Rs.MoveNext
        Loop
    End If

    Set Rs = Nothing

Exit Sub
ErrHandler:
    MsgBox "Error Number: " & Err.Number & vbCrLf & "Description: " & Err.Description, vbExclamation
    
    
End Sub



Private Sub CekData()
Dim Jam As String
Jam = Format(Now, "HH")
Select Case Jam
    Case "08"
        LoadData "08"
        LoadData_2 "09"
        
    Case "09"
        LoadData "08"
        LoadData_2 "09"
        
    Case "10"
        LoadData "09"
        LoadData_2 "10"

    Case "11"
        LoadData "10"
        LoadData_2 "11"

    Case "12"
        LoadData "11"
        LoadData_2 "12"

    Case "13"
        LoadData "12"
        LoadData_2 "13"

    Case "14"
        LoadData "13"
        LoadData_2 "14"

    Case "15"
        LoadData "14"
        LoadData_2 "15"

    Case "16"
        LoadData "15"
        LoadData_2 "16"

    Case "17"
        LoadData "16"
        LoadData_2 "17"

    Case "18"
        LoadData "17"
        LoadData_2 "18"

    Case "19"
        LoadData "18"
        LoadData_2 "19"

    Case "20"
        LoadData "19"
        LoadData_2 "20"

    Case "21"
        LoadData "20"
        LoadData_2 "21"

    Case "22"
        LoadData "21"
        LoadData_2 "22"

    Case "23"
        LoadData "22"
        LoadData_2 "23"

    Case "00"
        LoadData "23"
        LoadData_2 "00"

    Case "01"
        LoadData "00"
        LoadData_2 "01"

    Case "02"
        LoadData "01"
        LoadData_2 "02"

    Case "03"
        LoadData "02"
        LoadData_2 "03"

    Case "04"
        LoadData "03"
        LoadData_2 "04"

    Case "05"
        LoadData "04"
        LoadData_2 "05"

    Case "06"
        LoadData "05"
        LoadData_2 "06"

    Case "07"
        LoadData "06"
        LoadData_2 "07"

End Select

End Sub



Private Sub LoadData_2(sHour As String)
On Error GoTo ErrHandler

Dim Rs As New Recordset
Dim sSQL As String

Rs.CursorLocation = adUseClient

sSQL = "select a.plant_mark AS plant_mark,a.prod_machine_id AS prod_machine_id,b.number AS machine_no,"
sSQL = sSQL & " b.name AS machine_name,a.mkt_customer_id AS mkt_customer_id,"
sSQL = sSQL & " c.name AS customer_name,a.eng_product_1 AS eng_product_1,d.internal_part_id AS int_part_1,d.name AS prod_name_1,"
sSQL = sSQL & " d.cycle_time_ia AS cycle_time_ia_1,d.cavity AS cavity_1,d.weight_gr AS weight_gr_1,d.weight_runner_gr AS weight_runner_gr_1,"
sSQL = sSQL & " a.eng_product_2 AS eng_product_2,e.internal_part_id AS int_part_2,e.name AS prod_name_2,e.cavity AS cavity_2, a.description AS description,"
sSQL = sSQL & " a.status AS STATUS,a.machine_status AS machine_status,"
sSQL = sSQL & " round(3600/d.cycle_time_ia) as target_ct,"
sSQL = sSQL & " ifnull(f.counter_ok * d.cavity,'') as shot_1 ,"
sSQL = sSQL & " ifnull(g.counter_ok * e.cavity,'') as shot_2, "
sSQL = sSQL & " ifnull(h.ng_1,'') as ng_1,"
sSQL = sSQL & " ifnull(i.ng_2,'') as ng_2,"
sSQL = sSQL & " round((ifnull(f.counter_ok,0) /round(3600/d.cycle_time_ia)) * 100) as act_shot_1,"
sSQL = sSQL & " round((ifnull(g.counter_ok,0) /round(3600/d.cycle_time_ia)) * 100) as act_shot_2,"
sSQL = sSQL & " round((ifnull(h.ng_1,0)/ ifnull(f.counter_ok * d.cavity,0))*100) as act_ng_1,"
sSQL = sSQL & " round((ifnull(i.ng_2,0)/ ifnull(g.counter_ok * e.cavity,0))*100) As act_ng_2"
sSQL = sSQL & " from ((((prod_running_products a"
sSQL = sSQL & " left join (select * from prod_runnings x where x.period_shift = '" & Format(p_shift, "yyyy-mm-dd") & "' and x.period_hour = '" & sHour & "') f"
                sSQL = sSQL & " on a.plant_mark = f.plant_mark and a.prod_machine_id = f.prod_machine_id and a.eng_product_1 = f.eng_product_id"
                
sSQL = sSQL & " left join (select * from prod_runnings x where x.period_shift = '" & Format(p_shift, "yyyy-mm-dd") & "' and x.period_hour = '" & sHour & "') g"
                sSQL = sSQL & " on a.plant_mark = g.plant_mark and a.prod_machine_id = g.prod_machine_id and a.eng_product_2 = g.eng_product_id"

sSQL = sSQL & " left join (select x.plant_mark,x.prod_machine_id,x.mkt_customer_id,x.eng_product_id,x.period_shift,x.period_hour,sum(x.counter_ng) as ng_1"
                sSQL = sSQL & " from prod_data_ngs x where x.period_shift = '" & Format(p_shift, "yyyy-mm-dd") & "' and x.period_hour = '" & sHour & "'"
                sSQL = sSQL & " group by x.plant_mark,x.prod_machine_id,x.mkt_customer_id,x.eng_product_id,x.period_shift,x.period_hour) h"
                sSQL = sSQL & " on a.plant_mark = h.plant_mark and a.prod_machine_id = h.prod_machine_id and a.eng_product_1 = h.eng_product_id"
                
sSQL = sSQL & " left join (select x.plant_mark,x.prod_machine_id,x.mkt_customer_id,x.eng_product_id,x.period_shift,x.period_hour,sum(x.counter_ng) as ng_2"
                sSQL = sSQL & " from prod_data_ngs x where x.period_shift = '" & Format(p_shift, "yyyy-mm-dd") & "' and x.period_hour = '" & sHour & "'"
                sSQL = sSQL & " group by x.plant_mark,x.prod_machine_id,x.mkt_customer_id,x.eng_product_id,x.period_shift,x.period_hour) i"
                sSQL = sSQL & " on a.plant_mark = i.plant_mark and a.prod_machine_id = i.prod_machine_id and a.eng_product_2 = i.eng_product_id"
sSQL = sSQL & " left join prod_machines b on(a.prod_machine_id = b.id))"
sSQL = sSQL & " left join mkt_customers c on(a.mkt_customer_id = c.id))"
sSQL = sSQL & " left join eng_products d on(a.eng_product_1 = d.id))"
sSQL = sSQL & " left join eng_products e on(a.eng_product_2 = e.id))"
sSQL = sSQL & " where a.status = 'active' and a.plant_mark = '" & p_plant_mark & "' order by b.number"


Rs.Open sSQL, CN, adOpenStatic, adLockOptimistic

    If Rs.RecordCount > 0 Then
        Rs.MoveFirst
        Do While Not Rs.EOF
            Select Case Rs.Fields("machine_no")
                Case 1
                    txtShot(2).text = Rs.Fields("shot_1")
                    txtShot(3).text = Rs.Fields("shot_2")
                    txtng(2).text = Rs.Fields("ng_1")
                    txtng(3).text = Rs.Fields("ng_2")
                                    
                    CheckColor_2 Rs, txtng(2), txtng(3)
                    
        
                Case 2
                    txtShot(5).text = Rs.Fields("shot_1")
                    txtShot(4).text = Rs.Fields("shot_2")
                    txtng(5).text = Rs.Fields("ng_1")
                    txtng(4).text = Rs.Fields("ng_2")
                    
                    CheckColor_2 Rs, txtng(5), txtng(4)
                    
                 Case 3
                    txtShot(10).text = Rs.Fields("shot_1")
                    txtShot(11).text = Rs.Fields("shot_2")
                    txtng(10).text = Rs.Fields("ng_1")
                    txtng(11).text = Rs.Fields("ng_2")
                    
                    CheckColor_2 Rs, txtng(10), txtng(11)
                    
                 Case 4
                    txtShot(14).text = Rs.Fields("shot_1")
                    txtShot(15).text = Rs.Fields("shot_2")
                    txtng(14).text = Rs.Fields("ng_1")
                    txtng(15).text = Rs.Fields("ng_2")
                    
                    CheckColor_2 Rs, txtng(14), txtng(15)
                                
               Case 5
                    txtShot(18).text = Rs.Fields("shot_1")
                    txtShot(19).text = Rs.Fields("shot_2")
                    txtng(18).text = Rs.Fields("ng_1")
                    txtng(19).text = Rs.Fields("ng_2")
                    
                    CheckColor_2 Rs, txtng(18), txtng(19)
                    
                Case 6
                    txtShot(21).text = Rs.Fields("shot_1")
                    txtShot(20).text = Rs.Fields("shot_2")
                    txtng(21).text = Rs.Fields("ng_1")
                    txtng(20).text = Rs.Fields("ng_2")
                    
                     CheckColor_2 Rs, txtng(21), txtng(20)
                     
                 Case 7
                    txtShot(26).text = Rs.Fields("shot_1")
                    txtShot(27).text = Rs.Fields("shot_2")
                    txtng(26).text = Rs.Fields("ng_1")
                    txtng(27).text = Rs.Fields("ng_2")
                    
                     CheckColor_2 Rs, txtng(26), txtng(27)
                     
                 Case 8
                    txtShot(29).text = Rs.Fields("shot_1")
                    txtShot(28).text = Rs.Fields("shot_2")
                    txtng(29).text = Rs.Fields("ng_1")
                    txtng(28).text = Rs.Fields("ng_2")
                    
                     CheckColor_2 Rs, txtng(29), txtng(28)
                     
                 Case 9
                    txtShot(33).text = Rs.Fields("shot_1")
                    txtShot(32).text = Rs.Fields("shot_2")
                    txtng(33).text = Rs.Fields("ng_1")
                    txtng(32).text = Rs.Fields("ng_2")
                    
                     CheckColor_2 Rs, txtng(33), txtng(32)
                     
                Case 10
                    txtShot(38).text = Rs.Fields("shot_1")
                    txtShot(39).text = Rs.Fields("shot_2")
                    txtng(38).text = Rs.Fields("ng_1")
                    txtng(39).text = Rs.Fields("ng_2")
                    
                     CheckColor_2 Rs, txtng(38), txtng(39)
                     
                Case 11
                    txtShot(42).text = Rs.Fields("shot_1")
                    txtShot(43).text = Rs.Fields("shot_2")
                    txtng(42).text = Rs.Fields("ng_1")
                    txtng(43).text = Rs.Fields("ng_2")
                    
                     CheckColor_2 Rs, txtng(42), txtng(43)
                     
                Case 12
                    txtShot(45).text = Rs.Fields("shot_1")
                    txtShot(44).text = Rs.Fields("shot_2")
                    txtng(45).text = Rs.Fields("ng_1")
                    txtng(44).text = Rs.Fields("ng_2")
                    
                     CheckColor_2 Rs, txtng(45), txtng(44)
                     
                Case 13
                    txtShot(50).text = Rs.Fields("shot_1")
                    txtShot(51).text = Rs.Fields("shot_2")
                    txtng(50).text = Rs.Fields("ng_1")
                    txtng(51).text = Rs.Fields("ng_2")
                    
                     CheckColor_2 Rs, txtng(50), txtng(51)
                     
                Case 14
                    txtShot(54).text = Rs.Fields("shot_1")
                    txtShot(55).text = Rs.Fields("shot_2")
                    txtng(54).text = Rs.Fields("ng_1")
                    txtng(55).text = Rs.Fields("ng_2")
                    
                     CheckColor_2 Rs, txtng(54), txtng(55)
                     
                     
                Case 15
                    txtShot(57).text = Rs.Fields("shot_1")
                    txtShot(56).text = Rs.Fields("shot_2")
                    txtng(57).text = Rs.Fields("ng_1")
                    txtng(56).text = Rs.Fields("ng_2")
                    
                     CheckColor_2 Rs, txtng(57), txtng(56)
                     
                     
                Case 16
                    txtShot(62).text = Rs.Fields("shot_1")
                    txtShot(63).text = Rs.Fields("shot_2")
                    txtng(62).text = Rs.Fields("ng_1")
                    txtng(63).text = Rs.Fields("ng_2")
                    
                     CheckColor_2 Rs, txtng(62), txtng(63)
                     
                Case 17
                    txtShot(65).text = Rs.Fields("shot_1")
                    txtShot(64).text = Rs.Fields("shot_2")
                    txtng(65).text = Rs.Fields("ng_1")
                    txtng(64).text = Rs.Fields("ng_2")
                    
                     CheckColor_2 Rs, txtng(65), txtng(64)
                     
                Case 18
                    txtShot(70).text = Rs.Fields("shot_1")
                    txtShot(71).text = Rs.Fields("shot_2")
                    txtng(70).text = Rs.Fields("ng_1")
                    txtng(71).text = Rs.Fields("ng_2")
                    
                     CheckColor_2 Rs, txtng(70), txtng(71)
                     
                Case 19
                    txtShot(74).text = Rs.Fields("shot_1")
                    txtShot(75).text = Rs.Fields("shot_2")
                    txtng(74).text = Rs.Fields("ng_1")
                    txtng(75).text = Rs.Fields("ng_2")
                    
                     CheckColor_2 Rs, txtng(74), txtng(75)
                     
                Case 20
                    txtShot(77).text = Rs.Fields("shot_1")
                    txtShot(76).text = Rs.Fields("shot_2")
                    txtng(77).text = Rs.Fields("ng_1")
                    txtng(76).text = Rs.Fields("ng_2")
                    
                     CheckColor_2 Rs, txtng(77), txtng(76)
                     
                Case 21
                    txtShot(82).text = Rs.Fields("shot_1")
                    txtShot(83).text = Rs.Fields("shot_2")
                    txtng(82).text = Rs.Fields("ng_1")
                    txtng(83).text = Rs.Fields("ng_2")
                    
                     CheckColor_2 Rs, txtng(82), txtng(83)
                     
                Case 22
                    txtShot(85).text = Rs.Fields("shot_1")
                    txtShot(84).text = Rs.Fields("shot_2")
                    txtng(85).text = Rs.Fields("ng_1")
                    txtng(84).text = Rs.Fields("ng_2")
                    
                     CheckColor_2 Rs, txtng(85), txtng(84)
                     
                Case 23
                    txtShot(90).text = Rs.Fields("shot_1")
                    txtShot(91).text = Rs.Fields("shot_2")
                    txtng(90).text = Rs.Fields("ng_1")
                    txtng(91).text = Rs.Fields("ng_2")
                    
                     CheckColor_2 Rs, txtng(90), txtng(91)
                     
                Case 24
                    txtShot(94).text = Rs.Fields("shot_1")
                    txtShot(95).text = Rs.Fields("shot_2")
                    txtng(94).text = Rs.Fields("ng_1")
                    txtng(95).text = Rs.Fields("ng_2")
                    
                     CheckColor_2 Rs, txtng(94), txtng(95)
                     
                Case 25
                    txtShot(97).text = Rs.Fields("shot_1")
                    txtShot(96).text = Rs.Fields("shot_2")
                    txtng(97).text = Rs.Fields("ng_1")
                    txtng(96).text = Rs.Fields("ng_2")
                    
                     CheckColor_2 Rs, txtng(97), txtng(96)
                     
                Case 26
                    txtShot(102).text = Rs.Fields("shot_1")
                    txtShot(103).text = Rs.Fields("shot_2")
                    txtng(102).text = Rs.Fields("ng_1")
                    txtng(103).text = Rs.Fields("ng_2")
                    
                     CheckColor_2 Rs, txtng(102), txtng(103)
                     
                Case 27
                    txtShot(105).text = Rs.Fields("shot_1")
                    txtShot(104).text = Rs.Fields("shot_2")
                    txtng(105).text = Rs.Fields("ng_1")
                    txtng(104).text = Rs.Fields("ng_2")
                    
                     CheckColor_2 Rs, txtng(105), txtng(104)
                     
                Case 28
                    txtShot(110).text = Rs.Fields("shot_1")
                    txtShot(111).text = Rs.Fields("shot_2")
                    txtng(110).text = Rs.Fields("ng_1")
                    txtng(111).text = Rs.Fields("ng_2")
                    
                     CheckColor_2 Rs, txtng(110), txtng(111)
                     
                Case 29
                    txtShot(114).text = Rs.Fields("shot_1")
                    txtShot(115).text = Rs.Fields("shot_2")
                    txtng(114).text = Rs.Fields("ng_1")
                    txtng(115).text = Rs.Fields("ng_2")
                    
                     CheckColor_2 Rs, txtng(114), txtng(115)
                     
                Case 30
                    txtShot(118).text = Rs.Fields("shot_1")
                    txtShot(119).text = Rs.Fields("shot_2")
                    txtng(118).text = Rs.Fields("ng_1")
                    txtng(119).text = Rs.Fields("ng_2")
                    
                     CheckColor_2 Rs, txtng(118), txtng(119)
                     
                Case 31
                    txtShot(122).text = Rs.Fields("shot_1")
                    txtShot(123).text = Rs.Fields("shot_2")
                    txtng(122).text = Rs.Fields("ng_1")
                    txtng(122).text = Rs.Fields("ng_2")
                    
                     CheckColor_2 Rs, txtng(122), txtng(122)
                     
                Case 32
                    txtShot(125).text = Rs.Fields("shot_1")
                    txtShot(124).text = Rs.Fields("shot_2")
                    txtng(125).text = Rs.Fields("ng_1")
                    txtng(124).text = Rs.Fields("ng_2")
                    
                     CheckColor_2 Rs, txtng(125), txtng(124)
                     
                Case 33
                    txtShot(130).text = Rs.Fields("shot_1")
                    txtShot(131).text = Rs.Fields("shot_2")
                    txtng(130).text = Rs.Fields("ng_1")
                    txtng(131).text = Rs.Fields("ng_2")
                    
                     CheckColor_2 Rs, txtng(130), txtng(131)
                     

                Case 34
                    txtShot(134).text = Rs.Fields("shot_1")
                    txtShot(135).text = Rs.Fields("shot_2")
                    txtng(134).text = Rs.Fields("ng_1")
                    txtng(135).text = Rs.Fields("ng_2")
                    
                     CheckColor_2 Rs, txtng(134), txtng(135)
                     
                Case 35
                    txtShot(139).text = Rs.Fields("shot_1")
                    txtShot(138).text = Rs.Fields("shot_2")
                    txtng(139).text = Rs.Fields("ng_1")
                    txtng(138).text = Rs.Fields("ng_2")
                    
                     CheckColor_2 Rs, txtng(139), txtng(138)
                     
                Case 36
                    txtShot(142).text = Rs.Fields("shot_1")
                    txtShot(143).text = Rs.Fields("shot_2")
                    txtng(142).text = Rs.Fields("ng_1")
                    txtng(143).text = Rs.Fields("ng_2")
                    
                     CheckColor_2 Rs, txtng(142), txtng(143)
                     
                Case 37
                    txtShot(147).text = Rs.Fields("shot_1")
                    txtShot(146).text = Rs.Fields("shot_2")
                    txtng(147).text = Rs.Fields("ng_1")
                    txtng(146).text = Rs.Fields("ng_2")
                    
                     CheckColor_2 Rs, txtng(147), txtng(146)
                     
                Case 38
                    txtShot(150).text = Rs.Fields("shot_1")
                    txtShot(151).text = Rs.Fields("shot_2")
                    txtng(150).text = Rs.Fields("ng_1")
                    txtng(151).text = Rs.Fields("ng_2")
                    
                     CheckColor_2 Rs, txtng(150), txtng(151)
                     
                Case 39
                    txtShot(154).text = Rs.Fields("shot_1")
                    txtShot(155).text = Rs.Fields("shot_2")
                    txtng(154).text = Rs.Fields("ng_1")
                    txtng(155).text = Rs.Fields("ng_2")
                    
                     CheckColor_2 Rs, txtng(154), txtng(155)
                     
                Case 40
                    txtShot(157).text = Rs.Fields("shot_1")
                    txtShot(156).text = Rs.Fields("shot_2")
                    txtng(157).text = Rs.Fields("ng_1")
                    txtng(156).text = Rs.Fields("ng_2")
                    
                     CheckColor_2 Rs, txtng(157), txtng(156)
                     

                Case 41
                    txtShot(162).text = Rs.Fields("shot_1")
                    txtShot(163).text = Rs.Fields("shot_2")
                    txtng(162).text = Rs.Fields("ng_1")
                    txtng(163).text = Rs.Fields("ng_2")
                    
                     CheckColor_2 Rs, txtng(162), txtng(163)


                Case 42
                    txtShot(166).text = Rs.Fields("shot_1")
                    txtShot(167).text = Rs.Fields("shot_2")
                    txtng(166).text = Rs.Fields("ng_1")
                    txtng(167).text = Rs.Fields("ng_2")
                    
                     CheckColor_2 Rs, txtng(166), txtng(167)
                     
                     
            End Select

            Rs.MoveNext
        Loop
    End If

    Set Rs = Nothing

Exit Sub
ErrHandler:
    MsgBox "Error Number: " & Err.Number & vbCrLf & "Description: " & Err.Description, vbExclamation
    
    
End Sub

Private Sub CheckColor_1(ByVal Rs As Recordset, ByRef txtshot1 As TextBox, ByRef txtshot2 As TextBox, ByRef txtng1 As TextBox, ByRef txtng2 As TextBox)
    If Rs.Fields("act_shot_1") >= 98 Then
        txtshot1.BackColor = &HFF00&
    ElseIf Rs.Fields("act_shot_1") >= 1 And Rs.Fields("act_shot_1") < 98 Then
        txtshot1.BackColor = &HFF&
    Else
        txtshot1.BackColor = &HFFFFFF
    End If

    If Rs.Fields("act_shot_2") >= 98 Then
        txtshot2.BackColor = &HFF00&
    ElseIf Rs.Fields("act_shot_2") >= 1 And Rs.Fields("act_shot_2") < 98 Then
        txtshot2.BackColor = &HFF&
    Else
        txtshot2.BackColor = &HFFFFFF
    End If
    
    If Rs.Fields("act_ng_1") >= 5 Then
        txtng1.BackColor = &HFF&
    ElseIf Rs.Fields("act_ng_1") >= 1 And Rs.Fields("act_ng_1") < 5 Then
        txtng1.BackColor = &HFF00&
    Else
        txtng1.BackColor = &HFFFFFF
    End If
    
    If Rs.Fields("act_ng_2") >= 5 Then
        txtng2.BackColor = &HFF&
    ElseIf Rs.Fields("act_ng_2") >= 1 And Rs.Fields("act_ng_2") < 5 Then
        txtng2.BackColor = &HFF00&
    Else
        txtng2.BackColor = &HFFFFFF
    End If
                              
End Sub

Private Sub CheckColor_2(ByVal Rs As Recordset, ByRef txtng1 As TextBox, ByRef txtng2 As TextBox)

    If Rs.Fields("act_ng_1") >= 5 Then
        txtng1.BackColor = &HFF&
    ElseIf Rs.Fields("act_ng_1") >= 1 And Rs.Fields("act_ng_1") < 5 Then
        txtng1.BackColor = &HFF00&
    Else
        txtng1.BackColor = &HFFFFFF
    End If
    
    If Rs.Fields("act_ng_2") >= 5 Then
        txtng2.BackColor = &HFF&
    ElseIf Rs.Fields("act_ng_2") >= 1 And Rs.Fields("act_ng_2") < 5 Then
        txtng2.BackColor = &HFF00&
    Else
        txtng2.BackColor = &HFFFFFF
    End If
                              
End Sub

Private Sub LoadIdle()
On Error GoTo ErrHandler
Dim Rs As New Recordset
Dim sSQL As String

Rs.CursorLocation = adUseClient

sSQL = "select A.plant_mark,A.prod_machine_id,B.number as machine_no,A.mkt_customer_id,X.idle_shift1,idle_shift2,idle_shift3 from prod_running_products A"
sSQL = sSQL & " left join prod_machines B on A.prod_machine_id = B.id"
sSQL = sSQL & " Left Join"
            sSQL = sSQL & " (select a.plant_mark,a.prod_machine_id,a.mkt_customer_id,a.period_shift,"
            sSQL = sSQL & " sec_to_time(sum(time_to_sec(a.idle_time))) AS idle_shift1,sum(time_to_sec(a.idle_time)) AS jumlah_idle_sec"
            sSQL = sSQL & " from prod_machine_idles a"
            sSQL = sSQL & " where a.period_shift = '" & Format(p_shift, "yyyy-mm-dd") & "' and time_format(a.start_idle,'%H') between '08' and '15'"
            sSQL = sSQL & " group by a.plant_mark,a.prod_machine_id,a.period_shift) X"
            sSQL = sSQL & " on A.plant_mark = X.plant_mark and A.prod_machine_id = X.prod_machine_id"
sSQL = sSQL & " Left Join"
            sSQL = sSQL & " (select a.plant_mark,a.prod_machine_id,a.mkt_customer_id,a.period_shift,"
            sSQL = sSQL & " sec_to_time(sum(time_to_sec(a.idle_time))) AS idle_shift2,sum(time_to_sec(a.idle_time)) AS jumlah_idle_sec"
            sSQL = sSQL & " from prod_machine_idles a"
            sSQL = sSQL & " where a.period_shift = '" & Format(p_shift, "yyyy-mm-dd") & "' and time_format(a.start_idle,'%H') between '16' and '23'"
            sSQL = sSQL & " group by a.plant_mark,a.prod_machine_id,a.period_shift) Y"
            sSQL = sSQL & " on A.plant_mark = Y.plant_mark and A.prod_machine_id = Y.prod_machine_id"
sSQL = sSQL & " Left Join"
            sSQL = sSQL & " (select a.plant_mark,a.prod_machine_id,a.mkt_customer_id,a.period_shift,"
            sSQL = sSQL & " sec_to_time(sum(time_to_sec(a.idle_time))) AS idle_shift3,sum(time_to_sec(a.idle_time)) AS jumlah_idle_sec"
            sSQL = sSQL & " from prod_machine_idles a"
            sSQL = sSQL & " where a.period_shift = '" & Format(p_shift, "yyyy-mm-dd") & "' and time_format(a.start_idle,'%H') between '00' and '07'"
            sSQL = sSQL & " group by a.plant_mark,a.prod_machine_id,a.period_shift) Z"
            sSQL = sSQL & " on A.plant_mark = Z.plant_mark and A.prod_machine_id = Z.prod_machine_id"
sSQL = sSQL & " where A.`status` = 'active' and A.plant_mark = '" & p_plant_mark & "'"
sSQL = sSQL & " Order by B.number ASC"

    Rs.Open sSQL, CN, adOpenStatic, adLockOptimistic
    
    If Rs.RecordCount > 0 Then

        Rs.MoveFirst
        Do While Not Rs.EOF
            Select Case Rs.Fields("machine_no")
                Case 1
                    txtidle(0).text = Format(IIf(IsNull(Rs.Fields("idle_shift1")), "", Rs.Fields("idle_shift1")), "hh:mm")
                    txtidle(1).text = Format(IIf(IsNull(Rs.Fields("idle_shift2")), "", Rs.Fields("idle_shift2")), "hh:mm")
                    txtidle(2).text = Format(IIf(IsNull(Rs.Fields("idle_shift3")), "", Rs.Fields("idle_shift3")), "hh:mm")

                Case 2
                    txtidle(5).text = Format(IIf(IsNull(Rs.Fields("idle_shift1")), "", Rs.Fields("idle_shift1")), "hh:mm")
                    txtidle(4).text = Format(IIf(IsNull(Rs.Fields("idle_shift2")), "", Rs.Fields("idle_shift2")), "hh:mm")
                    txtidle(3).text = Format(IIf(IsNull(Rs.Fields("idle_shift3")), "", Rs.Fields("idle_shift3")), "hh:mm")
        
                Case 3
                    txtidle(6).text = Format(IIf(IsNull(Rs.Fields("idle_shift1")), "", Rs.Fields("idle_shift1")), "hh:mm")
                    txtidle(7).text = Format(IIf(IsNull(Rs.Fields("idle_shift2")), "", Rs.Fields("idle_shift2")), "hh:mm")
                    txtidle(8).text = Format(IIf(IsNull(Rs.Fields("idle_shift3")), "", Rs.Fields("idle_shift3")), "hh:mm")

                Case 4
                    txtidle(9).text = Format(IIf(IsNull(Rs.Fields("idle_shift1")), "", Rs.Fields("idle_shift1")), "hh:mm")
                    txtidle(10).text = Format(IIf(IsNull(Rs.Fields("idle_shift2")), "", Rs.Fields("idle_shift2")), "hh:mm")
                    txtidle(11).text = Format(IIf(IsNull(Rs.Fields("idle_shift3")), "", Rs.Fields("idle_shift3")), "hh:mm")

                Case 5
                    txtidle(12).text = Format(IIf(IsNull(Rs.Fields("idle_shift1")), "", Rs.Fields("idle_shift1")), "hh:mm")
                    txtidle(13).text = Format(IIf(IsNull(Rs.Fields("idle_shift2")), "", Rs.Fields("idle_shift2")), "hh:mm")
                    txtidle(14).text = Format(IIf(IsNull(Rs.Fields("idle_shift3")), "", Rs.Fields("idle_shift3")), "hh:mm")

                Case 6
                    txtidle(17).text = Format(IIf(IsNull(Rs.Fields("idle_shift1")), "", Rs.Fields("idle_shift1")), "hh:mm")
                    txtidle(16).text = Format(IIf(IsNull(Rs.Fields("idle_shift2")), "", Rs.Fields("idle_shift2")), "hh:mm")
                    txtidle(15).text = Format(IIf(IsNull(Rs.Fields("idle_shift3")), "", Rs.Fields("idle_shift3")), "hh:mm")

                Case 7
                    txtidle(18).text = Format(IIf(IsNull(Rs.Fields("idle_shift1")), "", Rs.Fields("idle_shift1")), "hh:mm")
                    txtidle(19).text = Format(IIf(IsNull(Rs.Fields("idle_shift2")), "", Rs.Fields("idle_shift2")), "hh:mm")
                    txtidle(20).text = Format(IIf(IsNull(Rs.Fields("idle_shift3")), "", Rs.Fields("idle_shift3")), "hh:mm")

                Case 8
                    txtidle(23).text = Format(IIf(IsNull(Rs.Fields("idle_shift1")), "", Rs.Fields("idle_shift1")), "hh:mm")
                    txtidle(22).text = Format(IIf(IsNull(Rs.Fields("idle_shift2")), "", Rs.Fields("idle_shift2")), "hh:mm")
                    txtidle(21).text = Format(IIf(IsNull(Rs.Fields("idle_shift3")), "", Rs.Fields("idle_shift3")), "hh:mm")

                Case 9
                    txtidle(26).text = Format(IIf(IsNull(Rs.Fields("idle_shift1")), "", Rs.Fields("idle_shift1")), "hh:mm")
                    txtidle(25).text = Format(IIf(IsNull(Rs.Fields("idle_shift2")), "", Rs.Fields("idle_shift2")), "hh:mm")
                    txtidle(24).text = Format(IIf(IsNull(Rs.Fields("idle_shift3")), "", Rs.Fields("idle_shift3")), "hh:mm")
                    
                Case 10
                    txtidle(27).text = Format(IIf(IsNull(Rs.Fields("idle_shift1")), "", Rs.Fields("idle_shift1")), "hh:mm")
                    txtidle(28).text = Format(IIf(IsNull(Rs.Fields("idle_shift2")), "", Rs.Fields("idle_shift2")), "hh:mm")
                    txtidle(29).text = Format(IIf(IsNull(Rs.Fields("idle_shift3")), "", Rs.Fields("idle_shift3")), "hh:mm")

                Case 11
                    txtidle(30).text = Format(IIf(IsNull(Rs.Fields("idle_shift1")), "", Rs.Fields("idle_shift1")), "hh:mm")
                    txtidle(31).text = Format(IIf(IsNull(Rs.Fields("idle_shift2")), "", Rs.Fields("idle_shift2")), "hh:mm")
                    txtidle(32).text = Format(IIf(IsNull(Rs.Fields("idle_shift3")), "", Rs.Fields("idle_shift3")), "hh:mm")
                    
                Case 12
                    txtidle(35).text = Format(IIf(IsNull(Rs.Fields("idle_shift1")), "", Rs.Fields("idle_shift1")), "hh:mm")
                    txtidle(34).text = Format(IIf(IsNull(Rs.Fields("idle_shift2")), "", Rs.Fields("idle_shift2")), "hh:mm")
                    txtidle(33).text = Format(IIf(IsNull(Rs.Fields("idle_shift3")), "", Rs.Fields("idle_shift3")), "hh:mm")
                    
                Case 13
                    txtidle(36).text = Format(IIf(IsNull(Rs.Fields("idle_shift1")), "", Rs.Fields("idle_shift1")), "hh:mm")
                    txtidle(37).text = Format(IIf(IsNull(Rs.Fields("idle_shift2")), "", Rs.Fields("idle_shift2")), "hh:mm")
                    txtidle(38).text = Format(IIf(IsNull(Rs.Fields("idle_shift3")), "", Rs.Fields("idle_shift3")), "hh:mm")
                    
                Case 14
                    txtidle(39).text = Format(IIf(IsNull(Rs.Fields("idle_shift1")), "", Rs.Fields("idle_shift1")), "hh:mm")
                    txtidle(40).text = Format(IIf(IsNull(Rs.Fields("idle_shift2")), "", Rs.Fields("idle_shift2")), "hh:mm")
                    txtidle(41).text = Format(IIf(IsNull(Rs.Fields("idle_shift3")), "", Rs.Fields("idle_shift3")), "hh:mm")
    
                Case 15
                    txtidle(44).text = Format(IIf(IsNull(Rs.Fields("idle_shift1")), "", Rs.Fields("idle_shift1")), "hh:mm")
                    txtidle(43).text = Format(IIf(IsNull(Rs.Fields("idle_shift2")), "", Rs.Fields("idle_shift2")), "hh:mm")
                    txtidle(42).text = Format(IIf(IsNull(Rs.Fields("idle_shift3")), "", Rs.Fields("idle_shift3")), "hh:mm")
                    
                Case 16
                    txtidle(45).text = Format(IIf(IsNull(Rs.Fields("idle_shift1")), "", Rs.Fields("idle_shift1")), "hh:mm")
                    txtidle(46).text = Format(IIf(IsNull(Rs.Fields("idle_shift2")), "", Rs.Fields("idle_shift2")), "hh:mm")
                    txtidle(47).text = Format(IIf(IsNull(Rs.Fields("idle_shift3")), "", Rs.Fields("idle_shift3")), "hh:mm")
                    
                Case 17
                    txtidle(50).text = Format(IIf(IsNull(Rs.Fields("idle_shift1")), "", Rs.Fields("idle_shift1")), "hh:mm")
                    txtidle(49).text = Format(IIf(IsNull(Rs.Fields("idle_shift2")), "", Rs.Fields("idle_shift2")), "hh:mm")
                    txtidle(48).text = Format(IIf(IsNull(Rs.Fields("idle_shift3")), "", Rs.Fields("idle_shift3")), "hh:mm")
                    
                Case 18
                    txtidle(51).text = Format(IIf(IsNull(Rs.Fields("idle_shift1")), "", Rs.Fields("idle_shift1")), "hh:mm")
                    txtidle(52).text = Format(IIf(IsNull(Rs.Fields("idle_shift2")), "", Rs.Fields("idle_shift2")), "hh:mm")
                    txtidle(53).text = Format(IIf(IsNull(Rs.Fields("idle_shift3")), "", Rs.Fields("idle_shift3")), "hh:mm")
                    
                Case 19
                    txtidle(54).text = Format(IIf(IsNull(Rs.Fields("idle_shift1")), "", Rs.Fields("idle_shift1")), "hh:mm")
                    txtidle(55).text = Format(IIf(IsNull(Rs.Fields("idle_shift2")), "", Rs.Fields("idle_shift2")), "hh:mm")
                    txtidle(56).text = Format(IIf(IsNull(Rs.Fields("idle_shift3")), "", Rs.Fields("idle_shift3")), "hh:mm")
                    
                Case 20
                    txtidle(59).text = Format(IIf(IsNull(Rs.Fields("idle_shift1")), "", Rs.Fields("idle_shift1")), "hh:mm")
                    txtidle(58).text = Format(IIf(IsNull(Rs.Fields("idle_shift2")), "", Rs.Fields("idle_shift2")), "hh:mm")
                    txtidle(57).text = Format(IIf(IsNull(Rs.Fields("idle_shift3")), "", Rs.Fields("idle_shift3")), "hh:mm")
                    
                Case 21
                    txtidle(60).text = Format(IIf(IsNull(Rs.Fields("idle_shift1")), "", Rs.Fields("idle_shift1")), "hh:mm")
                    txtidle(61).text = Format(IIf(IsNull(Rs.Fields("idle_shift2")), "", Rs.Fields("idle_shift2")), "hh:mm")
                    txtidle(62).text = Format(IIf(IsNull(Rs.Fields("idle_shift3")), "", Rs.Fields("idle_shift3")), "hh:mm")
                    
                Case 22
                    txtidle(65).text = Format(IIf(IsNull(Rs.Fields("idle_shift1")), "", Rs.Fields("idle_shift1")), "hh:mm")
                    txtidle(64).text = Format(IIf(IsNull(Rs.Fields("idle_shift2")), "", Rs.Fields("idle_shift2")), "hh:mm")
                    txtidle(63).text = Format(IIf(IsNull(Rs.Fields("idle_shift3")), "", Rs.Fields("idle_shift3")), "hh:mm")
                    
                Case 23
                    txtidle(66).text = Format(IIf(IsNull(Rs.Fields("idle_shift1")), "", Rs.Fields("idle_shift1")), "hh:mm")
                    txtidle(67).text = Format(IIf(IsNull(Rs.Fields("idle_shift2")), "", Rs.Fields("idle_shift2")), "hh:mm")
                    txtidle(68).text = Format(IIf(IsNull(Rs.Fields("idle_shift3")), "", Rs.Fields("idle_shift3")), "hh:mm")
                
                Case 24
                    txtidle(69).text = Format(IIf(IsNull(Rs.Fields("idle_shift1")), "", Rs.Fields("idle_shift1")), "hh:mm")
                    txtidle(70).text = Format(IIf(IsNull(Rs.Fields("idle_shift2")), "", Rs.Fields("idle_shift2")), "hh:mm")
                    txtidle(71).text = Format(IIf(IsNull(Rs.Fields("idle_shift3")), "", Rs.Fields("idle_shift3")), "hh:mm")
                                   
                Case 25
                    txtidle(74).text = Format(IIf(IsNull(Rs.Fields("idle_shift1")), "", Rs.Fields("idle_shift1")), "hh:mm")
                    txtidle(73).text = Format(IIf(IsNull(Rs.Fields("idle_shift2")), "", Rs.Fields("idle_shift2")), "hh:mm")
                    txtidle(72).text = Format(IIf(IsNull(Rs.Fields("idle_shift3")), "", Rs.Fields("idle_shift3")), "hh:mm")
                                                 
                Case 26
                    txtidle(75).text = Format(IIf(IsNull(Rs.Fields("idle_shift1")), "", Rs.Fields("idle_shift1")), "hh:mm")
                    txtidle(76).text = Format(IIf(IsNull(Rs.Fields("idle_shift2")), "", Rs.Fields("idle_shift2")), "hh:mm")
                    txtidle(77).text = Format(IIf(IsNull(Rs.Fields("idle_shift3")), "", Rs.Fields("idle_shift3")), "hh:mm")
                
                Case 27
                    txtidle(80).text = Format(IIf(IsNull(Rs.Fields("idle_shift1")), "", Rs.Fields("idle_shift1")), "hh:mm")
                    txtidle(79).text = Format(IIf(IsNull(Rs.Fields("idle_shift2")), "", Rs.Fields("idle_shift2")), "hh:mm")
                    txtidle(78).text = Format(IIf(IsNull(Rs.Fields("idle_shift3")), "", Rs.Fields("idle_shift3")), "hh:mm")
                
                Case 28
                    txtidle(81).text = Format(IIf(IsNull(Rs.Fields("idle_shift1")), "", Rs.Fields("idle_shift1")), "hh:mm")
                    txtidle(82).text = Format(IIf(IsNull(Rs.Fields("idle_shift2")), "", Rs.Fields("idle_shift2")), "hh:mm")
                    txtidle(83).text = Format(IIf(IsNull(Rs.Fields("idle_shift3")), "", Rs.Fields("idle_shift3")), "hh:mm")
                                   
                Case 29
                    txtidle(84).text = Format(IIf(IsNull(Rs.Fields("idle_shift1")), "", Rs.Fields("idle_shift1")), "hh:mm")
                    txtidle(85).text = Format(IIf(IsNull(Rs.Fields("idle_shift2")), "", Rs.Fields("idle_shift2")), "hh:mm")
                    txtidle(86).text = Format(IIf(IsNull(Rs.Fields("idle_shift3")), "", Rs.Fields("idle_shift3")), "hh:mm")
                                                 
                Case 30
                    txtidle(89).text = Format(IIf(IsNull(Rs.Fields("idle_shift1")), "", Rs.Fields("idle_shift1")), "hh:mm")
                    txtidle(88).text = Format(IIf(IsNull(Rs.Fields("idle_shift2")), "", Rs.Fields("idle_shift2")), "hh:mm")
                    txtidle(87).text = Format(IIf(IsNull(Rs.Fields("idle_shift3")), "", Rs.Fields("idle_shift3")), "hh:mm")
                    
                Case 31
                    txtidle(90).text = Format(IIf(IsNull(Rs.Fields("idle_shift1")), "", Rs.Fields("idle_shift1")), "hh:mm")
                    txtidle(91).text = Format(IIf(IsNull(Rs.Fields("idle_shift2")), "", Rs.Fields("idle_shift2")), "hh:mm")
                    txtidle(92).text = Format(IIf(IsNull(Rs.Fields("idle_shift3")), "", Rs.Fields("idle_shift3")), "hh:mm")
                                                 
                Case 32
                    txtidle(95).text = Format(IIf(IsNull(Rs.Fields("idle_shift1")), "", Rs.Fields("idle_shift1")), "hh:mm")
                    txtidle(94).text = Format(IIf(IsNull(Rs.Fields("idle_shift2")), "", Rs.Fields("idle_shift2")), "hh:mm")
                    txtidle(93).text = Format(IIf(IsNull(Rs.Fields("idle_shift3")), "", Rs.Fields("idle_shift3")), "hh:mm")
                
                Case 33
                    txtidle(96).text = Format(IIf(IsNull(Rs.Fields("idle_shift1")), "", Rs.Fields("idle_shift1")), "hh:mm")
                    txtidle(79).text = Format(IIf(IsNull(Rs.Fields("idle_shift2")), "", Rs.Fields("idle_shift2")), "hh:mm")
                    txtidle(78).text = Format(IIf(IsNull(Rs.Fields("idle_shift3")), "", Rs.Fields("idle_shift3")), "hh:mm")
                
                Case 34
                    txtidle(99).text = Format(IIf(IsNull(Rs.Fields("idle_shift1")), "", Rs.Fields("idle_shift1")), "hh:mm")
                    txtidle(97).text = Format(IIf(IsNull(Rs.Fields("idle_shift2")), "", Rs.Fields("idle_shift2")), "hh:mm")
                    txtidle(98).text = Format(IIf(IsNull(Rs.Fields("idle_shift3")), "", Rs.Fields("idle_shift3")), "hh:mm")
                                   
                Case 35
                    txtidle(104).text = Format(IIf(IsNull(Rs.Fields("idle_shift1")), "", Rs.Fields("idle_shift1")), "hh:mm")
                    txtidle(103).text = Format(IIf(IsNull(Rs.Fields("idle_shift2")), "", Rs.Fields("idle_shift2")), "hh:mm")
                    txtidle(102).text = Format(IIf(IsNull(Rs.Fields("idle_shift3")), "", Rs.Fields("idle_shift3")), "hh:mm")
                                                 
                Case 36
                    txtidle(105).text = Format(IIf(IsNull(Rs.Fields("idle_shift1")), "", Rs.Fields("idle_shift1")), "hh:mm")
                    txtidle(106).text = Format(IIf(IsNull(Rs.Fields("idle_shift2")), "", Rs.Fields("idle_shift2")), "hh:mm")
                    txtidle(107).text = Format(IIf(IsNull(Rs.Fields("idle_shift3")), "", Rs.Fields("idle_shift3")), "hh:mm")

                Case 37
                    txtidle(110).text = Format(IIf(IsNull(Rs.Fields("idle_shift1")), "", Rs.Fields("idle_shift1")), "hh:mm")
                    txtidle(109).text = Format(IIf(IsNull(Rs.Fields("idle_shift2")), "", Rs.Fields("idle_shift2")), "hh:mm")
                    txtidle(108).text = Format(IIf(IsNull(Rs.Fields("idle_shift3")), "", Rs.Fields("idle_shift3")), "hh:mm")
                    
                Case 38
                    txtidle(111).text = Format(IIf(IsNull(Rs.Fields("idle_shift1")), "", Rs.Fields("idle_shift1")), "hh:mm")
                    txtidle(112).text = Format(IIf(IsNull(Rs.Fields("idle_shift2")), "", Rs.Fields("idle_shift2")), "hh:mm")
                    txtidle(113).text = Format(IIf(IsNull(Rs.Fields("idle_shift3")), "", Rs.Fields("idle_shift3")), "hh:mm")
                    
                Case 39
                    txtidle(114).text = Format(IIf(IsNull(Rs.Fields("idle_shift1")), "", Rs.Fields("idle_shift1")), "hh:mm")
                    txtidle(115).text = Format(IIf(IsNull(Rs.Fields("idle_shift2")), "", Rs.Fields("idle_shift2")), "hh:mm")
                    txtidle(116).text = Format(IIf(IsNull(Rs.Fields("idle_shift3")), "", Rs.Fields("idle_shift3")), "hh:mm")
                    
                Case 40
                    txtidle(119).text = Format(IIf(IsNull(Rs.Fields("idle_shift1")), "", Rs.Fields("idle_shift1")), "hh:mm")
                    txtidle(118).text = Format(IIf(IsNull(Rs.Fields("idle_shift2")), "", Rs.Fields("idle_shift2")), "hh:mm")
                    txtidle(117).text = Format(IIf(IsNull(Rs.Fields("idle_shift3")), "", Rs.Fields("idle_shift3")), "hh:mm")

                Case 41
                    txtidle(120).text = Format(IIf(IsNull(Rs.Fields("idle_shift1")), "", Rs.Fields("idle_shift1")), "hh:mm")
                    txtidle(121).text = Format(IIf(IsNull(Rs.Fields("idle_shift2")), "", Rs.Fields("idle_shift2")), "hh:mm")
                    txtidle(122).text = Format(IIf(IsNull(Rs.Fields("idle_shift3")), "", Rs.Fields("idle_shift3")), "hh:mm")
                    
                Case 42
                    txtidle(123).text = Format(IIf(IsNull(Rs.Fields("idle_shift1")), "", Rs.Fields("idle_shift1")), "hh:mm")
                    txtidle(124).text = Format(IIf(IsNull(Rs.Fields("idle_shift2")), "", Rs.Fields("idle_shift2")), "hh:mm")
                    txtidle(125).text = Format(IIf(IsNull(Rs.Fields("idle_shift3")), "", Rs.Fields("idle_shift3")), "hh:mm")
                    
            End Select

            Rs.MoveNext
        Loop
    End If

    Set Rs = Nothing

Exit Sub
ErrHandler:
    MsgBox "Error Number: " & Err.Number & vbCrLf & "Description: " & Err.Description, vbExclamation
    
End Sub

Private Sub Timer2_Timer()
On Error Resume Next
    ProgressBar1.Max = 60
    ProgressBar1.Value = ProgressBar1.Value + 1
End Sub

