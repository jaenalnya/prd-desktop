VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form frmTotalYield 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Total Yield"
   ClientHeight    =   9855
   ClientLeft      =   2040
   ClientTop       =   2865
   ClientWidth     =   16470
   Icon            =   "frmTotalYield.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9855
   ScaleWidth      =   16470
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer3 
      Interval        =   1000
      Left            =   16920
      Top             =   45
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "PRODUCTION YIELD"
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
      Left            =   19620
      TabIndex        =   591
      Top             =   12330
      Width           =   6630
      Begin VB.Shape Shape14 
         BackColor       =   &H0080FF80&
         BackStyle       =   1  'Opaque
         Height          =   240
         Left            =   270
         Top             =   270
         Width           =   240
      End
      Begin VB.Label Label15 
         BackColor       =   &H00FFFFFF&
         Caption         =   "> 95 %"
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
         TabIndex        =   595
         Top             =   315
         Width           =   735
      End
      Begin VB.Shape Shape13 
         BackColor       =   &H0000FFFF&
         BackStyle       =   1  'Opaque
         Height          =   240
         Left            =   1575
         Top             =   270
         Width           =   240
      End
      Begin VB.Label Label14 
         BackColor       =   &H00FFFFFF&
         Caption         =   "85 % - 94 %"
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
         Left            =   1935
         TabIndex        =   594
         Top             =   315
         Width           =   1230
      End
      Begin VB.Shape Shape12 
         BackColor       =   &H00C000C0&
         BackStyle       =   1  'Opaque
         Height          =   240
         Left            =   3375
         Top             =   270
         Width           =   240
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFFFF&
         Caption         =   "80% - 84%"
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
         Left            =   3735
         TabIndex        =   593
         Top             =   315
         Width           =   1455
      End
      Begin VB.Shape Shape11 
         BackColor       =   &H000000FF&
         BackStyle       =   1  'Opaque
         Height          =   240
         Left            =   5265
         Top             =   270
         Width           =   240
      End
      Begin VB.Label Label12 
         BackColor       =   &H00FFFFFF&
         Caption         =   "< 80 %"
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
         Left            =   5625
         TabIndex        =   592
         Top             =   315
         Width           =   825
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "PERSENT TARGET"
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
      Left            =   11835
      TabIndex        =   586
      Top             =   12330
      Width           =   6630
      Begin VB.Label Label11 
         BackColor       =   &H00FFFFFF&
         Caption         =   "< 90 %"
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
         Left            =   5625
         TabIndex        =   590
         Top             =   315
         Width           =   825
      End
      Begin VB.Shape Shape10 
         BackColor       =   &H000000FF&
         BackStyle       =   1  'Opaque
         Height          =   240
         Left            =   5265
         Top             =   270
         Width           =   240
      End
      Begin VB.Label Label10 
         BackColor       =   &H00FFFFFF&
         Caption         =   "90% - 94%"
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
         Left            =   3690
         TabIndex        =   589
         Top             =   315
         Width           =   1455
      End
      Begin VB.Shape Shape9 
         BackColor       =   &H00C000C0&
         BackStyle       =   1  'Opaque
         Height          =   240
         Left            =   3330
         Top             =   270
         Width           =   240
      End
      Begin VB.Label Label9 
         BackColor       =   &H00FFFFFF&
         Caption         =   "95% - 97%"
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
         Left            =   1935
         TabIndex        =   588
         Top             =   315
         Width           =   1230
      End
      Begin VB.Shape Shape8 
         BackColor       =   &H0000FFFF&
         BackStyle       =   1  'Opaque
         Height          =   240
         Left            =   1575
         Top             =   270
         Width           =   240
      End
      Begin VB.Label Label8 
         BackColor       =   &H00FFFFFF&
         Caption         =   "> 98 %"
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
         TabIndex        =   587
         Top             =   315
         Width           =   510
      End
      Begin VB.Shape Shape7 
         BackColor       =   &H0080FF80&
         BackStyle       =   1  'Opaque
         Height          =   240
         Left            =   270
         Top             =   270
         Width           =   240
      End
   End
   Begin VB.Frame frameDesign 
      BackColor       =   &H00FFFFFF&
      Height          =   2175
      Index           =   44
      Left            =   23490
      TabIndex        =   585
      Top             =   10035
      Width           =   2760
   End
   Begin VB.Frame frameDesign 
      BackColor       =   &H00FFFFFF&
      Height          =   2175
      Index           =   43
      Left            =   23490
      TabIndex        =   584
      Top             =   7695
      Width           =   2760
   End
   Begin VB.Frame frameDesign 
      BackColor       =   &H00FFFFFF&
      Height          =   2175
      Index           =   42
      Left            =   23490
      TabIndex        =   583
      Top             =   3015
      Width           =   2760
      Begin VB.TextBox txtProdYiled_2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   42
         Left            =   1890
         Locked          =   -1  'True
         TabIndex        =   603
         Text            =   "100"
         Top             =   900
         Width           =   735
      End
      Begin VB.TextBox txtPerctarget_2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   42
         Left            =   1890
         Locked          =   -1  'True
         TabIndex        =   602
         Text            =   "100"
         Top             =   1305
         Width           =   735
      End
      Begin VB.TextBox txtTotalYield_2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   42
         Left            =   1890
         Locked          =   -1  'True
         TabIndex        =   601
         Text            =   "100"
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
         Index           =   42
         Left            =   585
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   600
         Top             =   135
         Width           =   2085
      End
      Begin VB.TextBox txtProdYiled_1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   42
         Left            =   1035
         Locked          =   -1  'True
         TabIndex        =   599
         Text            =   "100"
         Top             =   900
         Width           =   735
      End
      Begin VB.TextBox txtPerctarget_1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   42
         Left            =   1035
         Locked          =   -1  'True
         TabIndex        =   598
         Text            =   "100"
         Top             =   1305
         Width           =   735
      End
      Begin VB.TextBox txtTotalYield_1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   42
         Left            =   1035
         Locked          =   -1  'True
         TabIndex        =   597
         Text            =   "100"
         Top             =   1710
         Width           =   735
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "R"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   209
         Left            =   1935
         TabIndex        =   609
         Top             =   675
         Width           =   645
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "L"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   208
         Left            =   1035
         TabIndex        =   608
         Top             =   675
         Width           =   645
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
         Left            =   45
         TabIndex        =   607
         Top             =   135
         Width           =   510
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "TRG YIELD"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   207
         Left            =   90
         TabIndex        =   606
         Top             =   990
         Width           =   960
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "% TARGET"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   206
         Left            =   90
         TabIndex        =   605
         Top             =   1395
         Width           =   870
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "PRD YIELD"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   205
         Left            =   90
         TabIndex        =   604
         Top             =   1800
         Width           =   870
      End
   End
   Begin VB.Frame frameDesign 
      BackColor       =   &H00FFFFFF&
      Height          =   2175
      Index           =   1
      Left            =   23490
      TabIndex        =   582
      Top             =   5355
      Width           =   2760
   End
   Begin VB.Frame frameDesign 
      BackColor       =   &H00FFFFFF&
      Height          =   2175
      Index           =   41
      Left            =   23490
      TabIndex        =   568
      Top             =   675
      Width           =   2760
      Begin VB.TextBox txtTotalYield_1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   41
         Left            =   1035
         Locked          =   -1  'True
         TabIndex        =   575
         Text            =   "100"
         Top             =   1710
         Width           =   735
      End
      Begin VB.TextBox txtPerctarget_1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   41
         Left            =   1035
         Locked          =   -1  'True
         TabIndex        =   574
         Text            =   "100"
         Top             =   1305
         Width           =   735
      End
      Begin VB.TextBox txtProdYiled_1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   41
         Left            =   1035
         Locked          =   -1  'True
         TabIndex        =   573
         Text            =   "100"
         Top             =   900
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
         Left            =   585
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   572
         Top             =   135
         Width           =   2085
      End
      Begin VB.TextBox txtTotalYield_2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   41
         Left            =   1890
         Locked          =   -1  'True
         TabIndex        =   571
         Text            =   "100"
         Top             =   1710
         Width           =   735
      End
      Begin VB.TextBox txtPerctarget_2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   41
         Left            =   1890
         Locked          =   -1  'True
         TabIndex        =   570
         Text            =   "100"
         Top             =   1305
         Width           =   735
      End
      Begin VB.TextBox txtProdYiled_2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   41
         Left            =   1890
         Locked          =   -1  'True
         TabIndex        =   569
         Text            =   "100"
         Top             =   900
         Width           =   735
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "PRD YIELD"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   9
         Left            =   90
         TabIndex        =   581
         Top             =   1800
         Width           =   870
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "% TARGET"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   8
         Left            =   90
         TabIndex        =   580
         Top             =   1395
         Width           =   870
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "TRG YIELD"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   7
         Left            =   90
         TabIndex        =   579
         Top             =   990
         Width           =   960
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
         Left            =   45
         TabIndex        =   578
         Top             =   135
         Width           =   510
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "L"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   6
         Left            =   1035
         TabIndex        =   577
         Top             =   675
         Width           =   645
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "R"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   5
         Left            =   1935
         TabIndex        =   576
         Top             =   675
         Width           =   645
      End
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
      Left            =   90
      TabIndex        =   561
      Top             =   12330
      Width           =   10725
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
         TabIndex        =   567
         Top             =   315
         Width           =   510
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00004080&
         BackStyle       =   1  'Opaque
         Height          =   240
         Left            =   1575
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
         Left            =   1935
         TabIndex        =   566
         Top             =   315
         Width           =   870
      End
      Begin VB.Shape Shape3 
         BackColor       =   &H00008080&
         BackStyle       =   1  'Opaque
         Height          =   240
         Left            =   3195
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
         Left            =   3555
         TabIndex        =   565
         Top             =   315
         Width           =   1455
      End
      Begin VB.Shape Shape4 
         BackColor       =   &H00808000&
         BackStyle       =   1  'Opaque
         Height          =   240
         Left            =   5265
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
         Left            =   5625
         TabIndex        =   564
         Top             =   315
         Width           =   825
      End
      Begin VB.Shape Shape5 
         BackColor       =   &H00C000C0&
         BackStyle       =   1  'Opaque
         Height          =   240
         Left            =   6750
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
         Left            =   7110
         TabIndex        =   563
         Top             =   315
         Width           =   1455
      End
      Begin VB.Shape Shape6 
         BackColor       =   &H000000FF&
         BackStyle       =   1  'Opaque
         Height          =   240
         Left            =   8685
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
         Left            =   9045
         TabIndex        =   562
         Top             =   315
         Width           =   1455
      End
   End
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   630
      Top             =   11925
   End
   Begin VB.Timer Timer1 
      Interval        =   60000
      Left            =   180
      Top             =   11925
   End
   Begin VB.Frame frameDesign 
      BackColor       =   &H00FFFFFF&
      Height          =   2175
      Index           =   40
      Left            =   20565
      TabIndex        =   546
      Top             =   10035
      Width           =   2760
      Begin VB.TextBox txtProdYiled_2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   40
         Left            =   1890
         Locked          =   -1  'True
         TabIndex        =   553
         Text            =   "100"
         Top             =   900
         Width           =   735
      End
      Begin VB.TextBox txtPerctarget_2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   40
         Left            =   1890
         Locked          =   -1  'True
         TabIndex        =   552
         Text            =   "100"
         Top             =   1305
         Width           =   735
      End
      Begin VB.TextBox txtTotalYield_2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   40
         Left            =   1890
         Locked          =   -1  'True
         TabIndex        =   551
         Text            =   "100"
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
         Index           =   40
         Left            =   585
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   550
         Top             =   135
         Width           =   2085
      End
      Begin VB.TextBox txtProdYiled_1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   40
         Left            =   1035
         Locked          =   -1  'True
         TabIndex        =   549
         Text            =   "100"
         Top             =   900
         Width           =   735
      End
      Begin VB.TextBox txtPerctarget_1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   40
         Left            =   1035
         Locked          =   -1  'True
         TabIndex        =   548
         Text            =   "100"
         Top             =   1305
         Width           =   735
      End
      Begin VB.TextBox txtTotalYield_1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   40
         Left            =   1035
         Locked          =   -1  'True
         TabIndex        =   547
         Text            =   "100"
         Top             =   1710
         Width           =   735
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "R"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   200
         Left            =   1935
         TabIndex        =   559
         Top             =   675
         Width           =   645
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "L"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   201
         Left            =   1035
         TabIndex        =   558
         Top             =   675
         Width           =   645
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
         Left            =   45
         TabIndex        =   557
         Top             =   135
         Width           =   510
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "TRG YIELD"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   202
         Left            =   90
         TabIndex        =   556
         Top             =   990
         Width           =   960
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "% TARGET"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   203
         Left            =   90
         TabIndex        =   555
         Top             =   1395
         Width           =   870
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "PRD YIELD"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   204
         Left            =   90
         TabIndex        =   554
         Top             =   1800
         Width           =   870
      End
   End
   Begin VB.Frame frameDesign 
      BackColor       =   &H00FFFFFF&
      Height          =   2175
      Index           =   39
      Left            =   20565
      TabIndex        =   532
      Top             =   7695
      Width           =   2760
      Begin VB.TextBox txtProdYiled_2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   39
         Left            =   1890
         Locked          =   -1  'True
         TabIndex        =   539
         Text            =   "100"
         Top             =   900
         Width           =   735
      End
      Begin VB.TextBox txtPerctarget_2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   39
         Left            =   1890
         Locked          =   -1  'True
         TabIndex        =   538
         Text            =   "100"
         Top             =   1305
         Width           =   735
      End
      Begin VB.TextBox txtTotalYield_2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   39
         Left            =   1890
         Locked          =   -1  'True
         TabIndex        =   537
         Text            =   "100"
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
         Left            =   585
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   536
         Top             =   135
         Width           =   2085
      End
      Begin VB.TextBox txtProdYiled_1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   39
         Left            =   1035
         Locked          =   -1  'True
         TabIndex        =   535
         Text            =   "100"
         Top             =   900
         Width           =   735
      End
      Begin VB.TextBox txtPerctarget_1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   39
         Left            =   1035
         Locked          =   -1  'True
         TabIndex        =   534
         Text            =   "100"
         Top             =   1305
         Width           =   735
      End
      Begin VB.TextBox txtTotalYield_1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   39
         Left            =   1035
         Locked          =   -1  'True
         TabIndex        =   533
         Text            =   "100"
         Top             =   1710
         Width           =   735
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "R"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   195
         Left            =   1935
         TabIndex        =   545
         Top             =   675
         Width           =   645
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "L"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   196
         Left            =   1035
         TabIndex        =   544
         Top             =   675
         Width           =   645
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
         Left            =   45
         TabIndex        =   543
         Top             =   135
         Width           =   510
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "TRG YIELD"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   197
         Left            =   90
         TabIndex        =   542
         Top             =   990
         Width           =   960
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "% TARGET"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   198
         Left            =   90
         TabIndex        =   541
         Top             =   1395
         Width           =   870
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "PRD YIELD"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   199
         Left            =   90
         TabIndex        =   540
         Top             =   1800
         Width           =   870
      End
   End
   Begin VB.Frame frameDesign 
      BackColor       =   &H00FFFFFF&
      Height          =   2175
      Index           =   38
      Left            =   20565
      TabIndex        =   518
      Top             =   5355
      Width           =   2760
      Begin VB.TextBox txtProdYiled_2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   38
         Left            =   1890
         Locked          =   -1  'True
         TabIndex        =   525
         Text            =   "100"
         Top             =   900
         Width           =   735
      End
      Begin VB.TextBox txtPerctarget_2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   38
         Left            =   1890
         Locked          =   -1  'True
         TabIndex        =   524
         Text            =   "100"
         Top             =   1305
         Width           =   735
      End
      Begin VB.TextBox txtTotalYield_2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   38
         Left            =   1890
         Locked          =   -1  'True
         TabIndex        =   523
         Text            =   "100"
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
         Index           =   38
         Left            =   585
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   522
         Top             =   135
         Width           =   2085
      End
      Begin VB.TextBox txtProdYiled_1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   38
         Left            =   1035
         Locked          =   -1  'True
         TabIndex        =   521
         Text            =   "100"
         Top             =   900
         Width           =   735
      End
      Begin VB.TextBox txtPerctarget_1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   38
         Left            =   1035
         Locked          =   -1  'True
         TabIndex        =   520
         Text            =   "100"
         Top             =   1305
         Width           =   735
      End
      Begin VB.TextBox txtTotalYield_1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   38
         Left            =   1035
         Locked          =   -1  'True
         TabIndex        =   519
         Text            =   "100"
         Top             =   1710
         Width           =   735
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "R"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   190
         Left            =   1935
         TabIndex        =   531
         Top             =   675
         Width           =   645
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "L"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   191
         Left            =   1035
         TabIndex        =   530
         Top             =   675
         Width           =   645
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
         Left            =   45
         TabIndex        =   529
         Top             =   135
         Width           =   510
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "TRG YIELD"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   192
         Left            =   90
         TabIndex        =   528
         Top             =   990
         Width           =   960
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "% TARGET"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   193
         Left            =   90
         TabIndex        =   527
         Top             =   1395
         Width           =   870
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "PRD YIELD"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   194
         Left            =   90
         TabIndex        =   526
         Top             =   1800
         Width           =   870
      End
   End
   Begin VB.Frame frameDesign 
      BackColor       =   &H00FFFFFF&
      Height          =   2175
      Index           =   37
      Left            =   20565
      TabIndex        =   504
      Top             =   3015
      Width           =   2760
      Begin VB.TextBox txtProdYiled_2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   37
         Left            =   1890
         Locked          =   -1  'True
         TabIndex        =   511
         Text            =   "100"
         Top             =   900
         Width           =   735
      End
      Begin VB.TextBox txtPerctarget_2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   37
         Left            =   1890
         Locked          =   -1  'True
         TabIndex        =   510
         Text            =   "100"
         Top             =   1305
         Width           =   735
      End
      Begin VB.TextBox txtTotalYield_2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   37
         Left            =   1890
         Locked          =   -1  'True
         TabIndex        =   509
         Text            =   "100"
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
         Index           =   37
         Left            =   585
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   508
         Top             =   135
         Width           =   2085
      End
      Begin VB.TextBox txtProdYiled_1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   37
         Left            =   1035
         Locked          =   -1  'True
         TabIndex        =   507
         Text            =   "100"
         Top             =   900
         Width           =   735
      End
      Begin VB.TextBox txtPerctarget_1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   37
         Left            =   1035
         Locked          =   -1  'True
         TabIndex        =   506
         Text            =   "100"
         Top             =   1305
         Width           =   735
      End
      Begin VB.TextBox txtTotalYield_1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   37
         Left            =   1035
         Locked          =   -1  'True
         TabIndex        =   505
         Text            =   "100"
         Top             =   1710
         Width           =   735
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "R"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   185
         Left            =   1935
         TabIndex        =   517
         Top             =   675
         Width           =   645
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "L"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   186
         Left            =   1035
         TabIndex        =   516
         Top             =   675
         Width           =   645
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
         Left            =   45
         TabIndex        =   515
         Top             =   135
         Width           =   510
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "TRG YIELD"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   187
         Left            =   90
         TabIndex        =   514
         Top             =   990
         Width           =   960
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "% TARGET"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   188
         Left            =   90
         TabIndex        =   513
         Top             =   1395
         Width           =   870
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "PRD YIELD"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   189
         Left            =   90
         TabIndex        =   512
         Top             =   1800
         Width           =   870
      End
   End
   Begin VB.Frame frameDesign 
      BackColor       =   &H00FFFFFF&
      Height          =   2175
      Index           =   36
      Left            =   20565
      TabIndex        =   490
      Top             =   675
      Width           =   2760
      Begin VB.TextBox txtTotalYield_1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   36
         Left            =   1035
         Locked          =   -1  'True
         TabIndex        =   497
         Text            =   "100"
         Top             =   1710
         Width           =   735
      End
      Begin VB.TextBox txtPerctarget_1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   36
         Left            =   1035
         Locked          =   -1  'True
         TabIndex        =   496
         Text            =   "100"
         Top             =   1305
         Width           =   735
      End
      Begin VB.TextBox txtProdYiled_1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   36
         Left            =   1035
         Locked          =   -1  'True
         TabIndex        =   495
         Text            =   "100"
         Top             =   900
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
         Left            =   585
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   494
         Top             =   135
         Width           =   2085
      End
      Begin VB.TextBox txtTotalYield_2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   36
         Left            =   1890
         Locked          =   -1  'True
         TabIndex        =   493
         Text            =   "100"
         Top             =   1710
         Width           =   735
      End
      Begin VB.TextBox txtPerctarget_2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   36
         Left            =   1890
         Locked          =   -1  'True
         TabIndex        =   492
         Text            =   "100"
         Top             =   1305
         Width           =   735
      End
      Begin VB.TextBox txtProdYiled_2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   36
         Left            =   1890
         Locked          =   -1  'True
         TabIndex        =   491
         Text            =   "100"
         Top             =   900
         Width           =   735
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "PRD YIELD"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   180
         Left            =   90
         TabIndex        =   503
         Top             =   1800
         Width           =   870
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "% TARGET"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   181
         Left            =   90
         TabIndex        =   502
         Top             =   1395
         Width           =   870
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "TRG YIELD"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   182
         Left            =   90
         TabIndex        =   501
         Top             =   990
         Width           =   960
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
         Left            =   45
         TabIndex        =   500
         Top             =   135
         Width           =   510
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "L"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   183
         Left            =   1035
         TabIndex        =   499
         Top             =   675
         Width           =   645
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "R"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   184
         Left            =   1935
         TabIndex        =   498
         Top             =   675
         Width           =   645
      End
   End
   Begin VB.Frame frameDesign 
      BackColor       =   &H00FFFFFF&
      Height          =   2175
      Index           =   35
      Left            =   17640
      TabIndex        =   476
      Top             =   10035
      Width           =   2760
      Begin VB.TextBox txtProdYiled_2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   35
         Left            =   1890
         Locked          =   -1  'True
         TabIndex        =   483
         Text            =   "100"
         Top             =   900
         Width           =   735
      End
      Begin VB.TextBox txtPerctarget_2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   35
         Left            =   1890
         Locked          =   -1  'True
         TabIndex        =   482
         Text            =   "100"
         Top             =   1305
         Width           =   735
      End
      Begin VB.TextBox txtTotalYield_2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   35
         Left            =   1890
         Locked          =   -1  'True
         TabIndex        =   481
         Text            =   "100"
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
         Index           =   35
         Left            =   585
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   480
         Top             =   135
         Width           =   2085
      End
      Begin VB.TextBox txtProdYiled_1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   35
         Left            =   1035
         Locked          =   -1  'True
         TabIndex        =   479
         Text            =   "100"
         Top             =   900
         Width           =   735
      End
      Begin VB.TextBox txtPerctarget_1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   35
         Left            =   1035
         Locked          =   -1  'True
         TabIndex        =   478
         Text            =   "100"
         Top             =   1305
         Width           =   735
      End
      Begin VB.TextBox txtTotalYield_1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   35
         Left            =   1035
         Locked          =   -1  'True
         TabIndex        =   477
         Text            =   "100"
         Top             =   1710
         Width           =   735
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "R"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   175
         Left            =   1935
         TabIndex        =   489
         Top             =   675
         Width           =   645
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "L"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   176
         Left            =   1035
         TabIndex        =   488
         Top             =   675
         Width           =   645
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
         Left            =   45
         TabIndex        =   487
         Top             =   135
         Width           =   510
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "TRG YIELD"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   177
         Left            =   90
         TabIndex        =   486
         Top             =   990
         Width           =   960
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "% TARGET"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   178
         Left            =   90
         TabIndex        =   485
         Top             =   1395
         Width           =   870
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "PRD YIELD"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   179
         Left            =   90
         TabIndex        =   484
         Top             =   1800
         Width           =   870
      End
   End
   Begin VB.Frame frameDesign 
      BackColor       =   &H00FFFFFF&
      Height          =   2175
      Index           =   34
      Left            =   17640
      TabIndex        =   462
      Top             =   7695
      Width           =   2760
      Begin VB.TextBox txtProdYiled_2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   34
         Left            =   1890
         Locked          =   -1  'True
         TabIndex        =   469
         Text            =   "100"
         Top             =   900
         Width           =   735
      End
      Begin VB.TextBox txtPerctarget_2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   34
         Left            =   1890
         Locked          =   -1  'True
         TabIndex        =   468
         Text            =   "100"
         Top             =   1305
         Width           =   735
      End
      Begin VB.TextBox txtTotalYield_2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   34
         Left            =   1890
         Locked          =   -1  'True
         TabIndex        =   467
         Text            =   "100"
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
         Left            =   585
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   466
         Top             =   135
         Width           =   2085
      End
      Begin VB.TextBox txtProdYiled_1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   34
         Left            =   1035
         Locked          =   -1  'True
         TabIndex        =   465
         Text            =   "100"
         Top             =   900
         Width           =   735
      End
      Begin VB.TextBox txtPerctarget_1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   34
         Left            =   1035
         Locked          =   -1  'True
         TabIndex        =   464
         Text            =   "100"
         Top             =   1305
         Width           =   735
      End
      Begin VB.TextBox txtTotalYield_1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   34
         Left            =   1035
         Locked          =   -1  'True
         TabIndex        =   463
         Text            =   "100"
         Top             =   1710
         Width           =   735
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "R"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   170
         Left            =   1935
         TabIndex        =   475
         Top             =   675
         Width           =   645
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "L"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   171
         Left            =   1035
         TabIndex        =   474
         Top             =   675
         Width           =   645
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
         Left            =   45
         TabIndex        =   473
         Top             =   135
         Width           =   510
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "TRG YIELD"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   172
         Left            =   90
         TabIndex        =   472
         Top             =   990
         Width           =   960
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "% TARGET"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   173
         Left            =   90
         TabIndex        =   471
         Top             =   1395
         Width           =   870
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "PRD YIELD"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   174
         Left            =   90
         TabIndex        =   470
         Top             =   1800
         Width           =   870
      End
   End
   Begin VB.Frame frameDesign 
      BackColor       =   &H00FFFFFF&
      Height          =   2175
      Index           =   33
      Left            =   17640
      TabIndex        =   448
      Top             =   5355
      Width           =   2760
      Begin VB.TextBox txtProdYiled_2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   33
         Left            =   1890
         Locked          =   -1  'True
         TabIndex        =   455
         Text            =   "100"
         Top             =   900
         Width           =   735
      End
      Begin VB.TextBox txtPerctarget_2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   33
         Left            =   1890
         Locked          =   -1  'True
         TabIndex        =   454
         Text            =   "100"
         Top             =   1305
         Width           =   735
      End
      Begin VB.TextBox txtTotalYield_2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   33
         Left            =   1890
         Locked          =   -1  'True
         TabIndex        =   453
         Text            =   "100"
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
         Index           =   33
         Left            =   585
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   452
         Top             =   135
         Width           =   2085
      End
      Begin VB.TextBox txtProdYiled_1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   33
         Left            =   1035
         Locked          =   -1  'True
         TabIndex        =   451
         Text            =   "100"
         Top             =   900
         Width           =   735
      End
      Begin VB.TextBox txtPerctarget_1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   33
         Left            =   1035
         Locked          =   -1  'True
         TabIndex        =   450
         Text            =   "100"
         Top             =   1305
         Width           =   735
      End
      Begin VB.TextBox txtTotalYield_1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   33
         Left            =   1035
         Locked          =   -1  'True
         TabIndex        =   449
         Text            =   "100"
         Top             =   1710
         Width           =   735
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "R"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   165
         Left            =   1935
         TabIndex        =   461
         Top             =   675
         Width           =   645
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "L"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   166
         Left            =   1035
         TabIndex        =   460
         Top             =   675
         Width           =   645
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
         Left            =   45
         TabIndex        =   459
         Top             =   135
         Width           =   510
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "TRG YIELD"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   167
         Left            =   90
         TabIndex        =   458
         Top             =   990
         Width           =   960
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "% TARGET"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   168
         Left            =   90
         TabIndex        =   457
         Top             =   1395
         Width           =   870
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "PRD YIELD"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   169
         Left            =   90
         TabIndex        =   456
         Top             =   1800
         Width           =   870
      End
   End
   Begin VB.Frame frameDesign 
      BackColor       =   &H00FFFFFF&
      Height          =   2175
      Index           =   32
      Left            =   17640
      TabIndex        =   434
      Top             =   3015
      Width           =   2760
      Begin VB.TextBox txtProdYiled_2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   32
         Left            =   1890
         Locked          =   -1  'True
         TabIndex        =   441
         Text            =   "100"
         Top             =   900
         Width           =   735
      End
      Begin VB.TextBox txtPerctarget_2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   32
         Left            =   1890
         Locked          =   -1  'True
         TabIndex        =   440
         Text            =   "100"
         Top             =   1305
         Width           =   735
      End
      Begin VB.TextBox txtTotalYield_2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   32
         Left            =   1890
         Locked          =   -1  'True
         TabIndex        =   439
         Text            =   "100"
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
         Index           =   32
         Left            =   585
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   438
         Top             =   135
         Width           =   2085
      End
      Begin VB.TextBox txtProdYiled_1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   32
         Left            =   1035
         Locked          =   -1  'True
         TabIndex        =   437
         Text            =   "100"
         Top             =   900
         Width           =   735
      End
      Begin VB.TextBox txtPerctarget_1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   32
         Left            =   1035
         Locked          =   -1  'True
         TabIndex        =   436
         Text            =   "100"
         Top             =   1305
         Width           =   735
      End
      Begin VB.TextBox txtTotalYield_1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   32
         Left            =   1035
         Locked          =   -1  'True
         TabIndex        =   435
         Text            =   "100"
         Top             =   1710
         Width           =   735
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "R"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   160
         Left            =   1935
         TabIndex        =   447
         Top             =   675
         Width           =   645
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "L"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   161
         Left            =   1035
         TabIndex        =   446
         Top             =   675
         Width           =   645
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
         Left            =   45
         TabIndex        =   445
         Top             =   135
         Width           =   510
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "TRG YIELD"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   162
         Left            =   90
         TabIndex        =   444
         Top             =   990
         Width           =   960
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "% TARGET"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   163
         Left            =   90
         TabIndex        =   443
         Top             =   1395
         Width           =   870
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "PRD YIELD"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   164
         Left            =   90
         TabIndex        =   442
         Top             =   1800
         Width           =   870
      End
   End
   Begin VB.Frame frameDesign 
      BackColor       =   &H00FFFFFF&
      Height          =   2175
      Index           =   31
      Left            =   17640
      TabIndex        =   420
      Top             =   675
      Width           =   2760
      Begin VB.TextBox txtTotalYield_1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   31
         Left            =   1035
         Locked          =   -1  'True
         TabIndex        =   427
         Text            =   "100"
         Top             =   1710
         Width           =   735
      End
      Begin VB.TextBox txtPerctarget_1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   31
         Left            =   1035
         Locked          =   -1  'True
         TabIndex        =   426
         Text            =   "100"
         Top             =   1305
         Width           =   735
      End
      Begin VB.TextBox txtProdYiled_1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   31
         Left            =   1035
         Locked          =   -1  'True
         TabIndex        =   425
         Text            =   "100"
         Top             =   900
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
         Left            =   585
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   424
         Top             =   135
         Width           =   2085
      End
      Begin VB.TextBox txtTotalYield_2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   31
         Left            =   1890
         Locked          =   -1  'True
         TabIndex        =   423
         Text            =   "100"
         Top             =   1710
         Width           =   735
      End
      Begin VB.TextBox txtPerctarget_2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   31
         Left            =   1890
         Locked          =   -1  'True
         TabIndex        =   422
         Text            =   "100"
         Top             =   1305
         Width           =   735
      End
      Begin VB.TextBox txtProdYiled_2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   31
         Left            =   1890
         Locked          =   -1  'True
         TabIndex        =   421
         Text            =   "100"
         Top             =   900
         Width           =   735
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "PRD YIELD"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   155
         Left            =   90
         TabIndex        =   433
         Top             =   1800
         Width           =   870
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "% TARGET"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   156
         Left            =   90
         TabIndex        =   432
         Top             =   1395
         Width           =   870
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "TRG YIELD"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   157
         Left            =   90
         TabIndex        =   431
         Top             =   990
         Width           =   960
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
         Left            =   45
         TabIndex        =   430
         Top             =   135
         Width           =   510
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "L"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   158
         Left            =   1035
         TabIndex        =   429
         Top             =   675
         Width           =   645
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "R"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   159
         Left            =   1935
         TabIndex        =   428
         Top             =   675
         Width           =   645
      End
   End
   Begin VB.Frame frameDesign 
      BackColor       =   &H00FFFFFF&
      Height          =   2175
      Index           =   30
      Left            =   14715
      TabIndex        =   406
      Top             =   10035
      Width           =   2760
      Begin VB.TextBox txtProdYiled_2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   30
         Left            =   1890
         Locked          =   -1  'True
         TabIndex        =   413
         Text            =   "100"
         Top             =   900
         Width           =   735
      End
      Begin VB.TextBox txtPerctarget_2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   30
         Left            =   1890
         Locked          =   -1  'True
         TabIndex        =   412
         Text            =   "100"
         Top             =   1305
         Width           =   735
      End
      Begin VB.TextBox txtTotalYield_2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   30
         Left            =   1890
         Locked          =   -1  'True
         TabIndex        =   411
         Text            =   "100"
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
         Index           =   30
         Left            =   585
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   410
         Top             =   135
         Width           =   2085
      End
      Begin VB.TextBox txtProdYiled_1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   30
         Left            =   1035
         Locked          =   -1  'True
         TabIndex        =   409
         Text            =   "100"
         Top             =   900
         Width           =   735
      End
      Begin VB.TextBox txtPerctarget_1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   30
         Left            =   1035
         Locked          =   -1  'True
         TabIndex        =   408
         Text            =   "100"
         Top             =   1305
         Width           =   735
      End
      Begin VB.TextBox txtTotalYield_1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   30
         Left            =   1035
         Locked          =   -1  'True
         TabIndex        =   407
         Text            =   "100"
         Top             =   1710
         Width           =   735
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "R"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   150
         Left            =   1935
         TabIndex        =   419
         Top             =   675
         Width           =   645
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "L"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   151
         Left            =   1035
         TabIndex        =   418
         Top             =   675
         Width           =   645
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
         Left            =   45
         TabIndex        =   417
         Top             =   135
         Width           =   510
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "TRG YIELD"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   152
         Left            =   90
         TabIndex        =   416
         Top             =   990
         Width           =   960
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "% TARGET"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   153
         Left            =   90
         TabIndex        =   415
         Top             =   1395
         Width           =   870
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "PRD YIELD"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   154
         Left            =   90
         TabIndex        =   414
         Top             =   1800
         Width           =   870
      End
   End
   Begin VB.Frame frameDesign 
      BackColor       =   &H00FFFFFF&
      Height          =   2175
      Index           =   29
      Left            =   14715
      TabIndex        =   392
      Top             =   7695
      Width           =   2760
      Begin VB.TextBox txtProdYiled_2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   29
         Left            =   1890
         Locked          =   -1  'True
         TabIndex        =   399
         Text            =   "100"
         Top             =   900
         Width           =   735
      End
      Begin VB.TextBox txtPerctarget_2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   29
         Left            =   1890
         Locked          =   -1  'True
         TabIndex        =   398
         Text            =   "100"
         Top             =   1305
         Width           =   735
      End
      Begin VB.TextBox txtTotalYield_2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   29
         Left            =   1890
         Locked          =   -1  'True
         TabIndex        =   397
         Text            =   "100"
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
         Left            =   585
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   396
         Top             =   135
         Width           =   2085
      End
      Begin VB.TextBox txtProdYiled_1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   29
         Left            =   1035
         Locked          =   -1  'True
         TabIndex        =   395
         Text            =   "100"
         Top             =   900
         Width           =   735
      End
      Begin VB.TextBox txtPerctarget_1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   29
         Left            =   1035
         Locked          =   -1  'True
         TabIndex        =   394
         Text            =   "100"
         Top             =   1305
         Width           =   735
      End
      Begin VB.TextBox txtTotalYield_1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   29
         Left            =   1035
         Locked          =   -1  'True
         TabIndex        =   393
         Text            =   "100"
         Top             =   1710
         Width           =   735
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "R"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   145
         Left            =   1935
         TabIndex        =   405
         Top             =   675
         Width           =   645
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "L"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   146
         Left            =   1035
         TabIndex        =   404
         Top             =   675
         Width           =   645
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
         Left            =   45
         TabIndex        =   403
         Top             =   135
         Width           =   510
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "TRG YIELD"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   147
         Left            =   90
         TabIndex        =   402
         Top             =   990
         Width           =   960
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "% TARGET"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   148
         Left            =   90
         TabIndex        =   401
         Top             =   1395
         Width           =   870
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "PRD YIELD"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   149
         Left            =   90
         TabIndex        =   400
         Top             =   1800
         Width           =   870
      End
   End
   Begin VB.Frame frameDesign 
      BackColor       =   &H00FFFFFF&
      Height          =   2175
      Index           =   28
      Left            =   14715
      TabIndex        =   378
      Top             =   5355
      Width           =   2760
      Begin VB.TextBox txtProdYiled_2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   28
         Left            =   1890
         Locked          =   -1  'True
         TabIndex        =   385
         Text            =   "100"
         Top             =   900
         Width           =   735
      End
      Begin VB.TextBox txtPerctarget_2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   28
         Left            =   1890
         Locked          =   -1  'True
         TabIndex        =   384
         Text            =   "100"
         Top             =   1305
         Width           =   735
      End
      Begin VB.TextBox txtTotalYield_2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   28
         Left            =   1890
         Locked          =   -1  'True
         TabIndex        =   383
         Text            =   "100"
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
         Index           =   28
         Left            =   585
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   382
         Top             =   135
         Width           =   2085
      End
      Begin VB.TextBox txtProdYiled_1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   28
         Left            =   1035
         Locked          =   -1  'True
         TabIndex        =   381
         Text            =   "100"
         Top             =   900
         Width           =   735
      End
      Begin VB.TextBox txtPerctarget_1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   28
         Left            =   1035
         Locked          =   -1  'True
         TabIndex        =   380
         Text            =   "100"
         Top             =   1305
         Width           =   735
      End
      Begin VB.TextBox txtTotalYield_1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   28
         Left            =   1035
         Locked          =   -1  'True
         TabIndex        =   379
         Text            =   "100"
         Top             =   1710
         Width           =   735
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "R"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   140
         Left            =   1935
         TabIndex        =   391
         Top             =   675
         Width           =   645
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "L"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   141
         Left            =   1035
         TabIndex        =   390
         Top             =   675
         Width           =   645
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
         Left            =   45
         TabIndex        =   389
         Top             =   135
         Width           =   510
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "TRG YIELD"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   142
         Left            =   90
         TabIndex        =   388
         Top             =   990
         Width           =   960
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "% TARGET"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   143
         Left            =   90
         TabIndex        =   387
         Top             =   1395
         Width           =   870
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "PRD YIELD"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   144
         Left            =   90
         TabIndex        =   386
         Top             =   1800
         Width           =   870
      End
   End
   Begin VB.Frame frameDesign 
      BackColor       =   &H00FFFFFF&
      Height          =   2175
      Index           =   27
      Left            =   14715
      TabIndex        =   364
      Top             =   3015
      Width           =   2760
      Begin VB.TextBox txtProdYiled_2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   27
         Left            =   1890
         Locked          =   -1  'True
         TabIndex        =   371
         Text            =   "100"
         Top             =   900
         Width           =   735
      End
      Begin VB.TextBox txtPerctarget_2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   27
         Left            =   1890
         Locked          =   -1  'True
         TabIndex        =   370
         Text            =   "100"
         Top             =   1305
         Width           =   735
      End
      Begin VB.TextBox txtTotalYield_2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   27
         Left            =   1890
         Locked          =   -1  'True
         TabIndex        =   369
         Text            =   "100"
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
         Index           =   27
         Left            =   585
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   368
         Top             =   135
         Width           =   2085
      End
      Begin VB.TextBox txtProdYiled_1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   27
         Left            =   1035
         Locked          =   -1  'True
         TabIndex        =   367
         Text            =   "100"
         Top             =   900
         Width           =   735
      End
      Begin VB.TextBox txtPerctarget_1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   27
         Left            =   1035
         Locked          =   -1  'True
         TabIndex        =   366
         Text            =   "100"
         Top             =   1305
         Width           =   735
      End
      Begin VB.TextBox txtTotalYield_1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   27
         Left            =   1035
         Locked          =   -1  'True
         TabIndex        =   365
         Text            =   "100"
         Top             =   1710
         Width           =   735
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "R"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   135
         Left            =   1935
         TabIndex        =   377
         Top             =   675
         Width           =   645
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "L"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   136
         Left            =   1035
         TabIndex        =   376
         Top             =   675
         Width           =   645
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
         Left            =   45
         TabIndex        =   375
         Top             =   135
         Width           =   510
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "TRG YIELD"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   137
         Left            =   90
         TabIndex        =   374
         Top             =   990
         Width           =   960
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "% TARGET"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   138
         Left            =   90
         TabIndex        =   373
         Top             =   1395
         Width           =   870
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "PRD YIELD"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   139
         Left            =   90
         TabIndex        =   372
         Top             =   1800
         Width           =   870
      End
   End
   Begin VB.Frame frameDesign 
      BackColor       =   &H00FFFFFF&
      Height          =   2175
      Index           =   26
      Left            =   14715
      TabIndex        =   350
      Top             =   675
      Width           =   2760
      Begin VB.TextBox txtTotalYield_1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   26
         Left            =   1035
         Locked          =   -1  'True
         TabIndex        =   357
         Text            =   "100"
         Top             =   1710
         Width           =   735
      End
      Begin VB.TextBox txtPerctarget_1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   26
         Left            =   1035
         Locked          =   -1  'True
         TabIndex        =   356
         Text            =   "100"
         Top             =   1305
         Width           =   735
      End
      Begin VB.TextBox txtProdYiled_1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   26
         Left            =   1035
         Locked          =   -1  'True
         TabIndex        =   355
         Text            =   "100"
         Top             =   900
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
         Left            =   585
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   354
         Top             =   135
         Width           =   2085
      End
      Begin VB.TextBox txtTotalYield_2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   26
         Left            =   1890
         Locked          =   -1  'True
         TabIndex        =   353
         Text            =   "100"
         Top             =   1710
         Width           =   735
      End
      Begin VB.TextBox txtPerctarget_2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   26
         Left            =   1890
         Locked          =   -1  'True
         TabIndex        =   352
         Text            =   "100"
         Top             =   1305
         Width           =   735
      End
      Begin VB.TextBox txtProdYiled_2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   26
         Left            =   1890
         Locked          =   -1  'True
         TabIndex        =   351
         Text            =   "100"
         Top             =   900
         Width           =   735
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "PRD YIELD"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   130
         Left            =   90
         TabIndex        =   363
         Top             =   1800
         Width           =   870
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "% TARGET"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   131
         Left            =   90
         TabIndex        =   362
         Top             =   1395
         Width           =   870
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "TRG YIELD"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   132
         Left            =   90
         TabIndex        =   361
         Top             =   990
         Width           =   960
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
         Left            =   45
         TabIndex        =   360
         Top             =   135
         Width           =   510
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "L"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   133
         Left            =   1035
         TabIndex        =   359
         Top             =   675
         Width           =   645
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "R"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   134
         Left            =   1935
         TabIndex        =   358
         Top             =   675
         Width           =   645
      End
   End
   Begin VB.Frame frameDesign 
      BackColor       =   &H00FFFFFF&
      Height          =   2175
      Index           =   25
      Left            =   11790
      TabIndex        =   336
      Top             =   10035
      Width           =   2760
      Begin VB.TextBox txtProdYiled_2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   25
         Left            =   1890
         Locked          =   -1  'True
         TabIndex        =   343
         Text            =   "100"
         Top             =   900
         Width           =   735
      End
      Begin VB.TextBox txtPerctarget_2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   25
         Left            =   1890
         Locked          =   -1  'True
         TabIndex        =   342
         Text            =   "100"
         Top             =   1305
         Width           =   735
      End
      Begin VB.TextBox txtTotalYield_2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   25
         Left            =   1890
         Locked          =   -1  'True
         TabIndex        =   341
         Text            =   "100"
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
         Index           =   25
         Left            =   585
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   340
         Top             =   135
         Width           =   2085
      End
      Begin VB.TextBox txtProdYiled_1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   25
         Left            =   1035
         Locked          =   -1  'True
         TabIndex        =   339
         Text            =   "100"
         Top             =   900
         Width           =   735
      End
      Begin VB.TextBox txtPerctarget_1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   25
         Left            =   1035
         Locked          =   -1  'True
         TabIndex        =   338
         Text            =   "100"
         Top             =   1305
         Width           =   735
      End
      Begin VB.TextBox txtTotalYield_1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   25
         Left            =   1035
         Locked          =   -1  'True
         TabIndex        =   337
         Text            =   "100"
         Top             =   1710
         Width           =   735
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "R"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   125
         Left            =   1935
         TabIndex        =   349
         Top             =   675
         Width           =   645
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "L"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   126
         Left            =   1035
         TabIndex        =   348
         Top             =   675
         Width           =   645
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
         Left            =   45
         TabIndex        =   347
         Top             =   135
         Width           =   510
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "TRG YIELD"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   127
         Left            =   90
         TabIndex        =   346
         Top             =   990
         Width           =   960
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "% TARGET"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   128
         Left            =   90
         TabIndex        =   345
         Top             =   1395
         Width           =   870
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "PRD YIELD"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   129
         Left            =   90
         TabIndex        =   344
         Top             =   1800
         Width           =   870
      End
   End
   Begin VB.Frame frameDesign 
      BackColor       =   &H00FFFFFF&
      Height          =   2175
      Index           =   24
      Left            =   11790
      TabIndex        =   322
      Top             =   7695
      Width           =   2760
      Begin VB.TextBox txtProdYiled_2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   24
         Left            =   1890
         Locked          =   -1  'True
         TabIndex        =   329
         Text            =   "100"
         Top             =   900
         Width           =   735
      End
      Begin VB.TextBox txtPerctarget_2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   24
         Left            =   1890
         Locked          =   -1  'True
         TabIndex        =   328
         Text            =   "100"
         Top             =   1305
         Width           =   735
      End
      Begin VB.TextBox txtTotalYield_2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   24
         Left            =   1890
         Locked          =   -1  'True
         TabIndex        =   327
         Text            =   "100"
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
         Left            =   585
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   326
         Top             =   135
         Width           =   2085
      End
      Begin VB.TextBox txtProdYiled_1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   24
         Left            =   1035
         Locked          =   -1  'True
         TabIndex        =   325
         Text            =   "100"
         Top             =   900
         Width           =   735
      End
      Begin VB.TextBox txtPerctarget_1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   24
         Left            =   1035
         Locked          =   -1  'True
         TabIndex        =   324
         Text            =   "100"
         Top             =   1305
         Width           =   735
      End
      Begin VB.TextBox txtTotalYield_1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   24
         Left            =   1035
         Locked          =   -1  'True
         TabIndex        =   323
         Text            =   "100"
         Top             =   1710
         Width           =   735
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "R"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   120
         Left            =   1935
         TabIndex        =   335
         Top             =   675
         Width           =   645
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "L"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   121
         Left            =   1035
         TabIndex        =   334
         Top             =   675
         Width           =   645
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
         Left            =   45
         TabIndex        =   333
         Top             =   135
         Width           =   510
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "TRG YIELD"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   122
         Left            =   90
         TabIndex        =   332
         Top             =   990
         Width           =   960
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "% TARGET"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   123
         Left            =   90
         TabIndex        =   331
         Top             =   1395
         Width           =   870
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "PRD YIELD"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   124
         Left            =   90
         TabIndex        =   330
         Top             =   1800
         Width           =   870
      End
   End
   Begin VB.Frame frameDesign 
      BackColor       =   &H00FFFFFF&
      Height          =   2175
      Index           =   23
      Left            =   11790
      TabIndex        =   308
      Top             =   5355
      Width           =   2760
      Begin VB.TextBox txtProdYiled_2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   23
         Left            =   1890
         Locked          =   -1  'True
         TabIndex        =   315
         Text            =   "100"
         Top             =   900
         Width           =   735
      End
      Begin VB.TextBox txtPerctarget_2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   23
         Left            =   1890
         Locked          =   -1  'True
         TabIndex        =   314
         Text            =   "100"
         Top             =   1305
         Width           =   735
      End
      Begin VB.TextBox txtTotalYield_2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   23
         Left            =   1890
         Locked          =   -1  'True
         TabIndex        =   313
         Text            =   "100"
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
         Index           =   23
         Left            =   585
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   312
         Top             =   135
         Width           =   2085
      End
      Begin VB.TextBox txtProdYiled_1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   23
         Left            =   1035
         Locked          =   -1  'True
         TabIndex        =   311
         Text            =   "100"
         Top             =   900
         Width           =   735
      End
      Begin VB.TextBox txtPerctarget_1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   23
         Left            =   1035
         Locked          =   -1  'True
         TabIndex        =   310
         Text            =   "100"
         Top             =   1305
         Width           =   735
      End
      Begin VB.TextBox txtTotalYield_1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   23
         Left            =   1035
         Locked          =   -1  'True
         TabIndex        =   309
         Text            =   "100"
         Top             =   1710
         Width           =   735
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "R"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   115
         Left            =   1935
         TabIndex        =   321
         Top             =   675
         Width           =   645
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "L"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   116
         Left            =   1035
         TabIndex        =   320
         Top             =   675
         Width           =   645
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
         Left            =   45
         TabIndex        =   319
         Top             =   135
         Width           =   510
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "TRG YIELD"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   117
         Left            =   90
         TabIndex        =   318
         Top             =   990
         Width           =   960
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "% TARGET"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   118
         Left            =   90
         TabIndex        =   317
         Top             =   1395
         Width           =   870
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "PRD YIELD"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   119
         Left            =   90
         TabIndex        =   316
         Top             =   1800
         Width           =   870
      End
   End
   Begin VB.Frame frameDesign 
      BackColor       =   &H00FFFFFF&
      Height          =   2175
      Index           =   22
      Left            =   11790
      TabIndex        =   294
      Top             =   3015
      Width           =   2760
      Begin VB.TextBox txtProdYiled_2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   22
         Left            =   1890
         Locked          =   -1  'True
         TabIndex        =   301
         Text            =   "100"
         Top             =   900
         Width           =   735
      End
      Begin VB.TextBox txtPerctarget_2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   22
         Left            =   1890
         Locked          =   -1  'True
         TabIndex        =   300
         Text            =   "100"
         Top             =   1305
         Width           =   735
      End
      Begin VB.TextBox txtTotalYield_2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   22
         Left            =   1890
         Locked          =   -1  'True
         TabIndex        =   299
         Text            =   "100"
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
         Index           =   22
         Left            =   585
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   298
         Top             =   135
         Width           =   2085
      End
      Begin VB.TextBox txtProdYiled_1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   22
         Left            =   1035
         Locked          =   -1  'True
         TabIndex        =   297
         Text            =   "100"
         Top             =   900
         Width           =   735
      End
      Begin VB.TextBox txtPerctarget_1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   22
         Left            =   1035
         Locked          =   -1  'True
         TabIndex        =   296
         Text            =   "100"
         Top             =   1305
         Width           =   735
      End
      Begin VB.TextBox txtTotalYield_1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   22
         Left            =   1035
         Locked          =   -1  'True
         TabIndex        =   295
         Text            =   "100"
         Top             =   1710
         Width           =   735
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "R"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   110
         Left            =   1935
         TabIndex        =   307
         Top             =   675
         Width           =   645
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "L"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   111
         Left            =   1035
         TabIndex        =   306
         Top             =   675
         Width           =   645
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
         Left            =   45
         TabIndex        =   305
         Top             =   135
         Width           =   510
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "TRG YIELD"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   112
         Left            =   90
         TabIndex        =   304
         Top             =   990
         Width           =   960
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "% TARGET"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   113
         Left            =   90
         TabIndex        =   303
         Top             =   1395
         Width           =   870
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "PRD YIELD"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   114
         Left            =   90
         TabIndex        =   302
         Top             =   1800
         Width           =   870
      End
   End
   Begin VB.Frame frameDesign 
      BackColor       =   &H00FFFFFF&
      Height          =   2175
      Index           =   21
      Left            =   11790
      TabIndex        =   280
      Top             =   675
      Width           =   2760
      Begin VB.TextBox txtTotalYield_1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   21
         Left            =   1035
         Locked          =   -1  'True
         TabIndex        =   287
         Text            =   "100"
         Top             =   1710
         Width           =   735
      End
      Begin VB.TextBox txtPerctarget_1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   21
         Left            =   1035
         Locked          =   -1  'True
         TabIndex        =   286
         Text            =   "100"
         Top             =   1305
         Width           =   735
      End
      Begin VB.TextBox txtProdYiled_1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   21
         Left            =   1035
         Locked          =   -1  'True
         TabIndex        =   285
         Text            =   "100"
         Top             =   900
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
         Left            =   585
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   284
         Top             =   135
         Width           =   2085
      End
      Begin VB.TextBox txtTotalYield_2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   21
         Left            =   1890
         Locked          =   -1  'True
         TabIndex        =   283
         Text            =   "100"
         Top             =   1710
         Width           =   735
      End
      Begin VB.TextBox txtPerctarget_2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   21
         Left            =   1890
         Locked          =   -1  'True
         TabIndex        =   282
         Text            =   "100"
         Top             =   1305
         Width           =   735
      End
      Begin VB.TextBox txtProdYiled_2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   21
         Left            =   1890
         Locked          =   -1  'True
         TabIndex        =   281
         Text            =   "100"
         Top             =   900
         Width           =   735
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "PRD YIELD"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   105
         Left            =   90
         TabIndex        =   293
         Top             =   1800
         Width           =   870
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "% TARGET"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   106
         Left            =   90
         TabIndex        =   292
         Top             =   1395
         Width           =   870
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "TRG YIELD"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   107
         Left            =   90
         TabIndex        =   291
         Top             =   990
         Width           =   960
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
         Left            =   45
         TabIndex        =   290
         Top             =   135
         Width           =   510
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "L"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   108
         Left            =   1035
         TabIndex        =   289
         Top             =   675
         Width           =   645
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "R"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   109
         Left            =   1935
         TabIndex        =   288
         Top             =   675
         Width           =   645
      End
   End
   Begin VB.Frame frameDesign 
      BackColor       =   &H00FFFFFF&
      Height          =   2175
      Index           =   20
      Left            =   8865
      TabIndex        =   266
      Top             =   10035
      Width           =   2760
      Begin VB.TextBox txtProdYiled_2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   20
         Left            =   1890
         Locked          =   -1  'True
         TabIndex        =   273
         Text            =   "100"
         Top             =   900
         Width           =   735
      End
      Begin VB.TextBox txtPerctarget_2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   20
         Left            =   1890
         Locked          =   -1  'True
         TabIndex        =   272
         Text            =   "100"
         Top             =   1305
         Width           =   735
      End
      Begin VB.TextBox txtTotalYield_2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   20
         Left            =   1890
         Locked          =   -1  'True
         TabIndex        =   271
         Text            =   "100"
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
         Index           =   20
         Left            =   585
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   270
         Top             =   135
         Width           =   2085
      End
      Begin VB.TextBox txtProdYiled_1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   20
         Left            =   1035
         Locked          =   -1  'True
         TabIndex        =   269
         Text            =   "100"
         Top             =   900
         Width           =   735
      End
      Begin VB.TextBox txtPerctarget_1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   20
         Left            =   1035
         Locked          =   -1  'True
         TabIndex        =   268
         Text            =   "100"
         Top             =   1305
         Width           =   735
      End
      Begin VB.TextBox txtTotalYield_1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   20
         Left            =   1035
         Locked          =   -1  'True
         TabIndex        =   267
         Text            =   "100"
         Top             =   1710
         Width           =   735
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "R"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   100
         Left            =   1935
         TabIndex        =   279
         Top             =   675
         Width           =   645
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "L"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   101
         Left            =   1035
         TabIndex        =   278
         Top             =   675
         Width           =   645
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
         Left            =   45
         TabIndex        =   277
         Top             =   135
         Width           =   510
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "TRG YIELD"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   102
         Left            =   90
         TabIndex        =   276
         Top             =   990
         Width           =   960
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "% TARGET"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   103
         Left            =   90
         TabIndex        =   275
         Top             =   1395
         Width           =   870
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "PRD YIELD"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   104
         Left            =   90
         TabIndex        =   274
         Top             =   1800
         Width           =   870
      End
   End
   Begin VB.Frame frameDesign 
      BackColor       =   &H00FFFFFF&
      Height          =   2175
      Index           =   19
      Left            =   8865
      TabIndex        =   252
      Top             =   7695
      Width           =   2760
      Begin VB.TextBox txtProdYiled_2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   19
         Left            =   1890
         Locked          =   -1  'True
         TabIndex        =   259
         Text            =   "100"
         Top             =   900
         Width           =   735
      End
      Begin VB.TextBox txtPerctarget_2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   19
         Left            =   1890
         Locked          =   -1  'True
         TabIndex        =   258
         Text            =   "100"
         Top             =   1305
         Width           =   735
      End
      Begin VB.TextBox txtTotalYield_2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   19
         Left            =   1890
         Locked          =   -1  'True
         TabIndex        =   257
         Text            =   "100"
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
         Left            =   585
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   256
         Top             =   135
         Width           =   2085
      End
      Begin VB.TextBox txtProdYiled_1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   19
         Left            =   1035
         Locked          =   -1  'True
         TabIndex        =   255
         Text            =   "100"
         Top             =   900
         Width           =   735
      End
      Begin VB.TextBox txtPerctarget_1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   19
         Left            =   1035
         Locked          =   -1  'True
         TabIndex        =   254
         Text            =   "100"
         Top             =   1305
         Width           =   735
      End
      Begin VB.TextBox txtTotalYield_1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   19
         Left            =   1035
         Locked          =   -1  'True
         TabIndex        =   253
         Text            =   "100"
         Top             =   1710
         Width           =   735
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "R"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   95
         Left            =   1935
         TabIndex        =   265
         Top             =   675
         Width           =   645
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "L"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   96
         Left            =   1035
         TabIndex        =   264
         Top             =   675
         Width           =   645
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
         Left            =   45
         TabIndex        =   263
         Top             =   135
         Width           =   510
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "TRG YIELD"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   97
         Left            =   90
         TabIndex        =   262
         Top             =   990
         Width           =   960
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "% TARGET"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   98
         Left            =   90
         TabIndex        =   261
         Top             =   1395
         Width           =   870
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "PRD YIELD"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   99
         Left            =   90
         TabIndex        =   260
         Top             =   1800
         Width           =   870
      End
   End
   Begin VB.Frame frameDesign 
      BackColor       =   &H00FFFFFF&
      Height          =   2175
      Index           =   18
      Left            =   8865
      TabIndex        =   238
      Top             =   5355
      Width           =   2760
      Begin VB.TextBox txtProdYiled_2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   18
         Left            =   1890
         Locked          =   -1  'True
         TabIndex        =   245
         Text            =   "100"
         Top             =   900
         Width           =   735
      End
      Begin VB.TextBox txtPerctarget_2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   18
         Left            =   1890
         Locked          =   -1  'True
         TabIndex        =   244
         Text            =   "100"
         Top             =   1305
         Width           =   735
      End
      Begin VB.TextBox txtTotalYield_2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   18
         Left            =   1890
         Locked          =   -1  'True
         TabIndex        =   243
         Text            =   "100"
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
         Index           =   18
         Left            =   585
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   242
         Top             =   135
         Width           =   2085
      End
      Begin VB.TextBox txtProdYiled_1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   18
         Left            =   1035
         Locked          =   -1  'True
         TabIndex        =   241
         Text            =   "100"
         Top             =   900
         Width           =   735
      End
      Begin VB.TextBox txtPerctarget_1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   18
         Left            =   1035
         Locked          =   -1  'True
         TabIndex        =   240
         Text            =   "100"
         Top             =   1305
         Width           =   735
      End
      Begin VB.TextBox txtTotalYield_1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   18
         Left            =   1035
         Locked          =   -1  'True
         TabIndex        =   239
         Text            =   "100"
         Top             =   1710
         Width           =   735
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "R"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   90
         Left            =   1935
         TabIndex        =   251
         Top             =   675
         Width           =   645
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "L"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   91
         Left            =   1035
         TabIndex        =   250
         Top             =   675
         Width           =   645
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
         Left            =   45
         TabIndex        =   249
         Top             =   135
         Width           =   510
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "TRG YIELD"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   92
         Left            =   90
         TabIndex        =   248
         Top             =   990
         Width           =   960
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "% TARGET"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   93
         Left            =   90
         TabIndex        =   247
         Top             =   1395
         Width           =   870
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "PRD YIELD"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   94
         Left            =   90
         TabIndex        =   246
         Top             =   1800
         Width           =   870
      End
   End
   Begin VB.Frame frameDesign 
      BackColor       =   &H00FFFFFF&
      Height          =   2175
      Index           =   17
      Left            =   8865
      TabIndex        =   224
      Top             =   3015
      Width           =   2760
      Begin VB.TextBox txtProdYiled_2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   17
         Left            =   1890
         Locked          =   -1  'True
         TabIndex        =   231
         Text            =   "100"
         Top             =   900
         Width           =   735
      End
      Begin VB.TextBox txtPerctarget_2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   17
         Left            =   1890
         Locked          =   -1  'True
         TabIndex        =   230
         Text            =   "100"
         Top             =   1305
         Width           =   735
      End
      Begin VB.TextBox txtTotalYield_2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   17
         Left            =   1890
         Locked          =   -1  'True
         TabIndex        =   229
         Text            =   "100"
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
         Index           =   17
         Left            =   585
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   228
         Top             =   135
         Width           =   2085
      End
      Begin VB.TextBox txtProdYiled_1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   17
         Left            =   1035
         Locked          =   -1  'True
         TabIndex        =   227
         Text            =   "100"
         Top             =   900
         Width           =   735
      End
      Begin VB.TextBox txtPerctarget_1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   17
         Left            =   1035
         Locked          =   -1  'True
         TabIndex        =   226
         Text            =   "100"
         Top             =   1305
         Width           =   735
      End
      Begin VB.TextBox txtTotalYield_1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   17
         Left            =   1035
         Locked          =   -1  'True
         TabIndex        =   225
         Text            =   "100"
         Top             =   1710
         Width           =   735
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "R"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   85
         Left            =   1935
         TabIndex        =   237
         Top             =   675
         Width           =   645
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "L"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   86
         Left            =   1035
         TabIndex        =   236
         Top             =   675
         Width           =   645
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
         Left            =   45
         TabIndex        =   235
         Top             =   135
         Width           =   510
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "TRG YIELD"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   87
         Left            =   90
         TabIndex        =   234
         Top             =   990
         Width           =   960
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "% TARGET"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   88
         Left            =   90
         TabIndex        =   233
         Top             =   1395
         Width           =   870
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "PRD YIELD"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   89
         Left            =   90
         TabIndex        =   232
         Top             =   1800
         Width           =   870
      End
   End
   Begin VB.Frame frameDesign 
      BackColor       =   &H00FFFFFF&
      Height          =   2175
      Index           =   16
      Left            =   8865
      TabIndex        =   210
      Top             =   675
      Width           =   2760
      Begin VB.TextBox txtTotalYield_1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   16
         Left            =   1035
         Locked          =   -1  'True
         TabIndex        =   217
         Text            =   "100"
         Top             =   1710
         Width           =   735
      End
      Begin VB.TextBox txtPerctarget_1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   16
         Left            =   1035
         Locked          =   -1  'True
         TabIndex        =   216
         Text            =   "100"
         Top             =   1305
         Width           =   735
      End
      Begin VB.TextBox txtProdYiled_1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   16
         Left            =   1035
         Locked          =   -1  'True
         TabIndex        =   215
         Text            =   "100"
         Top             =   900
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
         Left            =   585
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   214
         Top             =   135
         Width           =   2085
      End
      Begin VB.TextBox txtTotalYield_2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   16
         Left            =   1890
         Locked          =   -1  'True
         TabIndex        =   213
         Text            =   "100"
         Top             =   1710
         Width           =   735
      End
      Begin VB.TextBox txtPerctarget_2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   16
         Left            =   1890
         Locked          =   -1  'True
         TabIndex        =   212
         Text            =   "100"
         Top             =   1305
         Width           =   735
      End
      Begin VB.TextBox txtProdYiled_2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   16
         Left            =   1890
         Locked          =   -1  'True
         TabIndex        =   211
         Text            =   "100"
         Top             =   900
         Width           =   735
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "PRD YIELD"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   80
         Left            =   90
         TabIndex        =   223
         Top             =   1800
         Width           =   870
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "% TARGET"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   81
         Left            =   90
         TabIndex        =   222
         Top             =   1395
         Width           =   870
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "TRG YIELD"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   82
         Left            =   90
         TabIndex        =   221
         Top             =   990
         Width           =   960
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
         Left            =   45
         TabIndex        =   220
         Top             =   135
         Width           =   510
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "L"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   83
         Left            =   1035
         TabIndex        =   219
         Top             =   675
         Width           =   645
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "R"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   84
         Left            =   1935
         TabIndex        =   218
         Top             =   675
         Width           =   645
      End
   End
   Begin VB.Frame frameDesign 
      BackColor       =   &H00FFFFFF&
      Height          =   2175
      Index           =   15
      Left            =   5940
      TabIndex        =   196
      Top             =   10035
      Width           =   2760
      Begin VB.TextBox txtTotalYield_1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   15
         Left            =   1035
         Locked          =   -1  'True
         TabIndex        =   203
         Text            =   "100"
         Top             =   1710
         Width           =   735
      End
      Begin VB.TextBox txtPerctarget_1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   15
         Left            =   1035
         Locked          =   -1  'True
         TabIndex        =   202
         Text            =   "100"
         Top             =   1305
         Width           =   735
      End
      Begin VB.TextBox txtProdYiled_1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   15
         Left            =   1035
         Locked          =   -1  'True
         TabIndex        =   201
         Text            =   "100"
         Top             =   900
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
         Index           =   15
         Left            =   585
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   200
         Top             =   135
         Width           =   2085
      End
      Begin VB.TextBox txtTotalYield_2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   15
         Left            =   1890
         Locked          =   -1  'True
         TabIndex        =   199
         Text            =   "100"
         Top             =   1710
         Width           =   735
      End
      Begin VB.TextBox txtPerctarget_2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   15
         Left            =   1890
         Locked          =   -1  'True
         TabIndex        =   198
         Text            =   "100"
         Top             =   1305
         Width           =   735
      End
      Begin VB.TextBox txtProdYiled_2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   15
         Left            =   1890
         Locked          =   -1  'True
         TabIndex        =   197
         Text            =   "100"
         Top             =   900
         Width           =   735
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "PRD YIELD"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   75
         Left            =   90
         TabIndex        =   209
         Top             =   1800
         Width           =   870
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "% TARGET"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   76
         Left            =   90
         TabIndex        =   208
         Top             =   1395
         Width           =   870
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "TRG YIELD"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   77
         Left            =   90
         TabIndex        =   207
         Top             =   990
         Width           =   960
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
         Left            =   45
         TabIndex        =   206
         Top             =   135
         Width           =   510
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "L"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   78
         Left            =   1035
         TabIndex        =   205
         Top             =   675
         Width           =   645
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "R"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   79
         Left            =   1935
         TabIndex        =   204
         Top             =   675
         Width           =   645
      End
   End
   Begin VB.Frame frameDesign 
      BackColor       =   &H00FFFFFF&
      Height          =   2175
      Index           =   14
      Left            =   5940
      TabIndex        =   182
      Top             =   7695
      Width           =   2760
      Begin VB.TextBox txtTotalYield_1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   14
         Left            =   1035
         Locked          =   -1  'True
         TabIndex        =   189
         Text            =   "100"
         Top             =   1710
         Width           =   735
      End
      Begin VB.TextBox txtPerctarget_1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   14
         Left            =   1035
         Locked          =   -1  'True
         TabIndex        =   188
         Text            =   "100"
         Top             =   1305
         Width           =   735
      End
      Begin VB.TextBox txtProdYiled_1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   14
         Left            =   1035
         Locked          =   -1  'True
         TabIndex        =   187
         Text            =   "100"
         Top             =   900
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
         Left            =   585
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   186
         Top             =   135
         Width           =   2085
      End
      Begin VB.TextBox txtTotalYield_2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   14
         Left            =   1890
         Locked          =   -1  'True
         TabIndex        =   185
         Text            =   "100"
         Top             =   1710
         Width           =   735
      End
      Begin VB.TextBox txtPerctarget_2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   14
         Left            =   1890
         Locked          =   -1  'True
         TabIndex        =   184
         Text            =   "100"
         Top             =   1305
         Width           =   735
      End
      Begin VB.TextBox txtProdYiled_2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   14
         Left            =   1890
         Locked          =   -1  'True
         TabIndex        =   183
         Text            =   "100"
         Top             =   900
         Width           =   735
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "PRD YIELD"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   70
         Left            =   90
         TabIndex        =   195
         Top             =   1800
         Width           =   870
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "% TARGET"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   71
         Left            =   90
         TabIndex        =   194
         Top             =   1395
         Width           =   870
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "TRG YIELD"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   72
         Left            =   90
         TabIndex        =   193
         Top             =   990
         Width           =   960
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
         Left            =   45
         TabIndex        =   192
         Top             =   135
         Width           =   510
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "L"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   73
         Left            =   1035
         TabIndex        =   191
         Top             =   675
         Width           =   645
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "R"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   74
         Left            =   1935
         TabIndex        =   190
         Top             =   675
         Width           =   645
      End
   End
   Begin VB.Frame frameDesign 
      BackColor       =   &H00FFFFFF&
      Height          =   2175
      Index           =   13
      Left            =   5940
      TabIndex        =   168
      Top             =   5355
      Width           =   2760
      Begin VB.TextBox txtTotalYield_1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   13
         Left            =   1035
         Locked          =   -1  'True
         TabIndex        =   175
         Text            =   "100"
         Top             =   1710
         Width           =   735
      End
      Begin VB.TextBox txtPerctarget_1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   13
         Left            =   1035
         Locked          =   -1  'True
         TabIndex        =   174
         Text            =   "100"
         Top             =   1305
         Width           =   735
      End
      Begin VB.TextBox txtProdYiled_1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   13
         Left            =   1035
         Locked          =   -1  'True
         TabIndex        =   173
         Text            =   "100"
         Top             =   900
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
         Index           =   13
         Left            =   585
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   172
         Top             =   135
         Width           =   2085
      End
      Begin VB.TextBox txtTotalYield_2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   13
         Left            =   1890
         Locked          =   -1  'True
         TabIndex        =   171
         Text            =   "100"
         Top             =   1710
         Width           =   735
      End
      Begin VB.TextBox txtPerctarget_2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   13
         Left            =   1890
         Locked          =   -1  'True
         TabIndex        =   170
         Text            =   "100"
         Top             =   1305
         Width           =   735
      End
      Begin VB.TextBox txtProdYiled_2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   13
         Left            =   1890
         Locked          =   -1  'True
         TabIndex        =   169
         Text            =   "100"
         Top             =   900
         Width           =   735
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "PRD YIELD"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   65
         Left            =   90
         TabIndex        =   181
         Top             =   1800
         Width           =   870
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "% TARGET"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   66
         Left            =   90
         TabIndex        =   180
         Top             =   1395
         Width           =   870
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "TRG YIELD"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   67
         Left            =   90
         TabIndex        =   179
         Top             =   990
         Width           =   960
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
         Left            =   45
         TabIndex        =   178
         Top             =   135
         Width           =   510
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "L"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   68
         Left            =   1035
         TabIndex        =   177
         Top             =   675
         Width           =   645
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "R"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   69
         Left            =   1935
         TabIndex        =   176
         Top             =   675
         Width           =   645
      End
   End
   Begin VB.Frame frameDesign 
      BackColor       =   &H00FFFFFF&
      Height          =   2175
      Index           =   12
      Left            =   5940
      TabIndex        =   154
      Top             =   3015
      Width           =   2760
      Begin VB.TextBox txtTotalYield_1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   12
         Left            =   1035
         Locked          =   -1  'True
         TabIndex        =   161
         Text            =   "100"
         Top             =   1710
         Width           =   735
      End
      Begin VB.TextBox txtPerctarget_1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   12
         Left            =   1035
         Locked          =   -1  'True
         TabIndex        =   160
         Text            =   "100"
         Top             =   1305
         Width           =   735
      End
      Begin VB.TextBox txtProdYiled_1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   12
         Left            =   1035
         Locked          =   -1  'True
         TabIndex        =   159
         Text            =   "100"
         Top             =   900
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
         Index           =   12
         Left            =   585
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   158
         Top             =   135
         Width           =   2085
      End
      Begin VB.TextBox txtTotalYield_2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   12
         Left            =   1890
         Locked          =   -1  'True
         TabIndex        =   157
         Text            =   "100"
         Top             =   1710
         Width           =   735
      End
      Begin VB.TextBox txtPerctarget_2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   12
         Left            =   1890
         Locked          =   -1  'True
         TabIndex        =   156
         Text            =   "100"
         Top             =   1305
         Width           =   735
      End
      Begin VB.TextBox txtProdYiled_2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   12
         Left            =   1890
         Locked          =   -1  'True
         TabIndex        =   155
         Text            =   "100"
         Top             =   900
         Width           =   735
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "PRD YIELD"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   60
         Left            =   90
         TabIndex        =   167
         Top             =   1800
         Width           =   870
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "% TARGET"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   61
         Left            =   90
         TabIndex        =   166
         Top             =   1395
         Width           =   870
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "TRG YIELD"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   62
         Left            =   90
         TabIndex        =   165
         Top             =   990
         Width           =   960
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
         Left            =   45
         TabIndex        =   164
         Top             =   135
         Width           =   510
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "L"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   63
         Left            =   1035
         TabIndex        =   163
         Top             =   675
         Width           =   645
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "R"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   64
         Left            =   1935
         TabIndex        =   162
         Top             =   675
         Width           =   645
      End
   End
   Begin VB.Frame frameDesign 
      BackColor       =   &H00FFFFFF&
      Height          =   2175
      Index           =   11
      Left            =   5940
      TabIndex        =   140
      Top             =   675
      Width           =   2760
      Begin VB.TextBox txtProdYiled_2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   11
         Left            =   1890
         Locked          =   -1  'True
         TabIndex        =   147
         Text            =   "100"
         Top             =   900
         Width           =   735
      End
      Begin VB.TextBox txtPerctarget_2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   11
         Left            =   1890
         Locked          =   -1  'True
         TabIndex        =   146
         Text            =   "100"
         Top             =   1305
         Width           =   735
      End
      Begin VB.TextBox txtTotalYield_2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   11
         Left            =   1890
         Locked          =   -1  'True
         TabIndex        =   145
         Text            =   "100"
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
         Left            =   585
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   144
         Top             =   135
         Width           =   2085
      End
      Begin VB.TextBox txtProdYiled_1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   11
         Left            =   1035
         Locked          =   -1  'True
         TabIndex        =   143
         Text            =   "100"
         Top             =   900
         Width           =   735
      End
      Begin VB.TextBox txtPerctarget_1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   11
         Left            =   1035
         Locked          =   -1  'True
         TabIndex        =   142
         Text            =   "100"
         Top             =   1305
         Width           =   735
      End
      Begin VB.TextBox txtTotalYield_1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   11
         Left            =   1035
         Locked          =   -1  'True
         TabIndex        =   141
         Text            =   "100"
         Top             =   1710
         Width           =   735
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "R"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   55
         Left            =   1935
         TabIndex        =   153
         Top             =   675
         Width           =   645
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "L"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   56
         Left            =   1035
         TabIndex        =   152
         Top             =   675
         Width           =   645
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
         Left            =   45
         TabIndex        =   151
         Top             =   135
         Width           =   510
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "TRG YIELD"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   57
         Left            =   90
         TabIndex        =   150
         Top             =   990
         Width           =   960
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "% TARGET"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   58
         Left            =   90
         TabIndex        =   149
         Top             =   1395
         Width           =   870
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "PRD YIELD"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   59
         Left            =   90
         TabIndex        =   148
         Top             =   1800
         Width           =   870
      End
   End
   Begin VB.Frame frameDesign 
      BackColor       =   &H00FFFFFF&
      Height          =   2175
      Index           =   10
      Left            =   3015
      TabIndex        =   126
      Top             =   10035
      Width           =   2760
      Begin VB.TextBox txtTotalYield_1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   10
         Left            =   1035
         Locked          =   -1  'True
         TabIndex        =   133
         Text            =   "100"
         Top             =   1710
         Width           =   735
      End
      Begin VB.TextBox txtPerctarget_1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   10
         Left            =   1035
         Locked          =   -1  'True
         TabIndex        =   132
         Text            =   "100"
         Top             =   1305
         Width           =   735
      End
      Begin VB.TextBox txtProdYiled_1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   10
         Left            =   1035
         Locked          =   -1  'True
         TabIndex        =   131
         Text            =   "100"
         Top             =   900
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
         Index           =   10
         Left            =   585
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   130
         Top             =   135
         Width           =   2085
      End
      Begin VB.TextBox txtTotalYield_2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   10
         Left            =   1890
         Locked          =   -1  'True
         TabIndex        =   129
         Text            =   "100"
         Top             =   1710
         Width           =   735
      End
      Begin VB.TextBox txtPerctarget_2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   10
         Left            =   1890
         Locked          =   -1  'True
         TabIndex        =   128
         Text            =   "100"
         Top             =   1305
         Width           =   735
      End
      Begin VB.TextBox txtProdYiled_2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   10
         Left            =   1890
         Locked          =   -1  'True
         TabIndex        =   127
         Text            =   "100"
         Top             =   900
         Width           =   735
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "PRD YIELD"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   50
         Left            =   90
         TabIndex        =   139
         Top             =   1800
         Width           =   870
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "% TARGET"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   51
         Left            =   90
         TabIndex        =   138
         Top             =   1395
         Width           =   870
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "TRG YIELD"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   52
         Left            =   90
         TabIndex        =   137
         Top             =   990
         Width           =   960
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
         Left            =   45
         TabIndex        =   136
         Top             =   135
         Width           =   510
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "L"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   53
         Left            =   1035
         TabIndex        =   135
         Top             =   675
         Width           =   645
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "R"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   54
         Left            =   1935
         TabIndex        =   134
         Top             =   675
         Width           =   645
      End
   End
   Begin VB.Frame frameDesign 
      BackColor       =   &H00FFFFFF&
      Height          =   2175
      Index           =   9
      Left            =   3015
      TabIndex        =   112
      Top             =   7695
      Width           =   2760
      Begin VB.TextBox txtTotalYield_1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   9
         Left            =   1035
         Locked          =   -1  'True
         TabIndex        =   119
         Text            =   "100"
         Top             =   1710
         Width           =   735
      End
      Begin VB.TextBox txtPerctarget_1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   9
         Left            =   1035
         Locked          =   -1  'True
         TabIndex        =   118
         Text            =   "100"
         Top             =   1305
         Width           =   735
      End
      Begin VB.TextBox txtProdYiled_1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   9
         Left            =   1035
         Locked          =   -1  'True
         TabIndex        =   117
         Text            =   "100"
         Top             =   900
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
         Index           =   9
         Left            =   585
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   116
         Top             =   135
         Width           =   2085
      End
      Begin VB.TextBox txtTotalYield_2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   9
         Left            =   1890
         Locked          =   -1  'True
         TabIndex        =   115
         Text            =   "100"
         Top             =   1710
         Width           =   735
      End
      Begin VB.TextBox txtPerctarget_2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   9
         Left            =   1890
         Locked          =   -1  'True
         TabIndex        =   114
         Text            =   "100"
         Top             =   1305
         Width           =   735
      End
      Begin VB.TextBox txtProdYiled_2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   9
         Left            =   1890
         Locked          =   -1  'True
         TabIndex        =   113
         Text            =   "100"
         Top             =   900
         Width           =   735
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "PRD YIELD"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   45
         Left            =   90
         TabIndex        =   125
         Top             =   1800
         Width           =   870
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "% TARGET"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   46
         Left            =   90
         TabIndex        =   124
         Top             =   1395
         Width           =   870
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "TRG YIELD"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   47
         Left            =   90
         TabIndex        =   123
         Top             =   990
         Width           =   960
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
         Left            =   45
         TabIndex        =   122
         Top             =   135
         Width           =   510
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "L"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   48
         Left            =   1035
         TabIndex        =   121
         Top             =   675
         Width           =   645
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "R"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   49
         Left            =   1935
         TabIndex        =   120
         Top             =   675
         Width           =   645
      End
   End
   Begin VB.Frame frameDesign 
      BackColor       =   &H00FFFFFF&
      Height          =   2175
      Index           =   8
      Left            =   3015
      TabIndex        =   98
      Top             =   5355
      Width           =   2760
      Begin VB.TextBox txtTotalYield_1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   8
         Left            =   1035
         Locked          =   -1  'True
         TabIndex        =   105
         Text            =   "100"
         Top             =   1710
         Width           =   735
      End
      Begin VB.TextBox txtPerctarget_1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   8
         Left            =   1035
         Locked          =   -1  'True
         TabIndex        =   104
         Text            =   "100"
         Top             =   1305
         Width           =   735
      End
      Begin VB.TextBox txtProdYiled_1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   8
         Left            =   1035
         Locked          =   -1  'True
         TabIndex        =   103
         Text            =   "100"
         Top             =   900
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
         Left            =   585
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   102
         Top             =   135
         Width           =   2085
      End
      Begin VB.TextBox txtTotalYield_2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   8
         Left            =   1890
         Locked          =   -1  'True
         TabIndex        =   101
         Text            =   "100"
         Top             =   1710
         Width           =   735
      End
      Begin VB.TextBox txtPerctarget_2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   8
         Left            =   1890
         Locked          =   -1  'True
         TabIndex        =   100
         Text            =   "100"
         Top             =   1305
         Width           =   735
      End
      Begin VB.TextBox txtProdYiled_2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   8
         Left            =   1890
         Locked          =   -1  'True
         TabIndex        =   99
         Text            =   "100"
         Top             =   900
         Width           =   735
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "PRD YIELD"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   40
         Left            =   90
         TabIndex        =   111
         Top             =   1800
         Width           =   870
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "% TARGET"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   41
         Left            =   90
         TabIndex        =   110
         Top             =   1395
         Width           =   870
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "TRG YIELD"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   42
         Left            =   90
         TabIndex        =   109
         Top             =   990
         Width           =   960
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
         Left            =   45
         TabIndex        =   108
         Top             =   135
         Width           =   510
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "L"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   43
         Left            =   1035
         TabIndex        =   107
         Top             =   675
         Width           =   645
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "R"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   44
         Left            =   1935
         TabIndex        =   106
         Top             =   675
         Width           =   645
      End
   End
   Begin VB.Frame frameDesign 
      BackColor       =   &H00FFFFFF&
      Height          =   2175
      Index           =   7
      Left            =   3015
      TabIndex        =   84
      Top             =   3015
      Width           =   2760
      Begin VB.TextBox txtTotalYield_1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   7
         Left            =   1035
         Locked          =   -1  'True
         TabIndex        =   91
         Text            =   "100"
         Top             =   1710
         Width           =   735
      End
      Begin VB.TextBox txtPerctarget_1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   7
         Left            =   1035
         Locked          =   -1  'True
         TabIndex        =   90
         Text            =   "100"
         Top             =   1305
         Width           =   735
      End
      Begin VB.TextBox txtProdYiled_1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   7
         Left            =   1035
         Locked          =   -1  'True
         TabIndex        =   89
         Text            =   "100"
         Top             =   900
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
         Left            =   585
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   88
         Top             =   135
         Width           =   2085
      End
      Begin VB.TextBox txtTotalYield_2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   7
         Left            =   1890
         Locked          =   -1  'True
         TabIndex        =   87
         Text            =   "100"
         Top             =   1710
         Width           =   735
      End
      Begin VB.TextBox txtPerctarget_2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   7
         Left            =   1890
         Locked          =   -1  'True
         TabIndex        =   86
         Text            =   "100"
         Top             =   1305
         Width           =   735
      End
      Begin VB.TextBox txtProdYiled_2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   7
         Left            =   1890
         Locked          =   -1  'True
         TabIndex        =   85
         Text            =   "100"
         Top             =   900
         Width           =   735
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "PRD YIELD"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   35
         Left            =   90
         TabIndex        =   97
         Top             =   1800
         Width           =   870
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "% TARGET"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   36
         Left            =   90
         TabIndex        =   96
         Top             =   1395
         Width           =   870
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "TRG YIELD"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   37
         Left            =   90
         TabIndex        =   95
         Top             =   990
         Width           =   960
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
         Left            =   45
         TabIndex        =   94
         Top             =   135
         Width           =   510
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "L"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   38
         Left            =   1035
         TabIndex        =   93
         Top             =   675
         Width           =   645
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "R"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   39
         Left            =   1935
         TabIndex        =   92
         Top             =   675
         Width           =   645
      End
   End
   Begin VB.Frame frameDesign 
      BackColor       =   &H00FFFFFF&
      Height          =   2175
      Index           =   6
      Left            =   3015
      TabIndex        =   70
      Top             =   675
      Width           =   2760
      Begin VB.TextBox txtProdYiled_2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   6
         Left            =   1890
         Locked          =   -1  'True
         TabIndex        =   77
         Text            =   "100"
         Top             =   900
         Width           =   735
      End
      Begin VB.TextBox txtPerctarget_2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   6
         Left            =   1890
         Locked          =   -1  'True
         TabIndex        =   76
         Text            =   "100"
         Top             =   1305
         Width           =   735
      End
      Begin VB.TextBox txtTotalYield_2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   6
         Left            =   1890
         Locked          =   -1  'True
         TabIndex        =   75
         Text            =   "100"
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
         Index           =   6
         Left            =   585
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   74
         Top             =   135
         Width           =   2085
      End
      Begin VB.TextBox txtProdYiled_1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   6
         Left            =   1035
         Locked          =   -1  'True
         TabIndex        =   73
         Text            =   "100"
         Top             =   900
         Width           =   735
      End
      Begin VB.TextBox txtPerctarget_1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   6
         Left            =   1035
         Locked          =   -1  'True
         TabIndex        =   72
         Text            =   "100"
         Top             =   1305
         Width           =   735
      End
      Begin VB.TextBox txtTotalYield_1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   6
         Left            =   1035
         Locked          =   -1  'True
         TabIndex        =   71
         Text            =   "100"
         Top             =   1710
         Width           =   735
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "R"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   30
         Left            =   1935
         TabIndex        =   83
         Top             =   675
         Width           =   645
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "L"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   31
         Left            =   1035
         TabIndex        =   82
         Top             =   675
         Width           =   645
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
         Left            =   45
         TabIndex        =   81
         Top             =   135
         Width           =   510
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "TRG YIELD"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   32
         Left            =   90
         TabIndex        =   80
         Top             =   990
         Width           =   960
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "% TARGET"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   33
         Left            =   90
         TabIndex        =   79
         Top             =   1395
         Width           =   870
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "PRD YIELD"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   34
         Left            =   90
         TabIndex        =   78
         Top             =   1800
         Width           =   870
      End
   End
   Begin VB.Frame frameDesign 
      BackColor       =   &H00FFFFFF&
      Height          =   2175
      Index           =   5
      Left            =   90
      TabIndex        =   56
      Top             =   10035
      Width           =   2760
      Begin VB.TextBox txtProdYiled_2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   5
         Left            =   1890
         Locked          =   -1  'True
         TabIndex        =   63
         Text            =   "0"
         Top             =   900
         Width           =   735
      End
      Begin VB.TextBox txtPerctarget_2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   5
         Left            =   1890
         Locked          =   -1  'True
         TabIndex        =   62
         Text            =   "0"
         Top             =   1305
         Width           =   735
      End
      Begin VB.TextBox txtTotalYield_2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   5
         Left            =   1890
         Locked          =   -1  'True
         TabIndex        =   61
         Text            =   "0"
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
         Left            =   585
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   60
         Top             =   135
         Width           =   2085
      End
      Begin VB.TextBox txtProdYiled_1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   5
         Left            =   1035
         Locked          =   -1  'True
         TabIndex        =   59
         Text            =   "0"
         Top             =   900
         Width           =   735
      End
      Begin VB.TextBox txtPerctarget_1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   5
         Left            =   1035
         Locked          =   -1  'True
         TabIndex        =   58
         Text            =   "0"
         Top             =   1305
         Width           =   735
      End
      Begin VB.TextBox txtTotalYield_1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   5
         Left            =   1035
         Locked          =   -1  'True
         TabIndex        =   57
         Text            =   "0"
         Top             =   1710
         Width           =   735
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "R"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   25
         Left            =   1935
         TabIndex        =   69
         Top             =   675
         Width           =   645
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "L"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   26
         Left            =   1035
         TabIndex        =   68
         Top             =   675
         Width           =   645
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
         Left            =   45
         TabIndex        =   67
         Top             =   135
         Width           =   510
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "TRG YIELD"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   27
         Left            =   90
         TabIndex        =   66
         Top             =   990
         Width           =   960
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "% TARGET"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   28
         Left            =   90
         TabIndex        =   65
         Top             =   1395
         Width           =   870
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "PRD YIELD"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   29
         Left            =   90
         TabIndex        =   64
         Top             =   1800
         Width           =   870
      End
   End
   Begin VB.Frame frameDesign 
      BackColor       =   &H00FFFFFF&
      Height          =   2175
      Index           =   4
      Left            =   90
      TabIndex        =   42
      Top             =   7695
      Width           =   2760
      Begin VB.TextBox txtProdYiled_2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   4
         Left            =   1890
         Locked          =   -1  'True
         TabIndex        =   49
         Text            =   "0"
         Top             =   900
         Width           =   735
      End
      Begin VB.TextBox txtPerctarget_2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   4
         Left            =   1890
         Locked          =   -1  'True
         TabIndex        =   48
         Text            =   "0"
         Top             =   1305
         Width           =   735
      End
      Begin VB.TextBox txtTotalYield_2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   4
         Left            =   1890
         Locked          =   -1  'True
         TabIndex        =   47
         Text            =   "0"
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
         Index           =   4
         Left            =   585
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   46
         Top             =   135
         Width           =   2085
      End
      Begin VB.TextBox txtProdYiled_1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   4
         Left            =   1035
         Locked          =   -1  'True
         TabIndex        =   45
         Text            =   "0"
         Top             =   900
         Width           =   735
      End
      Begin VB.TextBox txtPerctarget_1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   4
         Left            =   1035
         Locked          =   -1  'True
         TabIndex        =   44
         Text            =   "0"
         Top             =   1305
         Width           =   735
      End
      Begin VB.TextBox txtTotalYield_1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   4
         Left            =   1035
         Locked          =   -1  'True
         TabIndex        =   43
         Text            =   "0"
         Top             =   1710
         Width           =   735
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "R"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   20
         Left            =   1935
         TabIndex        =   55
         Top             =   675
         Width           =   645
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "L"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   21
         Left            =   1035
         TabIndex        =   54
         Top             =   675
         Width           =   645
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
         Left            =   45
         TabIndex        =   53
         Top             =   135
         Width           =   510
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "TRG YIELD"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   22
         Left            =   90
         TabIndex        =   52
         Top             =   990
         Width           =   960
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "% TARGET"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   23
         Left            =   90
         TabIndex        =   51
         Top             =   1395
         Width           =   870
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "PRD YIELD"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   24
         Left            =   90
         TabIndex        =   50
         Top             =   1800
         Width           =   870
      End
   End
   Begin VB.Frame frameDesign 
      BackColor       =   &H00FFFFFF&
      Height          =   2175
      Index           =   3
      Left            =   90
      TabIndex        =   28
      Top             =   5355
      Width           =   2760
      Begin VB.TextBox txtProdYiled_2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   3
         Left            =   1890
         Locked          =   -1  'True
         TabIndex        =   35
         Text            =   "0"
         Top             =   900
         Width           =   735
      End
      Begin VB.TextBox txtPerctarget_2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   3
         Left            =   1890
         Locked          =   -1  'True
         TabIndex        =   34
         Text            =   "0"
         Top             =   1305
         Width           =   735
      End
      Begin VB.TextBox txtTotalYield_2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   3
         Left            =   1890
         Locked          =   -1  'True
         TabIndex        =   33
         Text            =   "0"
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
         Index           =   3
         Left            =   585
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   32
         Top             =   135
         Width           =   2085
      End
      Begin VB.TextBox txtProdYiled_1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   3
         Left            =   1035
         Locked          =   -1  'True
         TabIndex        =   31
         Text            =   "0"
         Top             =   900
         Width           =   735
      End
      Begin VB.TextBox txtPerctarget_1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   3
         Left            =   1035
         Locked          =   -1  'True
         TabIndex        =   30
         Text            =   "0"
         Top             =   1305
         Width           =   735
      End
      Begin VB.TextBox txtTotalYield_1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   3
         Left            =   1035
         Locked          =   -1  'True
         TabIndex        =   29
         Text            =   "0"
         Top             =   1710
         Width           =   735
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "R"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   15
         Left            =   1935
         TabIndex        =   41
         Top             =   675
         Width           =   645
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "L"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   16
         Left            =   1035
         TabIndex        =   40
         Top             =   675
         Width           =   645
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
         Left            =   45
         TabIndex        =   39
         Top             =   135
         Width           =   510
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "TRG YIELD"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   17
         Left            =   90
         TabIndex        =   38
         Top             =   990
         Width           =   960
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "% TARGET"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   18
         Left            =   90
         TabIndex        =   37
         Top             =   1395
         Width           =   870
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "PRD YIELD"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   19
         Left            =   90
         TabIndex        =   36
         Top             =   1800
         Width           =   870
      End
   End
   Begin VB.Frame frameDesign 
      BackColor       =   &H00FFFFFF&
      Height          =   2175
      Index           =   2
      Left            =   90
      TabIndex        =   14
      Top             =   3015
      Width           =   2760
      Begin VB.TextBox txtProdYiled_2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   2
         Left            =   1890
         Locked          =   -1  'True
         TabIndex        =   21
         Text            =   "0"
         Top             =   900
         Width           =   735
      End
      Begin VB.TextBox txtPerctarget_2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   2
         Left            =   1890
         Locked          =   -1  'True
         TabIndex        =   20
         Text            =   "0"
         Top             =   1305
         Width           =   735
      End
      Begin VB.TextBox txtTotalYield_2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   2
         Left            =   1890
         Locked          =   -1  'True
         TabIndex        =   19
         Text            =   "0"
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
         Index           =   2
         Left            =   585
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   18
         Top             =   135
         Width           =   2085
      End
      Begin VB.TextBox txtProdYiled_1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   2
         Left            =   1035
         Locked          =   -1  'True
         TabIndex        =   17
         Text            =   "0"
         Top             =   900
         Width           =   735
      End
      Begin VB.TextBox txtPerctarget_1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   2
         Left            =   1035
         Locked          =   -1  'True
         TabIndex        =   16
         Text            =   "0"
         Top             =   1305
         Width           =   735
      End
      Begin VB.TextBox txtTotalYield_1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   2
         Left            =   1035
         Locked          =   -1  'True
         TabIndex        =   15
         Text            =   "0"
         Top             =   1710
         Width           =   735
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "R"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   10
         Left            =   1935
         TabIndex        =   27
         Top             =   675
         Width           =   645
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "L"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   11
         Left            =   1035
         TabIndex        =   26
         Top             =   675
         Width           =   645
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
         Left            =   45
         TabIndex        =   25
         Top             =   135
         Width           =   510
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "TRG YIELD"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   12
         Left            =   90
         TabIndex        =   24
         Top             =   990
         Width           =   960
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "% TARGET"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   13
         Left            =   90
         TabIndex        =   23
         Top             =   1395
         Width           =   870
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "PRD YIELD"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   14
         Left            =   90
         TabIndex        =   22
         Top             =   1800
         Width           =   870
      End
   End
   Begin VB.Frame frameDesign 
      BackColor       =   &H00FFFFFF&
      Height          =   2175
      Index           =   0
      Left            =   90
      TabIndex        =   0
      Top             =   675
      Width           =   2760
      Begin VB.TextBox txtTotalYield_1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   1
         Left            =   1035
         Locked          =   -1  'True
         TabIndex        =   7
         Text            =   "0"
         Top             =   1710
         Width           =   735
      End
      Begin VB.TextBox txtPerctarget_1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   1
         Left            =   1035
         Locked          =   -1  'True
         TabIndex        =   6
         Text            =   "0"
         Top             =   1305
         Width           =   735
      End
      Begin VB.TextBox txtProdYiled_1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   1
         Left            =   1035
         Locked          =   -1  'True
         TabIndex        =   5
         Text            =   "0"
         Top             =   900
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
         Left            =   585
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   4
         Top             =   135
         Width           =   2085
      End
      Begin VB.TextBox txtTotalYield_2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   1
         Left            =   1890
         Locked          =   -1  'True
         TabIndex        =   3
         Text            =   "0"
         Top             =   1710
         Width           =   735
      End
      Begin VB.TextBox txtPerctarget_2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   1
         Left            =   1890
         Locked          =   -1  'True
         TabIndex        =   2
         Text            =   "0"
         Top             =   1305
         Width           =   735
      End
      Begin VB.TextBox txtProdYiled_2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Index           =   1
         Left            =   1890
         Locked          =   -1  'True
         TabIndex        =   1
         Text            =   "0"
         Top             =   900
         Width           =   735
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "PRD YIELD"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   90
         TabIndex        =   13
         Top             =   1800
         Width           =   870
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "% TARGET"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   90
         TabIndex        =   12
         Top             =   1395
         Width           =   870
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "TRG YIELD"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   11
         Top             =   990
         Width           =   960
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
         Left            =   45
         TabIndex        =   10
         Top             =   135
         Width           =   510
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "L"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   1035
         TabIndex        =   9
         Top             =   675
         Width           =   645
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "R"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   4
         Left            =   1935
         TabIndex        =   8
         Top             =   675
         Width           =   645
      End
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   285
      Left            =   17730
      TabIndex        =   596
      Top             =   90
      Width           =   3345
      _ExtentX        =   5900
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000080&
      Caption         =   "PRODUCTION  YIELD  MONITORING"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   555
      Left            =   0
      TabIndex        =   560
      Top             =   0
      Width           =   26250
   End
End
Attribute VB_Name = "frmTotalYield"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub SetMachine()
On Error Resume Next
    Dim Rs As New ADODB.Recordset
    Dim sSQL As String
    Dim i As Integer

    Rs.CursorLocation = adUseClient

sSQL = "SELECT X.plant_mark,X.prod_machine_id,A.number AS machine_no,X.machine_status, X.eng_product_1,Y.product_name,"
sSQL = sSQL & " Y.target_yield,Y.period_shift,Y.jumlah_hour,Y.target_shot,Y.jumlah_shot,Y.gross_produksi,Y.net_produksi,"
sSQL = sSQL & " Y.prod_yield_1 , Y.persen_target_1, Z.prod_yield_2, Z.persen_target_2"
sSQL = sSQL & " FROM prod_running_products X"
sSQL = sSQL & " LEFT JOIN prod_machines A ON X.prod_machine_id = A.id"
sSQL = sSQL & " Left Join"
        sSQL = sSQL & " (SELECT a.plant_mark,a.prod_machine_id,a.eng_product_id,b.internal_part_id,b.name AS product_name,"
        sSQL = sSQL & " b.cavity,b.prod_yield AS target_yield,b.cycle_time_ia,a.period_shift,"
        sSQL = sSQL & " count(a.period_hour) AS jumlah_hour, (floor(3600/b.cycle_time_ia) * count(a.period_hour))  AS target_shot,"
        sSQL = sSQL & " sum(a.counter_ok) jumlah_shot,sum(a.counter_ok) * b.cavity AS gross_produksi,"
        sSQL = sSQL & " data_ngs.jumlah_ng,((sum(a.counter_ok) * b.cavity) - data_ngs.jumlah_ng) AS net_produksi,"
        sSQL = sSQL & " ROUND((((sum(a.counter_ok) * b.cavity) - data_ngs.jumlah_ng) / (sum(a.counter_ok) * b.cavity)) * 100,2) AS prod_yield_1,"
        sSQL = sSQL & " ROUND(Sum(a.counter_ok) / (floor((3600 / b.cycle_time_ia) * ROUND(Sum(a.counter_ok) / floor(3600 / b.cycle_time_ia), 1))) * 100, 2) As persen_target_1"
        sSQL = sSQL & " FROM prod_runnings a"
        sSQL = sSQL & " LEFT JOIN eng_products b ON a.eng_product_id = b.id"
        sSQL = sSQL & " LEFT JOIN (SELECT d.plant_mark,d.prod_machine_id,d.eng_product_id,d.period_shift,"
                    sSQL = sSQL & " sum(d.counter_ng) jumlah_ng FROM prod_data_ngs d"
                    sSQL = sSQL & " GROUP BY d.plant_mark,d.prod_machine_id,d.eng_product_id,d.period_shift) AS data_ngs"
                    sSQL = sSQL & " ON a.plant_mark = data_ngs.plant_mark AND a.prod_machine_id = data_ngs.prod_machine_id"
                    sSQL = sSQL & " AND a.eng_product_id = data_ngs.eng_product_id AND a.period_shift = data_ngs.period_shift"
        sSQL = sSQL & " WHERE a.period_shift = '" & Format(Now, "yyyy-mm-dd") & "'"
        sSQL = sSQL & " GROUP BY a.plant_mark,a.prod_machine_id,a.mkt_customer_id,a.eng_product_id,a.period_shift) Y"
sSQL = sSQL & " ON X.plant_mark = Y.plant_mark AND X.prod_machine_id = Y.prod_machine_id AND X.eng_product_1 = Y.eng_product_id"
sSQL = sSQL & " Left Join"
        sSQL = sSQL & " (SELECT a.plant_mark,a.prod_machine_id,a.eng_product_id,b.internal_part_id,b.name AS product_name,"
        sSQL = sSQL & " b.cavity,b.prod_yield AS target_yield,b.cycle_time_ia,a.period_shift,"
        sSQL = sSQL & " count(a.period_hour) AS jumlah_hour, (floor(3600/b.cycle_time_ia) * count(a.period_hour))  AS target_shot,"
        sSQL = sSQL & " sum(a.counter_ok) jumlah_shot,sum(a.counter_ok) * b.cavity AS gross_produksi,"
        sSQL = sSQL & " data_ngs.jumlah_ng,((sum(a.counter_ok) * b.cavity) - data_ngs.jumlah_ng) AS net_produksi,"
        sSQL = sSQL & " ROUND((((sum(a.counter_ok) * b.cavity) - data_ngs.jumlah_ng) / (sum(a.counter_ok) * b.cavity)) * 100,2) AS prod_yield_2,"
        sSQL = sSQL & " ROUND(Sum(a.counter_ok) / (floor((3600 / b.cycle_time_ia) * ROUND(Sum(a.counter_ok) / floor(3600 / b.cycle_time_ia), 1))) * 100, 2) As persen_target_2"
        sSQL = sSQL & " FROM prod_runnings a"
        sSQL = sSQL & " LEFT JOIN eng_products b ON a.eng_product_id = b.id"
        sSQL = sSQL & " LEFT JOIN (SELECT d.plant_mark,d.prod_machine_id,d.eng_product_id,d.period_shift,"
                    sSQL = sSQL & " sum(d.counter_ng) jumlah_ng FROM prod_data_ngs d"
                    sSQL = sSQL & " GROUP BY d.plant_mark,d.prod_machine_id,d.eng_product_id,d.period_shift) AS data_ngs"
                    sSQL = sSQL & " ON a.plant_mark = data_ngs.plant_mark AND a.prod_machine_id = data_ngs.prod_machine_id"
                    sSQL = sSQL & " AND a.eng_product_id = data_ngs.eng_product_id AND a.period_shift = data_ngs.period_shift"
        sSQL = sSQL & " WHERE a.period_shift = '" & Format(Now, "yyyy-mm-dd") & "'"
        sSQL = sSQL & " GROUP BY a.plant_mark,a.prod_machine_id,a.mkt_customer_id,a.eng_product_id,a.period_shift) Z"
sSQL = sSQL & " ON X.plant_mark = Z.plant_mark AND X.prod_machine_id = Z.prod_machine_id AND X.eng_product_2 = Z.eng_product_id"
sSQL = sSQL & " WHERE X.status = 'active' AND X.plant_mark = '" & p_plant_mark & "'"
sSQL = sSQL & " ORDER BY A.number ASC"



    
    Rs.Open sSQL, CN, adOpenStatic, adLockOptimistic

    If Rs.RecordCount > 0 Then
        Rs.MoveFirst
        Do While Not Rs.EOF
            i = i + 1
            txtInfo(i).Text = IIf(IsNull(Rs.Fields("product_name")), "", Rs.Fields("product_name"))
            txtProdYiled_1(i).Text = IIf(IsNull(Rs.Fields("target_yield")), "", Rs.Fields("target_yield"))
            txtProdYiled_2(i).Text = IIf(IsNull(Rs.Fields("target_yield")), "", Rs.Fields("target_yield"))
            
            txtPerctarget_1(i).Text = IIf(IsNull(Rs.Fields("persen_target_1")), "", Rs.Fields("persen_target_1"))
            txtPerctarget_2(i).Text = IIf(IsNull(Rs.Fields("persen_target_2")), "", Rs.Fields("persen_target_2"))
            
            txtTotalYield_1(i).Text = IIf(IsNull(Rs.Fields("prod_yield_1")), "", Rs.Fields("prod_yield_1"))
            txtTotalYield_2(i).Text = IIf(IsNull(Rs.Fields("prod_yield_2")), "", Rs.Fields("prod_yield_2"))
            
            If Rs.Fields("machine_status") = "loaded" Then
                lblMesin(i).BackColor = &H8000&
            ElseIf Rs.Fields("machine_status") = "no_load" Then
                lblMesin(i).BackColor = &H4080&
            ElseIf Rs.Fields("machine_status") = "maintenance" Then
                lblMesin(i).BackColor = &H8080&
            ElseIf Rs.Fields("machine_status") = "broken" Then
                lblMesin(i).BackColor = &H808000
            ElseIf Rs.Fields("machine_status") = "trial" Then
                lblMesin(i).BackColor = &HC000C0
            End If

            'ProdYiledColor txtProdYiled_1(i)
            'ProdYiledColor txtProdYiled_2(i)
            
            PersentColor txtPerctarget_1(i)
            PersentColor txtPerctarget_2(i)
            
'            TotalYieldColor txtTotalYield_1(i)
'            TotalYieldColor txtTotalYield_2(i)

            ProdYiledColor txtTotalYield_1(i)
            ProdYiledColor txtTotalYield_2(i)
            
            
            Rs.MoveNext
        Loop
    End If
    Set Rs = Nothing
 
End Sub

Private Sub ProdYiledColor(tBox As TextBox)
    If Val(tBox.Text) <> "0" Then
        If tBox.Text <> "" Then
            If Val(tBox.Text) >= 95 Then
                tBox.BackColor = &H80FF80
            ElseIf Val(tBox.Text) > 85 And Val(tBox.Text) < 95 Then
                tBox.BackColor = &HFFFF&
            ElseIf Val(tBox.Text) >= 81 And Val(Val(tBox.Text)) < 85 Then
                tBox.BackColor = &HC000C0
            ElseIf Val(tBox.Text) < 80 Then
                tBox.BackColor = &HFF&
            End If
        Else
            tBox.BackColor = &HFFFFFF
        End If
    Else
        tBox.BackColor = &HFFFFFF
    End If
End Sub

Private Sub PersentColor(tBox As TextBox)
    If tBox.Text <> "" Then
        If Val(tBox.Text) <> "0" Then
            If Val(tBox.Text) >= 98 Then
                tBox.BackColor = &H80FF80
            ElseIf Val(tBox.Text) >= 95 And Val(tBox.Text) < 98 Then
                tBox.BackColor = &HFFFF&
            ElseIf Val(tBox.Text) >= 90 And Val(tBox.Text) < 95 Then
                tBox.BackColor = &HC000C0
            ElseIf Val(tBox.Text) < 90 Then
                tBox.BackColor = &HFF&
            End If
        Else
            tBox.BackColor = &HFFFFFF
        End If
    Else
        tBox.BackColor = &HFFFFFF
    End If
End Sub



Private Sub DTDate_Change()
    Call SetMachine
    ProgressBar1.Value = 0
End Sub


Private Sub DtDate_Click()
    Call SetMachine
    ProgressBar1.Value = 0
End Sub

'Private Sub TotalYieldColor(tBox As TextBox)
'    If val(tBox.Text) <> "" Then
'        If val(tBox.Text) <> "0" Then
'            If val(tBox.Text) >= 85 Then
'                tBox.BackColor = &H80FF80
'            ElseIf val(tBox.Text) >= 82 And val(tBox.Text) < 85 Then
'                tBox.BackColor = &HFFFF&
'            ElseIf val(tBox.Text) >= 80 And val(tBox.Text) < 81 Then
'                tBox.BackColor = &HFF&
'            ElseIf val(tBox.Text) < 80 Then
'                tBox.BackColor = &H000040C0&
'                tBox.ForeColor = &HFFFFFF
'            End If
'        Else
'            tBox.BackColor = &HFFFFFF
'        End If
'    Else
'        tBox.BackColor = &HFFFFFF
'    End If
'End Sub

Private Sub Form_Load()
    Call SetMachine
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
qSQL = qSQL & " and a.period_shift = '" & Format(Now, "yyyy-mm-dd") & "'"
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

'    Xtimer = Xtimer + 1
'    If Xtimer > 2 Then
'       Xtimer = 0
'        Call SetMachine
'        Call CheckIdle
'    End If
'
    Call SetMachine

ProgressBar1.Value = 0
End Sub

Private Sub Timer2_Timer()
On Error Resume Next
    ProgressBar1.Max = 60
    ProgressBar1.Value = ProgressBar1.Value + 1
End Sub
